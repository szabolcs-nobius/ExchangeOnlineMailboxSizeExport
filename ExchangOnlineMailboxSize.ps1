[CmdletBinding()]
param(
	[string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("ExchangeMailboxReport_{0:yyyyMMdd_HHmmss}.xlsx" -f (Get-Date)))
)

$ErrorActionPreference = 'Stop'

foreach ($moduleName in @('ExchangeOnlineManagement', 'ImportExcel')) {
	if (-not (Get-Module -ListAvailable -Name $moduleName)) {
		throw "A(z) '$moduleName' modul nincs telepítve. Telepítés: Install-Module $moduleName -Scope CurrentUser"
	}

	Import-Module $moduleName
}

if ([System.IO.Path]::GetExtension($OutputPath) -ne '.xlsx') {
	throw 'Az OutputPath kiterjesztése .xlsx legyen.'
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
	Connect-ExchangeOnline -ShowBanner:$false
}

function Convert-TotalItemSizeToBytes {
	param(
		[Parameter(Mandatory = $false)]
		$TotalItemSize
	)

	if ($null -eq $TotalItemSize) {
		return 0.0
	}

	$value = $TotalItemSize.Value
	if ($null -ne $value -and $value.PSObject.Methods['ToBytes']) {
		return [double]$value.ToBytes()
	}

	if ($TotalItemSize.PSObject.Methods['ToBytes']) {
		return [double]$TotalItemSize.ToBytes()
	}

	$sizeText = [string]$TotalItemSize
	if ($sizeText -match '\(([\d,\.\s]+)\s*bytes?\)') {
		$byteDigits = $Matches[1] -replace '\D', ''
		if ($byteDigits) {
			return [double]::Parse($byteDigits, [Globalization.CultureInfo]::InvariantCulture)
		}
	}

	if ($sizeText -match '^\s*([\d.,]+)\s*(B|KB|MB|GB|TB)\b') {
		$number = [double]::Parse($Matches[1].Replace(',', '.'), [Globalization.CultureInfo]::InvariantCulture)
		$unitPower = switch ($Matches[2].ToUpperInvariant()) {
			'B'  { 0 }
			'KB' { 1 }
			'MB' { 2 }
			'GB' { 3 }
			'TB' { 4 }
		}
		return $number * [math]::Pow(1024, $unitPower)
	}

	throw "Nem sikerült bájtra alakítani a postaládaméretet: '$sizeText'"
}

$mailboxes = @(
	Get-EXOMailbox -RecipientTypeDetails UserMailbox, SharedMailbox `
		-ResultSize Unlimited `
		-Properties DisplayName, PrimarySmtpAddress, RecipientTypeDetails, UserPrincipalName
)

if ($mailboxes.Count -eq 0) {
	throw 'A tenantben nem található felhasználói vagy megosztott postaláda.'
}

$report = [System.Collections.Generic.List[object]]::new()
$index = 0

foreach ($mailbox in $mailboxes) {
	$index++
	Write-Progress -Activity 'Postaládastatisztikák lekérdezése' `
		-Status "$index / $($mailboxes.Count): $($mailbox.PrimarySmtpAddress)" `
		-PercentComplete (($index / $mailboxes.Count) * 100)

	$row = [ordered]@{
		DisplayName        = $mailbox.DisplayName
		PrimarySmtpAddress = [string]$mailbox.PrimarySmtpAddress
		UserPrincipalName  = $mailbox.UserPrincipalName
		MailboxType        = [string]$mailbox.RecipientTypeDetails
		ItemCount          = $null
		SizeGB             = $null
		Status             = 'OK'
		Error              = $null
	}

	try {
		$statistics = Get-EXOMailboxStatistics -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop
		$totalBytes = Convert-TotalItemSizeToBytes -TotalItemSize $statistics.TotalItemSize

		$row.ItemCount = $statistics.ItemCount
		$row.SizeGB = [math]::Round(($totalBytes / 1GB), 3)
	}
	catch {
		$row.Status = 'Hiba'
		$row.Error = $_.Exception.Message
		Write-Warning "Nem sikerült lekérdezni: $($mailbox.PrimarySmtpAddress) - $($row.Error)"
	}

	$report.Add([pscustomobject]$row)
}

Write-Progress -Activity 'Postaládastatisztikák lekérdezése' -Completed

$outputDirectory = Split-Path -Parent $OutputPath
if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
	New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

$report | Export-Excel -Path $OutputPath `
	-WorksheetName 'Postaladak' `
	-TableName 'MailboxReport' `
	-AutoSize `
	-FreezeTopRow `
	-BoldTopRow

Write-Host "Riport elkészült: $((Resolve-Path -LiteralPath $OutputPath).Path)"
Write-Host "Postaládák: $($report.Count); sikeres: $(($report | Where-Object Status -eq 'OK').Count); hibás: $(($report | Where-Object Status -eq 'Hiba').Count)"
