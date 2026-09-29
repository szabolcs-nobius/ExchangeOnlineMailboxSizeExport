[CmdletBinding()]
param(
	[string]$OutputPath
)

$ErrorActionPreference = 'Stop'

foreach ($moduleName in @('ExchangeOnlineManagement', 'ImportExcel')) {
	if (-not (Get-Module -ListAvailable -Name $moduleName)) {
		throw "A(z) '$moduleName' modul nincs telepítve. Telepítés: Install-Module $moduleName -Scope CurrentUser"
	}

	Import-Module $moduleName
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
	Connect-ExchangeOnline -ShowBanner:$false
}


$organization = Get-OrganizationConfig
$tenantName = [string]$organization.DisplayName
if ([string]::IsNullOrWhiteSpace($tenantName)) {
	$tenantName = [string]$organization.Name
}
if ([string]::IsNullOrWhiteSpace($tenantName)) {
	throw 'Nem sikerült meghatározni a tenant nevét a riport fájlnevéhez.'
}

$tenantName = $tenantName -replace '[<>:"/\\|?*\x00-\x1F]', '_'
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
	$OutputPath = Join-Path -Path (Get-Location) -ChildPath ("{0}_ExchangeMailboxReport_{1:yyyyMMdd_HHmmss}.xlsx" -f $tenantName, (Get-Date))
}
else {
	$outputDirectory = Split-Path -Path $OutputPath -Parent
	$outputFileName = Split-Path -Path $OutputPath -Leaf
	$tenantPrefix = "{0}_" -f $tenantName
	if (-not $outputFileName.StartsWith($tenantPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
		$outputFileName = "{0}_{1}" -f $tenantName, $outputFileName
	}
	if ($outputDirectory) {
		$OutputPath = Join-Path -Path $outputDirectory -ChildPath $outputFileName
	}
	else {
		$OutputPath = $outputFileName
	}
}

if ([System.IO.Path]::GetExtension($OutputPath) -ne '.xlsx') {
	throw 'Az OutputPath kiterjesztése .xlsx legyen.'
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
		$row.SizeGB = $totalBytes / 1GB
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

$sortedReport = @(
	$report | Sort-Object -Property @{ Expression = { if ($null -eq $_.SizeGB) { -1.0 } else { $_.SizeGB } }; Descending = $true }
)

$excelPackage = $sortedReport | Export-Excel -Path $OutputPath `
	-WorksheetName 'Postaladak' `
	-TableName 'MailboxReport' `
	-AutoSize `
	-FreezeTopRow `
	-BoldTopRow `
	-PassThru

try {
	$worksheet = $excelPackage.Workbook.Worksheets['Postaladak']
	$worksheet.Column(6).Style.Numberformat.Format = '0'
	$highlightColor = [System.Drawing.Color]::FromArgb(255, 0, 0)

	for ($rowIndex = 0; $rowIndex -lt $sortedReport.Count; $rowIndex++) {
		if ($null -ne $sortedReport[$rowIndex].SizeGB -and $sortedReport[$rowIndex].SizeGB -gt 30) {
			$excelRow = $rowIndex + 2
			$rowCells = $worksheet.Cells[$excelRow, 1, $excelRow, 8]
			$rowCells.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
			$rowCells.Style.Fill.BackgroundColor.SetColor($highlightColor)
		}
	}
}
finally {
	Close-ExcelPackage $excelPackage
}

Write-Host "Riport elkészült: $((Resolve-Path -LiteralPath $OutputPath).Path)"
Write-Host "Postaládák: $($report.Count); sikeres: $(($report | Where-Object Status -eq 'OK').Count); hibás: $(($report | Where-Object Status -eq 'Hiba').Count)"
