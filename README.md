

# Exchange Online postaládaméret-riport

A szkript az Exchange Online tenant összes felhasználói (`UserMailbox`) és megosztott (`SharedMailbox`) postaládájához lekéri az elsődleges postaláda méretét és elemszámát, majd Excel-munkafüzetbe (`.xlsx`) exportálja az eredményt.

## Előfeltételek

- PowerShell 7.2 vagy újabb
- Exchange Online lekérdezésére jogosító fiók
- Az `ExchangeOnlineManagement` és az `ImportExcel` PowerShell-modul

Telepítés:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

## Futtatás

```powershell
./ExchangOnlineMailboxSize.ps1
```

Az első futtatáskor a szkript interaktívan csatlakozik az Exchange Online-hoz. Alapértelmezés szerint az aktuális mappába ír egy, a tenant nevével kezdődő fájlt, például `Contoso_ExchangeMailboxReport_20260929_120000.xlsx`. Egyedi útvonal esetén a megadott mappát és fájlnevet használja, a tenant nevét pedig automatikusan a fájlnév elé teszi:

```powershell
./ExchangOnlineMailboxSize.ps1 -OutputPath ./MailboxReport.xlsx
```

A `SizeGB` oszlop értéke 1 GB = 1024³ bájttal számolódik, egész GB-ra kerekítve jelenik meg, és csökkenő méret szerint rendezi a sorokat. A 30 GB-nál nagyobb postaládák teljes sora kiemelt színt kap. A postaládánkénti lekérdezési hibák külön sorban, `Hiba` státusszal szerepelnek, így a többi postaláda exportja ettől még elkészül. Az archív postaládák mérete nem része ennek a riportnak.