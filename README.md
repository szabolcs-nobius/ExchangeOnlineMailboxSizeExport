

# Exchange Online postaládaméret-riport

A `ExchangOnlineMailboxSize.ps1` lekéri az Exchange Online tenant összes felhasználói (`UserMailbox`) és megosztott (`SharedMailbox`) postaládájának statisztikáját, majd Excel-munkafüzetbe exportálja. A riport postaládánként egy sort tartalmaz.

## Követelmények

- PowerShell 7.2 vagy újabb
- Exchange Online-ba bejelentkezni és postaláda-statisztikákat olvasni jogosult fiók
- Az `ExchangeOnlineManagement` és az `ImportExcel` PowerShell-modul

A modulok telepítése:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

## Futtatás

Az alapértelmezett riport létrehozásához futtasd a szkriptet:

```powershell
./ExchangOnlineMailboxSize.ps1
```

Ha nincs aktív Exchange Online-kapcsolat, a szkript interaktívan bejelentkeztet. Az Excel-fájl az aktuális mappába kerül, és a tenant nevével kezdődik, például `Contoso_ExchangeMailboxReport_20260929_120000.xlsx`.

Egyedi fájl vagy mappa megadásakor a tenant neve automatikusan a fájlnév elé kerül. A fájl kiterjesztése `.xlsx` legyen:

```powershell
./ExchangOnlineMailboxSize.ps1 -OutputPath ./MailboxReport.xlsx
```

Ebben a példában a létrejövő fájl neve `Contoso_MailboxReport.xlsx`.

## Riport tartalma

- Megjelenített név, elsődleges e-mail-cím, felhasználónév és postaládatípus
- Elemszám és postaládaméret GB-ban; a méret 1 GB = 1024³ bájt alapján számolódik, egész GB-ra kerekítve jelenik meg
- Csökkenő sorrend a pontos méret szerint
- A 30 GB-nál nagyobb postaládák teljes sora piros háttérrel kiemelve
- `Status` és `Error` oszlop az egyes postaládák lekérdezési hibáihoz; egy postaláda hibája nem akadályozza meg a többi postaláda exportját

A riport az elsődleges postaládákat tartalmazza; az archív postaládák mérete nem része az exportnak.