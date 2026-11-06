# ROADMAP — MicrosoftAccess bibliotheek

Open punten van de bibliotheek. Gedaan = regel weg (de geschiedenis staat in `git log`).
Besluit Raoul 2026-10-10: deze lijst woont hier, niet in CM of Customers.

## Open

| Punt | Wat | Opmerking |
|---|---|---|
| Push naar GitHub | De lokale commits (de historie is herschreven: namen, wachtwoordregels en oude databaseversies eruit) staan nog niet op GitHub; vermoedelijk force-push nodig | Alleen op woord van Raoul |
| Excel-bibliotheek aanvullen | 8 modules gebouwd en getest in een verborgen Excel; open: een proef in een klantwerkmap en `omHashAlgorithms` voor Excel (zonder `Option Compare Database`) | een klantwerkmap als eerste proef, na akkoord |
| Optimalisatie-audit | `omOutlook` Sender-voorwaarde (omgekeerd?), PK-substring in `TableCreateDDL`, `VBAFindReplace` en `omMSAccessFunctions` DMax-lus | Elk met test; `omOutlook` kan alleen met Outlook |
| Kandidaten pc-zoektocht | 19 Access- en 5 Excel/Outlook-kandidaten beoordelen | Klantprojecten waar we niet aan raken alleen lezen |
| Klantcode (klant A) | `omFile`/`omFolder` vervangen door `DeleteFile`/`RenameFile`/`ReadFilesIntoTable`/`SetFilesSelected`; `hbPrinterFunctions` op `omPrinterFunctions`; `hbReportColors` in de `.mdb` bouwen | Raakt de klantrepo: eerst met Raoul |
| Klantcode (klant B) | `GetWerkgebied` vervangen door `GetWorkArea` | |
| `omKernalFunctions` | Oudere variant in 8 klanten (Sleep/GetTickCount/RunExeAndWait/RunExeAndGetPID) | Geen actieve aanroepers; per klant opruimen |
| Git-hook voor stempels | `libstamp.py apply` automatisch | Nieuwe config: eerst voorleggen (wat/waarom/default/wie/wanneer weg) |
| Niet getest | `ConnectorTest` foutpad, `savePictEMF`, `omOutlook`, echte SFTP/Seq, volledige `LinkUsingSSMA` | Hebben een echte server of Outlook nodig |

## Besluiten

- Klantprojecten waar we niet aan raken: niets wijzigen, alleen overnemen wat bruikbaar is.
- Naamgeving: `om*` = bibliotheek (gestempeld), `tst*` = tests, `ex*` = voorbeelden.
- Een functie die een tabel nodig heeft, beheert die tabel zelf (`omPrinters`, `omFiles`).
