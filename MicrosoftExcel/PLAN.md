# Plan — Excel-bibliotheek `MicrosoftExcel`

Besluit Raoul 2026-10-10: eigen map naast `MicrosoftAccess`, omdat `Option Compare Database` in Excel niet geldig is.
Bron: de 8 algemene `om*`-modules uit client (`Ideeds\client\Src\ExportedVBA`) en SBDInc (`API-EPM`).
Dit is een plan; er staat nog geen module in deze map.

## Indeling

```
MicrosoftExcel/
  src/modules/*.bas      één bestand per module, platte tekst
  library-stamps.txt     manifest (naam;datum;sha)
  README.md              uitleg + stempel
  tests/                 testwerkmap + testscript
```

- Geen `.xlsm` in git: de bibliotheek zijn de `.bas`-bestanden. Een klantwerkmap importeert ze (en `omExcelVersionControl` doet dat zelf).
- Alleen `Option Explicit`. Geen `Option Compare Database`, en `Option Compare Text` uit SBD vervalt (vergelijkingen dan expliciet met `vbTextCompare`).
- Zelfde stempelregel als Access: `'@lib <datum> sha=<8 hex>` na de Option-regels, zelfde normalisatie. `scripts/libstamp.py` krijgt één argument voor de map (nu staat de Access-map er vast in, regel 19).

## Eerst zelfstandig maken

| Module | Wat er weg of anders moet |
|---|---|
| `omLockingFunctions` | `IsNullOrEmpty` en de globale `gFso` vervangen door `Len()` en een lokale late-bound FileSystemObject. Dan geen afhankelijkheid van client-`omStringFunctions`/`omFileFunctions` en geen naamconflict met onze Access-modules. Testblokken (regels 272-489) weg. |
| `omWorksheetFunctions` | A en B samenvoegen. `omArrayFunctions.IsArrayInitialized` en `InitializeStringArrayFromString` inline (met `Split`). Testsub A weg. |
| `omListObjectFunctions` | Dode code (regels 16-39) weg, `xlPinYin` nakijken, `Integer` → `Long`. |

## Fouten die mee gefixt worden

- `omExcelApplicationFunctions.SheetExists` (regel 72): `StrComp` is omgekeerd. Een workbook-parameter erbij in plaats van vast `ThisWorkbook`.
- `omExcelVBAFunctions` (regel 45): `CVErr` in een `Double` geeft stil 0.
- `omExcelRecordsetFunctions`: verbinding blijft open zonder `closeConnection`; `Recordset` zonder `ADODB.`-prefix.
- `omExcelVersionControl`: `MsgBox` na export en import (blokkeert automatisch gebruik), import maakt dubbele modules, `DisplayAlerts` altijd `True`.
- `omLockingFunctions`: de `.lock` wordt niet atomair gemaakt (regels 30-39).
- `omOnedriveFunctions` (regel 76): bij geen treffer blijft een niet-bestaand pad over.

## Volgorde (elk: kopiëren, fixen, testen, stempelen, committen)

1. `omExcelApplicationFunctions`
2. `omOnedriveFunctions`
3. `omExcelVBAFunctions`
4. `omExcelRecordsetFunctions`
5. `omListObjectFunctions`
6. `omWorksheetFunctions`
7. `omLockingFunctions`
8. `omExcelVersionControl`

Dit is de volgorde van afhankelijkheden: 8 heeft 1 en 2 nodig, 5 heeft 2 en 4 nodig.

## Testen

Zoals bij Access: een verborgen Excel via COM, een tijdelijke werkmap, de `.bas` importeren en de functies laten lopen. Voor het importeren van VBA moet "Toegang tot het VBA-projectmodel vertrouwen" aan staan (alleen op deze pc, voor de test).

## Wat niet in deze map komt

- `omCSHA256` en `omCSHA256Functions`: vervallen. `omHashAlgorithms` (Access) doet het in UTF-8. Voor Excel is later een aparte kopie nodig, zonder `Option Compare Database`.
- `omDataFunctions` en `omDatabaseSync`: SBD-specifiek, blijven bij SBD als `sbd*`.
- De client- en SBD-werkmappen worden niet aangepast; hun eigen kopieën blijven.

## Open vragen voor Raoul

1. Mag ik op deze pc de Excel-instelling "VBA-projectmodel vertrouwen" aanzetten voor de tests? (Een beveiligingsinstelling van Office, alleen voor de gebruiker.)
2. `omOnedriveFunctions` is code van Ricardo Gerbaudo. Overnemen met zijn naam erbij (en de licentie van de bron nakijken), of herschrijven?
