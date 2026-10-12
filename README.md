# LedgerFrame

A small WinForms spreadsheet tool built with [Syncfusion](https://www.syncfusion.com/) XlsIO / Spreadsheet controls.

Opens a workbook, lets the user edit input cells while formula cells stay
protected, resolves dynamic (`INDIRECT`-based) data-validation dropdowns,
and can refresh selected input ranges from an external data file.

## Contents

- `src/` — the application source
- `samples/` — a sample workbook (tree volume & stand tables) used as the development model

## Build

Requires .NET SDK 8.0+. Builds for both .NET Framework 4.6.2 and .NET 8:

```
cd src
dotnet build
```

Run the produced `LedgerFrame.exe` — with no embedded model resource it
falls back to the sample workbook copied next to the executable.

## License

MIT (see `LICENSE`). Syncfusion components require their own license
(Community License works for qualifying users).
