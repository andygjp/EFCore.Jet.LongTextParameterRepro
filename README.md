# EFCore.Jet long text parameter repro

Shows that [EntityFrameworkCore.Jet](https://github.com/CirrusRedOrg/EntityFrameworkCore.Jet), over ODBC, can't save
a string longer than 255 characters, even into a Long Text (memo) column. The Access ODBC driver refuses the parameter
while binding it:

```
ERROR [HY104] [Microsoft][ODBC Microsoft Access Driver]Invalid precision value
```

The provider sends every string as `OdbcType.NVarChar`, which the driver takes only up to 255 characters. A longer
value has to be sent as long text (`OdbcType.NText`). The program also runs each case with a `DbCommandInterceptor`
that does that, which is the workaround.

## What it does

1. Creates a new, empty Access 2000 format `.mdb` in the temp folder, with the ODBC driver's own `CREATE_DBV4`
   (`SQLConfigDataSourceW`). It doesn't use `EnsureCreated`, which needs ADOX or DAO, which Office installed with
   Click-to-Run may not register for other programs.
2. Creates a table `Notes (Id COUNTER, Body MEMO)`, and maps `Body` with `HasColumnType("longchar")`.
3. Without the workaround and then with it: inserts notes of 255, 256 and 300 characters, and updates a note from
   5 characters to 300. After each save, it reads the note back and checks the whole value was saved.

## Requirements

- Windows, with the Microsoft Access Database Engine ODBC driver, "Microsoft Access Driver (*.mdb, *.accdb)". Office
  installs it, as does the Access Database Engine redistributable.
- The .NET 11 SDK.

The process must have the same architecture as the driver, which has the architecture of the Office that installed
it. The project builds for `win-x64`, self-contained, for 64-bit Office. On Windows on ARM, it runs under x64
emulation.

## Run

```shell
dotnet run
```

For 32-bit Office:

```shell
dotnet run -r win-x86
```

To use an existing empty database instead of creating one, pass its path:

```shell
dotnet run -- C:\path\to\Empty.mdb
```

If you get `ERROR [IM003] Specified driver could not be loaded due to system error 193`, the process and the driver
have different architectures: use the other `-r`.

## Expected output

Without the workaround, the 256 and 300 character saves fail. With it, every save keeps the whole value:

```
Operation  Length  Workaround  Outcome
Insert        255  no          saved whole
Insert        256  no          FAILED: ERROR [HY104] [Microsoft][ODBC Microsoft Access Driver]Invalid precision value
Insert        300  no          FAILED: ERROR [HY104] [Microsoft][ODBC Microsoft Access Driver]Invalid precision value
Update        300  no          FAILED: ERROR [HY104] [Microsoft][ODBC Microsoft Access Driver]Invalid precision value
Insert        255  yes         saved whole
Insert        256  yes         saved whole
Insert        300  yes         saved whole
Update        300  yes         saved whole
```

## Versions

- EntityFrameworkCore.Jet.Odbc 11.0.0-alpha.5, with the EF Core and System.Data.Odbc versions it depends on
- .NET 11
