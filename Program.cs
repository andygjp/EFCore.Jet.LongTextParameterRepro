using System.Data.Common;
using System.Data.Odbc;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;

// Shows that EntityFrameworkCore.Jet, over ODBC, can't save a string longer than 255 characters, even into a Long Text
// (memo) column: the Access ODBC driver refuses the parameter with HY104 "Invalid precision value".
//
// Usage: EFCore.Jet.LongTextParameterRepro [path to an existing, empty .mdb or .accdb]
// Without a path, it creates a new Access 2000 format .mdb in the temp folder.

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("The Access ODBC driver only runs on Windows.");
    return 1;
}

return Repro.Run(args);

[SupportedOSPlatform("windows")]
internal static class Repro
{
    private const string Driver = "Microsoft Access Driver (*.mdb, *.accdb)";

    public static int Run(string[] args)
    {
        var path = args.Length > 0 ? Path.GetFullPath(args[0]) : NewDatabasePath();
        if (args.Length == 0)
        {
            AccessDatabase.Create(Driver, path);
        }

        var connectionString = $"Driver={{{Driver}}};DBQ={path};";
        AccessDatabase.CreateNotesTable(connectionString);

        Console.WriteLine($"Process architecture: {RuntimeInformation.ProcessArchitecture}");
        Console.WriteLine($"Database: {path}");
        Console.WriteLine();

        var results = new List<Result>();
        foreach (var withWorkaround in new[] { false, true })
        {
            foreach (var length in new[] { 255, 256, 300 })
            {
                results.Add(Insert(connectionString, length, withWorkaround));
            }

            results.Add(Update(connectionString, length: 300, withWorkaround));
        }

        Console.WriteLine($"{"Operation",-10} {"Length",6}  {"Workaround",-10}  Outcome");
        foreach (var result in results)
        {
            Console.WriteLine(
                $"{result.Operation,-10} {result.Length,6}  {(result.WithWorkaround ? "yes" : "no"),-10}  {result.Outcome}");
        }

        return 0;
    }

    private static string NewDatabasePath() =>
        Path.Combine(Path.GetTempPath(), $"LongTextParameterRepro-{DateTime.Now:yyyyMMdd-HHmmss}.mdb");

    private static Result Insert(string connectionString, int length, bool withWorkaround)
    {
        var body = Text(length);

        return Attempt("Insert", length, withWorkaround, () =>
        {
            int id;
            using (var context = NotesContext.Create(connectionString, withWorkaround))
            {
                var note = new Note { Body = body };
                context.Notes.Add(note);
                context.SaveChanges();
                id = note.Id;
            }

            return ReadBack(connectionString, id, body);
        });
    }

    private static Result Update(string connectionString, int length, bool withWorkaround)
    {
        var body = Text(length);

        return Attempt("Update", length, withWorkaround, () =>
        {
            int id;
            using (var context = NotesContext.Create(connectionString, withWorkaround: false))
            {
                var note = new Note { Body = "short" };
                context.Notes.Add(note);
                context.SaveChanges();
                id = note.Id;
            }

            using (var context = NotesContext.Create(connectionString, withWorkaround))
            {
                context.Notes.Single(x => x.Id == id).Body = body;
                context.SaveChanges();
            }

            return ReadBack(connectionString, id, body);
        });
    }

    private static Result Attempt(string operation, int length, bool withWorkaround, Func<string> attempt)
    {
        try
        {
            return new(operation, length, withWorkaround, attempt());
        }
        catch (DbUpdateException e) when (e.InnerException is OdbcException odbc)
        {
            return new(operation, length, withWorkaround, $"FAILED: {odbc.Message}");
        }
    }

    private static string ReadBack(string connectionString, int id, string expected)
    {
        using var context = NotesContext.Create(connectionString, withWorkaround: false);
        var actual = context.Notes.AsNoTracking().Single(x => x.Id == id).Body;

        return actual == expected ? "saved whole" : $"saved, but {actual?.Length ?? 0} characters came back";
    }

    /// <summary>
    /// Text whose characters differ along it, so a value that was cut shows
    /// </summary>
    private static string Text(int length) =>
        string.Concat(Enumerable.Range(0, length).Select(i => (char)('a' + i % 26)));

    private sealed record Result(string Operation, int Length, bool WithWorkaround, string Outcome);
}

public sealed class Note
{
    public int Id { get; set; }
    public string? Body { get; set; }
}

[SupportedOSPlatform("windows")]
public sealed class NotesContext(DbContextOptions<NotesContext> options) : DbContext(options)
{
    public DbSet<Note> Notes => Set<Note>();

    public static NotesContext Create(string connectionString, bool withWorkaround)
    {
        var builder = new DbContextOptionsBuilder<NotesContext>().UseJetOdbc(connectionString);
        if (withWorkaround)
        {
            builder.AddInterceptors(new LongTextParameterInterceptor());
        }

        return new(builder.Options);
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        var note = modelBuilder.Entity<Note>();
        note.ToTable("Notes");
        // The provider knows the column is Long Text, and still sends the value as NVarChar
        note.Property(x => x.Body).HasColumnType("longchar");
    }
}

/// <summary>
/// The workaround: sends each string parameter longer than 255 characters as long text
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class LongTextParameterInterceptor : DbCommandInterceptor
{
    public override InterceptionResult<DbDataReader> ReaderExecuting(DbCommand command, CommandEventData eventData,
        InterceptionResult<DbDataReader> result)
    {
        UseLongTextForLongStrings(command);
        return result;
    }

    public override InterceptionResult<int> NonQueryExecuting(DbCommand command, CommandEventData eventData,
        InterceptionResult<int> result)
    {
        UseLongTextForLongStrings(command);
        return result;
    }

    public override InterceptionResult<object> ScalarExecuting(DbCommand command, CommandEventData eventData,
        InterceptionResult<object> result)
    {
        UseLongTextForLongStrings(command);
        return result;
    }

    private static void UseLongTextForLongStrings(DbCommand command)
    {
        foreach (var parameter in command.Parameters.OfType<OdbcParameter>()
                     .Where(x => x.Value is string { Length: > 255 }))
        {
            parameter.OdbcType = OdbcType.NText;
        }
    }
}

/// <summary>
/// Creates the database with the ODBC driver itself, rather than EnsureCreated, which needs ADOX or DAO, which an
/// Office installed with Click-to-Run may not register for other programs
/// </summary>
[SupportedOSPlatform("windows")]
internal static partial class AccessDatabase
{
    private const ushort OdbcAddDsn = 1;

    public static void Create(string driver, string path)
    {
        // CREATE_DBV4 creates the Access 2000 (Jet 4) format
        if (!SQLConfigDataSourceW(IntPtr.Zero, OdbcAddDsn, driver, $"CREATE_DBV4=\"{path}\" General\0"))
        {
            throw new InvalidOperationException(
                $"Couldn't create {path} with the \"{driver}\" ODBC driver: {InstallerError()}");
        }
    }

    public static void CreateNotesTable(string connectionString)
    {
        using var connection = new OdbcConnection(connectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "CREATE TABLE Notes (Id COUNTER CONSTRAINT PK_Notes PRIMARY KEY, Body MEMO)";
        command.ExecuteNonQuery();
    }

    private static unsafe string InstallerError()
    {
        const ushort capacity = 512;
        var message = stackalloc char[capacity];
        return SQLInstallerErrorW(1, out var code, message, capacity, out var length) is 0 or 1
            ? $"{new string(message, 0, Math.Min((int)length, capacity))} (error {code})"
            : "no error was given";
    }

    [LibraryImport("odbccp32.dll", StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SQLConfigDataSourceW(IntPtr hwndParent, ushort fRequest, string lpszDriver,
        string lpszAttributes);

    [LibraryImport("odbccp32.dll")]
    private static unsafe partial short SQLInstallerErrorW(ushort iError, out uint pfErrorCode, char* lpszErrorMsg,
        ushort cbErrorMsgMax, out ushort pcbErrorMsg);
}
