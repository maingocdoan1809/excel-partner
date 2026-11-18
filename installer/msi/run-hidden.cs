// Starts one program with no console window at all, waits for it and returns its exit code.
//
// The MSI runs its PowerShell steps through this executable. A custom action that runs powershell.exe directly opens a
// console window: on Windows 11 the default terminal is Windows Terminal, so even `-WindowStyle Hidden` leaves a
// WindowsTerminal window on the desktop. This program is a GUI-subsystem executable (it has no console of its own) and
// creates the child with CreateNoWindow, which is the CREATE_NO_WINDOW flag: the only reliable way to keep the console
// off the desktop. It also keeps the custom action free of pipes, so the bridge and the gateway that the setup starts
// cannot hold the action open (WiX's QuietExec reads the child's output from a pipe and returns only when every process
// that inherited it has closed it).
using System;
using System.Diagnostics;
using System.Text;

internal static class RunHidden
{
    private static int Main(string[] args)
    {
        if (args.Length < 1) return 2;

        try
        {
            var start = new ProcessStartInfo(args[0])
            {
                Arguments = Arguments(args, 1),
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            using (var child = Process.Start(start))
            {
                if (child == null) return 3;
                child.WaitForExit();
                return child.ExitCode;
            }
        }
        catch (Exception error)
        {
            // Nobody is looking at a console, so leave the reason where the setup log lives.
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "office-partner-run-hidden.log"),
                    DateTime.Now.ToString("s") + "  " + error + Environment.NewLine);
            }
            catch { }
            return 3;
        }
    }

    private static string Arguments(string[] args, int first)
    {
        var line = new StringBuilder();
        for (var i = first; i < args.Length; i++)
        {
            if (i > first) line.Append(' ');
            line.Append(Quote(args[i]));
        }
        return line.ToString();
    }

    // The quoting rules of CommandLineToArgvW: surround with quotes when needed, double the backslashes that precede a
    // quote, and double the trailing backslashes of a quoted argument.
    private static string Quote(string value)
    {
        if (value.Length > 0 && value.IndexOfAny(new[] { ' ', '\t', '"' }) < 0) return value;

        var quoted = new StringBuilder("\"");
        var backslashes = 0;
        foreach (var character in value)
        {
            if (character == '\\') { backslashes++; continue; }
            if (character == '"')
            {
                quoted.Append('\\', backslashes * 2 + 1).Append('"');
                backslashes = 0;
                continue;
            }
            quoted.Append('\\', backslashes).Append(character);
            backslashes = 0;
        }
        quoted.Append('\\', backslashes * 2).Append('"');
        return quoted.ToString();
    }
}
