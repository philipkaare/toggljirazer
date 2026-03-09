using System.IO;
using System.Text;

namespace ToggJirazer;

/// <summary>
/// A <see cref="TextWriter"/> that routes each complete line to an
/// <see cref="Action{T}"/> delegate, making it easy to redirect
/// <see cref="Console.Out"/> / <see cref="Console.Error"/> to a WPF log control.
///
/// Carriage-return characters (<c>\r</c>) are stripped so that the
/// console-style progress-bar overwrite pattern (<c>\r...</c>) does not
/// produce garbled output in the log window.
/// </summary>
public sealed class LogTextWriter : TextWriter
{
    private readonly Action<string> _writeLine;
    private readonly StringBuilder _buffer = new();

    public LogTextWriter(Action<string> writeLine) => _writeLine = writeLine;

    public override Encoding Encoding => Encoding.UTF8;

    // --- character-level overrides -------------------------------------------

    public override void Write(char value)
    {
        if (value == '\r') return;      // strip; handled as part of \r\n or overwrite
        if (value == '\n') { EmitLine(); return; }
        _buffer.Append(value);
    }

    public override void Write(string? value)
    {
        if (value is null) return;
        foreach (var c in value)
            Write(c);
    }

    // --- line-level overrides ------------------------------------------------

    public override void WriteLine(string? value)
    {
        if (value is not null) Write(value);
        EmitLine();
    }

    public override void WriteLine() => EmitLine();

    // --- flush / dispose -----------------------------------------------------

    protected override void Dispose(bool disposing)
    {
        if (_buffer.Length > 0) EmitLine();
        base.Dispose(disposing);
    }

    // -------------------------------------------------------------------------

    private void EmitLine()
    {
        _writeLine(_buffer.ToString());
        _buffer.Clear();
    }
}
