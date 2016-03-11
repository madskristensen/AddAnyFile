using System;
using System.Diagnostics.CodeAnalysis;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

public static class Logger
{
    static IVsOutputWindowPane pane;
    static IServiceProvider _provider;
    static string _name;

    public static void Initialize(IServiceProvider provider, string name)
    {
        _provider = provider;
        _name = name;
    }

    [SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults", MessageId = "Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane.OutputString(System.String)")]
    public static void Log(string message)
    {
        if (string.IsNullOrEmpty(message))
            return;

        try
        {
            if (EnsurePane())
            {
                ThreadHelper.Generic.BeginInvoke(() =>
                {
                    pane.OutputStringThreadSafe(DateTime.Now + ": " + message + Environment.NewLine);
                });
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.Write(ex);
        }
    }

    public static void Log(Exception ex)
    {
        if (ex != null)
        {
            Log(ex.ToString());
            Telemetry.TrackException(ex);
        }
    }

    static bool EnsurePane()
    {
        if (pane == null)
        {
            ThreadHelper.Generic.Invoke(() =>
            {
                if (pane == null)
                {
                    Guid guid = Guid.NewGuid();
                    IVsOutputWindow output = (IVsOutputWindow)_provider.GetService(typeof(SVsOutputWindow));
                    output.CreatePane(ref guid, _name, 1, 1);
                    output.GetPane(ref guid, out pane);
                }
            });
        }

        return pane != null;
    }
}