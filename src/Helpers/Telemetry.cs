using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using EnvDTE;
using EnvDTE80;
using Microsoft.ApplicationInsights;

/// <summary>
/// Reports anonymous usage through ApplicationInsights
/// </summary>
public static class Telemetry
{
    private static TelemetryClient _telemetry;
    private static DTEEvents _events;

    /// <summary>
    /// Initializes the telemetry client.
    /// </summary>
    public static void Initialize(DTE2 dte, string version, string telemetryKey)
    {
        if (_telemetry != null)
            throw new NotSupportedException("The telemetry client is already initialized");

        _telemetry = new TelemetryClient();
        _telemetry.Context.Session.Id = Guid.NewGuid().ToString();
        _telemetry.Context.Device.Model = dte.Edition;
        _telemetry.InstrumentationKey = telemetryKey;
        _telemetry.Context.Component.Version = version;

        byte[] enc = Encoding.UTF8.GetBytes(Environment.UserName + Environment.MachineName);
        using (var crypto = new MD5CryptoServiceProvider())
        {
            byte[] hash = crypto.ComputeHash(enc);
            _telemetry.Context.User.Id = Convert.ToBase64String(hash);
        }

        _events = dte.Events.DTEEvents;
        _events.OnBeginShutdown += _telemetry.Flush;

        Enabled = true;
    }

    public static bool Enabled { get; set; }

    /// <summary>Tracks an event to ApplicationInsights.</summary>
    public static void TrackEvent(string key, IDictionary<string, string> properties = null, IDictionary<string, double> metrics = null)
    {
#if !DEBUG
            if (Enabled)
            {
                _telemetry.TrackEvent(key, properties, metrics);
            }
#endif
    }

    /// <summary>Tracks any exception.</summary>
    public static void TrackException(Exception ex)
    {
#if !DEBUG
            if (Enabled)
            {
                var telex = new Microsoft.ApplicationInsights.DataContracts.ExceptionTelemetry(ex);
                telex.HandledAt = Microsoft.ApplicationInsights.DataContracts.ExceptionHandledAt.UserCode;
                _telemetry.TrackException(telex);
            }
#endif
    }
}