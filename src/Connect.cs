using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: AssemblyTitle("AutoFrom Classic Outlook Add-in")]
[assembly: AssemblyVersion("1.2.6.0")]
[assembly: ComVisible(false)]

namespace AutoFrom
{
    [ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int mode,
            [MarshalAs(UnmanagedType.IDispatch)] object addIn, [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int mode, [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }
    [ComVisible(true), Guid("948FBFCA-513F-4BC4-8DA1-B83085E45B38"), ProgId("AutoFrom.Connect"), ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IAutoFromRibbon))]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility, IAutoFromRibbon
    {
        static readonly Guid ApplicationEvents = new Guid("0006302C-0000-0000-C000-000000000046");
        static readonly Guid InspectorsEvents = new Guid("00063079-0000-0000-C000-000000000046");
        delegate void NewInspectorHandler([MarshalAs(UnmanagedType.IDispatch)] object inspector);
        delegate void ItemSendHandler([MarshalAs(UnmanagedType.IDispatch)] object item, [In, Out] ref bool cancel);
        object application;
        ItemSendHandler sendHandler;
        object inspectorsSource;
        NewInspectorHandler newInspectorHandler;
        string windowEventError;
        Timer timer;
        Config config;
        string configError;
        bool busy;
        bool settingsOpen;
        string scanError;
        DateTime lastScan;
        Dictionary<object, string> draftErrors = new Dictionary<object, string>();
        Dictionary<object, DraftState> states = new Dictionary<object, DraftState>();
        public string GetCustomUI(string ribbonId) { return RibbonMarkup.For(ribbonId); }
        public void CheckDraft(object control)
        {
            var report = new System.Text.StringBuilder("AutoFrom 1.2.6 draft check (read-only)\n");
            report.AppendLine("Last automatic scan: " + (lastScan == default(DateTime) ? "not yet run" : lastScan.ToString("T")));
            report.AppendLine("Compose-open event: " + (windowEventError ?? (newInspectorHandler == null ? "not connected" : "connected")));
            try
            {
                if (configError != null) throw new InvalidOperationException(configError);
                if (config == null) throw new InvalidOperationException("Configuration is not loaded.");
                report.AppendLine("Enabled: " + config.Enabled + "; configured routing: " + config.HasRouting);
                if (scanError != null) report.AppendLine("Automatic scan error: " + scanError);
                dynamic app = application;
                object item = null;
                // The ribbon context selects this window's draft, not another open inspector.
                dynamic window = control == null ? app.ActiveWindow() : ((dynamic)control).Context;
                if (window != null)
                {
                    if ((int)window.Class == 35) item = (object)window.CurrentItem;
                    else if ((int)window.Class == 34) item = (object)window.ActiveInlineResponse;
                }
                if (!IsDraft(item)) throw new InvalidOperationException("Open a new email and click AutoFrom > Check draft in that email window.");
                string failure;
                report.AppendLine(draftErrors.TryGetValue(item, out failure) ? "Automatic selection error: " + failure : "No automatic selection error recorded for this draft.");
                var port = new OutlookPort(item, (object)app.Session);
                var current = port.Capture();
                DraftState state;
                bool tracked = states.TryGetValue(item, out state) && state.Original != null;
                string original = tracked ? state.Original.Account : current.Account;
                report.AppendLine("Draft evaluated previously: " + tracked);
                report.AppendLine("Original account: " + original);
                report.AppendLine("Current sending account: " + current.Account);
                report.AppendLine("Current represented name: " + current.Represented);
                object smtp;
                report.AppendLine("Current represented SMTP: " + (current.RepresentingProperties.TryGetValue(AliasIdentity.SmtpAddress, out smtp) ? Convert.ToString(smtp) : "(not set)"));
                bool complete;
                var recipients = port.Recipients(false, out complete);
                var decision = RuleEngine.Decide(config, original, recipients, complete);
                report.AppendLine("Selection: " + (decision.Error ?? (decision.Target == null ? "no matching default/rule" : decision.Target.FromAddress)));
                report.AppendLine("This check does not change or send the draft. Local properties do not prove delivered From or visible From-field refresh.");
            }
            catch (Exception ex) { report.AppendLine("Check stopped: " + ErrorDetail(ex)); }
            using (var dialog = new Form { Text = "AutoFrom draft check", Width = 760, Height = 440, StartPosition = FormStartPosition.CenterParent })
            {
                dialog.Controls.Add(new TextBox { Multiline = true, ReadOnly = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Both,
                    Text = report.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n") });
                dialog.ShowDialog(new OutlookWindow(GetForegroundWindow()));
            }
        }
        static string ErrorDetail(Exception ex)
        {
            return ex.Message + " (" + ex.GetType().Name + ", 0x" + ex.HResult.ToString("X8") + ")" +
                (ex.InnerException == null ? "" : "\nCause: " + ErrorDetail(ex.InnerException));
        }
        public void OpenSettings(object control)
        {
            if (settingsOpen || application == null) return;
            settingsOpen = true;
            busy = true;
            try
            {
                dynamic app = application;
                dynamic profile = app.Session.Accounts;
                var addresses = new List<string>();
                for (int i = 1; i <= (int)profile.Count; i++)
                {
                    string address = (string)profile.Item(i).SmtpAddress;
                    if (!string.IsNullOrWhiteSpace(address)) addresses.Add(Config.Address(address));
                }
                using (var dialog = new SettingsForm(config, addresses.ToArray(), configError, SaveSettings))
                    dialog.ShowDialog(new OutlookWindow(GetForegroundWindow()));
            }
            catch (Exception ex) { Show("Could not open settings.\n" + ex.Message); }
            finally { settingsOpen = false; busy = false; Tick(null, EventArgs.Empty); }
        }
        void SaveSettings(Config edited)
        {
            var next = SettingsStore.Validate(edited);
            dynamic app = application;
            // Reject invalid alias ownership before restoring drafts or saving configuration.
            foreach (Sender sender in next.Senders.Where(s => s.IsAlias))
            {
                dynamic profile = app.Session.Accounts;
                var matches = new List<object>();
                for (int i = 1; i <= (int)profile.Count; i++)
                {
                    dynamic account = profile.Item(i);
                    if (string.Equals((string)account.SmtpAddress, sender.Account, StringComparison.OrdinalIgnoreCase)) matches.Add(account);
                }
                if (matches.Count != 1) throw new InvalidOperationException("Choose the alias's owning account from the Outlook account list.");
                AliasIdentity.ValidateOwnership(matches[0], sender.Account, sender.Alias);
            }
            var open = OpenDrafts();
            foreach (object closed in states.Keys.Where(k => !open.Contains(k)).ToArray()) states.Remove(closed);
            // Restore before committing a new policy, including when disabling or removing rules.
            // If restoration or disk IO fails, retain the old policy and show an error in the dialog.
            foreach (var pair in states)
                Controller.RestoreManaged(new OutlookPort(pair.Key, (object)app.Session), pair.Value);
            SettingsStore.Save(SettingsStore.DefaultPath, next);
            config = next;
            configError = null;
            if (!config.Enabled || !config.HasRouting) states.Clear();
        }
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        sealed class OutlookWindow : IWin32Window
        {
            public OutlookWindow(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; private set; }
        }
        public void OnConnection(object app, int mode, object addIn, ref Array custom)
        {
            application = app;
            try
            {
                sendHandler = OnSend;
                ComEventsHelper.Combine(application, ApplicationEvents, 61442, sendHandler);
                try { config = Config.Load(SettingsStore.DefaultPath); }
                catch (Exception ex) { configError = "AutoFrom rules could not be loaded: " + ex.Message; }
                // Retain the collection RCW and delegate for the lifetime of the sink.
                // A failed optional event hookup must not disable send-time checking.
                try
                {
                    inspectorsSource = (object)((dynamic)application).Inspectors;
                    newInspectorHandler = OnNewInspector;
                    ComEventsHelper.Combine(inspectorsSource, InspectorsEvents, 61441, newInspectorHandler);
                }
                catch (Exception ex) { windowEventError = ErrorDetail(ex); }
                timer = new Timer { Interval = 250 };
                timer.Tick += Tick;
                timer.Start();
                if (configError != null) Show(configError + "\nMail sending is blocked. Open AutoFrom > Settings to repair and save your settings.");
            }
            catch (Exception ex)
            {
                Stop();
                Show("AutoFrom failed to load. Automatic sender selection and send checks are NOT active.\n" + ex.Message);
                throw;
            }
        }
        void OnNewInspector(object inspector)
        {
            if (busy || settingsOpen || config == null || !config.Enabled || !config.HasRouting || timer == null) return;
            // Writing before the Inspector is initialized can leave its visible From
            // stale even when the MAPI identity is correct. Defer to the UI message loop.
            timer.Stop();
            timer.Interval = 50;
            timer.Start();
        }
        void Tick(object sender, EventArgs args)
        {
            if (timer != null) timer.Interval = 250;
            if (busy || settingsOpen || config == null || !config.Enabled || !config.HasRouting) return;
            busy = true;
            try
            {
                lastScan = DateTime.Now;
                var open = OpenDrafts();
                scanError = null;
                foreach (object item in open) Preview(item);
                // Retain the original sender while a draft is open, even across managed GC.
                // Prune only after a complete scan, never after a transient COM enumeration failure.
                foreach (object closed in states.Keys.Where(k => !open.Contains(k)).ToArray()) states.Remove(closed);
                foreach (object closed in draftErrors.Keys.Where(k => !open.Contains(k)).ToArray()) draftErrors.Remove(closed);
            }
            catch (Exception ex) { scanError = ErrorDetail(ex); }
            finally { busy = false; }
        }
        HashSet<object> OpenDrafts()
        {
            dynamic app = application;
            var open = new HashSet<object>();
            dynamic inspectors = app.Inspectors;
            for (int i = 1; i <= (int)inspectors.Count; i++)
            {
                object item = (object)inspectors.Item(i).CurrentItem;
                if (IsDraft(item)) open.Add(item);
            }
            dynamic explorers = app.Explorers;
            for (int i = 1; i <= (int)explorers.Count; i++)
            {
                object item = (object)explorers.Item(i).ActiveInlineResponse;
                if (IsDraft(item)) open.Add(item);
            }
            return open;
        }
        void Preview(object item)
        {
            try { if (IsDraft(item)) { Evaluate(item, false); draftErrors.Remove(item); } }
            catch (Exception ex) { draftErrors[item] = ErrorDetail(ex); }
        }
        static bool IsDraft(object item)
        {
            if (item == null) return false;
            dynamic m = item;
            return (int)m.Class == 43 && !(bool)m.Sent && !(bool)m.Submitted;
        }
        void Evaluate(object item, bool sending)
        {
            dynamic app = application;
            DraftState state;
            if (!states.TryGetValue(item, out state)) { state = new DraftState(); states.Add(item, state); }
            Controller.Evaluate(config, new OutlookPort(item, (object)app.Session), state, sending);
        }
        void OnSend(object item, ref bool cancel)
        {
            if (cancel) return;
            try
            {
                dynamic m = item;
                if ((int)m.Class != 43) return;
                if (settingsOpen) throw new InvalidOperationException("Close the AutoFrom settings window before sending.");
                if (configError != null) throw new InvalidOperationException(configError);
                Evaluate(item, true);
            }
            catch (Exception ex)
            {
                cancel = true;
                Show("Message was not sent.\n\n" + ex.Message + "\n\nCorrect the recipients or open AutoFrom > Settings, then try again.");
            }
        }
        static void Show(string message) { MessageBox.Show(message, "AutoFrom", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        void Stop()
        {
            if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; }
            if (inspectorsSource != null && newInspectorHandler != null)
            {
                try { ComEventsHelper.Remove(inspectorsSource, InspectorsEvents, 61441, newInspectorHandler); } catch { }
            }
            newInspectorHandler = null;
            inspectorsSource = null;
            if (application != null && sendHandler != null)
            {
                try { ComEventsHelper.Remove(application, ApplicationEvents, 61442, sendHandler); } catch { }
            }
            sendHandler = null;
            states.Clear();
            draftErrors.Clear();
            application = null;
        }
        public void OnDisconnection(int mode, ref Array custom) { Stop(); }
        public void OnBeginShutdown(ref Array custom) { Stop(); }
        public void OnStartupComplete(ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
    }
}
