using System;
using System.Collections.Generic;

namespace AutoFrom
{
    public sealed class Snapshot
    {
        public object AccountObject;
        public string Account, Represented;
        public Dictionary<string, object> RepresentingProperties;
    }
    public interface IComposePort
    {
        Snapshot Capture();
        IList<string> Recipients(bool resolve, out bool complete);
        void Set(Sender target);
        void Restore(Snapshot snapshot);
    }
    public sealed class DraftState
    {
        public Snapshot Original;
        public bool Managed;
    }
    public static class Controller
    {
        public static void RestoreManaged(IComposePort port, DraftState state)
        {
            if (!state.Managed) return;
            port.Restore(state.Original);
            state.Managed = false;
        }
        public static void Evaluate(Config config, IComposePort port, DraftState state, bool sending)
        {
            if (!config.Enabled || !config.HasRouting) return;
            if (state.Original == null) state.Original = port.Capture();
            bool complete;
            var recipients = port.Recipients(sending, out complete);
            var decision = RuleEngine.Decide(config, state.Original.Account, recipients, complete);
            if (decision.Error != null) throw new InvalidOperationException(decision.Error);
            if (decision.Target != null)
            {
                // Mark first: even a partially failing write must be restored when the match disappears.
                state.Managed = true;
                port.Set(decision.Target);
            }
            else if (state.Managed)
            {
                RestoreManaged(port, state);
            }
        }
    }
}
