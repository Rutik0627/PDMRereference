using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PDMReReferenceAddin
{
    /// <summary>
    /// Read-only transition listing following the VBA macro's late-bound COM call pattern.
    /// Does not execute transitions or modify files.
    /// </summary>
    public static class PdmTransitionInspector
    {
        public static string Inspect(object edmFile)
        {
            if (edmFile == null)
                return "TRANSITIONS=UNAVAILABLE; Reason=File object is null";

            string stateName = "(unknown)";
            try
            {
                dynamic file = edmFile;
                dynamic state = file.CurrentState;
                if (state == null)
                    return "TRANSITIONS=UNAVAILABLE; Reason=CurrentState is null";

                try { stateName = Convert.ToString(state.Name) ?? stateName; } catch { }

                dynamic pos = state.GetFirstTransitionPosition();
                List<string> transitions = new List<string>();
                int guard = 0;

                while (pos != null && guard++ < 500)
                {
                    bool isNull = false;
                    try { isNull = Convert.ToBoolean(pos.IsNull); } catch { }
                    if (isNull) break;

                    // This intentionally mirrors the VBA: state.GetNextTransition(pos).
                    dynamic transition = state.GetNextTransition(pos);
                    if (transition == null) break;

                    string name = "(unnamed transition)";
                    string destinationName = "(destination unavailable)";
                    try { name = Convert.ToString(transition.Name) ?? name; } catch { }
                    // VBA uses trans.ToState.Name, not DestState/DestinationState.
                    try
                    {
                        dynamic destination = transition.ToState;
                        if (destination != null)
                            destinationName = Convert.ToString(destination.Name) ?? destinationName;
                    }
                    catch { }

                    transitions.Add(name + " -> " + destinationName);
                }

                if (guard >= 500)
                    return "State=" + Clean(stateName) + "; TRANSITIONS=INCOMPLETE; Reason=Safety limit reached";

                return "State=" + Clean(stateName) + "; OutgoingTransitionCount=" + transitions.Count +
                    "; OutgoingTransitions=" + (transitions.Count == 0 ? "(none returned)" : Clean(String.Join(" | ", transitions.ToArray())));
            }
            catch (Exception ex)
            {
                Exception root = ex;
                while (root.InnerException != null) root = root.InnerException;
                return "State=" + Clean(stateName) + "; TRANSITIONS=UNAVAILABLE; Method=VBA_STYLE_LATE_BOUND_COM; Reason=" +
                    Clean(root.GetType().Name + ": " + root.Message);
            }
        }

        private static string Clean(string value)
        {
            return (value ?? "").Replace("\r", " ").Replace("\n", " ").Replace(";", ",");
        }
    }
}
