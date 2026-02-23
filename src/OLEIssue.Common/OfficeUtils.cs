using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace OLEIssue.Common
{
    public static class OfficeUtils
    {
        /// <summary>
        /// When true, event handlers call Marshal.ReleaseComObject on COM parameters to fix
        /// OLE embedding errors (e.g. embedded Excel in Word). Set to false to reproduce the issue.
        /// </summary>
        public static bool UseComObjectReleaseWorkaround { get; set; } = true;

        /// <summary>
        /// Conditionally releases a COM object if UseComObjectReleaseWorkaround is enabled.
        /// Call this in event handler finally blocks for Workbook, Document, Window parameters.
        /// </summary>
        public static void ConditionalReleaseComObject(object comObject)
        {
            if (!UseComObjectReleaseWorkaround || comObject == null) return;
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // Ignore - object may already be released
            }
        }

        public static bool IsStartedByAutomation()
        {
            var commandLineArgs = Environment.GetCommandLineArgs();
            Logger.Log().Debug("Got command line args: {0}", string.Join(", ", commandLineArgs));

            var isStartedByAutomation =
                commandLineArgs.Any(x => x.ToLowerInvariant() == "/automation")
                || commandLineArgs.Any(x => x.ToLowerInvariant() == "-embedding");

            return isStartedByAutomation;
        }
    }
}