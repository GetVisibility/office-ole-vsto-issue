using System;
using System.Linq;

namespace OLEIssue.Common
{
    public static class OfficeUtils
    {
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