using Serilog;
using Serilog.Core;
using Serilog.Events;
using Serilog.Templates;
using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;

namespace OLEIssue.Common
{
    public static class Logger
    {
        private const string LogTemplate = "{@t:yyyy-MM-ddTHH:mm:ss.fffZ} {#if Caller is not null} [{Caller,50}]{#end}" + " {@l,15:u}: {@m}\n{@x}";

        private static readonly LoggingLevelSwitch LogLevelSwitch = new LoggingLevelSwitch(LogEventLevel.Debug);
        private static ILogger _logger;

        public static string LogFilePath { get; private set; }

        static Logger()
        {
            var logDirectory = AppDomain.CurrentDomain.BaseDirectory;

            var path = Path.Combine(logDirectory, $"log_{DateTime.Now:yyyy_MM_dd_HH_mm_ss}.txt");
            SetLogPath(path);
        }

        public static void SetLogPath(string logFilePath)
        {
            LogFilePath = logFilePath;
            InitLogger();
        }

        private static void InitLogger()
        {
            var loggerConfig = new LoggerConfiguration()
                .MinimumLevel.ControlledBy(LogLevelSwitch)
                .Enrich.FromLogContext()
                .WriteTo.File(new ExpressionTemplate(LogTemplate), LogFilePath, retainedFileCountLimit: 50)
                .WriteTo.Trace(new ExpressionTemplate(LogTemplate))
                ;

            _logger = loggerConfig.CreateLogger();
        }

        public static ILogger Log([CallerFilePath] string callerFilePath = "", [CallerMemberName] string memberName = "", [CallerLineNumber] int lineNumber = 0)
        {
            var callingMethod = GetCallerMethod(callerFilePath, memberName, lineNumber);
            return _logger?.ForContext("Caller", callingMethod);
        }

        private static string GetCallerMethod(string callerFilePath, string memberName, int lineNumber)
        {
            var callerClassName = Path.GetFileNameWithoutExtension(callerFilePath);
            return "T" + Thread.CurrentThread.ManagedThreadId + "\\" + callerClassName + "." + memberName + ":" + lineNumber;
        }
    }
}