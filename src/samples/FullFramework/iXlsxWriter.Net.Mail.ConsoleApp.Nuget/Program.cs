
using System;

using iTin.Logging;
using iTin.Logging.ComponentModel;

namespace iPdfWriter.Net.Mail.ConsoleApp
{
    using Code;
    using ComponentModel;

    internal static class Program
    {
        static void Main(string[] args)
        {
            Console.Title = Constants.AppName;

            ILogger logger = new Logger(Constants.AppName, new ILog[] { new FileLog(), new ColoredConsoleLog() }) { Deep = LogDeep.OnlyApplicationCalls, Status = LogStatus.Running };
            logger.Debug(">Start Logging<");

            // 01. Shows how to send pdf output result by email with gmail
            logger.Info("");
            logger.Info("> Start Sample 01");
            logger.Info(" > Shows how to send pdf output result by email with gmail");
            Sample01.Generate();

            // 02. Shows how to send pdf output result by email with Mailtrap (fake SMTP provider service)
            logger.Info("");
            logger.Info("> Start Sample 02");
            logger.Info(" > Shows how to send pdf output result by email with Mailtrap (fake SMTP provider service)");
            Sample02.Generate();

            // 03. Shows how to send pdf output result by email with Ethereal (fake SMTP provider service)
            logger.Info("");
            logger.Info("> Start Sample 03");
            logger.Info(" > Shows how to send pdf output result by email with Ethereal (fake SMTP provider service)");
            Sample03.Generate();

            logger.Info("");
            logger.Debug(">End Logging<");
            Console.ReadKey();
        }
    }
}
