
using System;

using iTin.Core.ComponentModel;
using iTin.Mail.Smtp.Net;

using iXlsxWriter.Abstractions.Writer.Operations.Actions;
using iXlsxWriter.Operations.Insert;

namespace iXlsxWriter.Net.Mail.ConsoleApp.Code;

/// <summary>
/// Shows how to send xlsx output result by email with <strong>Mailtrap</strong> (fake SMTP provider service).
/// <para>
/// For more information, please see <a href="https://mailtrap.io/home">https://mailtrap.io/home</a>.
/// </para>
/// </summary>
internal static class Sample02
{
    public static void Generate()
    {
        var xlsxCreationResult = BuildXlsxInput().CreateResult();
        if (!xlsxCreationResult.Success)
        {
            return;
        }

        var sendResult = xlsxCreationResult.Result.Action(
            new SendMail
            {
                AttachedFilename = "Sample-01",
                FromDisplayName = "",
                FromAddress = "",
                Settings = new SmtpMailSettings
                {
                    Credential = new SmtpCredential
                    {
                        Port = 2525,
                        UseSsl = true,
                        Host = SmtpMail.MailtrapSmtpHost,
                        Email = "",
                        UserName = "",
                        Password = ""
                    },
                    Templates = new TemplateSettings
                    {
                        IsBodyHtml = true,
                        BodyTemplate = "Hey!!",
                        SubjectTemplate = "Test > xlsx file"
                    },
                    Recipients = new RecipientsSettings
                    {
                        ToAddresses = new[] { "" }
                    }
                }
            });

        if (!sendResult.Success)
        {
            Console.WriteLine(sendResult.Errors.AsMessages().ToStringBuilder());
        }
    }


    private static XlsxInput BuildXlsxInput()
    {
        var doc = XlsxInput.Create("Sheet1");

        doc.Insert(
            new InsertText
            {
                SheetName = "Sheet1",
                Data = "Hello world! from iXlsxWriter"
            });

        return doc;
    }
}
