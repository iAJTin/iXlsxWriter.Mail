
using iTin.Core.ComponentModel;
using iTin.Mail.Smtp.Net;

using iXlsxWriter.Abstractions.Writer.Operations.Actions;
using iXlsxWriter.Operations.Insert;

namespace iXlsxWriter.Net.Mail.ConsoleApp.Code;

/// <summary>
/// Shows how to send xlsx output result by email with gmail asynchronously.
/// </summary>
internal static class Sample01
{
    public static async Task GenerateAsync(CancellationToken cancellationToken = default)
    {
        var xlsxCreationResult = await BuildXlsxInput()
            .CreateResultAsync(cancellationToken: cancellationToken)
            .ConfigureAwait(false);
        if (!xlsxCreationResult.Success)
        {
            return;
        }

        var sendResult = await xlsxCreationResult.Result.Action(
                new SendMailAsync
                {
                    AttachedFilename = "Sample-01",
                    FromDisplayName = "",
                    FromAddress = "",
                    Settings = new SmtpMailSettings
                    {
                        Credential = new SmtpCredential
                        {
                            Port = 465,
                            UseSsl = true,
                            Host = SmtpMail.GmailSmtpHost,
                            Email = "",
                            UserName = "",
                            Password = ""
                        },
                        Templates = new TemplateSettings
                        {
                            IsBodyHtml = true,
                            BodyTemplate = "Hey!!!",
                            SubjectTemplate = "Test > xlsx file"
                        },
                        Recipients = new RecipientsSettings
                        {
                            ToAddresses = new[] { "" }
                        }
                    }
                },
                cancellationToken)
            .ConfigureAwait(false);

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
