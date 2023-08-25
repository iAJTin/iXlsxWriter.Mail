
using System;
using System.Diagnostics;
using System.IO;

using iTin.Core.ComponentModel;
using iTin.Core.ComponentModel.Results;
using iTin.Mail.Smtp.Net;
using iTin.Mime;

using MimeKit;

using iXlsxWriter.Abstractions.Writer.Operations.Results;

using iTinIO = iTin.Core.IO;

namespace iXlsxWriter.Abstractions.Writer.Operations.Actions;

/// <summary>
/// Specialization of <see cref="IOutputAction"/> interface that send the file by email.
/// </summary>
/// <seealso cref="IOutputAction"/>
public class SendMail : IOutputAction
{
    #region private constants

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private const string XlsxExtension = "xlsx";

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private const string ZipExtension = "zip";

    #endregion

    #region interfaces

    #region IOutputAction

    /// <summary>
    /// Execute action for specified output result data.
    /// </summary>
    /// <param name="context">Target output result data.</param>
    /// <returns>
    /// <para>
    /// A <see cref="BooleanResult"/> which implements the <see cref="iTin.Core.ComponentModel.IResult{T}"/> interface reference that contains the result of the operation, to check if the operation is correct, the <b>Success</b>
    /// property will be <b>true</b> and the <b>Value</b> property will contain the value; Otherwise, the the <b>Success</b> property
    /// will be false and the <b>Errors</b> property will contain the errors associated with the operation, if they have been filled in.
    /// </para>
    /// <para>
    /// The type of the return value is <see cref="bool"/>, which contains the operation result
    /// </para>
    /// </returns>
    public IResult Execute(IOutputResultData context) => ExecuteImpl(context);

    #endregion

    #endregion

    #region public properties   

    /// <summary>
    /// Gets or sets the email settings
    /// </summary>
    /// <value>
    /// The email settings.
    /// </value>
    public SmtpMailSettings Settings { get; set; }

    /// <summary>
    /// Gets or sets the email settings
    /// </summary>
    /// <value>
    /// The email settings.
    /// </value>
    public string FromAddress { get; set; }

    /// <summary>
    /// Gets or sets the display name for <see cref="FromAddress"/> email address.
    /// </summary>
    /// <value>
    /// The display name.
    /// </value>
    public string FromDisplayName { get; set; }

    /// <summary>
    /// Gets or sets the attached filename.
    /// </summary>
    /// <value>
    /// The attached filename.
    /// </value>
    public string AttachedFilename { get; set; }

    #endregion

    #region private methods

    private IResult ExecuteImpl(IOutputResultData data)
    {
        if (data == null)
        {
            return BooleanResult.NullResult;
        }

        if (Settings == null)
        {
            return BooleanResult.CreateErrorResult("Missing a valid settings");
        }

        try
        {
            var message = new MimeMessage();
            message.Sender = MailboxAddress.Parse(FromAddress);
            message.Sender.Name = FromDisplayName;
            message.Subject = Settings.Templates.SubjectTemplate;

            foreach (var to in Settings.Recipients.ToAddresses)
            {
                message.To.Add(MailboxAddress.Parse(to));
            }

            foreach (var cc in Settings.Recipients.CCAddresses)
            {
                message.Cc.Add(MailboxAddress.Parse(cc));
            }

            var builder = new BodyBuilder();
            var fileExtension = data.IsZipped ? ZipExtension : XlsxExtension;
            var filename = Path.ChangeExtension(AttachedFilename, fileExtension);
            builder.Attachments.Add(
                filename, 
                data.GetOutputStream(), 
                ContentType.Parse(MimeMapping.Parse(fileExtension)));

            foreach (var attachment in Settings.Attachments)
            {
                var filenameNormalized = iTinIO.Path.PathResolver(attachment);
                var fi = new FileInfo(filenameNormalized);
                if (!fi.Exists)
                {
                    continue;
                }

                builder.Attachments.AddAsync(
                    fi.Name,
                    fi.Open(FileMode.Open, FileAccess.Read, FileShare.Read), 
                    ContentType.Parse(MimeMapping.Parse(fi.Extension)));
            }

            var isHtmlbody = Settings.Templates.IsBodyHtml;
            if (isHtmlbody)
            {
                builder.HtmlBody = Settings.Templates.BodyTemplate;
            }
            else
            {
                builder.TextBody = Settings.Templates.BodyTemplate;
            }

            message.Body = builder.ToMessageBody();
            var mail = new SmtpMail(Settings);

            return mail.SendMail(message);
        }
        catch (Exception e)
        {
            return BooleanResult.FromException(e);
        }
    }

    #endregion
}
