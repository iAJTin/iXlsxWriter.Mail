# What is iXlsxWriter.Mail?

**iXlsxWriter.Mail**, extends [iXlsxWriter](https://github.com/iAJTin/iXlsxWriter), contains extension methods to send by mail **XlsxInput** instances as well as **OutputResult**.

I hope it helps someone. :smirk:

## Usage

### Samples

#### Sample 1 - Shows the use of synchronous mail by SendMail action

``` csharp
var doc = XlsxInput.Create("Sheet1");
doc.Insert(
    new InsertText
    {
        SheetName = "Sheet1",
        Data = "Hello world! from iXlsxWriter"
    });

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
                BodyTemplate = "Hola",
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
    // Handle error(s)
}
```             

#### Sample 2 - Shows the use of asynchronous mail by SendMailAsync action

```csharp   
var doc = XlsxInput.Create("Sheet1");
doc.Insert(
    new InsertText
    {
        SheetName = "Sheet1",
        Data = "Hello world! from iXlsxWriter"
    });

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
    // Handle error(s)
}
```

## Documentation

 - Please see next link [documentation](https://github.com/iAJTin/iXlsxWriter.Mail/blob/main/documentation/iXlsxWriter.Mail.md).

## Changes

For more information, please visit the next link [CHANGELOG](https://github.com/iAJTin/iXlsxWriter.Mail/blob/main/CHANGELOG.md)
