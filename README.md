<p align="center">
  <img src="https://github.com/iAJTin/iXlsxWriter.Mail/blob/master/nuget/iXlsxWriter.Mail.png" height="32"/>
</p>
<p align="center">
  <a href="https://github.com/iAJTin/iXlsxWriter.Mail">
    <img src="https://img.shields.io/badge/iTin-iXlsxWriter.Mail-green.svg?style=flat"/>
  </a>
</p>

***

# What is iXlsxWriter.Mail?

**iXlsxWriter.Mail**, extends [iXlsxWriter](https://github.com/iAJTin/iXlsxWriter), contains extension methods to send by mail **XlsxInput** instances as well as **OutputResult**.

I hope it helps someone. :smirk:

# Install via NuGet

- From nuget gallery

<table>
  <tr>
    <td>
      <a href="https://github.com/iAJTin/iXlsxWriter.Mail">
        <img src="https://img.shields.io/badge/-iXlsxWriter.Mail-green.svg?style=flat"/>
      </a>
    </td>
    <td>
      <a href="https://www.nuget.org/packages/iXlsxWriter.Mail/">
        <img alt="NuGet Version" 
             src="https://img.shields.io/nuget/v/iXlsxWriter.Mail.svg" /> 
      </a>
    </td>  
  </tr>
</table>

- From package manager console

```PM> Install-Package iXlsxWriter.Mail```

# Usage

## Samples

### Sample 1 - Shows the use of synchronous mail by SendMail action

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

### Sample 2 - Shows the use of asynchronous mail by SendMailAsync action

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

# Documentation

 - Please see next link [documentation].

# How can I send feedback!!!

If you have found **iXlsxWriter.Mail** useful at work or in a personal project, I would love to hear about it. If you have decided not to use **iXlsxWriter.Mail**, please send me and email stating why this is so. I will use this feedback to improve **iXlsxWriter.Mail** in future releases.

My email address is 

![email.png][email] 


[email]: ./assets/email.png "email"
[documentation]: ./documentation/iXlsxWriter.Mail.md
