using System.IO;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

/// <summary>Sends the generated report as an email attachment via SMTP (MailKit).</summary>
public static class EmailService
{
    public static string? ValidateEmailConfig(EmailConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.SmtpHost))
            return "SMTP Host is required (Email tab).";
        if (config.SmtpPort <= 0)
            return "SMTP Port must be a positive number (Email tab).";
        if (string.IsNullOrWhiteSpace(config.FromAddress))
            return "From Address is required (Email tab).";
        if (ParseRecipients(config.Recipients).Count == 0)
            return "At least one recipient is required (Email tab).";
        return null;
    }

    public static List<string> ParseRecipients(string recipients) =>
        (recipients ?? string.Empty)
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

    public static async Task SendReportAsync(EmailConfig config, string subject, string body, string attachmentPath)
    {
        if (!File.Exists(attachmentPath))
            throw new FileNotFoundException($"Report file not found: {attachmentPath}");

        var message = new MimeMessage();
        message.From.Add(MailboxAddress.Parse(config.FromAddress));
        foreach (var recipient in ParseRecipients(config.Recipients))
            message.To.Add(MailboxAddress.Parse(recipient));
        message.Subject = subject;

        var builder = new BodyBuilder { TextBody = body };
        await builder.Attachments.AddAsync(attachmentPath);
        message.Body = builder.ToMessageBody();

        using var client = new SmtpClient();
        var socketOptions = config.UseSsl
            ? SecureSocketOptions.SslOnConnect
            : SecureSocketOptions.StartTlsWhenAvailable;
        await client.ConnectAsync(config.SmtpHost, config.SmtpPort, socketOptions);
        if (!string.IsNullOrWhiteSpace(config.Username))
            await client.AuthenticateAsync(config.Username, config.Password);
        await client.SendAsync(message);
        await client.DisconnectAsync(true);
    }
}
