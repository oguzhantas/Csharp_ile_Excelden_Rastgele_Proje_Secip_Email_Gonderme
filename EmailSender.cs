using System;
using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

public static class EmailSender
{
    private static readonly string SmtpHost    = "smtp.gmail.com";
    private static readonly int    SmtpPort    = 587;
    private static readonly string SenderEmail = "Gmail adresiniz";
    private static readonly string SenderPass  = "Şifreniz";

    public static async Task SendAsync(Student student)
    {
        var message = new MimeMessage();
        message.From.Add(new MailboxAddress("Nesne Yönelimli Dersi Uygulama Projesi", SenderEmail));
        message.To.Add(new MailboxAddress(student.Name+" "+student.Surname, student.Email));
        message.Subject = "Nesne Yönelimli Programlama Dersi Uygulama Projesi Atanması";
        message.Body    = new TextPart("html") { Text = BuildEmailBody(student) };


        using var smtp = new SmtpClient();

        // STARTTLS ile bağlan (SSL değil)  SecureSocketOptions.StartTls
        await smtp.ConnectAsync(SmtpHost, SmtpPort, SecureSocketOptions.Auto);
        await smtp.AuthenticateAsync(SenderEmail, SenderPass);
        await smtp.SendAsync(message);
        await smtp.DisconnectAsync(true);

        Console.WriteLine($"✅ E-posta gönderildi → {student.Name} {student.Surname}({student.Email})");  
     
    }

    private static string BuildEmailBody(Student student) => $"""
        <html><body style="font-family:Arial,sans-serif;padding:20px">
            <h2 style="color:#2c3e50">🎓 Proje Atama Bildirimi</h2>
            <p>Sayın <strong>{student.Name} {student.Surname}</strong>,</p>
            <p>Bu dönem için size aşağıdaki proje atanmıştır:</p>
            <div style="background:#f0f4ff;border-left:4px solid #3498db;padding:15px;margin:15px 0">
                <p style="margin:0;font-size:15px">{student.AssignedProject}</p>
                <p> <br>Dr. Oğuzhan TAŞ (Bu e-postayı kesinlikle cevaplamayınız.) <br>
                    Proje Başlangıç Tarihi: 28.03.2026 <br>
                    Proje Son Teslim Tarihi: 28.04.2026 Saat 23:59 
                </p>
            </div>
            <p>Başarılar dilerim! 🚀</p>
            <hr/>
            <small style="color:#999">Bu e-posta otomatik olarak gönderilmiştir.<br>
            
            
            </small>
        </body></html>
        """;
}