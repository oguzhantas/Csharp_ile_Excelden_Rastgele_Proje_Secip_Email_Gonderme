## Projeleri rastgele seçerek öğrencilere Toplu E-posta Gönderme

Öğrencilerin listesi ADI, SOYADI, EPOSTA olarak ogrenciler.xlsx dosyasında olacaktır.

![Öğrenciler Excel Dosyası ](https://github.com/oguzhantas/Csharp_ile_Excelden_Rastgele_Proje_Secip_Email_Gonderme/blob/main/ogrenciler_xlsx_dosyasi.png)

Projeler listesi A1'de tek sütun olarak projeler.xlsx dosyasında olacaktır.

Her öğrenciye bir proje verilecek ve rastgele seçilerek e-posta gönderilmektedir.

atama_sonuclari.xlsx dosyasında hangi öğrenciye hangi projenin atandığı görülmektedir.

## Ayarlar

MailKit kütüphanesini aşağıdaki komutla VS Studio Terminalden kurabilirsiniz.

dotnet add package MailKit

EmailSender.cs dosyasında aşağıdaki ayarları kendinize göre değiştiriniz.

    private static readonly string SmtpHost    = "smtp.gmail.com";
    
    private static readonly int    SmtpPort    = 587;
    
    private static readonly string SenderEmail = "Gmail adresiniz";
    
    private static readonly string SenderPass  = "Şifreniz";
    
