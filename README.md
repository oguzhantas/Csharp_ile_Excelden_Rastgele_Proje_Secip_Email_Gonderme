## Projeleri rastgele seçerek öğrencilere Toplu E-posta Gönderme

Öğrencilerin listesi ADI, SOYADI, EPOSTA olarak ogrenciler.xlsx dosyasında olacaktır.

![Öğrenciler Excel Dosyası ](https://github.com/oguzhantas/Csharp_ile_Excelden_Rastgele_Proje_Secip_Email_Gonderme/blob/main/ogrenciler_xlsx_dosyasi.png)

Projeler listesi A1'de tek sütun olarak projeler.xlsx dosyasında olacaktır.

![Projeler Excel Dosyası ](https://github.com/oguzhantas/Csharp_ile_Excelden_Rastgele_Proje_Secip_Email_Gonderme/blob/main/projeler_xlsx_dosyasi.png)

Her öğrenciye bir proje verilecek ve rastgele seçilerek e-posta gönderilmektedir.

Hangi öğrenciye, hangi projenin atandığı projenizin çalıştığı klasöre atama_sonuclari.xlsx adıyla kaydedilir.

![Atama Sonuclari Excel Dosyası ](https://github.com/oguzhantas/Csharp_ile_Excelden_Rastgele_Proje_Secip_Email_Gonderme/blob/main/atama_sonuclari_xlsx_dosyasi.png)

## Ayarlar
- Visual Studio Code'u aşağıdaki linkten indiriniz.

https://code.visualstudio.com/download

- .NET Kütüphanesini aşağıdaki linkten indirip kurunuz.

https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-8.0.25-windows-x64-installer?cid=getdotnetcore

- Visual Studio Code'da Terminal-New Terminal komut verdikten sonra MailKit kütüphanesini aşağıdaki komutla kurabilirsiniz.

dotnet add package MailKit

- EmailSender.cs dosyasında aşağıdaki ayarları kendinize göre değiştiriniz.

    private static readonly string SmtpHost    = "smtp.gmail.com";
    
    private static readonly int    SmtpPort    = 587;
    
    private static readonly string SenderEmail = "Gmail adresiniz";
    
    private static readonly string SenderPass  = "Şifreniz";

- Yine Visual Studio Code Terminal'den aşağıdaki komutu yazarak çalıştırınız.

   dotnet run

![Gönderilen E-posta ](https://github.com/oguzhantas/Csharp_ile_Excelden_Rastgele_Proje_Secip_Email_Gonderme/blob/main/gelen_email.png)
  
