using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;


// EPPlus 8+ için yeni lisans ayarı
ExcelPackage.License.SetNonCommercialPersonal("Oğuzhan");
static void SaveResultsToExcel(List<Student> students, string outputPath)
{
    using var package = new ExcelPackage();
    var sheet = package.Workbook.Worksheets.Add("Proje Atamaları");

    // Başlık satırı
    sheet.Cells[1, 1].Value = "Ad Soyad";
    sheet.Cells[1, 2].Value = "E-posta";
    sheet.Cells[1, 3].Value = "Atanan Proje";
    sheet.Cells[1, 4].Value = "Tarih";

    // Başlık stili
    using (var range = sheet.Cells[1, 1, 1, 4])
    {
        range.Style.Font.Bold      = true;
        range.Style.Font.Color.SetColor(System.Drawing.Color.White);
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(44, 62, 80));
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    }

    // Veri satırları
    for (int i = 0; i < students.Count; i++)
    {
        int row = i + 2;
        var s   = students[i];

        sheet.Cells[row, 1].Value = s.Name;
        sheet.Cells[row, 2].Value = s.Email;
        sheet.Cells[row, 3].Value = s.AssignedProject;
        sheet.Cells[row, 4].Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

        // Zebra renklendirme
        if (i % 2 == 0)
        {
            using var range = sheet.Cells[row, 1, row, 4];
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(235, 240, 255));
        }
    }

    // Sütun genişlikleri
    sheet.Column(1).Width = 22;
    sheet.Column(2).Width = 30;
    sheet.Column(3).Width = 60;
    sheet.Column(4).Width = 18;

    // Tüm hücrelere border
    using (var all = sheet.Cells[1, 1, students.Count + 1, 4])
    {
        all.Style.Border.Top.Style    = ExcelBorderStyle.Thin;
        all.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        all.Style.Border.Left.Style   = ExcelBorderStyle.Thin;
        all.Style.Border.Right.Style  = ExcelBorderStyle.Thin;
    }

    // Metni kaydır (uzun proje açıklamaları için)
    sheet.Cells[1, 3, students.Count + 1, 3].Style.WrapText = true;
    sheet.Column(3).Width = 60;

    package.SaveAs(new FileInfo(outputPath));
    Console.WriteLine($"\n📁 Atama sonuçları kaydedildi → {outputPath}");
}

Console.WriteLine("📊 Excel dosyaları okunuyor...");

var students = ExcelReader.ReadStudents("ogrenciler.xlsx");
var projects = ExcelReader.ReadProjects("projeler.xlsx");

Console.WriteLine($"👥 {students.Count} öğrenci, 📁 {projects.Count} proje bulundu.\n");

ProjectAssigner.Assign(students, projects);

Console.WriteLine("📋 Atama Sonuçları:");
Console.WriteLine(new string('-', 70));
foreach (var s in students)
    Console.WriteLine($"  {s.Name,-20}  {s.Surname,-20} → {s.AssignedProject[..Math.Min(45, s.AssignedProject.Length)]}...");
Console.WriteLine(new string('-', 70));

Console.WriteLine("\n📧 E-postalar gönderiliyor...");
foreach (var student in students)
{
    try   { await EmailSender.SendAsync(student); }
    catch (Exception ex) { Console.WriteLine($"❌ Hata ({student.Email}): {ex.Message}"); }
}

// ✅ Atama sonuçlarını Excel'e kaydet
SaveResultsToExcel(students, "atama_sonuclari.xlsx");

Console.WriteLine("\n✅ İşlem tamamlandı!");


