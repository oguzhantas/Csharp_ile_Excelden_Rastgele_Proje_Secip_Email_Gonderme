using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

public static class ExcelReader
{
    public static List<Student> ReadStudents(string filePath)
    {
        var students = new List<Student>();
        using var package = new ExcelPackage(new FileInfo(filePath));
        var sheet = package.Workbook.Worksheets[0];

        // Satır 1 başlık, veri 2. satırdan başlıyor
        for (int row = 2; row <= sheet.Dimension.End.Row; row++)
        {
            var name  = sheet.Cells[row, 1].Text.Trim();
            var surname  = sheet.Cells[row, 2].Text.Trim();
            var email = sheet.Cells[row, 3].Text.Trim();

            if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(email))
                students.Add(new Student { Name = name, Surname=surname, Email = email });
        }
        return students;
    }

    public static List<string> ReadProjects(string filePath)
    {
        var projects = new List<string>();
        using var package = new ExcelPackage(new FileInfo(filePath));
        var sheet = package.Workbook.Worksheets[0];

        for (int row = 2; row <= sheet.Dimension.End.Row; row++)
        {
            var desc = sheet.Cells[row, 1].Text.Trim();
            if (!string.IsNullOrEmpty(desc))
                projects.Add(desc);
        }
        return projects;
    }
}

public class Student
{
    public string Name  { get; set; } = "";
     public string Surname  { get; set; } = "";
    public string Email { get; set; } = "";
    public string AssignedProject { get; set; } = "";
}