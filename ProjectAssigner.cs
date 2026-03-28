using System;
using System.Collections.Generic;
using System.Linq;

public static class ProjectAssigner
{
    public static void Assign(List<Student> students, List<string> projects)
    {
        if (projects.Count < students.Count)
            throw new InvalidOperationException(
                $"Yetersiz proje sayısı! Öğrenci: {students.Count}, Proje: {projects.Count}");

        var random = new Random();

        // Projeleri karıştır, her öğrenciye benzersiz proje ver
        var shuffled = projects.OrderBy(_ => random.Next()).ToList();

        for (int i = 0; i < students.Count; i++)
            students[i].AssignedProject = shuffled[i];
    }
}