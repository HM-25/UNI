// Exercise 2: Student records with LINQ
//
// A list of students with grades, queried with LINQ:
// averages, filtering, grouping and ordering.

using System;
using System.Collections.Generic;
using System.Linq;

class Student
{
    public int Id { get; set; }
    public string Name { get; set; }
    public int Year { get; set; }
    public List<int> Grades { get; set; }

    public double Average
    {
        get { return Grades.Count == 0 ? 0 : Grades.Average(); }
    }
}

class Program
{
    static void Main()
    {
        List<Student> students = new List<Student>
        {
            new Student { Id = 1, Name = "Amir",  Year = 1, Grades = new List<int> { 8, 9, 7 } },
            new Student { Id = 2, Name = "Lejla", Year = 2, Grades = new List<int> { 10, 10, 9 } },
            new Student { Id = 3, Name = "Tarik", Year = 1, Grades = new List<int> { 6, 7, 6 } },
            new Student { Id = 4, Name = "Selma", Year = 3, Grades = new List<int> { 9, 8, 10 } },
            new Student { Id = 5, Name = "Kenan", Year = 2, Grades = new List<int> { 7, 8, 8 } }
        };

        Console.WriteLine("All students ordered by average:");
        foreach (var s in students.OrderByDescending(s => s.Average))
        {
            Console.WriteLine("  {0,-6} year {1}  average {2:F2}", s.Name, s.Year, s.Average);
        }

        var excellent = students.Where(s => s.Average >= 9).Select(s => s.Name);
        Console.WriteLine("\nStudents with average 9 or higher: " + string.Join(", ", excellent));

        Console.WriteLine("\nAverage per year:");
        var byYear = students
            .GroupBy(s => s.Year)
            .OrderBy(g => g.Key)
            .Select(g => new { Year = g.Key, Count = g.Count(), Average = g.Average(s => s.Average) });

        foreach (var g in byYear)
        {
            Console.WriteLine("  Year {0}: {1} student(s), average {2:F2}", g.Year, g.Count, g.Average);
        }

        Student best = students.OrderByDescending(s => s.Average).First();
        Console.WriteLine("\nBest student: {0} ({1:F2})", best.Name, best.Average);

        Console.Write("\nSearch student by name: ");
        string input = Console.ReadLine() ?? "";
        var found = students.FirstOrDefault(s => s.Name.Equals(input.Trim(), StringComparison.OrdinalIgnoreCase));
        Console.WriteLine(found != null
            ? string.Format("Found: {0}, grades: {1}", found.Name, string.Join(" ", found.Grades))
            : "Student not found.");
    }
}
