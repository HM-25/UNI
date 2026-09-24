"""
Exercise 1: Student grades from a CSV file

Read student grades from a CSV file, calculate each student's average,
find the best student and write a small report to a new file.

Topics: file I/O, csv module, dictionaries, functions
"""

import csv

SAMPLE_DATA = """name,math,programming,databases
Amir,8,9,10
Lejla,10,10,9
Tarik,6,7,8
Selma,9,8,9
"""


def create_sample_file(path):
    with open(path, "w", encoding="utf-8") as f:
        f.write(SAMPLE_DATA)


def read_grades(path):
    students = {}
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row.pop("name")
            students[name] = {subject: int(grade) for subject, grade in row.items()}
    return students


def average(grades):
    return sum(grades.values()) / len(grades)


def best_student(students):
    return max(students, key=lambda name: average(students[name]))


def write_report(students, path):
    with open(path, "w", encoding="utf-8") as f:
        f.write("Student report\n")
        f.write("==============\n")
        for name, grades in sorted(students.items()):
            f.write(f"{name:<10} average: {average(grades):.2f}\n")
        f.write(f"\nBest student: {best_student(students)}\n")


if __name__ == "__main__":
    create_sample_file("grades.csv")
    students = read_grades("grades.csv")

    for name, grades in students.items():
        print(f"{name}: {grades} -> average {average(grades):.2f}")

    print("Best student:", best_student(students))
    write_report(students, "report.txt")
    print("Report saved to report.txt")
