# DotNet (C#)

Console exercises from the C# / .NET course. Each file is a standalone program with its own `Main` method.

Compile and run one file:

```bash
csc 01_linq_even_numbers.cs && 01_linq_even_numbers.exe
```

| # | File | Task |
|---|------|------|
| 1 | [01_linq_even_numbers.cs](01_linq_even_numbers.cs) | LINQ query that selects even numbers from an array |
| 2 | [02_christmas_eve_checker.cs](02_christmas_eve_checker.cs) | Check if a date is Christmas Eve |
| 3 | [03_animal_farm_legs.cs](03_animal_farm_legs.cs) | Count all legs on a farm |
| 4 | [04_series_sum_1_11_111.cs](04_series_sum_1_11_111.cs) | Sum of the series 1 + 11 + 111 + ... |
| 5 | [05_diamond_pattern.cs](05_diamond_pattern.cs) | Draw a diamond of stars |
| 6 | [06_prime_check.cs](06_prime_check.cs) | Check if a number is prime |
| 7 | [07_menu_calculator.cs](07_menu_calculator.cs) | Menu driven calculator with `switch` |
| 8 | [08_month_name_switch.cs](08_month_name_switch.cs) | Month name from its number |
| 9 | [09_sort_string_alphabetically.cs](09_sort_string_alphabetically.cs) | Sort the characters of a string |
| 10 | [10_file_size.cs](10_file_size.cs) | File size with `FileInfo` |
| 11 | [11_static_calculator_class.cs](11_static_calculator_class.cs) | Calculator class with static methods and exceptions |
| 12 | [12_star_triangle.cs](12_star_triangle.cs) | Draw a triangle of stars |
| 13 | [13_minutes_to_seconds.cs](13_minutes_to_seconds.cs) | Convert minutes to seconds |
| 14 | [14_recursion_print_natural_numbers.cs](14_recursion_print_natural_numbers.cs) | Recursion: print the first n natural numbers |
| 15 | [15_recursion_countdown.cs](15_recursion_countdown.cs) | Recursion: count down from n to 1 |
| 17 | [17_recursion_sum.cs](17_recursion_sum.cs) | Recursion: sum of the first n numbers |
| 18 | [18_starts_with_csharp.cs](18_starts_with_csharp.cs) | Check if a string starts with "C#" |
| 19 | [19_max_of_three.cs](19_max_of_three.cs) | Largest of three numbers |
| 20 | [20_reverse_three_letters.cs](20_reverse_three_letters.cs) | Print three letters in reverse order |

## Cleanup notes

- Files were renamed from `1.cs` to `20.cs` to descriptive names, and a one line task description was added on top of each file.
- `16.cs` was an exact copy of `15.cs` and was removed (that's why number 16 is missing).
- Fixed in `02`: Christmas Eve was checked for month 11 instead of 12.
- Fixed in `07`: the menu did not list option 4 (division) and 5 (exit).
