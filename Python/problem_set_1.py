# 1.Basic Arithmetic with User Input

a = int(input("Enter first number: "))
b = int(input("Enter second number: "))

print("Sum:", a + b)
print("Difference:", a - b)
print("Product:", a * b)
print("Quotient:", a / b)
print("Modulus:", a % b)

# 2.Factorial of a Number with Error Handling

def factorial(n):
    if n < 0:
        return "Factorial is not defined for negative numbers"
    elif n == 0:
        return 1
    else:
        result = 1
        for i in range(1, n + 1):
            result *= i
        return result

num = int(input("Enter a number to find its factorial: "))
print("Factorial of", num, "is", factorial(num))

# 3.Fibonacci Sequence with User Input

def fibonacci(n):
    fib_sequence = [0, 1]
    while len(fib_sequence) < n:
        fib_sequence.append(fib_sequence[-1] + fib_sequence[-2])
    return fib_sequence

num = int(input("Enter the number of Fibonacci numbers to generate: "))
print(f"First {num} Fibonacci numbers:", fibonacci(num))

# 4. Prime Number Check with Multiple Inputs

def is_prime(n):
    if n <= 1:
        return False
    for i in range(2, int(n**0.5) + 1):
        if n % i == 0:
            return False
    return True

numbers = [int(x) for x in input("Enter numbers separated by spaces: ").split()]
results = {num: is_prime(num) for num in numbers}
print("Prime check results:", results)

# 5. List Comprehension with Conditional Filtering

numbers = range(20)
even_squares = [x**2 for x in numbers if x % 2 == 0]
print("Squares of even numbers from 0 to 19:", even_squares)

# 6. Reverse a String with User Input and Palindrome Check

def reverse_string(s):
    return s[::-1]

def is_palindrome(s):
    return s == reverse_string(s)

input_string = input("Enter a string: ")
reversed_string = reverse_string(input_string)
print("Reversed string:", reversed_string)
print("Is the original string a palindrome?", is_palindrome(input_string))

# 7. Sum of Digits with List Comprehension

def sum_of_digits(n):
    return sum(int(digit) for digit in str(n))

num = int(input("Enter a number: "))
print("Sum of digits of", num, "is", sum_of_digits(num))

# 8. Find Maximum in a List with Error Handling

def find_max(lst):
    if not lst:
        return "List is empty"
    return max(lst)

numbers = [int(x) for x in input("Enter numbers separated by spaces: ").split()]
print("Maximum number in the list is", find_max(numbers))

# 9. Simple Calculator with Multiple Operations

def calculator(a, b, operation):
    if operation == 'add':
        return a + b
    elif operation == 'subtract':
        return a - b
    elif operation == 'multiply':
        return a * b
    elif operation == 'divide':
        if b != 0:
            return a / b
        else:
            return "Cannot divide by zero"
    else:
        return "Invalid operation"

a = int(input("Enter first number: "))
b = int(input("Enter second number: "))
operation = input("Enter operation (add, subtract, multiply, divide): ")
print(f"Result of {operation} operation:", calculator(a, b, operation))

# 10. Count Vowels in a String with Case Insensitivity

def count_vowels(s):
    vowels = 'aeiou'
    return sum(1 for char in s.lower() if char in vowels)

input_string = input("Enter a string: ")
print("Number of vowels in the string is", count_vowels(input_string))

# 11. Generate Random Number within a Range

import random

low = int(input("Enter the lower bound: "))
high = int(input("Enter the upper bound: "))
print(f"Random number between {low} and {high}:", random.randint(low, high))

# 12. Find GCD of Two Numbers with Error Handling

import math

a = int(input("Enter first number: "))
b = int(input("Enter second number: "))
if a <= 0 or b <= 0:
    print("GCD is not defined for non-positive numbers")
else:
    print(f"GCD of {a} and {b} is", math.gcd(a, b))

# 13. Sorting a List with User Input

numbers = [int(x) for x in input("Enter numbers separated by spaces: ").split()]
numbers.sort()
print("Sorted list:", numbers)

# 14. Simple Dictionary Operations with User Input

student = {}
student["name"] = input("Enter student's name: ")
student["age"] = int(input("Enter student's age: "))
student["courses"] = input("Enter student's courses separated by commas: ").split(',')

print("Student name:", student["name"])
print("Student age:", student["age"])
print("Courses:", student["courses"])

# 15. Convert Celsius to Fahrenheit and Fahrenheit to Celsius

def celsius_to_fahrenheit(c):
    return (c * 9/5) + 32

def fahrenheit_to_celsius(f):
    return (f - 32) * 5/9

temp_c = float(input("Enter temperature in Celsius: "))
temp_f = float(input("Enter temperature in Fahrenheit: "))
print(f"{temp_c} degrees Celsius is {celsius_to_fahrenheit(temp_c)} Fahrenheit")
print(f"{temp_f} degrees Fahrenheit is {fahrenheit_to_celsius(temp_f)} Celsius")

# 16. Count Words in a String and Identify Longest Word

def count_words(s):
    words = s.split()
    longest_word = max(words, key=len)
    return len(words), longest_word

input_string = input("Enter a string: ")
word_count, longest_word = count_words(input_string)
print(f"Number of words: {word_count}")
print(f"Longest word: {longest_word}")

# 17. Merge Two Lists and Remove Duplicates

list1 = [int(x) for x in input("Enter first list of numbers separated by spaces: ").split()]
list2 = [int(x) for x in input("Enter second list of numbers separated by spaces: ").split()]
merged_list = list(set(list1 + list2))
print("Merged list without duplicates:", merged_list)

# 18. Check if List is Sorted and Sort if Not

def is_sorted(lst):
    return lst == sorted(lst)

numbers = [int(x) for x in input("Enter numbers separated by spaces: ").split()]
if is_sorted(numbers):
    print("The list is already sorted:", numbers)
else:
    print("The list is not sorted. Sorted list:", sorted(numbers))

# 19. Basic Matrix Addition

def matrix_addition(matrix1, matrix2):
    result = [[matrix1[i][j] + matrix2[i][j] for j in range(len(matrix1[0]))] for i in range(len(matrix1))]
    return result

matrix1 = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
matrix2 = [[9, 8, 7], [6, 5, 4], [3, 2, 1]]

result_matrix = matrix_addition(matrix1, matrix2)
print("Resultant Matrix after addition:")
for row in result_matrix:
    print(row)

# 20. Find Common Elements in Two Lists

def find_common_elements(list1, list2):
    return list(set(list1) & set(list2))

list1 = [int(x) for x in input("Enter first list of numbers separated by spaces: ").split()]
list2 = [int(x) for x in input("Enter second list of numbers separated by spaces: ").split()]
common_elements = find_common_elements(list1, list2)
print("Common elements:", common_elements)
