"""
Exercise 3: Sorting and searching algorithms

Own implementations of the classic algorithms from the Algorithms and
Data Structures course, checked against Python's built-in sorted().

Topics: algorithms, recursion, time complexity
"""

import random


def bubble_sort(items):
    a = list(items)
    n = len(a)
    for i in range(n):
        swapped = False
        for j in range(n - 1 - i):
            if a[j] > a[j + 1]:
                a[j], a[j + 1] = a[j + 1], a[j]
                swapped = True
        if not swapped:
            break
    return a


def selection_sort(items):
    a = list(items)
    for i in range(len(a)):
        smallest = i
        for j in range(i + 1, len(a)):
            if a[j] < a[smallest]:
                smallest = j
        a[i], a[smallest] = a[smallest], a[i]
    return a


def insertion_sort(items):
    a = list(items)
    for i in range(1, len(a)):
        key = a[i]
        j = i - 1
        while j >= 0 and a[j] > key:
            a[j + 1] = a[j]
            j -= 1
        a[j + 1] = key
    return a


def merge_sort(items):
    if len(items) <= 1:
        return list(items)
    middle = len(items) // 2
    left = merge_sort(items[:middle])
    right = merge_sort(items[middle:])

    result = []
    i = j = 0
    while i < len(left) and j < len(right):
        if left[i] <= right[j]:
            result.append(left[i])
            i += 1
        else:
            result.append(right[j])
            j += 1
    result.extend(left[i:])
    result.extend(right[j:])
    return result


def binary_search(sorted_items, target):
    """Return the index of target in sorted_items, or -1 if it is not there."""
    low, high = 0, len(sorted_items) - 1
    while low <= high:
        mid = (low + high) // 2
        if sorted_items[mid] == target:
            return mid
        if sorted_items[mid] < target:
            low = mid + 1
        else:
            high = mid - 1
    return -1


if __name__ == "__main__":
    numbers = [random.randint(1, 100) for _ in range(15)]
    print("Original:", numbers)

    for sort in (bubble_sort, selection_sort, insertion_sort, merge_sort):
        result = sort(numbers)
        assert result == sorted(numbers), sort.__name__
        print(f"{sort.__name__:<15} {result}")

    sorted_numbers = merge_sort(numbers)
    target = numbers[0]
    print(f"binary_search({target}) -> index {binary_search(sorted_numbers, target)}")
    print("binary_search(1000) ->", binary_search(sorted_numbers, 1000))
