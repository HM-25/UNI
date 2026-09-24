"""
Unit tests for the practice exercises.

Run with:  python -m unittest test_exercises.py
"""

import importlib
import unittest

sorting = importlib.import_module("03_sorting_and_searching")
bank = importlib.import_module("02_bank_account")
library = importlib.import_module("04_library_system")


class SortingTests(unittest.TestCase):
    data = [5, 3, 9, 1, 5, 0, -2, 8]

    def test_all_sorts(self):
        for sort in (sorting.bubble_sort, sorting.selection_sort,
                     sorting.insertion_sort, sorting.merge_sort):
            self.assertEqual(sort(self.data), sorted(self.data), sort.__name__)

    def test_empty_list(self):
        self.assertEqual(sorting.merge_sort([]), [])

    def test_binary_search(self):
        items = sorted(self.data)
        self.assertEqual(items[sorting.binary_search(items, 9)], 9)
        self.assertEqual(sorting.binary_search(items, 100), -1)


class BankTests(unittest.TestCase):
    def test_transfer(self):
        a = bank.BankAccount("A", 100)
        b = bank.BankAccount("B")
        a.transfer(b, 40)
        self.assertEqual(a.balance, 60)
        self.assertEqual(b.balance, 40)

    def test_insufficient_funds(self):
        a = bank.BankAccount("A", 10)
        with self.assertRaises(bank.InsufficientFundsError):
            a.withdraw(50)

    def test_negative_deposit(self):
        with self.assertRaises(ValueError):
            bank.BankAccount("A").deposit(-5)


class LibraryTests(unittest.TestCase):
    def test_borrow_and_return(self):
        lib = library.Library()
        lib.add_book(library.Book("1", "Clean Code", "Robert Martin"))
        lib.borrow("1", "Amir")
        self.assertFalse(lib.find("1").is_available())
        with self.assertRaises(ValueError):
            lib.borrow("1", "Lejla")
        lib.give_back("1")
        self.assertTrue(lib.find("1").is_available())

    def test_search(self):
        lib = library.Library()
        lib.add_book(library.Book("1", "Clean Code", "Robert Martin"))
        lib.add_book(library.Book("2", "The Pragmatic Programmer", "Hunt and Thomas"))
        self.assertEqual(len(lib.search("code")), 1)


if __name__ == "__main__":
    unittest.main()
