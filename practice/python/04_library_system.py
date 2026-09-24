"""
Exercise 4: Library management (console app)

Small menu driven console app to manage books and members.
Data is saved to a JSON file so it survives between runs.

Topics: classes, lists, JSON, menus, input validation
"""

import json
import os

DATA_FILE = "library.json"


class Book:
    def __init__(self, isbn, title, author, borrowed_by=None):
        self.isbn = isbn
        self.title = title
        self.author = author
        self.borrowed_by = borrowed_by

    def is_available(self):
        return self.borrowed_by is None

    def to_dict(self):
        return self.__dict__

    def __str__(self):
        status = "available" if self.is_available() else f"borrowed by {self.borrowed_by}"
        return f"[{self.isbn}] {self.title} by {self.author} ({status})"


class Library:
    def __init__(self):
        self.books = []

    def add_book(self, book):
        if self.find(book.isbn):
            raise ValueError("A book with this ISBN already exists")
        self.books.append(book)

    def find(self, isbn):
        for book in self.books:
            if book.isbn == isbn:
                return book
        return None

    def search(self, text):
        text = text.lower()
        return [b for b in self.books if text in b.title.lower() or text in b.author.lower()]

    def borrow(self, isbn, member):
        book = self.find(isbn)
        if book is None:
            raise ValueError("Book not found")
        if not book.is_available():
            raise ValueError("Book is already borrowed")
        book.borrowed_by = member

    def give_back(self, isbn):
        book = self.find(isbn)
        if book is None or book.is_available():
            raise ValueError("This book is not borrowed")
        book.borrowed_by = None

    def save(self, path=DATA_FILE):
        with open(path, "w", encoding="utf-8") as f:
            json.dump([b.to_dict() for b in self.books], f, indent=2)

    @classmethod
    def load(cls, path=DATA_FILE):
        library = cls()
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                for data in json.load(f):
                    library.books.append(Book(**data))
        return library


def menu():
    library = Library.load()
    actions = {
        "1": "List books",
        "2": "Add book",
        "3": "Search",
        "4": "Borrow book",
        "5": "Return book",
        "0": "Save and exit",
    }
    while True:
        print()
        for key, name in actions.items():
            print(f"{key}. {name}")
        choice = input("Choice: ").strip()

        try:
            if choice == "1":
                for book in library.books:
                    print(book)
                if not library.books:
                    print("No books yet.")
            elif choice == "2":
                library.add_book(Book(input("ISBN: "), input("Title: "), input("Author: ")))
                print("Book added.")
            elif choice == "3":
                for book in library.search(input("Search for: ")):
                    print(book)
            elif choice == "4":
                library.borrow(input("ISBN: "), input("Member name: "))
                print("Book borrowed.")
            elif choice == "5":
                library.give_back(input("ISBN: "))
                print("Book returned.")
            elif choice == "0":
                library.save()
                print("Saved. Bye!")
                break
            else:
                print("Unknown option.")
        except ValueError as e:
            print("Error:", e)


if __name__ == "__main__":
    menu()
