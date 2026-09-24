"""
Exercise 2: Bank account (OOP)

A simple bank account class with deposit, withdraw and transfer.
Invalid operations raise exceptions instead of failing silently.

Topics: classes, methods, custom exceptions, inheritance
"""


class InsufficientFundsError(Exception):
    pass


class BankAccount:
    def __init__(self, owner, balance=0.0):
        self.owner = owner
        self.balance = balance
        self.history = []

    def deposit(self, amount):
        if amount <= 0:
            raise ValueError("Deposit amount must be positive")
        self.balance += amount
        self.history.append(("deposit", amount))

    def withdraw(self, amount):
        if amount <= 0:
            raise ValueError("Withdraw amount must be positive")
        if amount > self.balance:
            raise InsufficientFundsError(
                f"{self.owner} has only {self.balance:.2f}, cannot withdraw {amount:.2f}"
            )
        self.balance -= amount
        self.history.append(("withdraw", amount))

    def transfer(self, other, amount):
        self.withdraw(amount)
        other.deposit(amount)

    def __str__(self):
        return f"{self.owner}: {self.balance:.2f} KM"


class SavingsAccount(BankAccount):
    def __init__(self, owner, balance=0.0, interest_rate=0.02):
        super().__init__(owner, balance)
        self.interest_rate = interest_rate

    def add_interest(self):
        interest = self.balance * self.interest_rate
        self.deposit(interest)
        return interest


if __name__ == "__main__":
    a = BankAccount("Amir", 100)
    b = SavingsAccount("Lejla", 500, interest_rate=0.05)

    a.deposit(50)
    a.transfer(b, 30)
    print(a)
    print(b)

    print("Interest added:", round(b.add_interest(), 2))
    print(b)

    try:
        a.withdraw(1000)
    except InsufficientFundsError as e:
        print("Error:", e)

    print("History of", a.owner, a.history)
