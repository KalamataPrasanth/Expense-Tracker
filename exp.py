import pandas as pd
import os

class Expense:
    slno = 0

    def __init__(self, spent_on, amount, date, note):
        Expense.slno += 1
        self.serial = Expense.slno
        self.spent_on = spent_on
        self.amount = amount
        self.date = date
        self.note = note

    def __repr__(self):
        return f"{self.serial} - {self.spent_on} - {self.amount} - {self.date} - {self.note}"
    
class ExpenseManager:
    def __init__(self):
        self.expenses = []
        self.total_amount = 0.0
        self.load()
    
    def add_expense(self, expense):
        self.expenses.append(expense)
        self.total_amount += expense.amount
    
    def remove_expense(self, index):
        if 0 <= index < len(self.expenses):
            removed_expense = self.expenses.pop(index)
            self.total_amount -= removed_expense.amount
    
    def display_expenses(self):
        for expense in self.expenses:
            print(expense)
        print(f"Total Amount Spent: {self.total_amount}")

    def save_to_file(self, filename = 'expenses.xlsx'):
        data = [{
            'Serial': expense.serial,
            'Spent_on': expense.spent_on,
            'Amount': expense.amount,
            'Date': expense.date,
            'Note': expense.note
        } for expense in self.expenses]

        df = pd.DataFrame(data)
        df.loc[len(df)] = ['Total', '', self.total_amount, '', '']

        while True:
            try:
                df.to_excel(filename, index = False)
                print(f"Expenses saved to {filename}")
                break
            except PermissionError:
                input(f"Please close the file {filename} and press Enter to continue...")
    
    def load(self, filename = 'expenses.xlsx'):
        if os.path.exists(filename):
            df = pd.read_excel(filename)
            for _, row in df.iterrows():
                if row['Serial'] != 'Total':
                    expense = Expense(
                        spent_on = row['Spent_on'],
                        amount= row['Amount'],
                        date = row['Date'],
                        note = row['Note']
                    )
                    expense.serial = row['Serial']
                    Expense.slno = max(Expense.slno, row['Serial'])
                    self.add_expense(expense)
    
    def prompt_for_expenses(self):
        while True:
            spent_on = input("Spent on: ")
            amount = float(input("Enter the amount: "))
            date = input("Date (dd-mm-yyyy): ")
            note = input("Description: ")

            expense = Expense(spent_on, amount, date, note)
            self.add_expense(expense)

            more_expenses = input("Do you have more expenses to enter? (y/n): ")
            if more_expenses.lower() != 'y':
                break

if __name__ == '__main__':
    manager = ExpenseManager()
    manager.prompt_for_expenses()
    manager.display_expenses()
    manager.save_to_file()
