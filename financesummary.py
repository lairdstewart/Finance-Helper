import openpyxl
from Transaction import Transaction
import matplotlib.pyplot as plt


class FinanceSummary:
    Transactions = []  # list storing each transaction

    def __init__(self, filename):
        self.n = 0  # number of entries

        self.filename = str(filename)  # get filename
        workbook = openpyxl.load_workbook(filename=self.filename)  # loads in an excel file
        sheet = workbook['Sheet1']  # assigns the first sheet to an object
        # Import data to Transaction classes, then combine into list
        for row in sheet.iter_rows(min_row=0, values_only=True):  # iterates through rows
            # first column: dates
            date = str(row[0])
            year = date[0:4]
            month = date[5:7]
            day = date[8:10]

            # second column: amounts
            amount = float(row[1])

            # third column: descriptions
            description = row[2]

            x = Transaction(day, month, year, description, amount)  # create transaction object
            self.Transactions.append(x)  # add new object to list

            self.n = self.n + 1  # update total number of entries

    def print_sheet(self, month=0):
        if month == 0:
            i = 0
            for x in self.Transactions:  # loop through all transactions
                date = self.Transactions[i].get_date()
                amount = str(self.Transactions[i].amount)
                description = str(self.Transactions[i].description)
                print("date: " + date + " amount: " + str(amount) + " description: " + description)
                i = i + 1
        else:
            print("report for month " + str(month) + ":")
            i = 0
            for x in self.Transactions:  # loop through all transactions
                month_i = self.Transactions[i].month
                if month_i == month:
                    date = self.Transactions[i].get_date()
                    amount = str(self.Transactions[i].amount)
                    description = str(self.Transactions[i].description)
                    print("date: " + date + " amount: " + str(amount) + " description: " + description)
                i = i + 1

    def net_balance(self, month=0):  # called with no arguments gives total net balance
        if month == 0:
            net_balance = 0
            i = 0
            for x in self.Transactions:  # loop through all transactions
                net_balance = net_balance + self.Transactions[i].amount  # update net balance
                i = i + 1
            print("net balance is: " + str(net_balance))
        else:
            net_balance = 0
            i = 0
            entry = False  # keep track if a month has entries
            for x in self.Transactions:  # loop through all transactions
                month_i = self.Transactions[i].month
                if month_i == month:  # check if months line up
                    entry = True
                    net_balance = net_balance + self.Transactions[i].amount  # update net balance
                i = i + 1
            if entry:
                print("net balance for month " + str(month) + " is: " + str(net_balance))
            else:
                print("no balance available for month: " + str(month))

    def average_withdraw(self, month=0):
        if month == 0:  # no argument given, calculate total average withdrawals
            total_withdraw = 0  # store total withdraws
            num_withdraws = 0  # store number of withdraws
            i = 0
            for x in self.Transactions:
                if self.Transactions[i].withdraw():  # check if it is a withdraw
                    total_withdraw = total_withdraw + self.Transactions[i].amount
                    num_withdraws = num_withdraws + 1
                i = i + 1
            if num_withdraws != 0:  # avoid divide by 0 error
                avg_withdraw = total_withdraw / num_withdraws  # compute average
                print("Average Withdrawal: " + str(avg_withdraw))
            else:
                print("No withdrawals")

        else:  # month given, calculate average withdrawals for that month
            total_withdraw = 0
            num_withdraws = 0
            i = 0
            for x in self.Transactions:
                month_i = self.Transactions[i].month
                if self.Transactions[i].withdraw() and month_i == month:  # also check month match
                    total_withdraw = total_withdraw + self.Transactions[i].amount
                    num_withdraws = num_withdraws + 1
                i = i + 1
            if num_withdraws != 0:
                avg_withdraw = total_withdraw / num_withdraws
                print("Average Withdrawal in month " + str(month_i) + ": " + str(avg_withdraw))
            else:
                print("No withdrawals in month " + str(month))

    def average_deposit(self, month=0):
        if month == 0:  # no argument given, calculate total average withdrawals
            total_deposit = 0  # store total withdraws
            num_deposits = 0  # store number of withdraws
            i = 0
            for x in self.Transactions:
                if not self.Transactions[i].withdraw():  # check if it is a deposit
                    total_deposit = total_deposit + self.Transactions[i].amount
                    num_deposits = num_deposits + 1
                i = i + 1
            if num_deposits != 0:  # avoid divide by 0 error
                avg_deposit = total_deposit / num_deposits  # compute average
                print("Average deposit: " + str(avg_deposit))
            else:
                print("No deposits")

        else:  # month given, calculate average withdrawals for that month
            total_deposit = 0
            num_deposits = 0
            i = 0
            for x in self.Transactions:
                month_i = self.Transactions[i].month
                if not self.Transactions[i].withdraw() and month_i == month:  # also check month match
                    total_deposit = total_deposit + self.Transactions[i].amount
                    num_deposits = num_deposits + 1
                i = i + 1
            if num_deposits != 0:
                avg_deposit = total_deposit / num_deposits
                print("Average deposit in month " + str(month_i) + ": " + str(avg_deposit))
            else:
                print("No deposits in month " + str(month))

    def scatter(self, month=0):
        if month == 0:
            i = 0
            days = []  # list to keep track of days and amounts
            amounts = []
            for x in self.Transactions:
                days.append(self.Transactions[i].datenum)  # append those lists
                amounts.append(self.Transactions[i].amount)
                i = i + 1
            plt.plot_date(days, amounts)  # print results
            plt.xlabel("day")
            plt.ylabel("withdrawal/deposit amount")
            plt.show()
        else:
            i = 0
            days = []
            amounts = []
            for x in self.Transactions:
                month_i = self.Transactions[i].month
                if month_i == month:  # check month
                    days.append(self.Transactions[i].datenum)
                    amounts.append(self.Transactions[i].amount)
                i = i + 1
            plt.plot_date(days, amounts)
            plt.xlabel("day")
            plt.ylabel("withdrawal/deposit amount")
            plt.show()

    def histogram(self, month=0):
        if month == 0:
            i = 0
            amounts = []  # list to keep track of amounts
            for x in self.Transactions:
                amounts.append(self.Transactions[i].amount)
                i = i + 1
            bins = 100  # for a usual range of expenses ~[-1000, 1000] this is reasonable
            plt.hist(amounts, bins)
            plt.xlabel("withdrawal/deposit amount")
            plt.ylabel("occurrences")
            plt.show()
        else:
            i = 0
            amounts = []  # list to keep track of amounts
            for x in self.Transactions:
                month_i = self.Transactions[i].month
                if month_i == month:
                    amounts.append(self.Transactions[i].amount)
                i = i + 1
            bins = 100
            plt.xlabel("withdrawal/deposit amount")
            plt.ylabel("occurrences")
            plt.hist(amounts, bins)
            plt.show()


summary = FinanceSummary("debit.xlsx")
# summary.print_sheet()
# summary.net_balance(2)
# summary.average_withdraw(3)
# summary.average_deposit(2)
# summary.scatter(1)
summary.histogram()
