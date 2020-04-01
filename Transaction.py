import matplotlib.dates as dt
import re


class Transaction:
    def __init__(self, day, month, year, addtl_info, amount):
        self.day = int(day)
        self.month = int(month)
        self.year = int(year)
        self.addtl_info = addtl_info
        self.amount = float(amount)
        self.date_authorized = ""  # get from addtl_info
        self.description = ""  # get from addtl_info

        #  create datenum object for each transaction
        self.date_string = str(month) + " " + str(day) + " " + str(year)
        self.datenum = dt.datestr2num(self.date_string)

        # date authorized ... all this below is specific regex to the sheet my bank uses
        try:
            x = re.search("\\d{2}/\\d{2}", self.addtl_info)
            self.date_authorized = x.group()
        except:
            self.date_authorized = "NA"

        # description
        try:
            if addtl_info[0:8] == "PURCHASE":
                cut_info = self.addtl_info[29:]
                x = re.search(".*[SP]\\d{10}", cut_info).group()
                length = len(x)
                self.description = x[-length: -12]
            else:
                self.description = addtl_info
        except:
            self.description = addtl_info

    def get_date(self):
        return str(self.month) + "-" + str(self.day) + "-" + str(self.year)

    def withdraw(self):
        # true if withdraw, false if deposit
        return self.amount < 0



