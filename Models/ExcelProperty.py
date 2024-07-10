from decimal import Decimal

class ExcelProperty:
    def __init__(self):
        self.date = None  # Assuming this will be set to a datetime object later
        self.credit_amount = Decimal('0.000000')
        self.debit_amount = Decimal('0.000000')
        self.remarks = ""
        

    # Getters
    def getDate(self):
        return self.date

    def getCreditAmount(self):
        return self.credit_amount

    def getDebitAmount(self):
        return self.debit_amount

    def getRemarks(self):
        return self.remarks

    # Setters
    def setDate(self, date):
        # Assuming Date will be set as a datetime object
        self.date = date

    def setCreditAmount(self, credit_amount):
        # Assuming credit_amount will be set as a Decimal object
        self.credit_amount = Decimal(str(credit_amount))

    def setDebitAmount(self, debit_amount):
        # Assuming debit_amount will be set as a Decimal object
        self.debit_amount = Decimal(str(debit_amount))

    def setRemarks(self, remarks):
        self.remarks = remarks
