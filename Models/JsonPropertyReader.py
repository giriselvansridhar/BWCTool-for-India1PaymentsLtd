class JsonPropertyReader:
    def __init__(self):
        self.FileName = ""
        self.HeadingRowNo = 0
        self.Date = None  # Assuming this will be set to a datetime object later
        self.DebitAmount = ""
        self.CreditAmount = ""
        self.Remarks = ""
        self.CrDrSeparator = False
        self.CrDrSeparatorColName = ""

    # Getters
    def getFileName(self):
        return self.FileName

    def getHeadingRowNo(self):
        return self.HeadingRowNo

    def getDate(self):
        return self.Date

    def getDebitAmount(self):
        return self.DebitAmount

    def getCreditAmount(self):
        return self.CreditAmount

    def getRemarks(self):
        return self.Remarks

    def getCrDrSeparator(self):
        return self.CrDrSeparator

    def getCrDrSeparatorColName(self):
        return self.CrDrSeparatorColName

    # Setters
    def setFileName(self, FileName):
        self.FileName = FileName

    def setHeadingRowNo(self, HeadingRowNo):
        self.HeadingRowNo = HeadingRowNo

    def setDate(self, Date):
        # Assuming Date will be set as a datetime object
        self.Date = Date

    def setDebitAmount(self, DebitAmount):
        self.DebitAmount = DebitAmount

    def setCreditAmount(self, CreditAmount):
        self.CreditAmount = CreditAmount

    def setRemarks(self, Remarks):
        self.Remarks = Remarks

    def setCrDrSeparator(self, CrDrSeparator):
        self.CrDrSeparator = CrDrSeparator

    def setCrDrSeparatorColName(self, CrDrSeparatorColName):
        self.CrDrSeparatorColName = CrDrSeparatorColName
