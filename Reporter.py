#...

# generiere 2 Reports UserReport / ReviewerReport oder Gesamtsicht / Usersicht 

class DataReporter():

    def __init__(self,violations:dict) -> None:
        self.violations = violations

    def generateReport(self,file:str):
        pass