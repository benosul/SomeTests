# A file to start with the code scanning

# We require a function to load in the .bas,.cls, and .frm files

# Then we have maybe a class that does the checking provided a file?
# We could have a class the logs the errors and creates two logs of must have violations and should have violations?

import Finder
import Scanner
import Reporter
import sys

dirPath     = sys.argv[1]
mode        = sys.argv[2]
print("path: " + dirPath)
print("mode: " + mode)
loader      = Finder.DataFinder(dirPath)
scanner     = Scanner.DataScanner(loader,mode)
reporter    = Reporter.DataReporter(scanner.getViolations(),scanner.getRuleSeverity(),scanner.getRuleName(),mode,dirPath)

# print(loader.getSourceCodeFiles())
# print(reporter.getViolations())
# print(reporter.getMode())
# print(reporter.getDirPath())
# print(reporter.getReportPath())

reporter.generateReport()
reporter.generateReportLogging()