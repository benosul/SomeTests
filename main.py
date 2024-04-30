# A file to start with the code scanning

# We require a function to load in the .bas,.cls, and .frm files

# Then we have maybe a class that does the checking provided a file?
# We could have a class the logs the errors and creates two logs of must have violations and should have violations?

import src.Finder
import src.Scanner
import src.Reporter
import sys
import os

if __name__=='__main__':

  with open("testFile1.txt",'w'):
    pass
  
  dirPath     = os.environ.get('GITHUB_WORKSPACE')
  loader      = Finder.DataFinder(dirPath)
  scanner     = Scanner.DataScanner(loader)
  reporter    = Reporter.DataReporter(scanner.getViolations(),scanner.getRuleSeverity(),scanner.getRuleName())
  
  print(loader.getSourceCodeFiles())
  print(loader.getRulesFiles())
  print(scanner.getRules())
  print(reporter.getViolations())
  print(reporter.getMode())
  print(reporter.getDirPath())
  print(reporter.getReportPath())
  
  reporter.generateReport()
  reporter.generateReportLogging()
