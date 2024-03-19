import Reporter
import Scanner
import Finder


if __name__ == "__main__":
    dirPath     = "."
    mode        = "User"
    loader      = Finder.DataFinder(dirPath)
    scanner     = Scanner.DataScanner(loader,mode)
    reporter    = Reporter.DataReporter(scanner.getViolations(),scanner.getRuleSeverity(),scanner.getRuleName(),mode,dirPath)
    
    # print(reporter.getViolations())
    # print(reporter.getMode())
    # print(reporter.getDirPath())
    # print(reporter.getReportPath())

    reporter.generateReport()
    reporter.generateReportLogging()