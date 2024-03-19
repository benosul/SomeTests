import Finder
import Scanner

def test_ParseRules(finder:Finder.DataFinder):
    scanner         = Scanner.DataScanner(finder)
    scanner.parseRules()
    rules           = scanner.getRules()
    expectedResult  = {"have":['Option Explicit'],"avoid":["tsda_"],"other":[]}
    for key in expectedResult:
        assert set(rules[key]) == set(expectedResult[key])
    for key in rules:
        assert set(rules[key]) == set(expectedResult[key])

def test_FindViolations(finder:Finder.DataFinder):
    scanner         = Scanner.DataScanner(finder)
    scanner.parseRules()
    violations      = scanner.findViolations()
    expected        = {'.\\testFilesAndDirs\\Sourcecode_vba\\emptyModule.bas': {'Option Explicit': [-1]}, 
                        '.\\testFilesAndDirs\\Sourcecode_vba\\noOptionExplicit.frm': {'Option Explicit': [-1]}, 
                        '.\\testFilesAndDirs\\Sourcecode_vba\\includes_tsda.bas': {'tsda_': [2]}}
    for file in expected:
        for rule in expected[file]:
            assert set(expected[file][rule]) == set(violations[file][rule])
    for file in violations:
        for rule in violations[file]:
            assert set(expected[file][rule]) == set(violations[file][rule])

if __name__ == "__main__":
    finder = Finder.DataFinder(".")

    test_ParseRules(finder)
    test_FindViolations(finder)
