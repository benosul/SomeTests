import src.Finder as Finder
import src.Scanner as Scanner

def test_ParseRules(rules:dict):
    expectedResult  = {"have":['Option\\sExplicit'],"avoid":["tsda_"]}    
    for key in expectedResult:
        assert set(rules[key]) == set(expectedResult[key])
    for key in rules:
        assert set(rules[key]) == set(expectedResult[key])

def test_getSeverity(severity:dict):
    expected = {'tsda_': ['High'], 'Option\\sExplicit': ['Medium']}
    for key in expected:
        assert set(severity[key]) == set(expected[key])
    for key in severity:
        assert set(severity[key]) == set(expected[key])

def test_FindViolations(violations,mode):
    if mode == "User":
        expected        = {'.\\CodeScanning\\testFilesAndDirs\\Sourcecode_vba\\emptyModule.bas': {'Medium': {'Option\\sExplicit': [-1]}},
                           '.\\CodeScanning\\testFilesAndDirs\\Sourcecode_vba\\noOptionExplicit.frm': {'Medium': {'Option\\sExplicit': [-1]}},
                           '.\\CodeScanning\\testFilesAndDirs\\Sourcecode_vba\\includes_tsda.bas': {'High': {'tsda_': [2]}}}
    elif mode == "Reviewer":
        expected        = {'Medium': {'Option\\sExplicit': {'.\\CodeScanning\\testFilesAndDirs\\Sourcecode_vba\\emptyModule.bas': [-1], 
                                                       '.\\CodeScanning\\testFilesAndDirs\\Sourcecode_vba\\noOptionExplicit.frm': [-1]}}, 
                           'High': {'tsda_': {'.\\CodeScanning\\testFilesAndDirs\\Sourcecode_vba\\includes_tsda.bas': [2]}}}
    for file in expected:
        for severity in expected[file]:
            for rule in expected[file][severity]:
                assert set(expected[file][severity][rule]) == set(violations[file][severity][rule])
    for file in violations:
        for severity in expected[file]:
            for rule in expected[file][severity]:
                assert set(expected[file][severity][rule]) == set(violations[file][severity][rule])

if __name__ == "__main__":
    finder      = Finder.DataFinder(".\\CodeScanning")
    # mode       = "User"
    mode        = "Reviewer"
    scanner     = Scanner.DataScanner(finder,mode)
    rules       = scanner.getRules()
    
    
    for key in rules:
        print(key)
        print(rules[key])
    print()
    severity    = scanner.getRuleSeverity()
    for key in severity:
        print(key)
        print(severity[key])
    print()
    violations  = scanner.getViolations()
    for key in violations:
        print(key)
        print(violations[key])
    print()
    print(scanner.getRuleName())
    print()
    # test_ParseRules(rules)
    # test_getSeverity(severity)
    # test_FindViolations(violations,mode)
