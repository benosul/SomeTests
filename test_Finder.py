import Finder

def test_FindingDirectories(foundDirectories):
    expectedDirectories = [".\\CodeScanning\\testFilesAndDirs\\Sourcecode_vba",".\\CodeScanning\\testFilesAndDirs\\workbook_vba"]
    assert set(foundDirectories) == set(expectedDirectories)

def test_FindingFiles(foundFiles):    
    expectedFiles   = ['.\\CodeScanning\\testFilesAndDirs\\Sourcecode_vba\\' + file for file in ['badFunctionDeclaration.bas',
                                                                                    'cleanModule.bas',
                                                                                    'emptyModule.bas',
                                                                                    'noOptionExplicit.frm',
                                                                                    'includes_tsda.bas']] + ['.\\CodeScanning\\testFilesAndDirs\\workbook_vba\\' + file for file in ["Module1.bas"]]
    assert set(foundFiles) == set(expectedFiles)

def test_FindingRulesFiles(foundFiles):
    expectedFiles   = ['.\\CodeScanning\\testFilesAndDirs\\Rules\\' + file + ".txt" for file in ['Rules_Have',"Rules_Avoid"]]
    assert set(foundFiles) == set(expectedFiles)

if __name__=='__main__':
    loader = Finder.DataFinder(".\\CodeScanning\\")      
    foundDirectories    = loader.getVBADirectories()
    foundFiles      = loader.getSourceCodeFiles()
    foundRulesFiles      = loader.getRulesFiles()

    test_FindingDirectories(foundDirectories)
    test_FindingFiles(foundFiles)
    test_FindingRulesFiles(foundRulesFiles)