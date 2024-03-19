import Finder

def test_FindingDirectories():
    loader = Finder.DataFinder(".")      
    foundDirectories    = loader.getVBADirectories()
    expectedDirectories = [".\\testFilesAndDirs\\Sourcecode_vba",".\\testFilesAndDirs\\workbook_vba"]
    assert set(foundDirectories) == set(expectedDirectories)

def test_FindingFiles():
    loader = Finder.DataFinder(".")
    foundFiles      = loader.getSourceCodeFiles()
    expectedFiles   = ['.\\testFilesAndDirs\\Sourcecode_vba\\' + file for file in ['badFunctionDeclaration.bas',
                                                                                    'cleanModule.bas',
                                                                                    'emptyModule.bas',
                                                                                    'noOptionExplicit.frm',
                                                                                    'includes_tsda.bas']] + ['.\\testFilesAndDirs\\workbook_vba\\' + file for file in ["Module1.bas"]]
    assert set(foundFiles) == set(expectedFiles)

def test_FindingRulesFiles():
    loader = Finder.DataFinder(".")
    foundFiles      = loader.getRulesFiles()
    expectedFiles   = ['.\\testFilesAndDirs\\Rules\\' + file + ".txt" for file in ['Rules_Have',"Rules_Avoid"]]
    assert set(foundFiles) == set(expectedFiles)

if __name__=='__main__':
    test_FindingDirectories()
    test_FindingFiles()
    test_FindingRulesFiles()