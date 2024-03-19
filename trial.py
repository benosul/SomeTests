testDict = {"layer1" : {}}

for key1 in testDict:
    if not "layer2" in testDict[key1]:
        testDict[key1].update({"layer2" : {}})
    for key2 in testDict[key1]:
        if not "layer3" in testDict[key1][key2]:
            testDict[key1][key2].update({ "layer3" : ["layer4"]})
        for key3 in testDict[key1][key2]:
            print(testDict[key1][key2][key3])
