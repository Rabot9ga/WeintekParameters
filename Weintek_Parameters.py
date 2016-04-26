screenName = 'Sensor'
fileName = 'Var.csv'

structsList = []
varList = []
enList = []
appStructs = 0
appVariables = 0
for line in open(fileName): # Использовать итератор файла
    print(line, end='')
    if line == 'Structs;\n':
        appStructs = 1
        continue

    if (line == 'Variables;Control\n'):
        appStructs = 0
        appVariables = 1
        continue

    if line == ';\n':
        continue

    if appStructs:
        structsList.append(line[:-2])

    if appVariables:
        varList.append(line[:-3])
        enList.append(line[-2:-1])



print (structsList)
print (varList)
print(enList)

startScript = """
macro_command main()

short LcParameterSetNumber, TempShort, ConstZero, TempOld
bool ConstTrue, ConstFalse, TempBool

ConstTrue  = true
ConstFalse = false
ConstZero  = 0

GetData(LcParameterSetNumber, "Local HMI", "Parameter_Set_Number", 1)

"""

outputFile = open('Output.txt','w')

print(startScript)
outputFile.write(startScript+'\n')

currStructName = 0
endStructs = len(structsList)
for j in structsList:
    currNumVar = 0
    if currStructName == 0:
        str1 = 'if LcParameterSetNumber == ' + str(currStructName + 1) +' then'
        print(str1)
        outputFile.write(str1+'\n')
    else:
        str1 = 'else if LcParameterSetNumber == ' + str(currStructName + 1) +' then'
        print(str1)
        outputFile.write(str1+'\n')
    for i in enList:
        if i == '1':
            str1 = '    GetData(TempOld, "Local HMI", "'+ screenName + '_Old_' + varList[currNumVar] + '", 1)\n'
            str2 = '    GetData(TempShort, "Local HMI", "'+ screenName + '_' + varList[currNumVar] + '", 1)\n'
            str3 = '    if 	(TempOld <> TempShort) then\n'
            str4 = '        SetData(TempShort, "Local HMI", "'+ structsList[currStructName] + '.' + varList[currNumVar] + '", 1)\n'
            str5 = '    else\n'
            str6 = '        GetData(TempShort, "Local HMI", "'+ structsList[currStructName] + '.' + varList[currNumVar] + '", 1)\n'
            str7 = '        SetData(TempShort, "Local HMI", "'+ screenName + '_' + varList[currNumVar] + '", 1)\n'
            str8 = '    GetData(TempShort, "Local HMI", "'+ screenName + '_' + varList[currNumVar] + '", 1)\n'
            str9 = '    SetData(TempShort, "Local HMI", "'+ screenName + '_Old_' + varList[currNumVar] + '", 1)\n'
            print(str1+str2+str3+str4+str5+str6+str7+str8+str9)
            outputFile.write(str1+str2+str3+str4+str5+str6+str7+str8+str9+'\n')
        else:
            str1 = '    GetData(TempShort, "Local HMI", "'+ structsList[currStructName] + '.' + varList[currNumVar] + '", 1)\n'
            str2 = '    SetData(TempShort, "Local HMI", "'+ screenName + '_' + varList[currNumVar] + '", 1)\n'
            print(str1+str2)
            outputFile.write(str1+str2+'\n')
        currNumVar += 1
    currStructName += 1
    if currStructName == endStructs:
        str1 = 'else\n'
        print(str1)
        outputFile.write(str1+'\n')

currStructName = 0
for j in structsList:
    currNumVar = 0
    for i in varList:
        str1 = '    SetData(ConstZero, "Local HMI", "'+ structsList[currStructName] + '.' + varList[currNumVar] + '", 1)'
        print(str1)
        outputFile.write(str1+'\n')
        currNumVar += 1
    currStructName += 1

str1 = 'end if\n'
str2 = 'end macro_command'
print('\n'+str1+'\n'+str2+'\n')
outputFile.write('\n'+str1+'\n'+str2+'\n')

outputFile.close()

input('Press any key')



