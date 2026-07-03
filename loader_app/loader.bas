Attribute VB_Name = "loaderMod"
'-----------------GLOBAL CONSTANTS-----------------------
Const c_date = 0
Const c_time = 1
Const c_partId = 2
Const c_partName = 3
Const c_partNumber = 4
Const c_featureId = 5
Const c_value = 6
Const c_toCep = 7
Const c_deviceId = 8

'-----------------GLOBAL VARS-----------------------
Public rootPath
Public databaseFolderPath
Public inputFolderPath
Public processedFolderPath
Public databaseFilePath
Public indexesFilePath
Public postReviewsFilePath
Public reporterFolderPath
Public backupFolderPath
Public currentFileName
Public currentDate
Public currentTime
Public currentPartId
Public currentPartName
Public currentPartNum
Public currentToCep

'-----------------GLOBAL OBJECTS-----------------------
Public fso As FileSystemObject
Public inputFiles

'-----------------MAIN-----------------------
Sub Main()
    Dim inputFiles, inputFileData As String, index
    
    setGlobals
    validateGlobalVars
    
    regenerateDataIndexes databaseFilePath, indexesFilePath, 76
    Set inputFiles = getInputFiles
    
    If inputFiles.Count > 0 Then
        makeBackup
    End If
    
    For Each file In inputFiles
        If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
        
            currentFileName = file.Name
        
            inputFileData = getDatafromInputFile(file)
            
            index = getIndexfromDate(currentDate, indexesFilePath)
            
            regenerateDataIndexes databaseFilePath, indexesFilePath, index
            
            appendToDataBase inputFileData, index, currentDate, currentTime
            
            moveToProcessed file
            
            postReviews inputFileData
            
        End If
    Next
End Sub

'-----------------SUBRUTINES-----------------------
Sub setGlobals()
    ' global vars
    rootPath = App.path
    databaseFolderPath = getConfig("DATABASE_PATH")
    inputFolderPath = getConfig("INPUT_FILES_PATH")
    processedFolderPath = getConfig("PROCESSED_FILES_PATH")
    backupFolderPath = getConfig("BACKUPS_PATH")
    reporterFolderPath = getConfig("REPORTER_PATH")
    databaseFilePath = databaseFolderPath & "data_features.csv"
    postReviewsFilePath = databaseFolderPath & "post_reviews.csv"
    indexesFilePath = databaseFolderPath & "dates_index.csv"
End Sub

Sub validateGlobalVars()
    Dim errorMessage
    
    If Not pathExist(databaseFolderPath) Then
        errorMessage = "Database Folder:" & vbCrLf & """" & databaseFolderPath & """" & vbCrLf & " Not Found"
    ElseIf Not pathExist(inputFolderPath) Then
        errorMessage = "Input Files Folder:" & vbCrLf & """" & inputFolderPath & """" & vbCrLf & " Not Found"
    ElseIf Not pathExist(processedFolderPath) Then
        errorMessage = "Processed Files Folder:" & vbCrLf & """" & processedFolderPath & """" & vbCrLf & " Not Found"
    ElseIf Not pathExist(databaseFilePath) Then
        errorMessage = "Database File:" & vbCrLf & """" & databaseFilePath & """" & vbCrLf & " Not Found"
    ElseIf Not pathExist(postReviewsFilePath) Then
        errorMessage = "Post Reviews File:" & vbCrLf & """" & postReviewsFilePath & """" & vbCrLf & " Not Found"
    ElseIf Not pathExist(indexesFilePath) Then
        errorMessage = "Dates Index File:" & vbCrLf & """" & indexesFilePath & """" & vbCrLf & " Not Found"
    ElseIf Not pathExist(backupFolderPath) Then
        errorMessage = "Backups Folder:" & vbCrLf & """" & backupFolderPath & """" & vbCrLf & " Not Found"
    ElseIf Not pathExist(reporterFolderPath) Then
        errorMessage = "Reporter Folder:" & vbCrLf & """" & reporterFolderPath & """" & vbCrLf & " Not Found"
    Else
        Exit Sub
    End If
    
    MsgBox errorMessage, vbExclamation, "Path not Found"
    End ' exit execution for VBA
End Sub

Sub makeBackup()
    fso.CopyFolder Left(databaseFolderPath, Len(databaseFolderPath) - 1), backupFolderPath & Format(Now, "yyyymmdd")
End Sub

Function getDatafromInputFile(file)
    Dim inputFile, fileContent, fields, totalLines, dataStart, validationError As String
    ' open for reading
    Set inputFile = file.OpenAsTextStream(1)

    fileContent = Split(inputFile.ReadAll, vbCrLf)
    
    inputFile.Close
    
    totalLines = UBound(fileContent) + 1
    
    dataStartLine = 14
    
    currentPartId = Replace(Split(fileContent(10), ",")(3), """", "")
    currentPartName = Replace(Split(fileContent(11), ",")(3), """", "")
    currentPartNum = Split(fileContent(12), ",")(3)
    currentDate = Split(fileContent(8), ",")(3)
    currentTime = Split(fileContent(9), ",")(3)
    currentToCep = Split(fileContent(13), ",")(3)
    v_device_id = "DEA_TORO"
    
    ' dataStartLine is the first data item and totalLines - 3 is the last
    For i = dataStartLine To (totalLines - 3)
        ' ningun feature_id puede contener comillas dobles ni espacios
        fields = Split(fileContent(i), ",")
        v_feature_id = Replace(fields(1), """", "")
        v_feature_id = Replace(v_feature_id, " ", "_")

        ' this have to do it because "Dimensions" add another comma to the line
        If UBound(fields) > 23 Then
            v_value = fields(8)
        Else
            v_value = fields(7)
        End If
        
        ' validate fields
        validationError = validateFields(currentDate, currentTime, currentPartId, currentPartName, currentPartNum, v_feature_id, v_value, currentToCep, v_device_id)
        
        If validationError <> "" Then
            MsgBox currentFileName & vbCrLf & validationError, vbCritical, "Fields Validation Failed"
            End ' terminate program
        End If
        
        getDatafromInputFile = getDatafromInputFile & currentDate & ";" & currentTime & ";" & currentPartId & ";" & currentPartName & ";" & _
                                                      currentPartNum & ";" & v_feature_id & ";" & v_value & ";" & _
                                                      currentToCep & ";" & v_device_id & vbCrLf
        
    Next
                
End Function

Sub postReviews(inputFileData As String)
    Dim fso As FileSystemObject
    Dim dataArr() As String, dataLen As Long, featureId As String, value As Currency
    Dim nominal As Currency, CLie As Currency, CLse As Currency, ILie As Currency, ILse As Currency, onAccuracy As Boolean
    Dim reviewsFile As Object, spectsFile As Object, spectsArr() As String, i As Long, j As Long
    Dim totalOnAcc As Long, totalOnAccNG As Long, status As String, rejectedBy As String, rejecteds() As String
    Dim v_date As String, v_time As String, partId As String, partNumber As String, device As String, accuracy As Currency
    'date;time;part_id;part_number;device_id;accuracy;status;rejected_by;value

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    status = "Pass"
    
    ReDim rejecteds(0)
    
    reporterFolderPath = getConfig("REPORTER_PATH")
    databaseFolderPath = getConfig("DATABASE_PATH")
    
    ' open for reading
    Set spectsFile = fso.OpenTextFile(reporterFolderPath & "database\spec_features.csv", ForReading)
    Set reviewsFile = fso.OpenTextFile(databaseFolderPath & "post_reviews.csv", ForAppending, True)
    
    spectsFile.SkipLine
    spectsArr = Split(spectsFile.ReadAll, vbCrLf)
    
    spectsFile.Close
    
    dataArr = Split(inputFileData, vbCrLf)
    
    v_date = Split(dataArr(0), ";")(0)
    v_time = Split(dataArr(0), ";")(1)
    partId = Split(dataArr(0), ";")(2)
    partNumber = Split(dataArr(0), ";")(4)
    device = Split(dataArr(0), ";")(8)
    
    
    
    ' read every data row
    For Each dataRow In dataArr
        DoEvents
        If dataRow = "" Then GoTo Continue
        featureId = Split(dataRow, ";")(5)
        value = strToDec(Split(dataRow, ";")(6), 1)
        ' search the feature on spects
        For Each specRow In spectsArr
            If specRow = "" Then GoTo ContinueData
            If featureId = Split(specRow, ";")(0) Then
                nominal = strToDec(Split(specRow, ";")(2), 1)
                ' first calc status
                ILie = nominal + strToDec(Split(specRow, ";")(5), 1)
                ILse = nominal + strToDec(Split(specRow, ";")(6), 1)
                If (value < ILie) Or (value > ILse) Then
                    status = "Rejected"
                    ReDim Preserve rejecteds(0 To UBound(rejecteds) + 1)
                    rejecteds(UBound(rejecteds)) = featureId & ";" & CStr(value)
                End If
                ' add data to accuracy calc
                CLie = nominal + strToDec(Split(specRow, ";")(3), 1)
                CLse = nominal + strToDec(Split(specRow, ";")(4), 1)
                onAccuracy = CBool(Split(specRow, ";")(8))
                If onAccuracy Then
                    totalOnAcc = totalOnAcc + 1
                    If (value < CLie) Or (value > CLse) Then
                        totalOnAccNG = totalOnAccNG + 1
                    End If
                End If
                GoTo ContinueData
            End If
        Next
ContinueData:
    Next
Continue:

    accuracy = (totalOnAcc - totalOnAccNG) / totalOnAcc * 100
    
    If status = "Rejected" Then
        For i = 1 To UBound(rejecteds)
            reviewsFile.WriteLine v_date & ";" & v_time & ";" & partId & ";" & partNumber & ";" & device & ";" & Round(accuracy, 1) & ";" & status & ";" & rejecteds(i)
        Next i
    Else
        reviewsFile.WriteLine v_date & ";" & v_time & ";" & partId & ";" & partNumber & ";" & device & ";" & Round(accuracy, 1) & ";" & status
    End If
    
    reviewsFile.Close
    
End Sub

Sub moveToProcessed(file)
    Dim fileName As String
    
    fileName = Split(currentDate, "/")(2) & Split(currentDate, "/")(1) & Split(currentDate, "/")(0) & "_" & Replace(currentTime, ":", "") & _
        "_" & currentPartId & "_" & currentPartName & "_" & currentPartNum & ".csv"
    
    FileCopy file.path, processedFolderPath & "\" & fileName
    
    Kill file.path
End Sub



Sub appendToDataBase(newData, ByVal fromIndex, fromDate, fromTime)
    Dim sourceFile, tempFile, currentLine, tempFilePath, isEqualDates
    
    tempFilePath = databaseFolderPath & "temp.csv"
    
    If pathExist(tempFilePath) Then Kill tempFilePath
    
    ' open for reading
    sourceFile = FreeFile
    Open databaseFilePath For Input As #sourceFile
    
    ' open for append
    tempFile = FreeFile
    Open tempFilePath For Output As #tempFile
    
    ' search the date
    Seek #sourceFile, fromIndex
    
    isEqualDates = False
    
    Do While Not EOF(sourceFile)
        Line Input #sourceFile, currentLine
        If strToDate(Split(currentLine, ";")(c_date)) >= strToDate(fromDate) Then
            If strToDate(Split(currentLine, ";")(c_date)) = strToDate(fromDate) Then
                isEqualDates = True
            End If
            Exit Do
        End If
        fromIndex = Seek(sourceFile)
    Loop
    
    ' if dates are equal search for time
    If isEqualDates Then
        Seek #sourceFile, fromIndex
        Do While Not EOF(sourceFile)
            Line Input #sourceFile, currentLine
            If CDate(Split(currentLine, ";")(c_time)) >= CDate(fromTime) Then
                If CDate(Split(currentLine, ";")(c_time)) = CDate(fromTime) Then
                    ' check for duplicated data (same partId, partNum, date and time)
                    If (Split(currentLine, ";")(c_partId) = currentPartId) And (Split(currentLine, ";")(c_partNumber) = currentPartNum) Then
                        MsgBox "The input file """ & currentFileName & """ data is duplicated on database", vbExclamation, "Duplicated input file data"
                        Close #sourceFile ' Siempre cierra el archivo al terminar
                        Close #tempFile ' Siempre cierra el archivo al terminar
                        End
                    End If
                End If
                Exit Do
            End If
            ' date is different exit loop
            If CDate(Split(currentLine, ";")(c_date)) <> CDate(fromDate) Then
                Exit Do
            End If
            fromIndex = Seek(sourceFile)
        Loop
    End If
    
    Seek #sourceFile, 1
    begining = Input(fromIndex - 1, #sourceFile)
    Print #tempFile, begining;
    
    If Not EOF(sourceFile) Then
        Print #tempFile, newData;
        
        ending = Input(LOF(sourceFile) - Seek(sourceFile) + 1, #sourceFile)
        Print #tempFile, ending;
    Else
        Print #tempFile, newData;
    End If
    
    Close #sourceFile ' Siempre cierra el archivo al terminar
    Close #tempFile ' Siempre cierra el archivo al terminar
    
    FileCopy tempFilePath, databaseFilePath
    
    Kill tempFilePath

End Sub

Sub regenerateDataIndexes(databaseFilePath, indexFilePath, fromDbIndex)
    Dim sourceFile, indexFile, index
    Dim prevMonth, currentDate
    Dim currentLine, oldIndexes, newIndexes
        
    ' open for read
    indexFile = FreeFile
    Open indexFilePath For Input As #indexFile
    
    'get the headers
    Line Input #indexFile, currentLine
    oldIndexes = oldIndexes & currentLine & vbCrLf
    
    Do Until EOF(indexFile)
        Line Input #indexFile, currentLine
        
        If Split(currentLine, ";")(1) = CStr(fromDbIndex) Then
            Exit Do
        End If
        If EOF(indexFile) Then
            Close #indexFile
            Exit Sub
        End If
        oldIndexes = oldIndexes & currentLine & vbCrLf
    Loop
    
    Close #indexFile ' Siempre cierra el archivo al terminar
    
    ' open for write
    indexFile = FreeFile
    Open indexFilePath For Output As #indexFile
    
    Print #indexFile, oldIndexes;
    
    ' open for reading
    sourceFile = FreeFile
    Open databaseFilePath For Input As #sourceFile
    
    Seek #sourceFile, fromDbIndex
    
    Line Input #sourceFile, currentLine
    currentDate = CDate(Split(currentLine, ";")(0))
    Print #indexFile, currentDate & ";" & fromDbIndex
    
    prevDate = currentDate
    
    ' Bucle hasta llegar al final del archivo (EOF)
    Do Until EOF(sourceFile)
        index = Seek(sourceFile)
        Line Input #sourceFile, currentLine
        currentDate = CDate(Split(currentLine, ";")(0))
        
        If currentDate >= (prevDate + 30) Then
            Print #indexFile, currentDate & ";" & index
            prevDate = currentDate
        End If
    Loop

    Close #sourceFile ' Siempre cierra el archivo al terminar
    Close #indexFile ' Siempre cierra el archivo al terminar
    
End Sub



'-----------------FUNCTIONS-----------------------
Function getInputFiles()
    Set getInputFiles = fso.GetFolder(inputFolderPath).Files
End Function

Function getIndexfromDate(fromDate, indexesFilePath)
    Dim indexesFile, currentDate, currentLine

    ' open for reading
    indexesFile = FreeFile
    Open indexesFilePath For Input As #indexesFile
    
    ' skip the headers
    Line Input #indexesFile, currentLine
    
    Line Input #indexesFile, currentLine
    currentDate = Split(currentLine, ";")(0)
       
    index = Split(currentLine, ";")(1)
    
    Do Until EOF(indexesFile)
        Line Input #indexesFile, currentLine
        currentDate = Split(currentLine, ";")(0)
        If CDate(currentDate) >= CDate(fromDate) Then
            getIndexfromDate = index
            Close #indexesFile ' Siempre cierra el archivo al terminar
            Exit Function
        End If
        index = Split(currentLine, ";")(1)
    Loop

    Close #indexesFile ' Siempre cierra el archivo al terminar
    
    ' if fromDate is the highest return the last index
    getIndexfromDate = index
    
End Function

'----------------- HELPER SUBS AND FUNCTIONS-----------------------
Function validateFields(f_date, f_time, f_partId, f_partName, f_partNumber, _
                        f_featureId, f_value, f_toCep, f_deviceId) As String
                        
    If Not IsDate(f_date) Then
        validateFields = "The ""date"" field (" & f_date & ") is not valid."
        Exit Function
    ElseIf Not IsDate(f_time) Then
        validateFields = "The ""time"" field (" & f_time & ") is not valid."
        Exit Function
    ElseIf Not IsNumeric(f_partNumber) Then
        validateFields = "The ""part_number"" field (" & f_partNumber & ") is not valid."
        Exit Function
    ElseIf Not IsNumeric(f_value) Then
        validateFields = "The ""value"" field (" & f_value & ") is not valid."
        Exit Function
    ElseIf Not ((f_toCep = "True") Or (f_toCep = "False")) Then
        validateFields = "The ""to_cep"" field (" & f_toCep & ") is not valid."
        Exit Function
        ' ADD MISSING VALIDATIONS HERE
    End If
    
End Function


Function pathExist(path)
    pathExist = fso.FileExists(path) Or fso.FolderExists(path)
End Function

Function strToDate(dd_mm_yyyy_StringDate)
    Dim parts
    parts = Split(dd_mm_yyyy_StringDate, "/")
    
    strToDate = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
End Function

Function strToDec(value As Variant, decimals As Long) As Double
    Dim pointPos As Long
    ' replace comma
    value = Replace(value, ",", ".")
    pointPos = InStr(value, ".")
    If pointPos > 0 Then
        value = Left$(value, pointPos + decimals)
        strToDec = Val(value)
        Exit Function
    End If
    
    strToDec = Val(value)
End Function

Function getConfig(configKey As String) As String
    Dim confFile As Object, v_confPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    v_confPath = rootPath & "\init.conf"
    
    
    ' open the file for read
    Set confFile = fso.OpenTextFile(v_confPath, 1, True)
    ' read the file
    v_FileContent = confFile.ReadAll
    
    confFile.Close
    
    v_FileLines = Split(v_FileContent, vbCrLf)
    
    v_key = 0
    v_value = 1
    
    ' from zero because not have headers
    For lineNum = 0 To UBound(v_FileLines)
        v_Field = Split(v_FileLines(lineNum), "=")
        If UBound(v_Field) = -1 Then GoTo ContinueA
        ' Start fields edits
        
        If v_Field(v_key) = configKey Then
            getConfig = v_Field(v_value)
            Exit Function
        End If
        
        ' End fields edit
ContinueA:
    Next lineNum
    
    getConfig = ""

End Function

