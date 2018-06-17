
'
' Dictionary Functions
'

Function createDict
    Dim objDict
    Set objDict = CreateObject("Scripting.Dictionary")
    objDict.CompareMode = TextMode
	set createDict = objDict
End Function

Function getDict(objDict, strKey) 
    If isObject(objDict) Then
        If objDict.Exists(strKey) Then
            getDict = objDict.Item(strKey)
        Else
            getDict = ""
        End If
    End If
End Function
Function add2Dict(objDict, strKey, strValue)
    If isObject(objDict) Then
	    If objDict.Exists(strKey) Then
		    add2Dict = False
		Else
	        objDict.Add strKey, strValue
		    add2Dict = True
		End If
	Else
        add2Dict = False
    End If		
End Function

Function updateDict(objDict, strKey, strValue)
    If isObject(objDict) Then
	    If objDict.Exists(strKey) Then
		    objDict.Item(strKey) = strValue
		    updateDict = True
		Else
	        objDict.Add strKey, strValue
		    updateDict = True
		End If
	Else
        updateDict = False
    End If		
End Function

Function mergeDict(objDict, strMessage)
    Dim strKey, strValue, strNewMsg
    strNewMsg = strMessage
    For Each strKey in objDict
        strValue = objDict(strKey)
        strKey = "##" & strKey & "##"
        strNewMsg = Replace(strNewMsg, strKey, strValue)
    Next    
    mergeDict = strNewMsg
End Function

Sub showDictObj(objDict)
    Dim strKey
    For Each strKey in objDict
        Wscript.Echo strKey & "==>" & objDict(strKey)
    Next
End Sub

Sub logDictObj(objLog, objDict)
    Dim strKey, bRes
    For Each strKey in objDict
        If Instr(LCase(strKey), "password") > 0 Then
            writeLog objLog, "  " & strKey & " ==> " & "****"
        Else
            writeLog objLog, "  " & strKey & " ==> " & objDict(strKey)
        End If
    Next
End Sub ' logDictObj

'
' Properties Functions
'
Function getAllProperties(objDict, strFileName)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")

    

    If fso.FileExists(strFileName) Then
        Set f = fso.OpenTextFile(strFileName, ForReading)
    Else
        getAllProperties = False
        Exit Function
    End If

    Do Until f.AtEndOfStream
        Dim i, strKey, strValue, myLine
        Dim arrValues
        Do
            myLine = f.ReadLine
            If Left(myLine, 1) = "#" Then
                Exit Do
            End If
            arrValues = Split(myLine, "=")
            strKey = arrValues(0)
            strValue = arrValues(1)
            For i = 2 to Ubound(arrValues)
                strValue = strValue & "=" & arrValues(i) 
            Next
            strValue = Replace(strValue, """", "") ' Ta bort eventuelle " 
            If ( objDict.Exists(strKey)) Then
                objDict.Item(strKey) = strValue
            Else
                objDict.Add strKey, strValue
            End IF
        Loop While False
    Loop
    f.Close
    getAllProperties = True
End Function ' getAllProperties

'
' FileSystem Functions
'
Function getFiles(folder, fileType)
    Dim objFolder, objFSO, objFile
    Dim colFiles
    Dim fileNames

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(folder)
    Set colFiles = objFolder.Files

    fileNames = ""
    For Each objFile in colFiles
        If Right(objFile.Name, Len(fileType)) = fileType Then
            fileNames = fileNames & objFile.Name & ";"
        End If
    Next

    If Right(fileNames, 1) = ";" Then
        fileNames = Left(fileNames, Len(fileNames) - 1)
    End If
    getFiles = Split(fileNames, ";")
End Function

Function moveFile(fromFilePath, toFilePath) 
    Dim objFS
    Set objFS = CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(fromFilePath) Then
        If objFS.FileExists(toFilePath) Then
            moveFile = False
        Else
            objFS.MoveFile fromFilePath, toFilePath
            moveFile = True
        End If
    Else
        moveFile = False
    End If
End Function

Function fileExists(filePath) 
    Dim objFS
    Set objFS = CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(filePath) Then
        fileExists = True
    Else
        fileExists = False
    End If
End Function

Function deleteFile(filePath)
    Dim objFS, bRes
    Set objFS = CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(filePath) Then
        objFS.DeleteFile filePath, True
        bRes = True
    Else
        bRes = False
    End If
    deleteFile = bRes
End Function

Function readAll(strFileName)
    Dim fso, f    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(strFileName, ForReading)
    readAll = f.ReadAll
    f.Close
End Function

Function writeNewFile(filePath, strData)
    Dim objFSO, objF
    Dim bResult

    Set objFSO = CreateObject("Scripting.FileSystemObject")    
    If objFSO.FileExists(filePath) Then
        bResult = False
    Else
        Set objF = objFSO.CreateTextFile(filePath, ForWriting)
        objF.Write(strData)
        objF.Close
        bResult = True
    End If
    writeNewFile = bResult
End Function ' writeNewFile

Function openLogFile(logFilePath)
    Dim objFS, objF
    Set objFS = CreateObject("Scripting.FileSystemObject")
    If objFS.FileExists(logFilePath) Then 
        Set objF = objFS.OpenTextFile(logFilePath, ForAppending)
    Else
        Set objF = objFS.CreateTextFile(logFilePath, ForWriting)
    End If
    Set openLogFile = objF
End Function

Sub writeLog(objLog, strLine)
    Dim strLogLine
    strLogLine = getHourMinuteSecondISO() & " " & strLine & vbCrLf
    If isObject(objLog) Then
        objLog.write(strLogLine)
    End If
End Sub

'
' Date and Time Functions
'
Function getYearMonthDay()
    getYearMonthDay = DatePart("yyyy", Now) &  _ 
                      Right("0" & DatePart("m", Now), 2) & _
                      Right("0" & DatePart("d", Now), 2)
End Function

Function getHourMinuteSecond()
    getHourMinuteSecond = Right("0" & DatePart("h", Now), 2) & _
                      Right("0" & DatePart("n", Now), 2) &  _
                      Right("0" & DatePart("s", Now), 2)
End Function

Function getYearMonthDayISO()
    getYearMonthDayISO = DatePart("yyyy", Now) & "-" &  _ 
                      Right("0" & DatePart("m", Now), 2) & "-" & _
                      Right("0" & DatePart("d", Now), 2)
End Function

Function getHourMinuteSecondISO()
    getHourMinuteSecondISO = Right("0" & DatePart("h", Now), 2) & ":" & _
                      Right("0" & DatePart("n", Now), 2) & ":" & _
                      Right("0" & DatePart("s", Now), 2)
End Function

Function convertToISO(ddmmyyyy)
    ' Input format: 23.04.2018
    ' Output format: 2018-04-23
    convertToISO = Right(ddmmyyyy, 4) & "-" & _
                   Mid(ddmmyyyy, 4, 2) & "-" & _
                   Left(ddmmyyyy, 2)
End Function

'
' Diverse hjelpefunksjoner
'
Sub responseDict(nodeList, oDic)
    Dim Nodes 
    Dim xNode 

    Set Nodes = nodeList

    For Each xNode In Nodes
        If xNode.nodeType = NODE_ELEMENT Then
            If xNode.hasChildNodes = True Then
                updateDict oDic, xNode.nodeName, xNode.text
            End If
       End If
       ' Traverse to next Child Node
       If xNode.hasChildNodes Then
            responseDict xNode.childNodes, oDic
       End If
    Next ' xNode
End Sub

Sub displayNode(nodeList, ind)
    Dim Nodes
    Dim xNode
    Set Nodes = nodeList
    'Wscript.Echo "displayNode Entry" & ind
    For Each xNode In Nodes
        'Wscript.Echo "  nodeType == " & xNode.nodeName
        If xNode.nodeType = NODE_ELEMENT Then
            'If xNode.hasChildNodes = True Then
                ' Element Node with no child
                WScript.Echo "Node type: " & xNode.nodeType & " Name=" _ 
                    & xNode.nodeName  & ", Value(" & Len(xNode.text) & ") " & xNode.text
            'End If
        End If
        ' Traverse to next Child Node
        If xNode.hasChildNodes Then
          'WScript.Echo "Indent=" & Indent
          displayNode xNode.childNodes, ind
        End If
    Next ' xNode
    'Wscript.Echo "displayNode Exit" & ind
End Sub

'
' Helper Classes
'
Class CSequenceNumber
    Private intSequence
    Private strSequenceFileName
    Private strDatePart
    Private strDCID
    Private intSeqPartLen
    Private strLatestSequence

    Private Sub Class_Initialize 
        Dim strLine, strElement
        strSequenceFileName = "sequenceNumber.txt"
        strDatePart = getYearMonthDay
        ' Hent inn sist brukte sekvensnummer i dag.
        If fileExists(strSequenceFileName) Then
            strLine = readAll(strSequenceFileName)
            strElement = Split(strLine, ";")
            If strDatePart <> strElement(0) Then
                intSequence = 0
            Else
                intSequence = strElement(1)
            End If
        Else
            intSequence = 0
        End If

    End Sub

    Private Sub Class_Terminate
       deleteFile strSequenceFileName
       writeNewFile strSequenceFileName, strDatePart & ";" & intSequence
    End Sub

    Public Function setDCID(dcid)
        strDCID = dcid
        intSeqPartLen = 8 - Len(strDCID)
        setDCID = True
    End Function

    Public Function getNextSequence
        intSequence = intSequence + 1
        getNextSequence = strDatePart & strDCID & _
            Right("00000" & intSequence, intSeqPartLen)
    End Function
End Class