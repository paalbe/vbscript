Option Explicit

Sub LesKurser
    Dim intFileNo, bResult
    Dim fso, f, myLine
    Dim inFileName, inTemplateFileName
    Dim priceFolder, priceFiles, priceLine
    Dim priceOutFolder, priceOutFiles, bKontroll
    Dim priceFileNamePrefix, logFileNamePrefix, logFilePath
    Dim arrKey, arrValue
    Dim objDict
    Dim objLog
    Dim strISOmsg, strMessage
    Dim tidspunkt

    tidspunkt = Now
    Wscript.Echo "LesKurser Startet " & tidspunkt

    ' Lag Dictionary objektet som skal inneholde
    ' - alle verdier fra properties.ini fila
    ' - prisopplysninger for det aktuelle ISIN

    set objDict = createDict

    ' Hent alle parametere som er oppgitt i properties.ini
    If getAllProperties(objDict, "properties.ini") Then
       Wscript.Echo "Properties.ini ble lest inn!"
       ' showDictObj(objDict)
    Else
        Wscript.Echo "properties.ini filen ble ikke lest inn"
        Wscript.Echo "LesKurser avslutter. Rett feilen og start igjen"
        Exit Sub
    End If

    ' Bygg opp unikt filnavnprefix
    priceFileNamePrefix = getDict(objDict, "IdInMsgId")  _ 
        & getYearMonthDay & getHourMinuteSecond
    logFileNamePrefix = "\Log" & getYearMonthDay
    priceOutFolder = getDict(objDict, "PriceOutFolder") & "\"

     ' Klargjør for logging, det er en loggfil fra hver dag.
    logFilePath = getDict(objDict, "LogFileFolder") & logFileNamePrefix & ".txt"
    Set objLog = openLogFile(logFilePath)

    writeLog objLog, "LesKurser starttidspunkt: " & tidspunkt
    writeLog objLog, "Poperties.ini ble lest inn:"
    ' List ut parameterene som ble lest inn fra properties.ini
    logDictObj objLog, objDict

    ' Det undersøkes om det ligger XML-filer i utkurven.
    ' Hvis det ligger filer der, avsluttes kjøringen.
    ' Dette kan hoppes over ved å endre på bKontroll
    ' bKontroll = False
    bKontroll = True
    bKontroll = False
    If bKontroll Then 
        priceOutFiles = getFiles(priceOutFolder, "xml")
        If Ubound(priceOutFiles) > -1 Then
            Wscript.Echo "Det ligger igjen filer i " & priceOutFolder 
            Wscript.Echo "Kontroller om alt har blitt sendt."
            Exit Sub
        End If
    End If

    ' Hent inn iso 20022 Price Report template
    inTemplateFileName = objDict.Item("PriceTemplateFile")
    strISOmsg = readAll(inTemplateFileName)
    writeLog objLog, "Mal for iso Price Report ble lest inn: " & inTemplateFileName

    ' Alle filene av type csv som ligger i
    ' folderen angitt i PriceDataFolder skal leses
    ' og for hvert ISIN lages det en iso 20022 Price Report
    ' XML fil, som senere skal sendes til Centevo
    intFileNo = 0
    priceFolder = objDict.Item("PriceDataFolder") & "\"
    priceFiles = getFiles(priceFolder, "csv")
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each inFileName in priceFiles
        writeLog objLog, "Kursfilen " & inFileName & " blir lest"
        Set f = fso.OpenTextFile(priceFolder & inFileName, ForReading)
        priceLine = f.ReadLine
        writeLog objLog, "    " & priceLine
        arrKey = Split(priceLine, ";")
        ' Les inn alle radene fra csv filen
        Do Until f.AtEndOfStream
            Dim i, strKey, strValue, arkivFolder, arkivFileName
            Dim strNewFileName, strDateTimeISO
            myLine = f.ReadLine
            writeLog objLog, "    " & myLine
            arrValue = Split(myLine, ";")
            ' Legg inn verdiene i Dictionary objektet
            For i = 0 to Ubound(arrKey)
                strKey = arrKey(i)
                strValue = arrValue(i)
                If strKey = "DATO" Then
                    strValue = convertToISO(strValue)
                End If
                bResult = updateDict(objDict, strKey, strValue)
            Next
            ' Legg inn alle verdiene i XML fila
            intFileNo = intFileNo + 1
            strNewFileName = priceFileNamePrefix & intFileNo & ".xml"
            bResult = updateDict(objDict, "pduref", priceFileNamePrefix & intFileNo)
            strDateTimeISO = getYearMonthDayISO & "T" & getHourMinuteSecondISO
            bResult = updateDict(objDict, "datetime", strDateTimeISO)
            strMessage = mergeDict(objDict, strISOmsg)
            If writeNewFile(priceOutFolder & strNewFileName, strMessage) Then
                writeLog objLog, "      Filen " & strNewFileName & " ble laget"
            Else
                writeLog objLog, "    Filen " & strNewFileName & " ble ikke laget"
            End If
        Loop
        f.Close
        arkivFolder = getDict(objDict, "Arkiv") & "\"
        arkivFileName = "A" & Mid(getYearMonthDay, 3, 2) & _ 
                Mid(getYearMonthDay, 5, 2) & Mid(getYearMonthDay, 7, 2) & _ 
                Left(getHourMinuteSecond, 4) & inFileName
        
        bResult = moveFile(priceFolder & inFileName, arkivFolder & arkivFileName)
        writeLog objLog, "Kursfilen " & arkivFileName & " ble flyttet til " & arkivFolder & " [" & bResult & "]"
    Next
    writeLog objLog, "LesKurser avsluttet normalt"
    objLog.Close
    Wscript.Echo "LesKurser avsluttet normalt " & Now
End Sub
