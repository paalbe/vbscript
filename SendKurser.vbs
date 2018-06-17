Option Explicit
Sub SendKurser
    Dim logFileNamePrefix, logFilePath
    Dim objSOAP, sResponse
    Dim objDict, objLog, objSeq
    Dim sPage, sSOAPmap
    Dim sSerialPrefix, requestNo
    Dim priceFilesFolder, priceFiles, priceFileName
    Dim StartTidspunkt
    
    StartTidspunkt = Now
    Wscript.Echo "SendPriser Startet " & StartTidspunkt

'    priceFileNamePrefix = "PR" & getYearMonthDay & getHourMinuteSecond
    logFileNamePrefix = "\Log" & getYearMonthDay

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
        Wscript.Echo "SendKurser avslutter. Rett feilen og start igjen"
        Exit Sub
    End If
     ' Klargjør for logging, det er en loggfil fra hver dag.
    logFilePath = getDict(objDict, "LogFileFolder") & logFileNamePrefix & ".txt"
    Set objLog = openLogFile(logFilePath)

    writeLog objLog, "SendKurser starttidspunkt: " & StartTidspunkt
    writeLog objLog, "Poperties.ini ble lest inn:"
    ' List ut parameterene som ble lest inn fra properties.ini
    logDictObj objLog, objDict

    ' Forberede for SOAP kall til Centevo
    Set objSOAP = New ServiceRequest 
    objSOAP.SetSoapAction(getDict(objDict, "VPSUrl"))
    sSOAPmap = readAll(getDict(objDict, "AddTemplateFile"))
    objSOAP.SetUidPw getDict(objDict, "SystemUser"), getDict(objDict, "SystemPassword")
    objSOAP.sSOAPRequest = sPage
    objSOAP.openConnection

    ' Lag prefix for sekvensnummer til Centevo
    sSerialPrefix = getYearMonthDay & getDict(objDict, "IdInMsgId")

    ' Lag prefix for sekvensnummer til Centevo
    Set objSeq = New CSequenceNumber
    objSeq.SetDCID getDict(objDict, "IdInMsgId")

    ' Les alle kursfilene som skal sendes
    priceFilesFolder = getDict(objDict, "PriceOutFolder") & "\"
    priceFiles = getFiles(priceFilesFolder, ".xml")

    For Each priceFileName in priceFiles
        Dim bResult, strValue
        strValue = objSeq.getNextSequence
        bResult = updateDict(objDict, "vpsmsgid", strValue)
        strValue = readAll(priceFilesFolder & priceFileName)
        bResult = updateDict(objDict, "datapdu", strValue)
        If bResult = False Then
            writeLog objLog, "Filen " & priceFilesFolder &  priceFileName & " kunne ikke leses"
            WScript.Echo "Filen " & priceFilesFolder & priceFileName & " kunne ikke leses"
            Exit Sub
        End If
        objSOAP.sSOAPRequest = mergeDict(objDict, sSOAPmap)
        objSOAP.SendRequest 
        'sResponse = objSOAP.getResponse
        sResponse = objSOAP.getElementText("addMessageToQueueResponse")
        writeLog objLog, sResponse
        deleteFile priceFilesFolder & priceFileName
        writeLog objLog, "Filen " & priceFilesFolder & priceFileName & " ble slettet"
    Next

    objSOAP.Close
    writeLog objLog, "SendKurser avsluttet normalt"
    objLog.Close
    Set objSeq = Nothing
    Wscript.Echo "SendKurser avsluttet normalt " & Now

 End Sub