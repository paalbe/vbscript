Option Explicit
Sub HentSvar
    Dim logFileNamePrefix, logFilePath
    Dim objSOAP, sResponse
    Dim objDict, objLog, objSeq
    Dim sPage, sSOAPmap
    Dim priceFilesFolder, priceFiles, priceFileName
    Dim StartTidspunkt
    Dim bDone
    
    StartTidspunkt = Now
    Wscript.Echo "HentSvar Startet " & StartTidspunkt

'    priceFileNamePrefix = "PR" & getYearMonthDay & getHourMinuteSecond
    logFileNamePrefix = "\Resp" & getYearMonthDay

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
        Wscript.Echo "HentSvar avslutter. Rett feilen og start igjen"
        Exit Sub
    End If
     ' Klargjør for logging, det er en loggfil fra hver dag.
    logFilePath = getDict(objDict, "LogFileFolder") & logFileNamePrefix & ".txt"
    Set objLog = openLogFile(logFilePath)

    writeLog objLog, "HentSvar starttidspunkt: " & StartTidspunkt
    writeLog objLog, "Poperties.ini ble lest inn:"
    ' List ut parameterene som ble lest inn fra properties.ini
    logDictObj objLog, objDict

    ' Forberede for SOAP kall til Centevo
    Set objSOAP = New ServiceRequest 
    objSOAP.SetSoapAction(getDict(objDict, "VPSUrl"))
    sSOAPmap = readAll(getDict(objDict, "GetTemplateFile"))
    objSOAP.SetUidPw getDict(objDict, "SystemUser"), getDict(objDict, "SystemPassword")
    objSOAP.sSOAPRequest = sPage
    objSOAP.openConnection

    ' Lag prefix for sekvensnummer til Centevo
    Set objSeq = New CSequenceNumber
    objSeq.SetDCID getDict(objDict, "IdInMsgId")    

    bDone = False
    While bDone = False
        Dim bResult, strValue
        strValue = objSeq.getNextSequence
        bResult = updateDict(objDict, "vpsmsgid", strValue)
        objSOAP.sSOAPRequest = mergeDict(objDict, sSOAPmap)
        objSOAP.SendRequest
        'writeLog objLog, objSOAP.getResponse()
        sResponse = objSOAP.getElementText("getMessageFromQueueResponse")
        If Len(sResponse) < 1 Then
            writeLog objLog, "Siste melding har blitt lest!"
            bDone = True
        Else
            Dim oXML, oDict, rFile
            Set oXML = CreateObject("Microsoft.XMLDOM")
            Set oDict = createDict
            If oXML.LoadXML(sResponse) Then
                responseDict oXML.childNodes, oDict
                rFile = getDict(objDict, "ResponseFolder") & "\" & getDict(oDict, "Ref") 
                If Left(getDict(oDict, "NS1:MessageIdentifier"), 4) = "semt" Then
                    rFile = rFile & ".semt.xml"
                Else
                    rFile = rFile & ".xml"
                End If
                writeNewFile rFile, sResponse
                writeLog objLog, "Mottatt melding lagt i fil " & rFile
            Else
                writeLog objLog, "Mottatt melding lot seg ikke parse " & oXML.parsed
            End If
        End If
    Wend

    objSOAP.Close
    writeLog objLog, "HentSvar avsluttet normalt"
    objLog.Close
    Set objSeq = Nothing ' Lagre sist brukte sekvensnummer
    Wscript.Echo "HentSvar avsluttet normalt " & Now

 End Sub