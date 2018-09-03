Imports System.Data.SqlClient


Module SendText

    Dim gShutdownRequested As Boolean
    Dim gCycleInterval As Integer
    Dim gEmailToUse As String
    Dim gTestEmail As String
    Dim gOutput As String
    Dim gProcess As String
    Dim gHeartbeat As Integer

    Public Const APP_NAME As String = "SendText"
    Const APP_VERSION As String = "v180902"

    Const EMAIL_TO_USE_LIVE As String = "LIVE"
    Const EMAIL_TO_USE_TEST As String = "TEST"

    Const OUTPUT_SEND As String = "SEND"
    Const OUTPUT_DRAFT As String = "DRAFT"

    Const PROCESS_ALL As String = "ALL"
    Const PROCESS_FIRST As String = "FIRST"

    Const BATCH_STATUS_INITIAL As String = "I"
    Const BATCH_STATUS_CREATED As String = "C"
    Const BATCH_STATUS_QUEUED As String = "Q"
    Const BATCH_STATUS_RETRY As String = "R"
    Const BATCH_STATUS_STARTED As String = "S"
    Const BATCH_STATUS_FINISHED As String = "F"
    Const BATCH_STATUS_ERROR As String = "E"
    Const BATCH_STATUS_PRIORITY As String = "P"

    Const MESSAGE_STATUS_CREATED As String = "C"
    Const MESSAGE_STATUS_RETRY As String = "R"
    Const MESSAGE_STATUS_STARTED As String = "S"
    Const MESSAGE_STATUS_FINISHED As String = "F"
    Const MESSAGE_STATUS_ERROR As String = "E"

    Const APP_STATUS_START As String = "Starting"
    Const APP_STATUS_HEARTBEAT As String = "Running"
    Const APP_STATUS_SHUTDOWN As String = "Stopping"


    Sub Main()

        Dim lError As Integer = 0

        lError = SendTextStartup()
        If lError <> 0 Then GoTo MAIN_EXIT

        LoadAppConfigs() ' Subroutine just uses defaults if it fails for whatever reason

        ShutdownRequestClear(APP_NAME)
        gShutdownRequested = False

        Do

            Main_Loop()
            Threading.Thread.Sleep(gCycleInterval)

            If Not gShutdownRequested Then gShutdownRequested = ShutdownRequestTest(APP_NAME, False)

        Loop Until gShutdownRequested

MAIN_EXIT:
        SendTextShutdown()

    End Sub


    Sub Main_Loop()

        Dim lError As Integer = 0
        Dim lBatches As SqlDataReader
        Dim lRecord As IDataRecord
        Dim lBatchID As Integer
        Dim lSubject, lBody, lBatchStatus As String

        lError = ImportTxtLog()
        If lError <> 0 Then
            gShutdownRequested = True ' No point in continuuing if we cannot find detect new messages
            Exit Sub
        End If

        lBatches = GetActiveTextBatches()
        lBatches.Read() 'Only want first row

        If lBatches.HasRows Then

            lRecord = CType(lBatches, IDataRecord)
            lBatchID = lRecord.GetInt32(0)
            lSubject = Left(lRecord.GetString(2), 26) 'Only display first 26 characters
            lBody = Left(lRecord.GetString(3), 40) ' Only display first 40 characters
            lBatchStatus = lRecord.GetString(4)

            If lBatchStatus = BATCH_STATUS_PRIORITY Then LogMessage("*** PRIORITY MESSAGE ***")
            LogMessage("Subject: " & lSubject & ", Body: " & lBody)

            ProcessBatch(lBatchID)

        Else

            LogHeartbeat()

        End If

        lBatches.Close()

        lRecord = Nothing
        lBatches = Nothing

    End Sub


    Sub LogMessage(pMessage As String)

        If gHeartbeat > 0 Then Console.WriteLine()
        gHeartbeat = 0

        Console.WriteLine(Now.ToLocalTime & "> " & pMessage)

    End Sub


    Sub LogHeartbeat()

        Console.Write(".")
        gHeartbeat = (gHeartbeat + 1) Mod 80
        If gHeartbeat = 0 Then Console.WriteLine() ' Line break after 80 characters

        UpdateAppStatus(APP_STATUS_HEARTBEAT)

    End Sub


    Sub ProcessBatch(pBatchID As Integer)

        Dim lError As Integer = 0
        Dim lAnyError As Integer = 0
        Dim lMessages As SqlDataReader
        Dim lRecord As IDataRecord
        Dim lMessageID As Integer
        Dim lFirstOfBatch, lProcessMessage As Boolean
        Dim lTo, lToEmail As String

        UpdateTextBatchStatus(pBatchID, BATCH_STATUS_STARTED)

        lFirstOfBatch = True

        lMessages = GetActiveTextMessages(pBatchID)
        While lMessages.Read

            lProcessMessage = (lFirstOfBatch Or (gProcess = PROCESS_ALL))
            lFirstOfBatch = False

            lRecord = CType(lMessages, IDataRecord)
            lMessageID = lRecord.GetInt32(0)
            lTo = lRecord.GetString(1)
            lToEmail = lRecord.GetString(2)

            LogMessage("To: " & lTo & ", Email: " & lToEmail)

            lError = ProcessMessage(lMessageID, lProcessMessage)
            If lError <> 0 Then lAnyError = -1 'If any message in batch fails then mark batch as failed but keep trying to send individual messages

        End While

        If lAnyError = 0 Then
            UpdateTextBatchStatus(pBatchID, BATCH_STATUS_FINISHED)
        Else
            UpdateTextBatchStatus(pBatchID, BATCH_STATUS_ERROR)
        End If

        lMessages.Close()

        lRecord = Nothing
        lMessages = Nothing

    End Sub


    Function ProcessMessage(pMessageID As Integer, lProcessMessage As Boolean) As Integer

        Dim lError As Integer = 0

        UpdateTextMessageStatus(pMessageID, MESSAGE_STATUS_STARTED)

        If lProcessMessage Then lError = SendMessage(pMessageID)

        If lError = 0 Then
            UpdateTextMessageStatus(pMessageID, MESSAGE_STATUS_FINISHED)
        Else
            UpdateTextMessageStatus(pMessageID, MESSAGE_STATUS_ERROR)
        End If

        ProcessMessage = lError

    End Function


    Function SendMessage(pMessageID As Integer) As Integer

        Dim lError As Integer = 0
        Dim lSendMessage As Boolean
        Dim lMessage As SqlDataReader
        Dim lRecord As IDataRecord
        Dim lTextTo, lTenantEmail, lTextToEmail, lSubject, lBody As String

        lSendMessage = (gOutput = OUTPUT_SEND)

        lMessage = GetTextMessage(pMessageID)
        lMessage.Read()

        If lMessage.HasRows Then

            lRecord = CType(lMessage, IDataRecord)

            lTextTo = lRecord.GetString(1)
            lTenantEmail = lRecord.GetString(2)
            lSubject = lRecord.GetString(3)
            lBody = lRecord.GetString(4)

            If gEmailToUse = EMAIL_TO_USE_LIVE Then
                lTextToEmail = lTenantEmail
            Else
                lTextToEmail = gTestEmail
            End If

            lError = SendMail(lTextToEmail, lSubject, lBody, pMessageID, lSendMessage)

        End If

        SendMessage = lError

        lMessage.Close()

        lRecord = Nothing
        lMessage = Nothing

    End Function


    Function SendTextStartup() As Integer

        Dim lError As Integer = 0
        Dim lAnyError As Integer = 0

        LogMessage(APP_NAME & ": Starting up (" & APP_VERSION & ")")

        lError = DBOpen()
        If lError <> 0 Then lAnyError = -1 'If we cannot open the database, flag the error but keep going in case there are more errors during initialization

        UpdateAppStatus(APP_STATUS_START)

        SendTextStartup = lAnyError

    End Function


    Sub SendTextShutdown()

        LogMessage(APP_NAME & ": Shutting down")
        UpdateAppStatus(APP_STATUS_SHUTDOWN)

        DBClose()

    End Sub


    Sub LoadAppConfigs()

        Dim lCycleInterval As String

        lCycleInterval = GetAppConfig(APP_NAME, "Cycle_Interval", "10000")
        If IsNumeric(lCycleInterval) Then
            gCycleInterval = Val(lCycleInterval)
        Else
            gCycleInterval = 10000
        End If
        LogMessage("Config: Cycle_Interval: " & gCycleInterval.ToString)

        If UCase(GetAppConfig(APP_NAME, "Email_to_Use", EMAIL_TO_USE_TEST)) = EMAIL_TO_USE_LIVE Then
            gEmailToUse = EMAIL_TO_USE_LIVE ' Send to Tenant (LIVE)
        Else
            gEmailToUse = EMAIL_TO_USE_TEST ' Send to Developer (TEST)
        End If
        LogMessage("Config: Email_to_Use: " & gEmailToUse)

        gTestEmail = GetAppConfig(APP_NAME, "Test_Email", "7208403074@messaging.sprintpcs.com")  ' Default to Scott Thorne's SMS email address
        LogMessage("Config: Test_Email: " & gTestEmail)

        'If UCase(GetAppConfig(APP_NAME, "Output", OUTPUT_DRAFT)) = OUTPUT_SEND Then
        gOutput = OUTPUT_SEND ' Send messages
        'Else
        '    gOutput = OUTPUT_DRAFT ' Create draft only
        'End If
        LogMessage("Config: Output: " & gOutput & " (hard-coded; ignoring App_Config)")

        If UCase(GetAppConfig(APP_NAME, "Process", PROCESS_FIRST)) = PROCESS_ALL Then
            gProcess = PROCESS_ALL ' All messages
        Else
            gProcess = PROCESS_FIRST ' First in batch
        End If
        LogMessage("Config: Process: " & gProcess)

        LogMessage(Strings.StrDup(57, "="))

    End Sub

End Module
