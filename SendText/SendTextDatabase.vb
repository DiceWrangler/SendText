Imports System.Data.SqlClient
Imports System.Configuration


Module SendTextDatabase

    Dim gDBConn As SqlConnection


    Public Function DBOpen() As Integer

        Dim lError As Integer = 0
        Dim lConnectionString As String

        Try

            '   gDBConn = New SqlConnection("Initial Catalog=SendText;Data Source=localhost;Integrated Security=SSPI;MultipleActiveResultSets=True;")
            lConnectionString = ConfigurationManager.ConnectionStrings("SendText").ConnectionString
            gDBConn = New SqlConnection(lConnectionString)
            gDBConn.Open()

        Catch ex As Exception

            lError = -1 ' Flag failure to open database
            LogMessage("*** ERROR *** DBOpen: " & ex.ToString)

        End Try

        DBOpen = lError

    End Function


    Public Sub DBClose()

        Try

            gDBConn.Close()
            gDBConn.Dispose()

        Catch

            ' No worries if error encountered while closing

        Finally

            gDBConn = Nothing

        End Try


    End Sub


    Public Function GetAppConfig(pAppName As String, pConfigName As String, pDefaultValue As String) As String

        Dim lCmd As New SqlCommand
        Dim lResults As String

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetAppConfig"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.VarChar)
            lCmd.Parameters("@AppName").Value = pAppName
            lCmd.Parameters.Add("@ConfigName", SqlDbType.VarChar)
            lCmd.Parameters("@ConfigName").Value = pConfigName

            lResults = lCmd.ExecuteScalar
            If IsNothing(lResults) Then lResults = pDefaultValue 'If configuration not defined just use default value

        Catch ex As Exception

            lResults = pDefaultValue ' If we encountered an error just use default value but log it anyway
            LogMessage("*** ERROR *** GetAppConfig: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        GetAppConfig = lResults

    End Function


    Public Function GetActiveTextBatches() As SqlDataReader

        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader

        Try
            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetActiveTextBatches"
            lCmd.CommandType = CommandType.StoredProcedure

            lReader = lCmd.ExecuteReader(CommandBehavior.SingleRow)

            GetActiveTextBatches = lReader

        Catch ex As Exception

            LogMessage("*** ERROR *** GetActiveTextBatches: " & ex.ToString)
            GetActiveTextBatches = Nothing

        Finally

            lReader = Nothing

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Function


    Public Function GetActiveTextMessages(pBatchID As Integer) As SqlDataReader

        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetActiveTextMessages"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@BatchID", SqlDbType.Int)
            lCmd.Parameters("@BatchID").Value = pBatchID

            lReader = lCmd.ExecuteReader

            GetActiveTextMessages = lReader

        Catch ex As Exception

            LogMessage("*** ERROR *** GetActiveTextMessages: " & ex.ToString)
            GetActiveTextMessages = Nothing

        Finally

            lReader = Nothing

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Function


    Public Function GetTextMessage(pMessageID As Integer) As SqlDataReader

        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "GetTextMessage"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@MessageID", SqlDbType.Int)
            lCmd.Parameters("@MessageID").Value = pMessageID

            lReader = lCmd.ExecuteReader(CommandBehavior.SingleRow)

            GetTextMessage = lReader

        Catch ex As Exception

            LogMessage("*** ERROR *** GetTextMessage: " & ex.ToString)
            GetTextMessage = Nothing

        Finally

            lReader = Nothing

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Function


    Public Sub UpdateTextBatchStatus(pBatchID As Integer, pBatchStatus As String)

        Dim lCmd As New SqlCommand

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "UpdateTextBatchStatus"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@BatchID", SqlDbType.Int)
            lCmd.Parameters("@BatchID").Value = pBatchID
            lCmd.Parameters.Add("@BatchStatus", SqlDbType.Char)
            lCmd.Parameters("@BatchStatus").Value = pBatchStatus

            lCmd.ExecuteNonQuery()

        Catch ex As Exception

            LogMessage("*** ERROR *** UpdateTextBatchStatus: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Sub


    Public Sub UpdateTextMessageStatus(pMessageID As Integer, pMessageStatus As String)

        Dim lCmd As New SqlCommand

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "UpdateTextMessageStatus"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@MessageID", SqlDbType.Int)
            lCmd.Parameters("@MessageID").Value = pMessageID
            lCmd.Parameters.Add("@MessageStatus", SqlDbType.Char)
            lCmd.Parameters("@MessageStatus").Value = pMessageStatus

            lCmd.ExecuteNonQuery()

        Catch ex As Exception

            LogMessage("*** ERROR *** UpdateTextMessageStatus: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Sub


    Public Function ImportTxtLog() As Integer

        Dim lError As Integer = 0
        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader
        Dim lRecord As IDataRecord
        Dim lBatchesImported, lMessagesImported As Integer

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "ImportTxtLog"
            lCmd.CommandType = CommandType.StoredProcedure

            lReader = lCmd.ExecuteReader(CommandBehavior.SingleRow)

            lReader.Read()
            If lReader.HasRows Then

                lRecord = CType(lReader, IDataRecord)

                lBatchesImported = lRecord.GetInt32(0)
                lMessagesImported = lRecord.GetInt32(1)

                If (lBatchesImported > 0) Or (lMessagesImported > 0) Then
                    LogMessage("Imported: " & lBatchesImported.ToString & " Batches and " & lMessagesImported.ToString & " Messages")
                End If

            End If

            lReader.Close()
            lReader = Nothing

        Catch ex As Exception

            lError = -1 ' Flag failure to import new messages
            LogMessage("*** ERROR *** ImportTxtLog: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        ImportTxtLog = lError

    End Function


    Public Sub ShutdownRequestClear(pAppName As String)

        Dim lCmd As New SqlCommand

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "ShutdownRequestClear"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.Char)
            lCmd.Parameters("@AppName").Value = pAppName

            lCmd.ExecuteNonQuery()

        Catch ex As Exception

            LogMessage("*** ERROR *** ShutdownRequestClear: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Sub


    Public Sub ShutdownRequestSet(pAppName As String)

        Dim lCmd As New SqlCommand

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "ShutdownRequestSet"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.Char)
            lCmd.Parameters("@AppName").Value = pAppName

            lCmd.ExecuteNonQuery()

        Catch ex As Exception

            LogMessage("*** ERROR *** ShutdownRequestSet: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

    End Sub


    Public Function ShutdownRequestTest(pAppName As String, pDefaultValue As Boolean) As Boolean

        Dim lCmd As New SqlCommand
        Dim lResults As String

        Try

            lCmd = gDBConn.CreateCommand

            lCmd.CommandText = "ShutdownRequestTest"
            lCmd.CommandType = CommandType.StoredProcedure

            lCmd.Parameters.Add("@AppName", SqlDbType.VarChar)
            lCmd.Parameters("@AppName").Value = pAppName

            lResults = lCmd.ExecuteScalar
            If IsNothing(lResults) Then lResults = pDefaultValue.ToString 'If configuration not defined just use default value

        Catch ex As Exception

            lResults = pDefaultValue ' If we encountered an error just use default value but log it anyway
            LogMessage("*** ERROR *** ShutdownRequestTest: " & ex.ToString)

        Finally

            lCmd.Dispose()
            lCmd = Nothing

        End Try

        ShutdownRequestTest = (lResults = True.ToString)

    End Function


    Public Function SendDBMail(pTo As String, pSubject As String, pBody As String, pSend As Boolean) As Integer

        Dim lCmd As New SqlCommand
        Dim lReader As SqlDataReader

        If pSend Then ' sp_send_dbmail does not support drafts so nothing to do unless sending

            Try

                lCmd = gDBConn.CreateCommand

                lCmd.CommandText = "msdb.dbo.sp_send_dbmail"
                lCmd.CommandType = CommandType.StoredProcedure

                lCmd.Parameters.Add("@recipients", SqlDbType.VarChar)
                lCmd.Parameters("@recipients").Value = pTo
                lCmd.Parameters.Add("@subject", SqlDbType.VarChar)
                lCmd.Parameters("@subject").Value = pSubject
                lCmd.Parameters.Add("@body", SqlDbType.VarChar)
                lCmd.Parameters("@body").Value = pBody

                lReader = lCmd.ExecuteReader

                SendDBMail = 0

            Catch ex As Exception

                LogMessage("*** ERROR *** SendDBMail: " & ex.ToString)
                SendDBMail = -1

            Finally

                lReader = Nothing

                lCmd.Dispose()
                lCmd = Nothing

            End Try

        Else

            SendDBMail = 0

        End If

    End Function

End Module
