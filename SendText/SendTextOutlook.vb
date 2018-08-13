Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Runtime.InteropServices


Module SendTextOutlook

    Dim gOutlook As Outlook.Application
    Dim gOutlookNS As Outlook.NameSpace


    Public Function OutlookOpen() As Integer

        Dim lError As Integer = 0

        '        Try
        '
        'gOutlook = DirectCast(Marshal.GetActiveObject("Outlook.Application.16"), Outlook.Application) ' Is there a current instance of Outlook 2016?
        '
        'Catch

        Try

            gOutlook = New Outlook.Application ' If no then instantiate one

        Catch ex As Exception

            lError = -1 ' Flag failure to open mail client
            LogMessage("*** ERROR *** OutlookOpen.gOutlook: " & ex.ToString)

        End Try

        'End Try

        If lError = 0 Then

            Try

                gOutlookNS = gOutlook.GetNamespace("MAPI")
                gOutlookNS.Logon("Outlook") ' Name of MAPI profile
                gOutlookNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) ' Initialize MAPI

            Catch ex As Exception

                lError = -1 ' Flag failure to initialize MAPI
                LogMessage("*** ERROR *** OutlookOpen.gOutlookNS: " & ex.ToString)

            End Try

        End If

        OutlookOpen = lError

    End Function


    Public Sub OutlookClose()

        gOutlookNS.Logoff()
        gOutlook.Quit()

        gOutlookNS = Nothing
        gOutlook = Nothing

    End Sub


    Public Function OutlookSend(pTo As String, pSubject As String, pBody As String, pSend As Boolean) As Integer

        Dim lError As Integer = 0
        Dim lItem As Outlook.MailItem

        Try

            lItem = gOutlook.CreateItem(Outlook.OlItemType.olMailItem)

            With lItem
                .To = pTo
                .Subject = pSubject
                .Body = pBody

                If pSend Then
                    .Send()
                Else
                    .Save()
                    .Close(Outlook.OlInspectorClose.olSave)
                End If

            End With

        Catch ex As Exception

            lError = -1 ' Flag failure to create an email
            LogMessage("*** ERROR *** OutlookSend: " & ex.ToString)

        Finally

            lItem = Nothing

        End Try

        OutlookSend = lError

    End Function


    Public Sub OutlookFlush()

        Dim lSync As Object

        Try

            GC.Collect() ' Force garbage collection to free-up old RPC connections
            GC.WaitForPendingFinalizers()

            lSync = gOutlookNS.SyncObjects.Item(1)
            lSync.Start

        Catch

            ' If we can't Send/Receive then don't worry about it because it will happen eventually anyway

        Finally

            lSync = Nothing

        End Try

    End Sub


End Module
