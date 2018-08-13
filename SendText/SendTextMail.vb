Module SendTextMail

    Const MAX_SUBJECT_LENGTH As Integer = 26
    Const MAX_MESSAGE_LENGTH As Integer = 113
    Const MAX_MESSAGE_PARTS As Integer = 10


    Private Function GetSubjectLength(pSubject As String) As Integer

        Dim lSubjectLength As Integer

        If Len(pSubject) > MAX_SUBJECT_LENGTH Then
            lSubjectLength = MAX_SUBJECT_LENGTH
        Else
            lSubjectLength = Len(pSubject)
        End If

        GetSubjectLength = lSubjectLength

    End Function


    Private Function GetMessagePart(pBodyText As String, pMaxLength As Integer) As String

        Dim lMessagePart As String
        Dim lLastWhitespace As Integer

        If Len(pBodyText) > pMaxLength Then

            lMessagePart = Left(pBodyText, pMaxLength)
            lLastWhitespace = InStrRev(lMessagePart, " ") - 1

            If lLastWhitespace > 0 Then
                lMessagePart = Left(lMessagePart, lLastWhitespace)
            End If

        Else

            lMessagePart = pBodyText

        End If

        lMessagePart = Trim(lMessagePart)
        GetMessagePart = lMessagePart

    End Function


    Private Function BreakupMessageParts(pSubjectLineLength As Integer, pBodyText As String) As String()

        Dim lMaxBodyTextLength As Integer
        Dim lProcessedText As Boolean
        Dim lRawBodyText As String
        Dim lRawMessagePart As String
        Dim lMessagePartIdx As Integer
        Dim lMessagePart() As String

        ReDim lMessagePart(MAX_MESSAGE_PARTS) ' Pre-allocate array larger than needed; will be resized at end

        lRawBodyText = Trim(Replace(pBodyText, Chr(13), " "))  ' Replace Carriage Returns with a Space; preserve Line Feeds
        lMaxBodyTextLength = MAX_MESSAGE_LENGTH - pSubjectLineLength

        lMessagePartIdx = 0
        lProcessedText = False

        Do

            lRawMessagePart = GetMessagePart(lRawBodyText, lMaxBodyTextLength)

            If Len(lRawMessagePart) > 0 Then

                lMessagePartIdx = lMessagePartIdx + 1
                lMessagePart(lMessagePartIdx) = lRawMessagePart
                lRawBodyText = Trim(Mid(lRawBodyText, Len(lRawMessagePart) + 1))

            Else

                lProcessedText = True

            End If

        Loop Until lProcessedText Or (lMessagePartIdx >= MAX_MESSAGE_PARTS)

        ReDim Preserve lMessagePart(lMessagePartIdx)

        BreakupMessageParts = lMessagePart

    End Function


    Private Function FormatSubjectLine(pSubject As String, pMessagePartIdx As Integer, pNumMessageParts As Integer, pMessageID As Integer) As String

        Dim lSubject, lMessagePartDisplay, lSubjectLine As String

        lSubject = Left(pSubject, MAX_SUBJECT_LENGTH)
        lMessagePartDisplay = Trim(Str(pMessagePartIdx)) + "/" + Trim(Str(pNumMessageParts))

        lSubjectLine = lSubject + " [" + lMessagePartDisplay + "]"

        FormatSubjectLine = lSubjectLine

    End Function


    Public Function SendMail(pTo As String, pSubject As String, pBody As String, pMessageID As Integer, pSend As Boolean) As Integer

        Dim lError As Integer = 0
        Dim lSubject As String
        Dim lBody As String
        Dim lSubjectLength As Integer
        Dim lMessagePart As String()
        Dim lNumMessageParts As Integer
        Dim lSubjectLine As String
        Dim lMessagePartIdx As Integer

        lSubject = pSubject
        lBody = pBody
        lSubjectLength = GetSubjectLength(lSubject)
        lMessagePart = BreakupMessageParts(lSubjectLength, lBody)
        lNumMessageParts = UBound(lMessagePart)

        For lMessagePartIdx = 1 To lNumMessageParts

            lSubjectLine = FormatSubjectLine(lSubject, lMessagePartIdx, lNumMessageParts, pMessageID)
            lBody = lMessagePart(lMessagePartIdx)

            'lError = OutlookSend(pTo, lSubjectLine, lBody, pSend)
            lError = SendDBMail(pTo, lSubjectLine, lBody, pSend)
            If lError <> 0 Then Exit For ' Stop processing this message if an error encounter when sending it

        Next

        SendMail = lError

    End Function


End Module
