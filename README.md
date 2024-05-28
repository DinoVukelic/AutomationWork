Sub CheckParticipantsByDateOnly()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Assumes data is on the first sheet; adjust as necessary

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Finds the last row in Column A

    ' Define the list of participants with trimmed spaces and standardized hyphens
    Dim participants As Variant
    participants = Array("ESB - ESBIE NI", "ESB - ESBIE", "ESB - Coolkeeragh", "ESB - Customer Supply", "ESB - PGEN", "ESB - Synergen")
    Dim j As Integer
    For j = LBound(participants) To UBound(participants)
        ' Replace any en-dash or em-dash with a standard hyphen-minus and trim spaces
        participants(j) = Trim(Replace(Replace(participants(j), ChrW(8211), "-"), ChrW(8212), "-"))
    Next j

    Dim dateParticipantsDict As Object
    Set dateParticipantsDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To lastRow ' Assuming Row 1 has headers
        Dim currentDate As String
        currentDate = Format(ws.Cells(i, "A").Value, "dd/mm/yyyy") ' Ensure the date format is consistent

        ' Add date if it doesn't exist
        If Not dateParticipantsDict.exists(currentDate) Then
            dateParticipantsDict.Add currentDate, CreateObject("Scripting.Dictionary")
            Dim part As Variant
            For Each part In participants
                dateParticipantsDict(currentDate).Add part, False
            Next part
        End If

        Dim currentParticipant As String
        currentParticipant = Trim(ws.Cells(i, "E").Value) ' Get the participant name and trim spaces
        currentParticipant = Replace(Replace(currentParticipant, ChrW(8211), "-"), ChrW(8212), "-") ' Standardize hyphens

        ' Mark participant as present if their name matches
        If dateParticipantsDict(currentDate).exists(currentParticipant) Then
            dateParticipantsDict(currentDate)(currentParticipant) = True
        End If
    Next i

    Dim missingParticipants As String
    missingParticipants = ""
    Dim allDatesCovered As Boolean
    allDatesCovered = True

    ' Check each date for missing participants
    Dim dateKey As Variant
    For Each dateKey In dateParticipantsDict.keys
        Dim isDateMissing As Boolean
        isDateMissing = False
        Dim participantList As String
        participantList = ""

        Dim participantKey As Variant
        For Each participantKey In dateParticipantsDict(dateKey).keys
            If dateParticipantsDict(dateKey)(participantKey) = False Then
                allDatesCovered = False
                isDateMissing = True
                participantList = participantList & "Participant " & participantKey & vbCrLf
            End If
        Next participantKey

        ' Add only if there are missing participants for this date
        If isDateMissing Then
            missingParticipants = missingParticipants & "Date: " & dateKey & vbCrLf & participantList & vbCrLf
        End If
    Next dateKey

    ' Display results based on the gathered information
    If allDatesCovered Then
        MsgBox "All participants are present for all dates!", vbInformation
    Else
        If Len(missingParticipants) > 0 Then
            MsgBox "Some participants are missing for certain dates:" & vbCrLf & missingParticipants, vbExclamation
        Else
            MsgBox "There is no missing data for any date.", vbInformation
        End If
    End If
End Sub
