Sub CheckParticipantsByDateOnly()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Assumes data is in the first sheet; adjust as necessary
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Finds the last row in Column A
    
    ' Adjust the array of participants as needed
    Dim participants As Variant
    participants = Array("ESB - ESBIE NI", "ESB - ESBIE", "ESB – Coolkeeragh", "ESB - Customer Supply", "ESB – PGEN", "ESB – Synergen")
    
    Dim dateParticipantsDict As Object
    Set dateParticipantsDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lastRow ' Assuming Row 1 has headers
        Dim currentDate As Date
        currentDate = Int(CDate(ws.Cells(i, "A").Value)) ' Get the start date, removing time
        
        If Not dateParticipantsDict.exists(currentDate) Then
            dateParticipantsDict.Add currentDate, CreateObject("Scripting.Dictionary")
            Dim part As Variant
            For Each part In participants
                dateParticipantsDict(currentDate).Add part, False
            Next part
        End If
        
        Dim currentParticipant As String
        currentParticipant = ws.Cells(i, "E").Value ' Get the participant
        
        If dateParticipantsDict(currentDate).exists(currentParticipant) Then
            dateParticipantsDict(currentDate)(currentParticipant) = True
        End If
    Next i
    
    Dim missingParticipants As String
    missingParticipants = ""
    Dim allDatesCovered As Boolean
    allDatesCovered = True
    
    Dim dateKey As Variant, participantKey As Variant
    For Each dateKey In dateParticipantsDict.keys
        For Each participantKey In dateParticipantsDict(dateKey).keys
            If dateParticipantsDict(dateKey)(participantKey) = False Then
                allDatesCovered = False
                missingParticipants = missingParticipants & "Participant " & participantKey & " is missing for date " & Format(dateKey, "mm/dd/yyyy") & "." & vbCrLf
            End If
        Next participantKey
    Next dateKey
    
    If allDatesCovered Then
        MsgBox "All participants are present for all dates!", vbInformation
    Else
        MsgBox "Some participants are missing for certain dates:" & vbCrLf & missingParticipants, vbExclamation
    End If
End Sub
