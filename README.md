Sub CheckParticipantsByDateOnly()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Assumes data is in the first sheet; adjust as necessary
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Finds the last row in Column A
    
    ' Define the list of participants with trimmed spaces and standardized hyphens
    Dim participants As Variant
    participants = Array("ESB - ESBIE NI", "ESB - ESBIE", "ESB – Coolkeeragh", "ESB - Customer Supply", "ESB – PGEN", "ESB – Synergen")
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
        
        If Not dateParticipantsDict.exists(currentDate) Then
            dateParticipantsDict.Add currentDate, CreateObject("Scripting.Dictionary")
            Dim part As Variant
            For Each part In participants
                dateParticipantsDict(currentDate).Add part, False
            Next part
        End If
        
        Dim currentParticipant As String
        currentParticipant = Trim(ws.Cells(i, "E").Value) ' Get the participant and trim spaces
        
        ' Standardize hyphens in participant name to avoid mismatch due to different hyphen characters
        currentParticipant = Replace(Replace(currentParticipant, ChrW(8211), "-"), ChrW(8212), "-")
        
        Dim partKey As Variant
        For Each partKey In dateParticipantsDict(currentDate).keys
            If UCase(currentParticipant) = UCase(partKey) Then
                dateParticipantsDict(currentDate)(partKey) = True
                Exit For
            End If
        Next partKey
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
                missingParticipants = missingParticipants & "Participant " & participantKey & " is missing for date " & dateKey & "." & vbCrLf
            End If
        Next participantKey
    Next dateKey
    
    If allDatesCovered Then
        MsgBox "All participants are present for all dates!", vbInformation
    Else
        MsgBox "Some participants are missing for certain dates:" & vbCrLf & missingParticipants, vbExclamation
    End If
End Sub
