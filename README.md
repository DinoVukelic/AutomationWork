Sub CheckParticipantsByDate()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Assumes data is in the first sheet; adjust as necessary
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Finds the last row in Column A
    
    Dim participants As Variant
    participants = Array("ESB - ESBIE NI", "ESB - ESBIE", "ESB – Coolkeeragh", "ESB - Customer Supply", "ESB – PGEN", "ESB – Synergen")
    
    Dim startDateDict As Object
    Set startDateDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lastRow ' Assuming Row 1 has headers
        Dim currentDate As Date
        currentDate = Int(ws.Cells(i, "A").Value) ' Get the start date, removing time
        
        If Not startDateDict.exists(currentDate) Then
            startDateDict.Add currentDate, CreateObject("Scripting.Dictionary")
            Dim participant
            For Each participant In participants
                startDateDict(currentDate).Add participant, False
            Next participant
        End If
        
        Dim currentParticipant As String
        currentParticipant = ws.Cells(i, "E").Value ' Get the participant
        If startDateDict(currentDate).exists(currentParticipant) Then
            startDateDict(currentDate)(currentParticipant) = True
        End If
    Next i
    
    Dim missingInfo As String
    missingInfo = ""
    Dim dateKey, partKey
    For Each dateKey In startDateDict.keys
        Dim allPresent As Boolean
        allPresent = True
        For Each partKey In startDateDict(dateKey).keys
            If startDateDict(dateKey)(partKey) = False Then
                allPresent = False
                missingInfo = missingInfo & "Missing participant: " & partKey & " on date: " & dateKey & vbCrLf
            End If
        Next partKey
        If allPresent Then
            missingInfo = missingInfo & "All participants are present on " & dateKey & "." & vbCrLf
        End If
    Next dateKey
    
    If Len(missingInfo) = 0 Then
        MsgBox "All participants are present for all dates.", vbInformation
    Else
        MsgBox "Check for missing participants:" & vbCrLf & missingInfo, vbExclamation
    End If
End Sub
