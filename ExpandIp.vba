Option Explicit ' Enforces variable declaration, good practice

Sub ExpandIPAddresses()
    ' --- Configuration ---
    ' !!! IMPORTANT: Adjust these sheet and column settings to match your Excel file !!!
    Const INPUT_SHEET_NAME As String = "Sheet1"    ' Name of the sheet containing your IP list
    Const INPUT_COLUMN As String = "A"             ' Column letter where IP ranges/addresses are
    Const START_ROW As Long = 1                    ' Row number where your IP list begins (1 if no header, 2 if header in row 1)
    Const OUTPUT_SHEET_NAME As String = "ExpandedIPs" ' Name for the new sheet with results
    Const OUTPUT_HEADER_TEXT As String = "Expanded IP Address"
    ' --- End Configuration ---

    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowInput As Long
    Dim outputRow As Long
    Dim i As Long
    Dim ipRangeOrAddress As String
    Dim ipParts() As String
    Dim startIPString As String
    Dim endIPString As String
    Dim startOctets(0 To 3) As Long
    Dim endOctets(0 To 3) As Long
    Dim currentOctets(0 To 3) As Long
    Dim octet1 As Long, octet2 As Long, octet3 As Long, octet4 As Long
    Dim newIP As String
    Dim tempSheet As Worksheet

    ' Error handling for sheet operations
    On Error GoTo ErrorHandler

    ' Get the input worksheet
    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET_NAME)
    If wsInput Is Nothing Then
        MsgBox "Input sheet '" & INPUT_SHEET_NAME & "' not found. Please check the INPUT_SHEET_NAME constant in the VBA code.", vbCritical, "Sheet Not Found"
        Exit Sub
    End If

    ' Prepare the output worksheet
    ' Delete the output sheet if it already exists to avoid errors/old data
    Application.DisplayAlerts = False ' Turn off confirmation dialogs
    On Error Resume Next ' Ignore error if sheet doesn't exist
    Set tempSheet = ThisWorkbook.Sheets(OUTPUT_SHEET_NAME)
    If Not tempSheet Is Nothing Then
        tempSheet.Delete
    End If
    On Error GoTo ErrorHandler ' Reinstate error handling
    Application.DisplayAlerts = True  ' Turn confirmation dialogs back on

    ' Add a new sheet for the output
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=wsInput)
    wsOutput.Name = OUTPUT_SHEET_NAME

    ' Add header to the output sheet
    wsOutput.Cells(1, 1).Value = OUTPUT_HEADER_TEXT
    outputRow = 2 ' Start writing data from row 2

    ' Find the last row with data in the input column
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, INPUT_COLUMN).End(xlUp).Row

    ' Check if there's any data to process
    If lastRowInput < START_ROW Then
        MsgBox "No data found in column " & INPUT_COLUMN & " starting from row " & START_ROW & " on sheet '" & INPUT_SHEET_NAME & "'.", vbInformation, "No Data"
        Exit Sub
    End If

    ' Loop through each cell in the input column
    For i = START_ROW To lastRowInput
        ipRangeOrAddress = Trim(CStr(wsInput.Cells(i, INPUT_COLUMN).Value))

        If ipRangeOrAddress = "" Then GoTo NextIteration ' Skip empty cells

        ' Check if it's a range (contains "-") or a single IP
        If InStr(1, ipRangeOrAddress, "-") > 0 Then
            ipParts = Split(ipRangeOrAddress, "-")
            If UBound(ipParts) <> 1 Then ' Should be exactly two parts for a valid range
                wsOutput.Cells(outputRow, 1).Value = "Invalid Range Format: " & ipRangeOrAddress
                outputRow = outputRow + 1
                GoTo NextIteration
            End If
            startIPString = Trim(ipParts(0))
            endIPString = Trim(ipParts(1))
        Else
            startIPString = ipRangeOrAddress
            endIPString = ipRangeOrAddress ' Single IP, so start and end are the same
        End If

        ' Validate and convert start and end IPs to octet arrays
        If Not ConvertIPStringToOctets(startIPString, startOctets) Then
            wsOutput.Cells(outputRow, 1).Value = "Invalid Start IP: " & startIPString & " (Original: " & ipRangeOrAddress & ")"
            outputRow = outputRow + 1
            GoTo NextIteration
        End If

        If Not ConvertIPStringToOctets(endIPString, endOctets) Then
            wsOutput.Cells(outputRow, 1).Value = "Invalid End IP: " & endIPString & " (Original: " & ipRangeOrAddress & ")"
            outputRow = outputRow + 1
            GoTo NextIteration
        End If
        
        ' Ensure start IP is not greater than end IP numerically
        If IPToLong(startIPString) > IPToLong(endIPString) Then
            wsOutput.Cells(outputRow, 1).Value = "Start IP is greater than End IP: " & ipRangeOrAddress
            outputRow = outputRow + 1
            GoTo NextIteration
        End If

        ' Loop through all IPs in the range
        For octet1 = startOctets(0) To endOctets(0)
            currentOctets(0) = octet1
            For octet2 = IIf(octet1 = startOctets(0), startOctets(1), 0) To IIf(octet1 = endOctets(0), endOctets(1), 255)
                currentOctets(1) = octet2
                For octet3 = IIf(octet1 = startOctets(0) And octet2 = startOctets(1), startOctets(2), 0) To IIf(octet1 = endOctets(0) And octet2 = endOctets(1), endOctets(2), 255)
                    currentOctets(2) = octet3
                    For octet4 = IIf(octet1 = startOctets(0) And octet2 = startOctets(1) And octet3 = startOctets(2), startOctets(3), 0) To IIf(octet1 = endOctets(0) And octet2 = endOctets(1) And octet3 = endOctets(2), endOctets(3), 255)
                        currentOctets(3) = octet4

                        newIP = currentOctets(0) & "." & currentOctets(1) & "." & currentOctets(2) & "." & currentOctets(3)
                        wsOutput.Cells(outputRow, 1).Value = newIP
                        outputRow = outputRow + 1

                        ' Safety break for extremely large ranges to prevent Excel from freezing
                        If outputRow > wsOutput.Rows.Count - 10 Then ' wsOutput.Rows.Count is max rows for the Excel version
                             wsOutput.Cells(outputRow, 1).Value = "WARNING: Maximum row limit reached. Expansion stopped."
                             MsgBox "Expansion stopped as it was about to exceed Excel's row limit." & vbCrLf & _
                                    "The last processed input was: " & ipRangeOrAddress, vbExclamation, "Row Limit Reached"
                             GoTo CleanUpAndExit
                        End If
                    Next octet4
                Next octet3
            Next octet2
        Next octet1
NextIteration:
    Next i

CleanUpAndExit:
    ' Auto-fit column width in output sheet
    wsOutput.Columns(1).AutoFit

    MsgBox "IP address expansion complete." & vbCrLf & "Results are in the sheet named '" & OUTPUT_SHEET_NAME & "'.", vbInformation, "Process Complete"
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True ' Ensure alerts are re-enabled
    If Err.Number = 9 And wsInput Is Nothing Then ' Subscript out of range (likely sheet not found)
         MsgBox "Error: Input sheet '" & INPUT_SHEET_NAME & "' not found. Please check the configuration at the top of the VBA code.", vbCritical, "Configuration Error"
    Else
        MsgBox "An unexpected error occurred:" & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Description: " & Err.Description, vbCritical, "VBA Runtime Error"
    End If
    ' Optionally, clean up any created objects if necessary
    Set wsInput = Nothing
    Set wsOutput = Nothing
End Sub

Private Function ConvertIPStringToOctets(ByVal ipString As String, ByRef octets() As Long) As Boolean
    ' Converts an IP string (e.g., "192.168.1.1") to an array of Longs.
    ' Returns True if successful, False otherwise.
    ' octets array should be Dim octets(0 To 3) As Long
    Dim parts() As String
    Dim i As Integer
    Dim tempVal As Variant

    ConvertIPStringToOctets = False ' Assume failure

    If ipString = "" Then Exit Function

    parts = Split(ipString, ".")
    If UBound(parts) <> 3 Then Exit Function ' Must have 4 parts

    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then Exit Function
        tempVal = CLng(parts(i)) ' Use CLng for conversion, can handle numbers as strings
        If tempVal < 0 Or tempVal > 255 Then Exit Function
        octets(i) = tempVal
    Next i

    ConvertIPStringToOctets = True
End Function

Private Function IPToLong(ByVal ipAddress As String) As Double
    ' Converts an IP address string to a sortable numeric representation.
    ' Using Double to ensure it can hold the large number.
    Dim octets() As String
    Dim i As Integer
    Dim num As Double

    IPToLong = 0 ' Default for invalid IP
    If ipAddress = "" Then Exit Function

    octets = Split(ipAddress, ".")
    If UBound(octets) <> 3 Then Exit Function

    For i = 0 To 3
        If Not IsNumeric(octets(i)) Then Exit Function
        If CLng(octets(i)) < 0 Or CLng(octets(i)) > 255 Then Exit Function
    Next i

    num = (CLng(octets(0)) * (256 ^ 3)) + _
          (CLng(octets(1)) * (256 ^ 2)) + _
          (CLng(octets(2)) * 256) + _
           CLng(octets(3))
    IPToLong = num
End Function
