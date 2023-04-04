Attribute VB_Name = "SplitText"
'*********************************************
'Tutorial: Split range values from cells to different rows
'Code: VBA
'Platform: Microsoft Excel
'Only for educational propouses
'https://www.youtube.com/channel/UCwJ8qS-Jr8h-BaCfIrpRzEQ/featured
'*********************************************
Option Explicit
'---------------------------------------------
'Main Sub
'---------------------------------------------
Public Sub SplitTextFromRange()
    'Vars
    Dim myRange, myColumn, myCell As String
    Dim myRow As Long
    Dim selRange As Range
    
    'Select Range of cells and convert to String
    Do While CheckRange(1, myRange) = False
        myRange = InputBox("Type your range (example A1 or A1:A10): ", "Selected Range")
        On Error Resume Next
            Set selRange = ActiveSheet.Range(myRange)
            myCell = RangeToString(selRange)
        On Error GoTo 0
    Loop
    
    'Select Column by an InputBox
    Do While CheckRange(2, myColumn) = False
        myColumn = InputBox("Type the column", "Print Results")
    Loop
    
    'Select Row by an InputBox
    Do While CheckRange(3, myRow) = False
        On Error Resume Next
            myRow = InputBox("Type the row", "Print Results")
        On Error GoTo 0
    Loop
    
    'Get all Data [myCell] into collection
    Dim UCol As New Collection
    Set UCol = AddStringToCollection(myCell)
    
    'Print data in a selected column from selected row
    Dim i As Integer
    For i = 1 To UCol.Count
        Range(myColumn & myRow).Value = UCol.Item(i)
        myRow = myRow + 1
    Next i
    
    'End main sub
    myRow = myRow - UCol.Count
    Range(myColumn & myRow).Select
End Sub
'*********************************************
'PRIVATE FUNCTIONS
'*********************************************
'---------------------------------------------
'Add string to collection by each Break Line
'---------------------------------------------
Private Function AddStringToCollection(ByVal oCell As String) As Collection
    Dim myCol As New Collection
    Dim myStr As String
    Dim i As Integer
    For i = 1 To Len(oCell)
        myStr = myStr & Mid(oCell, i, 1)
        If Mid(oCell, i, 1) = Chr(10) Or i = Len(oCell) Then
            myCol.Add (Replace(myStr, vbLf, ""))
            myStr = ""
        End If
    Next i
    Set AddStringToCollection = myCol
End Function
'---------------------------------------------
'Detect the selected range
'---------------------------------------------
Private Function CheckRange(ByVal oMode As Byte, ByVal oValue As Variant) As Boolean
    On Error Resume Next
        Select Case oMode
            Case 1
                'Check if selected range exists
                Range(oValue).Select
            Case 2
                'Check if selected column exists
                Range(oValue & 1).Select
            Case 3
                'Check if selected cell exists
                Rows(oValue).Select
        End Select
        'Detect errors and return values
        '1004:  Application-defined or object-defined error (for columns)
        '6:     Overflow (for rows)
        If Err.Number = 1004 Or Err.Number = 13 Or Err.Number = 6 Then CheckRange = False Else CheckRange = True
        Err.Clear
    On Error GoTo 0
End Function
'---------------------------------------------
'Get Range to String
'Thanks to: https://stackoverflow.com/questions/41777996/how-can-i-convert-a-range-to-a-string-vba
'---------------------------------------------
Function RangeToString(ByVal oRange As Range) As String
    RangeToString = ""
    If Not oRange Is Nothing Then
        Dim myCell As Range
        For Each myCell In oRange
            RangeToString = RangeToString & vbLf & myCell.Value
        Next myCell
        'Remove extra char
        RangeToString = Right(RangeToString, Len(RangeToString) - 1)
    End If
End Function
