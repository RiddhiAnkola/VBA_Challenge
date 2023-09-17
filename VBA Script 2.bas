Attribute VB_Name = "Module2"
Sub Main()
    Dim ws As Worksheet
    Sheets(2).Activate
    Set ws = ThisWorkbook.ActiveSheet
    
    Dim numRows As Long
    Dim i As Long
    Dim open_val As Double
    Dim close_val As Double
    Dim average_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Long
    
    Dim volume As Long
    
    numRows = ws.UsedRange.Rows.Count
    
    Dim randomStrings() As String
    randomStrings = GenerateDistinctRandomStrings(numRows - 1)
    
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Range("L:L").NumberFormat = "0"
    
    For i = 2 To numRows
    
        open_val = ws.Cells(i, "C").Value
        close_val = ws.Cells(i, "F").Value
        yearly_change = open_val - close_val
        average_price = (ws.Cells(i, "D").Value + ws.Cells(i, "E").Value) / 2
        volume = ws.Cells(i, "G").Value
        percent_change = (yearly_change / open_val)
        total_stock_volume = average_price * volume
        
        ws.Cells(i, "I").Value = randomStrings(i - 1)
        ws.Cells(i, "J").Value = yearly_change
        ws.Cells(i, "J").Interior.Color = IIf(yearly_change >= 0, RGB(0, 255, 0), RGB(255, 0, 0))
        ws.Cells(i, "K").Value = percent_change
        ws.Cells(i, "K").Interior.Color = IIf(percent_change >= 0, RGB(0, 255, 0), RGB(255, 0, 0))
        ws.Cells(i, "L").Value = total_stock_volume
    Next i
End Sub

Function GenerateDistinctRandomStrings(ByVal arraySize As Long) As String()
    Dim chars As String
    Dim result() As String
    Dim usedStrings() As String
    Dim i As Integer
    Dim generatedStr As String
    
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" ' Uppercase letters only
    ReDim result(1 To arraySize) ' Initialize the result array
    ReDim usedStrings(1 To arraySize) ' Initialize an array to keep track of used strings
    
    For i = 1 To arraySize ' Generate 'arraySize' number of strings
        Do
            generatedStr = ""
            Dim j As Integer
            For j = 1 To 3 ' Generate 3 characters for each string
                generatedStr = generatedStr & Mid(chars, Int((Len(chars) * Rnd) + 1), 1)
            Next j
        Loop While IsStringInArray(generatedStr, usedStrings) ' Check if the generated string is already used
        
        usedStrings(i) = generatedStr ' Add the generated string to the usedStrings array
        result(i) = generatedStr
    Next i
    
    GenerateDistinctRandomStrings = result
End Function

Function IsStringInArray(ByVal str As String, arr() As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(arr)
        If arr(i) = str Then
            IsStringInArray = True
            Exit Function
        End If
    Next i
    IsStringInArray = False
End Function




