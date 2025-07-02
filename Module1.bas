Attribute VB_Name = "Module1"
Option Explicit

Sub AddNumbersA()
'Place your code here
Dim x As Double, y As Double
x = InputBox("Enter a number:")
y = Range("D4").Value
Range("G12") = x + y
End Sub

Sub AddNumbersB()
'Place your code here
Dim x As Double, y As Double
x = InputBox("Enter a number:")
y = ActiveCell.Value
ActiveCell.Offset(-3, 2).Value = x + y
End Sub

Sub WherePutMe()
'Place your code here
Dim x As Double, y As String, z As Double
x = InputBox("Enter row number:")
y = InputBox("Enter column letter:")
z = Selection.Cells(2, 2).Value
Range(y & CStr(x)) = z
End Sub

Sub Swap()
'Place your code here
Dim x As Double, Temp As Double
Temp = Selection.Cells(1, 1).Value
Selection.Cells(1, 1) = Selection.Cells(1, 2).Value
Selection.Cells(1, 2) = Temp
End Sub
