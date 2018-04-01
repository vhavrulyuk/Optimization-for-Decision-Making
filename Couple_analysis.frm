VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Couple_analysis 
   Caption         =   "Вхідні дані"
   ClientHeight    =   3000
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5196
   OleObjectBlob   =   "Couple_analysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Couple_analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lastDataIndex As Integer

Private Sub CommandButton1_Click()
Dim alternatives As Integer
Call FillVertical
Call FillHorizontal
Call PairCompare
Call RangSum
Call FinalRang
End Sub

Sub FillVertical()
Sheets("PairComparison").UsedRange.ClearContents
For i = 1 To SpinButton1.Value
Range("A" & i + 2).Value = "A" & i
Range("B" & i + 2).Value = Controls("TextBox" & i).Value
Next i
End Sub

Sub FillHorizontal()
For i = 1 To SpinButton1.Value
Cells(1, i + 2).Value = "A" & i
Cells(2, i + 2).Value = Controls("TextBox" & i).Value
Next i
lastDataIndex = i + 2
End Sub

Private Sub PairCompare()
For i = 3 To SpinButton1.Value + 2
For j = 3 To SpinButton1.Value + 2
    If Cells(i, 2).Value >= Cells(2, j).Value Then
        Cells(i, j).Value = 1
    Else: Cells(i, j).Value = 0
    End If
Next j
Next i
End Sub

Private Sub RangSum()
Cells(SpinButton1.Value + 3, 2).Value = "Сума рангів"
For j = 3 To SpinButton1.Value + 2
Cells(SpinButton1.Value + 3, j).Formula = "=SUM(" & Range(Cells(3, j), Cells(lastDataIndex - 1, j)).Address & ")"
Next j
End Sub

Private Sub FinalRang()
Cells(SpinButton1.Value + 4, 2).Value = "Кінцевий ранг"
For j = 3 To SpinButton1.Value + 2
Cells(SpinButton1.Value + 4, j).Formula = "=RANK(" & Cells(SpinButton1.Value + 3, j).Address & "," & Range(Cells(SpinButton1.Value + 3, 3), Cells(SpinButton1.Value + 3, lastDataIndex - 1)).Address & ",1)" '"=RANK(R[-1]C,R[-1]C[& 2*j -2 &]:R[-1]C" & SpinButton1.Value & ",1)"
Next j


End Sub

Private Sub SpinButton1_SpinDown()
Count.Value = Format(0, "0")
Count.Value = Count.Value + SpinButton1.Value
Count.Value = Format(Count.Value, "0")
Controls("TextBox" & SpinButton1.Value + 1).Visible = False
End Sub

Private Sub SpinButton1_SpinUp()
Count.Value = Format(0, "0")
Count.Value = Count.Value + SpinButton1.Value
Count.Value = Format(Count.Value, "0")
Controls("TextBox" & SpinButton1.Value).Visible = True
End Sub

Private Sub BubbleSort(arr)
  Dim strTemp As String
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  lngMin = LBound(arr)
  lngMax = UBound(arr)
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If arr(i) > arr(j) Then
        strTemp = arr(i)
        arr(i) = arr(j)
        arr(j) = strTemp
      End If
    Next j
  Next i
End Sub

Private Function WhereInArray(rangArray As Variant, rang As Variant) As Variant
Dim i As Long
For i = LBound(rangArray) To UBound(rangArray)
    If rangArray(i) = rang Then
        WhereInArray = i
        Exit Function
    End If
Next i
'if you get here, vFind was not in the array. Set to null
WhereInArray = Null
End Function

Private Sub UserForm_Click()

End Sub
