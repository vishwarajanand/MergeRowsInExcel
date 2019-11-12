Option Explicit


Sub MergeRows()
'
' MergeRows Macro
'

' Declare vars for total row/col and num rows in original data per final data
Dim tR, tC, nRPer As Integer
nRPer = Range("B1").Value
tR = Range("B4").Value
tC = Range("B5").Value

' Declare vars for Original/Final data source row/col
Dim Odsr, Odsc, Fdsr, Fdsc As Integer
Odsr = Range(Range("B2").Value).Row
Odsc = Range(Range("B2").Value).Column
Fdsr = Range(Range("B3").Value).Row
Fdsc = Range(Range("B3").Value).Column

' Declare vars for Original data relative row/col
Dim Odrr, Odrc As Integer

' Declare source/final row/col
Dim Sr, Sc, Fr, Fc As Integer
For Odrr = 0 To (tR - 1)
    For Odrc = 0 To (tC - 1)
        Sr = Odsr + Odrr
        Sc = Odsc + Odrc
        Fr = Fdsr + (Odrr \ nRPer)
        Fc = (Fdsc + Odrc + (Odrr Mod nRPer) * nRPer)
        
        ' get OD
        ' Debug.Print "Reading from: (" & Odrr & "," & Odrc & ")"
        ' Debug.Print "Value found-> " & Cells(Odsr + Odrr, Odsc + Odrc).Value
        ' get FD
        ' Debug.Print "Save Location: (" & Fdsr + (Odrr \ nRPer) & "," & Fdsc + (Odrc + (Odrr Mod nRPer) * nRPer) & ")"
        ' set FD
        Cells(Fr, Fc).Value = Cells(Sr, Sc).Value
        
    Next Odrc
Next Odrr


'MsgBox ("Data Merge is completed")

End Sub
