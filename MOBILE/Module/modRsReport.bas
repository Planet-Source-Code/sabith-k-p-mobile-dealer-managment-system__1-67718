Attribute VB_Name = "modRsReport"
'=======================
'+Author:   Sabith Kp  +
'=======================
Public Function GetMont(sMonth As Integer, SelMonth As String)
Select Case sMonth
Case 1
SelMonth = "JAN"

Case 2
SelMonth = "FEB"

Case 3
SelMonth = "MAR"

Case 4
SelMonth = "APR"

Case 5
SelMonth = "MAY"

Case 6
SelMonth = "JUN"

Case 7
SelMonth = "JUL"

Case 8
SelMonth = "AUG"

Case 9
SelMonth = "SEP"

Case 10
SelMonth = "OCT"

Case 11
SelMonth = "NOV"

Case 12
SelMonth = "DEC"

End Select
End Function


