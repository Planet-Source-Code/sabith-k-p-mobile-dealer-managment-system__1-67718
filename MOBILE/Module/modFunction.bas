Attribute VB_Name = "modFunction"
'=======================
'+Author:   Sabith Kp  +
'=======================

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function AnyRecordExisted(ByRef vRS As ADODB.Recordset) As Boolean
    If vRS.State = adStateClosed Then
        AnyRecordExisted = False
        Exit Function
    End If
    
    
    vRS.Requery
    
    If (vRS.BOF = True) And (vRS.EOF = True) Then
        AnyRecordExisted = False
    Else
        On Error GoTo errh
        vRS.MoveFirst
        AnyRecordExisted = True
    End If

    Exit Function
    '--------------------------
    
errh:
    AnyRecordExisted = False
End Function

Public Function IsEmpty(s As String) As Boolean
    If Len(Trim(s)) < 1 Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
End Function
Public Function HLTxt(ByRef txt As Object)
On Error Resume Next
    txt.SelStart = 0
    txt.SelLength = Len(txt)
    txt.SetFocus
End Function

Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
Dim tmp_listtview As ListItem
Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem + lvwText, lvwPartial, lvwPartial)
If Not tmp_listtview Is Nothing Then
    tmp_listtview.EnsureVisible
    tmp_listtview.Selected = True
End If
End Sub
Public Function CheckAsci(txt As TextBox, KeyAscii As Integer)
Select Case KeyAscii
    Case Asc(0) To Asc(9), vbKeyBack
    Case Else
        KeyAscii = 0
    End Select
End Function
Public Function ReadField(ByRef vField As Field) As Variant
    
    On Error GoTo errh

    If Not IsNull(vField.Value) Then
        ReadField = vField.Value
    Else
        Select Case vField.Type
            Case adBigInt
                ReadField = 0
            Case adBinary
                ReadField = 0
            Case adBoolean
                ReadField = False
            Case adByRef 'temp
                ReadField = 0
            Case adBSTR
                ReadField = ""
            Case adChar
                ReadField = ""
            Case adCurrency
                ReadField = 0
            Case adDate
                ReadField = CDate(0)
            Case adDBDate
                ReadField = CDate(0)
            Case adDBTime
                ReadField = FormatDateTime(CDate(0), vbLongTime)
            Case adDBTimeStamp
                ReadField = CDate(0)
            Case adDecimal
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case adEmpty 'temp
                ReadField = ""
            Case adError
                ReadField = 0
            
                
                
                
            Case adNumeric
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case Else
                ReadField = ""
            End Select
    End If
    
    Exit Function
    
errh:
    ReadField = ""
End Function
Public Function LoadCategory(frm As Form, cmb As ComboBox)
    On Error GoTo RAE
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim rec_count, X
    
    'set SQL Expression
   ' sSQL = "SELECT * From tblItem where ItemName=" & "'" & cmb.Text & "'" & ""
     sSQL = "SELECT * From tblCategory "   ''& " where" & " field = " & " '" & cmb.Text & "'" & ""
             
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
   With vRS
     .MoveLast
            rec_count = .RecordCount
    .MoveFirst
            For X = 1 To rec_count
                 cmb.AddItem !Name
    .MoveNext
            Next
End With

RAE:
    Set vRS = Nothing
      'listStudent.Refresh

End Function
Public Sub ComboBuilder(strTable As String, strFDescription As String, _
                 strFID As String, cboGeneric As ComboBox, Optional blnMoveToFirst As Boolean = True)
Dim rs As New ADODB.Recordset
Dim strSql As String

'On Error GoTo ErrHandle
On Error Resume Next
'Variables
    '---strTable the Table name
    '---strFDescription is the name of the field with the description
    '---strFID is the identity field
    '---cboGeneric is a combobox target
    '---blnMoveToFirst = True move for the first position


'TO CALL this procedure USE
'ComboBuilder "Table", "Description_Field", "Identity_Field", ComboBox
'
strSql = "Select " & strFID & ", " & strFDescription & " from " & strTable

cboGeneric.Clear

 If ConnectRS(PrimeDB, rs, strSql) = False Then
        'WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        'GoTo RAE
   End If
    
    Do Until rs.EOF
        cboGeneric.AddItem rs(1)
        cboGeneric.ItemData(cboGeneric.NewIndex) = rs(0)
        rs.MoveNext
    Loop
    
    rs.Close
    
    If blnMoveToFirst = True Then
        If cboGeneric.ListCount > 0 Then
            'cboGeneric.ListIndex = 0
        End If
    End If
    
    GoTo EXIT_
    

ErrHandle:
MsgBox "Error on Building the Combo Box " & cboGeneric.Name & ". " & Chr(10) & _
            Err.Description, vbExclamation, "ComboBuilder Procedure"
            
EXIT_:

End Sub

Public Function ComNumZ(ByVal vVal As Variant, ByVal iWidth As Integer) As String
    If Len(Trim(vVal)) > iWidth Then
        ComNumZ = CStr(vVal)
    Else
        ComNumZ = String$(iWidth - Len(Trim(vVal)), "0") & Trim(vVal)
    End If
End Function
Public Function ClearFlex(Flx As MSFlexGrid)

Dim j As Integer

     For j = 1 To Flx.Rows - 1


       If Flx.Rows > 2 Then
       Flx.RemoveItem (Flx.Row)
       End If

    Next j

  End Function




Public Function CheckAsci_ForSigns(txt As TextBox, KeyAscii As Integer)
Select Case KeyAscii
    Case Asc(0) To Asc(9), vbKeyBack
    Case Asc("A") To Asc("Z"), vbKeyBack
    Case Asc("a") To Asc("z"), vbKeyBack
    Case Asc(" "), vbKeyBack
    Case Else
        KeyAscii = 0
    End Select
End Function
Public Function ShowReport(tbl As String, Datarpt As DataReport, Optional DataCon As String = "")
    Dim vRS As New ADODB.Recordset
        Dim sSQL As String
    sSQL = "SELECT * FROM " & tbl & DataCon
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       MsgBox "clsReport" & "Report" & "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
    End If
        Set Datarpt.DataSource = vRS
End Function
