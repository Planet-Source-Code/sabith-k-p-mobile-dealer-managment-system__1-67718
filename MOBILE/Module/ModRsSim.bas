Attribute VB_Name = "ModRsSim"
'=======================
'+Author:   Sabith Kp  +
'=======================


Public Type pSim
    ID As String
    EmployName As String
    OutletName As String
    ProductName As String
    NoofpSim As Long
    Amount As Long
    eDate As String
   
End Type
Public Function AddpSim(vpSim As pSim) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddpSim = False
    
    sSQL = "SELECT * FROM tblSim WHERE ID='" & vpSim.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddpSim = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WritepSim(vRS, vpSim) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddpSim = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EdipSim(vpSim As pSim) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EdipSim = False
    
    sSQL = "SELECT * FROM tblSim WHERE ID= '" & vpSim.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        GoTo RAE
    End If
    
    'edit
    If WritepSim(vRS, vpSim) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EdipSim = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeletepSim(ByVal ippSim As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblSim WHERE ID= '" & ippSim & "'"

    Dim sErrD As String
    Dim iErrN As Long
    If ConnectRS(PrimeDB, vRS, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
            'WriteErrorLog "modRAgent", "DeleteAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            'GoTo RAE
        End If
    End If
     
    DeletepSim = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function WritepSim(ByRef vRS As ADODB.Recordset, ByRef vpSim As pSim) As Boolean
    
    'default
    
    WritepSim = False
    
    'On Error GoTo RAE

    With vpSim
        vRS.Fields("Id") = .ID
        vRS.Fields("EmployName") = .EmployName
        vRS.Fields("OutletName") = .OutletName
        vRS.Fields("ProductName") = .ProductName
        vRS.Fields("NoofSim") = .NoofpSim
        vRS.Fields("Amount") = .Amount
        vRS.Fields("eDate") = .eDate
    End With

    WritepSim = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetpSimNo(sppSim As String, vpSim As pSim) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetpSimNo = False
    
    sSQL = "SELECT * FROM tblSim WHERE Id='" & sppSim & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadpSim(vRS, vpSim) = False Then
        GoTo RAE
    End If
    
    GetpSimNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function ReadpSim(ByRef vRS As ADODB.Recordset, ByRef vpSim As pSim) As Boolean
    
    'default
    ReadpSim = False
    
    On Error GoTo RAE
    
    With vpSim
        .ID = ReadField(vRS.Fields("ID"))
        .EmployName = ReadField(vRS.Fields("EmployName"))
        .OutletName = ReadField(vRS.Fields("OutletName"))
        .ProductName = ReadField(vRS.Fields("ProductName"))
        .Amount = ReadField(vRS.Fields("Amount"))
        .NoofpSim = ReadField(vRS.Fields("NoofSim"))
                .eDate = ReadField(vRS.Fields("eDate"))
    End With
    
    ReadpSim = True
    Exit Function
    
RAE:
    
End Function














