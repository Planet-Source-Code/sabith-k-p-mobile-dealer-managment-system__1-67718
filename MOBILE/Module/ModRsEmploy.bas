Attribute VB_Name = "ModRsEmploy"
'=======================
'+Author:   Sabith Kp  +
'=======================


Public Type aEmploy
   ID As String
   Name As String
   Address As String
   eDate As String
End Type
Public Function AddEmploy(vEmploy As aEmploy) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddEmploy = False
    
    sSQL = "SELECT * FROM tblEmploy WHERE ID='" & vEmploy.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddEmploy = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteEmploy(vRS, vEmploy) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddEmploy = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditEmploy(vEmploy As aEmploy) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditEmploy = False
    
    sSQL = "SELECT * FROM tblEmploy WHERE ID= '" & vEmploy.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        GoTo RAE
    End If
    
    'edit
    If WriteEmploy(vRS, vEmploy) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditEmploy = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteEmploy(ByVal iEmployid As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteEmploy = False
    
    sSQL = "DELETE * FROM tblEmploy WHERE ID= '" & iEmployid & "'"

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
     
    DeleteEmploy = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function WriteEmploy(ByRef vRS As ADODB.Recordset, ByRef vEmploy As aEmploy) As Boolean
    
    'default
    WriteEmploy = False
    
    'On Error GoTo RAE

    With vEmploy
        vRS.Fields("Id") = .ID
        vRS.Fields("Name") = .Name
        vRS.Fields("Address") = .Address
        vRS.Fields("eDate") = .eDate
    End With

    WriteEmploy = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetEmployNo(sEmployId As String, vEmploy As aEmploy) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetEmployNo = False
    
    sSQL = "SELECT * FROM tblEmploy WHERE Id='" & sEmployId & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadEmploy(vRS, vEmploy) = False Then
        GoTo RAE
    End If
    
    GetEmployNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function ReadEmploy(ByRef vRS As ADODB.Recordset, ByRef vEmploy As aEmploy) As Boolean
    
    'default
    ReadEmploy = False
    
    On Error GoTo RAE
    
    With vEmploy
        .ID = ReadField(vRS.Fields("ID"))
        .Name = ReadField(vRS.Fields("Name"))
        .Address = ReadField(vRS.Fields("Address"))
        .eDate = ReadField(vRS.Fields("eDate"))
    End With
    
    ReadEmploy = True
    Exit Function
    
RAE:
    
End Function










