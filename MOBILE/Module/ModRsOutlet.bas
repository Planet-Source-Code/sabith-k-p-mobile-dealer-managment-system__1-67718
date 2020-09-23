Attribute VB_Name = "ModRsOutlet"
'=======================
'+Author:   Sabith Kp  +
'=======================

Public Type aOutlet
   ID As String
   Name As String
   EmployName As String
   Place As String
   PhoneNo As String
   Date As String
End Type
Public Function Addoutlet(vOutlet As aOutlet) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddAgent = False
    
    sSQL = "SELECT * FROM tbloutlet WHERE ID='" & vOutlet.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        Addoutlet = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If Writeoutlet(vRS, vOutlet) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    Addoutlet = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function Editoutlet(vOutlet As aOutlet) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    Editoutlet = False
    
    sSQL = "SELECT * FROM tbloutlet WHERE ID= '" & vOutlet.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        GoTo RAE
    End If
    
    'edit
    If Writeoutlet(vRS, vOutlet) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    Editoutlet = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function Deleteoutlet(ByVal ioutletid As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblOutlet WHERE ID= '" & ioutletid & "'"

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
     
    Deleteoutlet = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function Writeoutlet(ByRef vRS As ADODB.Recordset, ByRef vOutlet As aOutlet) As Boolean
    
    'default
    Writeoutlet = False
    
    'On Error GoTo RAE

    With vOutlet
        vRS.Fields("Id") = .ID
        vRS.Fields("Name") = .Name
        vRS.Fields("EmployName") = .EmployName
        vRS.Fields("Place") = .Place
        vRS.Fields("Phone") = .PhoneNo
        vRS.Fields("oDate") = .Date
    End With

    Writeoutlet = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetoutletNo(soutletId As String, vOutlet As aOutlet) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetoutletNo = False
    
    sSQL = "SELECT * FROM tbloutlet WHERE Id='" & soutletId & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If Readoutlet(vRS, vOutlet) = False Then
        GoTo RAE
    End If
    
    GetoutletNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function Readoutlet(ByRef vRS As ADODB.Recordset, ByRef vOutlet As aOutlet) As Boolean
    
    'default
    Readoutlet = False
    
    On Error GoTo RAE
    
    With vOutlet
        .ID = ReadField(vRS.Fields("ID"))
        .Name = ReadField(vRS.Fields("Name"))
        .EmployName = ReadField(vRS.Fields("EmployName"))
        .PhoneNo = ReadField(vRS.Fields("Phone"))
        .Place = ReadField(vRS.Fields("Place"))
        .Date = ReadField(vRS.Fields("oDate"))
    End With
    
    Readoutlet = True
    Exit Function
    
RAE:
    
End Function






