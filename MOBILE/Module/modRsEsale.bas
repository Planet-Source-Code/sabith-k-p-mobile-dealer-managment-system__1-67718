Attribute VB_Name = "modRsEsale"
'=======================
'+Author:   Sabith Kp  +
'=======================


Public Type aESale
    ID As Long
    eDate As String
    OutletName As String
    Amount  As Currency
    FinalAmount As Currency
       
End Type
Public Function AddESale(vESale As aESale) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
   
    
    'default
    AddESale = False
    
    sSQL = "SELECT * FROM tblESale WHERE ID=" & vESale.ID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddESale = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteESale(vRS, vESale) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddESale = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditESale(vESale As aESale) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditESale = False
    
    sSQL = "SELECT * FROM tblESale WHERE ID= '" & vESale.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        GoTo RAE
    End If
    
    'edit
    If WriteESale(vRS, vESale) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditESale = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteESale(ByVal iESaleid As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblESale WHERE ID= '" & iESaleid & "'"

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
     
    DeleteESale = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function WriteESale(ByRef vRS As ADODB.Recordset, ByRef vESale As aESale) As Boolean
    
    'default
    WriteESale = False
    
    'On Error GoTo RAE

    With vESale
        vRS.Fields("Id") = .ID
        vRS.Fields("eDate") = .eDate
        vRS.Fields("Outletname") = .OutletName
        vRS.Fields("Amount") = .Amount
        vRS.Fields("Finalamount") = .FinalAmount
    End With

    WriteESale = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetESaleNo(sESaleId As String, vESale As aESale) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetESaleNo = False
    
    sSQL = "SELECT * FROM tblESale WHERE Id='" & sESaleId & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadESale(vRS, vESale) = False Then
        GoTo RAE
    End If
    
    GetESaleNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function ReadESale(ByRef vRS As ADODB.Recordset, ByRef vESale As aESale) As Boolean
    
    'default
    ReadESale = False
    
    On Error GoTo RAE
    
    With vESale
        .ID = ReadField(vRS.Fields("ID"))
        .eDate = ReadField(vRS.Fields("edate"))
        .OutletName = ReadField(vRS.Fields("outletname"))
        .Amount = ReadField(vRS.Fields("Amount"))
        .FinalAmount = ReadField(vRS.Fields("Finalamount"))
      End With
    
    ReadESale = True
    Exit Function
    
RAE:
    
End Function















