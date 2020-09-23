Attribute VB_Name = "modRsEPurchase"



Public Type aEPurchase
    ID As Long
    eDate As String
    Amount  As Currency
    FinalAmount As Currency
    Eval As Currency
   
End Type
Public Function AddEPurchase(vEPurchase As aEPurchase) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
   
    
    'default
    AddEPurchase = False
    
    sSQL = "SELECT * FROM tblEPurchase WHERE ID=" & vEPurchase.ID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddEPurchase = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteEPurchase(vRS, vEPurchase) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddEPurchase = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditEPurchase(vEPurchase As aEPurchase) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditEPurchase = False
    
    sSQL = "SELECT * FROM tblEPurchase WHERE ID= '" & vEPurchase.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        GoTo RAE
    End If
    
    'edit
    If WriteEPurchase(vRS, vEPurchase) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditEPurchase = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteEPurchase(ByVal iEPurchaseid As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblEPurchase WHERE ID= '" & iEPurchaseid & "'"

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
     
    DeleteEPurchase = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function WriteEPurchase(ByRef vRS As ADODB.Recordset, ByRef vEPurchase As aEPurchase) As Boolean
    
    'default
    WriteEPurchase = False
    
    'On Error GoTo RAE

    With vEPurchase
        vRS.Fields("Id") = .ID
        vRS.Fields("eDate") = .eDate
        vRS.Fields("Amount") = .Amount
        vRS.Fields("EVal") = .Eval
        vRS.Fields("Finalamount") = .FinalAmount
    End With

    WriteEPurchase = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetEPurchaseNo(sEPurchaseId As String, vEPurchase As aEPurchase) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetEPurchaseNo = False
    
    sSQL = "SELECT * FROM tblEPurchase WHERE Id='" & sEPurchaseId & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadEPurchase(vRS, vEPurchase) = False Then
        GoTo RAE
    End If
    
    GetEPurchaseNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function ReadEPurchase(ByRef vRS As ADODB.Recordset, ByRef vEPurchase As aEPurchase) As Boolean
    
    'default
    ReadEPurchase = False
    
    On Error GoTo RAE
    
    With vEPurchase
        .ID = ReadField(vRS.Fields("ID"))
        .eDate = ReadField(vRS.Fields("edate"))
        .Amount = ReadField(vRS.Fields("amount"))
        .Eval = ReadField(vRS.Fields("Eval"))
        .FinalAmount = ReadField(vRS.Fields("Finalamount"))
      End With
    
    ReadEPurchase = True
    Exit Function
    
RAE:
    
End Function













