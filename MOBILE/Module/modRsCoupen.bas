Attribute VB_Name = "modRsCoupon"
'=======================
'+Author:   Sabith Kp  +
'=======================


Public Type aCoupon
    ID As String
    EmployName As String
    OutletName As String
    ProductName As String
    NoofCoupon  As Long
    Amount As Long
    eDate As String
   
End Type
Public Function AddCoupon(vCoupon As aCoupon) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddCoupon = False
    
    sSQL = "SELECT * FROM tblCoupon WHERE ID='" & vCoupon.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddCoupon = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteCoupon(vRS, vCoupon) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddCoupon = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditCoupon(vCoupon As aCoupon) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditCoupon = False
    
    sSQL = "SELECT * FROM tblCoupon WHERE ID= '" & vCoupon.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        GoTo RAE
    End If
    
    'edit
    If WriteCoupon(vRS, vCoupon) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditCoupon = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteCoupon(ByVal iCouponid As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblCoupon WHERE ID= '" & iCouponid & "'"

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
     
    DeleteCoupon = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function WriteCoupon(ByRef vRS As ADODB.Recordset, ByRef vCoupon As aCoupon) As Boolean
    
    'default
    WriteCoupon = False
    
    'On Error GoTo RAE

    With vCoupon
        vRS.Fields("Id") = .ID
        vRS.Fields("EmployName") = .EmployName
        vRS.Fields("OutletName") = .OutletName
        vRS.Fields("ProductName") = .ProductName
        vRS.Fields("NoofC") = .NoofCoupon
        vRS.Fields("Amount") = .Amount
        vRS.Fields("eDate") = .eDate
    End With

    WriteCoupon = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetCouponNo(sCouponId As String, vCoupon As aCoupon) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCouponNo = False
    
    sSQL = "SELECT * FROM tblCoupon WHERE Id='" & sCouponId & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCoupon(vRS, vCoupon) = False Then
        GoTo RAE
    End If
    
    GetCouponNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function ReadCoupon(ByRef vRS As ADODB.Recordset, ByRef vCoupon As aCoupon) As Boolean
    
    'default
    ReadCoupon = False
    
    On Error GoTo RAE
    
    With vCoupon
        .ID = ReadField(vRS.Fields("ID"))
        .EmployName = ReadField(vRS.Fields("EmployName"))
        .OutletName = ReadField(vRS.Fields("OutletName"))
        .ProductName = ReadField(vRS.Fields("ProductName"))
        .Amount = ReadField(vRS.Fields("Amount"))
        .NoofCoupon = ReadField(vRS.Fields("NoofC"))
                .eDate = ReadField(vRS.Fields("eDate"))
    End With
    
    ReadCoupon = True
    Exit Function
    
RAE:
    
End Function












