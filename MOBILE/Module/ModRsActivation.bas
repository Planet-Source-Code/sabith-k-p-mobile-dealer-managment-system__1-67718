Attribute VB_Name = "ModRsActivation"
'=======================
'+Author:   Sabith Kp  +
'=======================


Public Type aActivation
    ID As String
    CurDate As String
    aDate As Date
    MobileNo As String
    Attest As String
    Photograph As String
    PhotoId As String
    APEF As String
    Retailseal As String
    Distributer As String
    OutletName As String
    CoustomerName As String
    Address As String
    MeffDate As Date
    DeliveryDate As String
    CompeletWizard As Integer
End Type
Public Function AddActivation(vActivation As aActivation) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddActivation = False
    
    sSQL = "SELECT * FROM tblActivation WHERE ID='" & vActivation.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddActivation = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteActivation(vRS, vActivation) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddActivation = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditActivation(vActivation As aActivation) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditActivation = False
    
    sSQL = "SELECT * FROM tblActivation WHERE ID= '" & vActivation.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        GoTo RAE
    End If
    
    'edit
    If WriteActivation(vRS, vActivation) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditActivation = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteActivation(ByVal iActivationid As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblActivation WHERE ID= '" & iActivationid & "'"

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
     
    DeleteActivation = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function WriteActivation(ByRef vRS As ADODB.Recordset, ByRef vActivation As aActivation) As Boolean
    
    'default
    WriteActivation = False
    
    'On Error GoTo RAE

    With vActivation
        vRS.Fields("Id") = .ID
        vRS.Fields("CurDate") = .CurDate
        vRS.Fields("aDate") = .aDate
        vRS.Fields("mobileNo") = .MobileNo
        vRS.Fields("Attestphoto") = .Attest
        vRS.Fields("photograph") = .Photograph
        vRS.Fields("PhotoId") = .PhotoId
        vRS.Fields("CSinAPEF") = .APEF
        vRS.Fields("retailSeal") = .Retailseal
        vRS.Fields("DistributerSeal") = .Distributer
        vRS.Fields("Outletname") = .OutletName
        vRS.Fields("CustomerName") = .CoustomerName
        vRS.Fields("Address") = .Address
        vRS.Fields("meffDate") = .MeffDate
        vRS.Fields("deliverydate") = .DeliveryDate
        vRS.Fields("Complete") = .CompeletWizard
    End With

    WriteActivation = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetActivationNo(sActivationNo As String, vActivation As aActivation) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetActivationNo = False
    
    sSQL = "SELECT * FROM tblActivation WHERE MobileNo='" & sActivationNo & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadActivation(vRS, vActivation) = False Then
        GoTo RAE
    End If
    
    GetActivationNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function ReadActivation(ByRef vRS As ADODB.Recordset, ByRef vActivation As aActivation) As Boolean
    
    'default
    ReadActivation = False
    
    On Error GoTo RAE
    
    With vActivation
        
    With vActivation
      .ID = vRS.Fields("Id")
      .CurDate = vRS.Fields("CurDate")
        .aDate = vRS.Fields("adate")
        .MobileNo = vRS.Fields("MobileNo")
        .Attest = vRS.Fields("Attestphoto")
        .Photograph = vRS.Fields("Photograph")
        .PhotoId = vRS.Fields("photoid")
        .APEF = vRS.Fields("CSinAPEF")
        .Retailseal = vRS.Fields("RetailSeal")
        .Distributer = vRS.Fields("DistributerSeal")
        .OutletName = vRS.Fields("Outletname")
        .CoustomerName = vRS.Fields("CustomerName")
        .Address = vRS.Fields("Address")
        .MeffDate = vRS.Fields("meffDate")
        .DeliveryDate = vRS.Fields("Deliverydate")
        .CompeletWizard = vRS.Fields("Complete")
    
    End With
    End With
    
    ReadActivation = True
    Exit Function
    
RAE:
    
End Function













