Attribute VB_Name = "modRsProduct"
'=======================
'+Author:   Sabith Kp  +
'=======================

Public Type aProduct
   ID As String
   Name As String
   Catogory As String
   Amount As String
   PurchaseRate  As Long
      oDate As String
End Type
Public Function AddProduct(vProduct As aProduct) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddProduct = False
    
    sSQL = "SELECT * FROM tblProduct WHERE ID='" & vProduct.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddProduct = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteProduct(vRS, vProduct) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddProduct = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditProduct(vProduct As aProduct) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditProduct = False
    
    sSQL = "SELECT * FROM tblProduct WHERE ID= '" & vProduct.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        'GoTo RAE
    End If
    
    'edit
    If WriteProduct(vRS, vProduct) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditProduct = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteProduct(ByVal iProductid As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblProduct WHERE ID='" & iProductid & "'"

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
     
    DeleteProduct = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function WriteProduct(ByRef vRS As ADODB.Recordset, ByRef vProduct As aProduct) As Boolean
    
    'default
    WriteProduct = False
    
    'On Error GoTo RAE

    With vProduct
        vRS.Fields("Id") = .ID
        vRS.Fields("Name") = .Name
        vRS.Fields("Category") = .Catogory
        vRS.Fields("prValue") = .PurchaseRate
        vRS.Fields("Amount") = .Amount
        vRS.Fields("oDate") = .oDate
    End With

    WriteProduct = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetProductNo(sProductId As String, vProduct As aProduct) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetProductNo = False
    
    sSQL = "SELECT * FROM tblProduct WHERE Id='" & sProductId & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadProduct(vRS, vProduct) = False Then
        GoTo RAE
    End If
    
    GetProductNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function ReadProduct(ByRef vRS As ADODB.Recordset, ByRef vProduct As aProduct) As Boolean
    
    'default
    ReadProduct = False
    
    On Error GoTo RAE
    
    With vProduct
        .ID = ReadField(vRS.Fields("ID"))
        .Name = ReadField(vRS.Fields("Name"))
        .Catogory = ReadField(vRS.Fields("Category"))
        .PurchaseRate = ReadField(vRS.Fields("prvalue"))
        .Amount = ReadField(vRS.Fields("Amount"))
        .oDate = ReadField(vRS.Fields("oDate"))
    End With
    
    ReadProduct = True
    Exit Function
    
RAE:
    
End Function


Public Function GetProductName(SProductName As String, vProduct As aProduct) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetProductName = False
    
    sSQL = "SELECT * FROM tblProduct WHERE Name='" & SProductName & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadProduct(vRS, vProduct) = False Then
        GoTo RAE
    End If
    
    GetProductName = True
    
RAE:
    Set vRS = Nothing
End Function





