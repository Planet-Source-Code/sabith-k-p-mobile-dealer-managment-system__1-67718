Attribute VB_Name = "modRsCategory"
'=======================
'+Author:   Sabith Kp  +
'=======================


Public Type aCategory
   ID As String
   Name As String
End Type
Public Function AddCategory(vCategory As aCategory) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String

    
    
    'default
    AddCategory = False
    
    sSQL = "SELECT * FROM tblCategory WHERE ID='" & vCategory.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "AddAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        'GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddCategory = True
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteCategory(vRS, vCategory) = False Then
        'GoTo RAE
    End If
    
    vRS.Update
   
    
    AddCategory = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function EditCategory(vCategory As aCategory) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditCategory = False
    
    sSQL = "SELECT * FROM tblCategory WHERE ID= '" & vCategory.ID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRAgent", "EditAgent", "AgentID does not exist.AgentID= " & vAgent.AgentID"
        GoTo RAE
    End If
    
    'edit
    If WriteCategory(vRS, vCategory) = False Then
        'GoTo RAE
    End If
    
    vRS.Update

    EditCategory = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function DeleteCategory(ByVal iCategoryid As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
        
    
    'On Error GoTo RAE
    'default
    DeleteAgent = False
    
    sSQL = "DELETE * FROM tblCategory WHERE ID= '" & iCategoryid & "'"

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
     
    DeleteCategory = True
    
'RAE:
 '   Set vRS = Nothing
End Function
Public Function WriteCategory(ByRef vRS As ADODB.Recordset, ByRef vCategory As aCategory) As Boolean
    
    'default
    WriteCategory = False
    
    'On Error GoTo RAE

    With vCategory
        vRS.Fields("Id") = .ID
        vRS.Fields("Name") = .Name
       
    End With

    WriteCategory = True
    Exit Function
    
'RAE:
 '  MsgBox Err.Address
End Function

Public Function GetCategoryNo(sCategoryId As String, vCategory As aCategory) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetCategoryNo = False
    
    sSQL = "SELECT * FROM tblCategory WHERE Id='" & sCategoryId & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadCategory(vRS, vCategory) = False Then
        GoTo RAE
    End If
    
    GetCategoryNo = True
    
RAE:
    Set vRS = Nothing
End Function
Public Function ReadCategory(ByRef vRS As ADODB.Recordset, ByRef vCategory As aCategory) As Boolean
    
    'default
    ReadCategory = False
    
    On Error GoTo RAE
    
    With vCategory
        .ID = ReadField(vRS.Fields("ID"))
        .Name = ReadField(vRS.Fields("Name"))
       
    End With
    
    ReadCategory = True
    Exit Function
    
RAE:
    
End Function












