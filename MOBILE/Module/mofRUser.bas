Attribute VB_Name = "modRSUser"
'=======================
'+Author:   Sabith Kp  +
'=======================
Option Explicit

Public Type tUser
    
    UserID As String
    Password As String
    CreationDate As Date
    ModifiedDate As Date
    CreatedBy As String
    ModifiedBy As String
    
End Type


Public Function AddUser(vUser As tUser) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddUser = False
    
    sSQL = "SELECT * FROM tblUser WHERE UserID='" & vUser.UserID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
       ' WriteErrorLog "modUser", "AddUser", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        MsgBox "Invalid User ID. It is already existed in record.", vbExclamation
        GoTo RAE
    End If
    
    'add new record
    vRS.AddNew
    
    If WriteUser(vRS, vUser) = False Then
       GoTo RAE
    End If
    
    vRS.Update
   
    
    AddUser = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditUser(vUser As tUser) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditUser = False
    
    sSQL = "SELECT * FROM tblUser WHERE UserID='" & vUser.UserID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
      '  WriteErrorLog "modUser", "EditUser", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "Invalid User ID. It does not exist in record.", vbExclamation
        GoTo RAE
    End If
    
    'edit
    
    If WriteUser(vRS, vUser) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    EditUser = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteUser(ByVal UserID As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    DeleteUser = False
    
    sSQL = "DELETE* FROM tblUser WHERE UserID='" & UserID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "DeleteUser", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    DeleteUser = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function GetUserByID(sUserID As String, vUser As tUser) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetUserByID = False
    
    sSQL = "SELECT * FROM tblUser WHERE UserID='" & sUserID & "'"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "GetUserByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadUser(vRS, vUser) = False Then
        GoTo RAE
    End If
    
    GetUserByID = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnyUserExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyUserExist = False
    
    sSQL = "SELECT * FROM tblUser"
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        'WriteErrorLog "modUser", "AnyUserExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyUserExist = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function ReadUser(ByRef vRS As ADODB.Recordset, ByRef vUser As tUser) As Boolean
    
    'default
    ReadUser = False
    
    On Error GoTo RAE
    
    With vUser
        .UserID = ReadField(vRS.Fields("UserID"))
        .Password = ReadField(vRS.Fields("Password"))
        .CreationDate = ReadField(vRS.Fields("CreationDate"))
        .ModifiedDate = ReadField(vRS.Fields("ModifiedDate"))
        .CreatedBy = ReadField(vRS.Fields("CreatedBy"))
        .ModifiedBy = ReadField(vRS.Fields("ModifiedBy"))
    End With
    
    ReadUser = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteUser(ByRef vRS As ADODB.Recordset, ByRef vUser As tUser) As Boolean
    
    'default
    WriteUser = False
    
    On Error GoTo RAE
    
    With vUser
        vRS.Fields("UserID") = .UserID
        vRS.Fields("Password") = .Password
        vRS.Fields("CreationDate") = .CreationDate
        vRS.Fields("ModifiedDate") = .ModifiedDate
        vRS.Fields("CreatedBy") = .CreatedBy
        vRS.Fields("ModifiedBy") = .ModifiedBy
    End With
    
    WriteUser = True
    Exit Function
    
RAE:
    
End Function

