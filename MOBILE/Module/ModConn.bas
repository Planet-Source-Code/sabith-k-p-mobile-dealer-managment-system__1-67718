Attribute VB_Name = "ModConn"
'=======================
'+Author:   Sabith Kp  +
'=======================

Public Const DBFileName = "Sim.mdb"
Public PrimeDB As New ADODB.Connection
Public DBPathFileName As String


'Check to the perticular file excist in that folder
Public Function InitDB() As Boolean
    
    Dim FSO As New FileSystemObject
    'default
    InitDB = False
   
   'check database file path
    If FSO.FileExists(App.Path & "\" & DBFileName) = False Then
        DBPathFileName = App.Path & "\" & DBFileName
        MsgBox "Error"
        GoTo RAE
    End If
    
    'set path file name
    DBPathFileName = App.Path & "\" & DBFileName

    'return true
    InitDB = True
    
RAE:
    Set FSO = Nothing
    
End Function

Public Function OpenDB() As Boolean

    OpenDB = False
    
    'open databse
    If ConnectDB(PrimeDB, DBPathFileName) = False Then
        'WriteErrorLog "modMain", "InitDB", "Unable to connect databse."
        GoTo RAE
    End If
    
    OpenDB = True
    
RAE:
End Function



Public Function ConnectDB(ByRef vDB As ADODB.Connection, PathFileName As String) As Boolean

On Error GoTo errh
 
    If vDB.State = adStateOpen Then vDB.Close
        
    vDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password=youarenotarealhacker"
    
    ConnectDB = True
    
    Exit Function
    
errh:

   MsgBox "modDBMain" & "," & "ConnectDB" & "," & Err.Description
    ConnectDB = False
    
End Function
Public Function CloseDB(ByRef vDB As ADODB.Connection)
    On Error GoTo errh
    vDB.Close
errh:
End Function


Public Function ConnectRS(ByRef vDB As ADODB.Connection, ByRef vRS As ADODB.Recordset, sSQL As String, Optional sHowMSG As Boolean = True, Optional ByRef iErrNumber As Variant, Optional ByRef sErrDescription As Variant) As Boolean
    
On Error GoTo errh

    
    Set vRS = Nothing
    Set vRS = New ADODB.Recordset
  
  
    vRS.Open sSQL, vDB, adOpenStatic, adLockOptimistic
    ConnectRS = True

    
    Exit Function
    
'-------------------------------------------
errh:
    If sHowMSG = True Then
        MsgBox "modDBMain" & "," & "ConnectRS" & "," & "Unable to connect Recordset / Err: " & Err.Description, vbExclamation
    End If
    If Not IsMissing(iErrNumber) Then
        iErrNumber = Err.Number
    End If
    If Not IsMissing(sErrDescription) Then
        sErrDescription = Err.Description
    End If
    ConnectRS = False
End Function

