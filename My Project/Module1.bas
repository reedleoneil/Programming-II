Attribute VB_Name = "Module1"
Public db As ADODB.Connection
Public rs As ADODB.Recordset

Public Const constring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Ferddie\Desktop\My Project\lyb.mdb"
Public msg As String

Public Sub ConnectDB()
    Set db = New ADODB.Connection
    db.Open (constring)
End Sub

Public Sub ConnectTable()
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    
End Sub
