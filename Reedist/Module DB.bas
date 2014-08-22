Attribute VB_Name = "Module1"
Public db As ADODB.Connection
Public rs As ADODB.Recordset

Public Const constring = "provider=microsoft.jet.oledb.4.0;data source = C:\WINDOWS\Reedist\DB\lyb.mdb"

Public Sub connectdb()
    Set db = New ADODB.Connection
    db.Open (constring)
End Sub

Public Sub connecttable()
    Set rs = New ADODB.Recordset
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
End Sub
