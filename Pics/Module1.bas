Attribute VB_Name = "Module1"
Public db As New ADODB.Connection
Public rs As New ADODB.Recordset
Public msstream As New ADODB.Stream
Public isEdit As Boolean
Public cn, msg As String

Public Sub InitDB()
    cn = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" & App.Path & "\pics.mdb"

    Set db = New ADODB.Connection
    db.Open (cn)

End Sub


Public Sub InitRS()
    Set rs = New ADODB.Recordset
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
End Sub
