Attribute VB_Name = "Module1"
Public Dtbs As ADODB.Connection
Public RcrdSt As ADODB.Recordset
Public ImgStream As ADODB.Stream

Public msg As String

Public Sub CnnctDtbs()
    Set Dtbs = New ADODB.Connection
    Dtbs.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\Database.mdb")
End Sub

Public Sub CnnctRcrdSt()
    Set RcrdSt = New ADODB.Recordset
    RcrdSt.CursorLocation = adUseClient
    RcrdSt.CursorType = adOpenDynamic
    RcrdSt.LockType = adLockOptimistic
End Sub

Public Sub ImgStrm()
    Set ImgStream = New ADODB.Stream
    ImgStream.Type = adTypeBinary
    ImgStream.Open
End Sub
