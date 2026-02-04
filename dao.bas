Attribute VB_Name = "Dao"
'Infrastructure / Data Access layer
Public Sub TestCon()

Dim db As clsDbSession
Set db = New clsDbSession


' робота з БД через db.Connection

Set db = Nothing   ' < тут гарантовано закриється конекшен
Debug.Print "done with connection"

End Sub

Sub TestRS()
Sql = TemplateSql( _
    "WHERE (IPN = {IPN} AND Period = {Period})", _
    "IPN", "1234567890", _
    "Period", 202401 _
)
Debug.Print (Sql)

End Sub

Public Function TemplateSql(ByVal template As String, ParamArray args() As Variant) As String
    Dim i As Long
    Dim key As String
    Dim value As Variant
    Dim result As String
    
    result = template
    
    If (UBound(args) + 1) Mod 2 <> 0 Then
        Err.Raise vbObjectError + 1, "TemplateSql", _
                  "Кількість аргументів має бути парною (ключ, значення)"
    End If
    
    For i = LBound(args) To UBound(args) Step 2
        key = CStr(args(i))
        value = args(i + 1)
        
        result = Replace(result, "{" & key & "}", SqlValue(value))
    Next i
    
    TemplateSql = result
End Function

Private Function SqlValue(v As Variant) As String
    Select Case True
        Case IsNull(v)
            SqlValue = "NULL"
        Case VarType(v) = vbString
            SqlValue = "'" & Replace(v, "'", "''") & "'"
        Case VarType(v) = vbDate
            SqlValue = "'" & format(v, "yyyy-mm-dd") & "'"
        Case Else
            SqlValue = CStr(v)
    End Select
End Function

