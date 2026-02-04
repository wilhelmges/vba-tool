VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDbSession
Private mConn As ADODB.Connection

Private Function OpenFESConnection() As ADODB.Connection
    On Error GoTo ErrHandler
    
    Dim sConnString As String
    sConnString = "Provider=SQLOLEDB;" & _
                      "Data Source=DellServer;" & _
                      "Initial Catalog=FES;" & _
                      "User ID=FES_R;" & _
                      "Password=Fes_User4007;"
        
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    conn.Open sConnString
    Set OpenFESConnection = conn
    Exit Function
    
ErrHandler:
    Err.Raise Err.number, "OpenFESConnection", Err.Description

End Function

Private Sub Class_Initialize()
    Set mConn = OpenFESConnection()
End Sub

Public Property Get OpenConnection() As ADODB.Connection
    Set Connection = mConn
End Property

Private Sub Class_Terminate()
    If Not mConn Is Nothing Then
        If mConn.State = adStateOpen Then mConn.Close
    End If
End Sub


