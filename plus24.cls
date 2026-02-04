VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "plus24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Class Module plus24
Public pId As Long

Public Sub LoadFromRecordset(ByVal rs As Object)
    ' rs = ADODB.Recordset

    If rs Is Nothing Then Err.Raise 5, , "Recordset is Nothing"
    If rs.EOF Then Err.Raise 5, , "Recordset is EOF"

    pDocDate = rs.Fields("DocDate").Value
    pFuelCode = rs.Fields("FuelCode").Value
    pQuantity = nz(rs.Fields("Quantity").Value, 0)
    pAmount = nz(rs.Fields("Amount").Value, 0)
End Sub
