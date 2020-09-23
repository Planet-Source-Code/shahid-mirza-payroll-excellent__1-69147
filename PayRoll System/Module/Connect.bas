Attribute VB_Name = "Connect"
Public Sub Main()
    DB_Conect.Open "DSN=Pay Roll", Admin, "G9194556", -1
    Load frmIndependance: frmIndependance.Show
End Sub

