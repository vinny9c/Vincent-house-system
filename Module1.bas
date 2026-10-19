Attribute VB_Name = "Module1"
Public con As ADODB.Connection
Public loggedinuser As String

Public Sub connectDB()

    Set con = New ADODB.Connection

    con.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=C:\Users\summit\Desktop\House rental\Management.mdb;"

    con.Open

End Sub







