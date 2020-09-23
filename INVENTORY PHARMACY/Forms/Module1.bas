Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rsbill As New ADODB.Recordset
Public Constring As String
Dim btb As String
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Sub Main()
Constring = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Pharmacy1.mdb"

'btb = "\\ADESINA-PC\PharmMan\Pharmacy1.mdb\"
'C:\Users\ADESINA\Desktop\VB Projects\PharmMan\Project\Pharmacy1.mdb

'Constring = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & btb & ""
End Sub



