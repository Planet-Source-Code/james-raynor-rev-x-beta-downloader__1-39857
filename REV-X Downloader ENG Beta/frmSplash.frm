VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2004
   LinkTopic       =   "Form8"
   ScaleHeight     =   600
   ScaleWidth      =   2004
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   324
      Top             =   888
   End
   Begin VB.Image Image1 
      Height          =   1788
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   5160
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

If Command <> "" Then startfile = Command

Me.Height = Image1.Height
Me.Width = Image1.Width
Form1.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\db2.mdb" & ";Jet OLEDB:Database Password=qwertyuiopè+;" & ";Persist Security Info=False"
Form6.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\db2.mdb" & ";Jet OLEDB:Database Password=qwertyuiopè+;" & ";Persist Security Info=False"
Form7.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\db2.mdb" & ";Jet OLEDB:Database Password=qwertyuiopè+;" & ";Persist Security Info=False"

Form1.Adodc1.RecordSource = "FTP1"
Form6.Adodc1.RecordSource = "FTP1"
Form7.Adodc1.RecordSource = "FTP1"

Form1.Adodc1.Refresh
Form6.Adodc1.Refresh
Form7.Adodc1.Refresh

Timer1.Enabled = True

End Sub

Private Sub Image1_Click()
Unload Me
Form1.Show
Form1.Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
Unload Me
Form1.Show
Form1.Timer3.Enabled = True
End Sub
