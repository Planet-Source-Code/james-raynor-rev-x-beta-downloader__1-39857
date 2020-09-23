VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REV-X - Configuratore"
   ClientHeight    =   2892
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8424
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   8424
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   672
      Left            =   72
      TabIndex        =   5
      Top             =   2184
      Width           =   8316
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   1536
         Top             =   300
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Help..."
         Height          =   372
         Left            =   96
         TabIndex        =   10
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnIndietro 
         Caption         =   "<== Indietro"
         Height          =   372
         Left            =   5556
         TabIndex        =   7
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnAvanti 
         Caption         =   "Avanti ==>"
         Height          =   372
         Left            =   6900
         TabIndex        =   6
         Top             =   216
         Width           =   1308
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1668
      Left            =   60
      TabIndex        =   4
      Top             =   516
      Width           =   8316
      Begin VB.TextBox txtDescrizione 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   756
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "Form2.frx":030A
         Top             =   732
         Width           =   6000
      End
      Begin VB.TextBox txtProgetto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1608
         TabIndex        =   0
         Text            =   "Nuovo progetto"
         Top             =   336
         Width           =   6564
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descrizione progetto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   876
         Left            =   108
         TabIndex        =   9
         Top             =   672
         Width           =   8112
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nome progetto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   108
         TabIndex        =   8
         Top             =   276
         Width           =   8112
      End
   End
   Begin VB.Frame Frame4 
      Height          =   444
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   8316
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Rev-X - Step 2: Descrizione progetto..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   72
         TabIndex        =   3
         Top             =   144
         Width           =   8184
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAvanti_Click()
nomeprogetto = txtProgetto.Text
descrizioneprogetto = txtDescrizione.Text
Me.Visible = False
Form3.Show
End Sub

Private Sub btnIndietro_Click()
Me.Visible = False: Form1.Show
End Sub

Private Sub Form_GotFocus()
txtProgetto.Text = nomeprogetto
txtDescrizione.Text = descrizioneprogetto


End Sub

Private Sub Form_Load()
txtProgetto.Text = nomeprogetto
txtDescrizione.Text = descrizioneprogetto
End Sub


Private Sub Form_Terminate()
On Error Resume Next
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload form9
Unload Form10
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload form9
Unload Form10

End Sub

Private Sub Timer1_Timer()
'txtProgetto.Text = nomeprogetto
'txtDescrizione.Text = descrizioneprogetto

End Sub
