VERSION 5.00
Begin VB.Form form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REV-X - Configuratore"
   ClientHeight    =   3144
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8424
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   3144
   ScaleWidth      =   8424
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   444
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   8316
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Rev-X - Step 8: Configura opzioni..."
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
         TabIndex        =   6
         Top             =   144
         Width           =   8184
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1836
      Left            =   60
      TabIndex        =   4
      Top             =   516
      Width           =   8316
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Height          =   264
         Left            =   6228
         TabIndex        =   14
         Top             =   1464
         Value           =   1  'Checked
         Width           =   192
      End
      Begin VB.TextBox directory 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3984
         TabIndex        =   13
         Text            =   "Img"
         Top             =   1068
         Width           =   2088
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Height          =   264
         Left            =   6228
         TabIndex        =   12
         Top             =   1056
         Value           =   1  'Checked
         Width           =   192
      End
      Begin VB.TextBox estensione 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4692
         TabIndex        =   9
         Text            =   "Jpg"
         Top             =   672
         Width           =   1716
      End
      Begin VB.TextBox fileseq 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4236
         TabIndex        =   7
         Text            =   "Thumbnail"
         Top             =   264
         Width           =   2172
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Falso Html ""Default.html"""
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
         Left            =   2076
         TabIndex        =   15
         Top             =   1428
         Width           =   4392
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Directory da creare:"
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
         Left            =   2076
         TabIndex        =   11
         Top             =   1020
         Width           =   4392
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Estensione file sequenziale:"
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
         Left            =   2076
         TabIndex        =   10
         Top             =   612
         Width           =   4392
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nome file sequenziale:"
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
         Left            =   2076
         TabIndex        =   8
         Top             =   204
         Width           =   4392
      End
   End
   Begin VB.Frame Frame1 
      Height          =   672
      Left            =   60
      TabIndex        =   0
      Top             =   2364
      Width           =   8316
      Begin VB.CommandButton btnAvanti 
         Caption         =   "Avanti ==>"
         Height          =   372
         Left            =   6900
         TabIndex        =   3
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnIndietro 
         Caption         =   "<== Indietro"
         Height          =   372
         Left            =   5556
         TabIndex        =   2
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Help..."
         Height          =   372
         Left            =   96
         TabIndex        =   1
         Top             =   216
         Width           =   1308
      End
   End
End
Attribute VB_Name = "form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAvanti_Click()
nomefilesequenziale = fileseq.Text
estensionefilesequenziale = estensione.Text
directorydacreare = directory.Text
ocio = False
Me.Visible = False
Form10.Show
End Sub

Private Sub btnIndietro_Click()
nomefilesequenziale = fileseq.Text
estensionefilesequenziale = estensione.Text
directorydacreare = directory.Text

Me.Visible = False
Form7.Show
End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then okCrearedirectory = True Else okCrearedirectory = False

End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then falsohtml = True Else falsohtml = False

End Sub

Private Sub Form_Load()
'okCrearedirectory = True
'falsohtml = True
If nomefilesequenziale = "" Then nomefilesequenziale = "Thumbnail"
fileseq.Text = nomefilesequenziale

If estensionefilesequenziale = "" Then estensionefilesequenziale = "Jpg"
estensione.Text = estensionefilesequenziale

If directorydacreare = "" Then directorydacreare = "Img"
directory.Text = directorydacreare


Check1.value = 1
Check2.value = 1

If Check1.value = 1 Then okCrearedirectory = True Else okCrearedirectory = False
If Check2.value = 1 Then falsohtml = True Else falsohtml = False


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
