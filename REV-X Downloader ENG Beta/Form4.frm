VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REV-X - Configuratore"
   ClientHeight    =   2460
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8400
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   672
      Left            =   60
      TabIndex        =   8
      Top             =   1668
      Width           =   8316
      Begin VB.TextBox txtLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   276
         Left            =   3324
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Attendere...."
         Top             =   252
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   1560
         Top             =   264
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "Help..."
         Height          =   372
         Left            =   96
         TabIndex        =   9
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnIndietro 
         Caption         =   "<== Indietro"
         Height          =   372
         Left            =   5556
         TabIndex        =   4
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnAvanti 
         Caption         =   "Avanti ==>"
         Enabled         =   0   'False
         Height          =   372
         Left            =   6900
         TabIndex        =   3
         Top             =   216
         Width           =   1308
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1164
      Left            =   60
      TabIndex        =   7
      Top             =   516
      Width           =   8316
      Begin VB.CommandButton btnSplitta 
         Caption         =   "Splitta!!"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   6972
         TabIndex        =   2
         Top             =   696
         Width           =   1200
      End
      Begin VB.CommandButton btnSalvaDove 
         Caption         =   "Salva dove..."
         Height          =   324
         Left            =   6972
         TabIndex        =   1
         Top             =   216
         Width           =   1200
      End
      Begin VB.TextBox txtnumerofiles 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   744
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4044
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Kb"
         Top             =   756
         Width           =   216
      End
      Begin VB.TextBox txtsplit 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3444
         TabIndex        =   0
         Text            =   "16384"
         Top             =   756
         Width           =   588
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1164
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Nessun file selezionato"
         Top             =   276
         Width           =   3012
      End
      Begin VB.TextBox txtfileSize 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   5784
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   276
         Width           =   1008
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Numero files creati:"
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
         Left            =   4320
         TabIndex        =   17
         Top             =   696
         Width           =   2532
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Grandezza singolo spezzettamento:"
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
         Left            =   156
         TabIndex        =   14
         Top             =   696
         Width           =   4128
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nome file:"
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
         Left            =   156
         TabIndex        =   13
         Top             =   216
         Width           =   4128
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Grandezza file:"
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
         Left            =   4320
         TabIndex        =   12
         Top             =   216
         Width           =   2532
      End
   End
   Begin VB.Frame Frame4 
      Height          =   444
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   8316
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Rev-X - Step 4: Selezionare grandezza spezzettamenti.."
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
         Left            =   84
         TabIndex        =   6
         Top             =   144
         Width           =   8184
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAvanti_Click()
Me.Visible = False
Form5.Show
End Sub

Private Sub btnIndietro_Click()
Me.Visible = False
Form3.Show
End Sub

Private Sub btnSalvaDove_Click()
folder = BrowseFolder("Salva dove...", Me)
directoryconglisplit = folder
End Sub

Private Sub btnSplitta_Click()
On Error Resume Next
txtLabel.Visible = True
Me.Refresh

    Dim err_descr As String

    If Not SplitFile(txtFileName.Text, 0, err_descr, CLng(txtsplit.Text)) Then
        MsgBox err_descr
    Else
        loading.Visible = False
        MsgBox "File splittato correttamente.", vbOKOnly
    End If
    txtLabel.Visible = False
    Exit Sub

End Sub

Private Sub Form_Load()
txtFileName.Text = FileDaSplittare
txtfileSize.Text = FileDaSplittareSize
txtsplit.Text = singlesplit
If singlesplit = 0 Then txtsplit.Text = "8192"
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
Me.Refresh

singlesplit = txtsplit.Text
txtnumerofiles.Text = Int(FileDaSplittareSize / txtsplit.Text) + 1
If directoryconglisplit <> "" Then btnSplitta.Enabled = True: btnAvanti.Enabled = True
End Sub
