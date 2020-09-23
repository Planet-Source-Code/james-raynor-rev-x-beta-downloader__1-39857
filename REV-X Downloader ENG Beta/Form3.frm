VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REV-X - Configuratore"
   ClientHeight    =   2220
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8424
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   8424
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   444
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   8316
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Rev-X - Step 3: Selezionare file da splittare..."
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
         TabIndex        =   7
         Top             =   144
         Width           =   8184
      End
   End
   Begin VB.Frame Frame3 
      Height          =   912
      Left            =   60
      TabIndex        =   5
      Top             =   456
      Width           =   8316
      Begin VB.TextBox txtfileSize 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   480
         Width           =   1008
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Nessun file selezionato"
         Top             =   480
         Width           =   3012
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apri file..."
         Height          =   444
         Left            =   108
         TabIndex        =   0
         Top             =   300
         Width           =   1296
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
         Left            =   5676
         TabIndex        =   11
         Top             =   420
         Width           =   2532
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
         Left            =   1512
         TabIndex        =   9
         Top             =   420
         Width           =   4128
      End
   End
   Begin VB.Frame Frame1 
      Height          =   672
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   8316
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   1656
         Top             =   228
      End
      Begin VB.CommandButton btnAvanti 
         Caption         =   "Avanti ==>"
         Enabled         =   0   'False
         Height          =   372
         Left            =   6900
         TabIndex        =   4
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnIndietro 
         Caption         =   "<== Indietro"
         Height          =   372
         Left            =   5556
         TabIndex        =   3
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "Help..."
         Height          =   372
         Left            =   96
         TabIndex        =   2
         Top             =   216
         Width           =   1308
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   324
      Top             =   1296
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Color           =   25
      DialogTitle     =   "Apri file da splittare"
      Filter          =   "*.*"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAvanti_Click()
Form4.Show
Form4.txtFileName.Text = FileDaSplittare
Form4.txtfileSize.Text = FileDaSplittareSize
Form4.txtsplit.Text = singlesplit
If singlesplit = 0 Then txtsplit.Text = "8192"

Me.Visible = False

End Sub

Private Sub btnIndietro_Click()
Me.Visible = False
Form2.Show
End Sub

Private Sub Command1_Click()
tempor = "Kb"

With CommonDialog1
    .CancelError = False
    .ShowOpen
        If Len(.FileName) = 0 Then
            FileDaSplittare = "Nessun file selezionato": Exit Sub
        End If
    FileDaSplittareConPercorso = .FileName
    FileDaSplittare = .FileTitle
    txtFileName.Text = FileDaSplittare
    GetFileSize FileDaSplittareConPercorso
    FileDaSplittareSize = tempsize
    
    If tempsize >= 1048576 Then tempsize = Int(tempsize / 1048576): tempor = "Mb"
    FileDaSplittareSize = tempsize
    txtfileSize.Text = FileDaSplittareSize & " " & tempor
    tempor = "Kb"
    
End With

End Sub

Private Sub Form_Load()
txtFileName.Text = FileDaSplittare
If FileDaSplittare = "" Then txtFileName.Text = "Nessun file selezionato"
txtfileSize.Text = FileDaSplittareSize
If FileDaSplittareSize = "" Then txtfileSize.Text = "0"

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
txtFileName.Text = FileDaSplittare
If FileDaSplittare = "" Then txtFileName.Text = "Nessun file selezionato"
txtfileSize.Text = FileDaSplittareSize
If FileDaSplittareSize = "" Then txtfileSize.Text = "0"

txtFileName.ToolTipText = FileDaSplittareConPercorso
If txtFileName.Text <> "Nessun file selezionato" Then btnAvanti.Enabled = True: Exit Sub
btnAvanti.Enabled = False: Exit Sub

End Sub
