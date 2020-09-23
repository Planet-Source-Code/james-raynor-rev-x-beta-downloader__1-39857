VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8FF0514F-A9CD-4CA9-AB6E-31D3B9591CA0}#1.0#0"; "ftpocx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REV-X - Configuratore"
   ClientHeight    =   3636
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8424
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   3636
   ScaleWidth      =   8424
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   60
      TabIndex        =   12
      Top             =   3192
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   672
      Left            =   60
      TabIndex        =   3
      Top             =   1848
      Width           =   8328
      Begin MSComDlg.CommonDialog CommonSalva 
         Left            =   2988
         Top             =   228
         _ExtentX        =   677
         _ExtentY        =   677
         _Version        =   393216
         DefaultExt      =   "ini"
         DialogTitle     =   "Salva configurazione..."
         Filter          =   "*.ini"
      End
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
         Left            =   3408
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Attendere...."
         Top             =   252
         Visible         =   0   'False
         Width           =   1200
      End
      Begin ftpOCX.FTP FTP1 
         Left            =   2292
         Top             =   156
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   1524
         Top             =   228
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Help..."
         Height          =   372
         Left            =   96
         TabIndex        =   6
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnIndietro 
         Caption         =   "<== Indietro"
         Height          =   372
         Left            =   5556
         TabIndex        =   5
         Top             =   216
         Width           =   1308
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
   End
   Begin VB.Frame Frame3 
      Height          =   1308
      Left            =   60
      TabIndex        =   2
      Top             =   516
      Width           =   8316
      Begin VB.CommandButton Command4 
         Caption         =   "UPLOAD.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   5856
         TabIndex        =   11
         Top             =   720
         Width           =   2352
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Crea Downloader FILE..."
         Height          =   384
         Left            =   2520
         TabIndex        =   10
         Top             =   708
         Width           =   2352
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salva configurazione..."
         Height          =   384
         Left            =   96
         TabIndex        =   9
         Top             =   708
         Width           =   2352
      End
      Begin VB.TextBox txtProgetto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1596
         TabIndex        =   7
         Top             =   288
         Width           =   6564
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
         Left            =   96
         TabIndex        =   8
         Top             =   228
         Width           =   8112
      End
   End
   Begin VB.Frame Frame4 
      Height          =   444
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8316
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Rev-X - Step 9: Passo finale..."
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
         Left            =   108
         TabIndex        =   1
         Top             =   144
         Width           =   8184
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   300
      Left            =   60
      TabIndex        =   13
      Top             =   2712
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String
Dim strFileName

Private Sub btnIndietro_Click()
Me.Visible = False
form9.Show
End Sub

Private Sub Command2_Click()
With CommonSalva
    .DefaultExt = "ini"
    .DialogTitle = "Salva configurazione..."
    .Filter = "*.ini"
    .CancelError = False
    .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
    strFileName = .FileName
End With


'REVX FIRMA
KeySection = "REVX"
KeyKey = "REVX"
KeyValue = 5
saveini

'NOME PROGETTO
KeySection = "Progetto"
KeyKey = "Nome"
KeyValue = nomeprogetto
saveini

'INFORMAZIONI PROGETTO
KeySection = "Informazioni"
KeyKey = "Info"
If informazioni = "" Then informazioni = "-"
KeyValue = descrizioneprogetto
saveini

'FileDaSplittareConPercorso
KeySection = "FileDaSplittareConPercorso"
KeyKey = "Info"
KeyValue = FileDaSplittareConPercorso
saveini

'FileDaSplittare
KeySection = "FileDaSplittare"
KeyKey = "Info"
KeyValue = FileDaSplittare
saveini

'FileDaSplittareSize
KeySection = "FileDaSplittareSize"
KeyKey = "Info"
KeyValue = FileDaSplittareSize
saveini

'singlesplit
KeySection = "singlesplit"
KeyKey = "Info"
KeyValue = singlesplit
saveini

'directoryconglisplit
KeySection = "directoryconglisplit"
KeyKey = "Info"
KeyValue = directoryconglisplit
saveini

'nomefilesequenziale
KeySection = "nomefilesequenziale"
KeyKey = "Info"
KeyValue = nomefilesequenziale
saveini

'estensionefilesequenziale
KeySection = "estensionefilesequenziale"
KeyKey = "Info"
KeyValue = estensionefilesequenziale
saveini

'directorydacreare
KeySection = "directorydacreare"
KeyKey = "Info"
KeyValue = directorydacreare
saveini

'okCrearedirectory
KeySection = "okCrearedirectory"
KeyKey = "Info"
If okCrearedirectory = True Then tempor = 1 Else tempor = 0
KeyValue = tempor
saveini

'falsohtml
KeySection = "falsohtml"
KeyKey = "Info"
If falsohtml = True Then tempor = 1 Else tempor = 0
KeyValue = tempor
saveini




Form6.Adodc1.Recordset.MoveLast
tempor5 = Form6.Adodc1.Recordset.AbsolutePosition

'INFORMAZIONI QUANTIFTP
KeySection = "QUANTIFTP"
KeyKey = "Quanti"
KeyValue = tempor5
saveini

Form6.Adodc1.Recordset.MoveFirst


On Error Resume Next
'INFORMAZIONI FTP
For v = 1 To tempor5

KeySection = "FTP"
KeyKey = "Url" & v
KeyValue = Form6.Url.Text
saveini
KeyKey = "Dir" & v
KeyValue = Form6.txtDir.Text
saveini
KeyKey = "Nome" & v
KeyValue = Form6.Nome.Text
saveini
KeyKey = "Passw" & v
KeyValue = Form6.Passw.Text
saveini
KeyKey = "Porta" & v
KeyValue = Form6.porta.Text
saveini
KeyKey = "Numero" & v
KeyValue = Form6.Numero.Text
saveini

Form6.Adodc1.Recordset.MoveNext
Next v


For v = 1 To 49
KeySection = "COUNT"
KeyKey = "Count" & v
KeyValue = Form7.List2(v).ListCount
saveini
Next v



For v = 1 To 49
KeySection = "LIST" & v
For I = 1 To Form7.List2(v).ListCount
KeyKey = "FileListed" & I
'a = Len(directoryconglisplit)

KeyValue = Form7.List2(v).List(I - 1)
saveini
Next I
Next v



Adodc1.Recordset.MoveFirst



End Sub

Private Sub Command3_Click()
Dim tempfilename As String
On Error Resume Next

With CommonSalva
    .DefaultExt = "RVX"
    .DialogTitle = "Crea file downloader..."
    .Filter = "*.RVX"
    .CancelError = False
    .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
    tempfilename = .FileName
    strFileName = "c:\temp.ini"
End With

'NOME PROGETTO
KeySection = "Progetto"
KeyKey = "Nome"
KeyValue = nomeprogetto
saveini

'INFORMAZIONI PROGETTO
KeySection = "Informazioni"
KeyKey = "Info"
If informazioni = "" Then informazioni = "-"
KeyValue = descrizioneprogetto
saveini

'FileDaSplittareConPercorso
KeySection = "FileDaSplittareConPercorso"
KeyKey = "Info"
KeyValue = FileDaSplittareConPercorso
saveini

'FileDaSplittare
KeySection = "FileDaSplittare"
KeyKey = "Info"
KeyValue = FileDaSplittare
saveini

'FileDaSplittareSize
KeySection = "FileDaSplittareSize"
KeyKey = "Info"
KeyValue = FileDaSplittareSize
saveini

'singlesplit
KeySection = "singlesplit"
KeyKey = "Info"
KeyValue = singlesplit
saveini

'directoryconglisplit
KeySection = "directoryconglisplit"
KeyKey = "Info"
KeyValue = directoryconglisplit
saveini

'nomefilesequenziale
KeySection = "nomefilesequenziale"
KeyKey = "Info"
KeyValue = nomefilesequenziale
saveini

'estensionefilesequenziale
KeySection = "estensionefilesequenziale"
KeyKey = "Info"
KeyValue = estensionefilesequenziale
saveini

'directorydacreare
KeySection = "directorydacreare"
KeyKey = "Info"
KeyValue = directorydacreare
saveini

'okCrearedirectory
KeySection = "okCrearedirectory"
KeyKey = "Info"
If okCrearedirectory = True Then tempor = 1 Else tempor = 0
KeyValue = tempor
saveini

'falsohtml
KeySection = "falsohtml"
KeyKey = "Info"
If falsohtml = True Then tempor = 1 Else tempor = 0
KeyValue = tempor
saveini




Form6.Adodc1.Recordset.MoveLast
tempor5 = Form6.Adodc1.Recordset.AbsolutePosition

'INFORMAZIONI QUANTIFTP
KeySection = "QUANTIFTP"
KeyKey = "Quanti"
KeyValue = tempor5
saveini

Form6.Adodc1.Recordset.MoveFirst


On Error Resume Next
'INFORMAZIONI FTP
For v = 1 To tempor5

KeySection = "FTP"
KeyKey = "Url" & v
KeyValue = Form6.Url.Text
saveini
KeyKey = "Dir" & v
KeyValue = Form6.txtDir.Text
saveini
KeyKey = "Nome" & v
KeyValue = Form6.Nome.Text
saveini
KeyKey = "Passw" & v
KeyValue = Form6.Passw.Text
saveini
KeyKey = "Porta" & v
KeyValue = Form6.porta.Text
saveini
KeyKey = "Numero" & v
KeyValue = Form6.Numero.Text
saveini

Form6.Adodc1.Recordset.MoveNext
Next v


For v = 1 To 49
KeySection = "COUNT"
KeyKey = "Count" & v
KeyValue = Form7.List2(v).ListCount
saveini
Next v



For v = 1 To 49
KeySection = "LIST" & v
For I = 1 To Form7.List2(v).ListCount
KeyKey = "FileListed" & I
a = Len(directoryconglisplit)

KeyValue = Mid(Form7.List2(v).List(I - 1), a + 2)
saveini
Next I
Next v



Adodc1.Recordset.MoveFirst

a = CompressFile(strFileName, tempfilename, 5)
DeleteFile ("c:\temp.ini")

End Sub

Private Sub Form_Load()
On Error Resume Next
ocio = False

txtProgetto.Text = nomeprogetto
Me.Height = 3000

ProgressBar1.Max = Form6.Adodc1.Recordset.RecordCount
ProgressBar2.Max = Form7.List2(1).ListCount

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

Private Sub Timer1_Timer()
Form10.Refresh
End Sub

Private Sub Command4_Click()
On Error Resume Next
txtLabel.Visible = True
ProgressBar1.Max = Form6.Adodc1.Recordset.RecordCount
ProgressBar2.Max = Form7.List2(1).ListCount
Me.Height = 4000
Me.Refresh


Form6.Adodc1.Recordset.MoveLast
tempor5 = Form6.Adodc1.Recordset.AbsolutePosition
tempor5 = tempor5 + 1

Form6.Adodc1.Recordset.MoveFirst


For v = 1 To tempor5 - 1
    
FTP1.Connect App.Title, Form6.Url.Text, Form6.porta.Text, Form6.Nome.Text, Form6.Passw.Text
If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub
    
If okCrearedirectory = True Then FTP1.MakeDIR Form6.txtDir.Text & directorydacreare
If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub
If falsohtml = True Then FTP1.UploadFile Form6.txtDir.Text & "default.html", App.Path & "\default.html"
If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub
For I = 0 To Form7.List2(v).ListCount - 1

'If frmConfig.txtdirectory.Text = True Then FTP1.UploadFile Text1.Text & Text2.Text & i & "." & Text3.Text, List2.List(i)
If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub

If okCrearedirectory = True Then
    FTP1.UploadFile Form6.txtDir.Text & directorydacreare & nomefilesequenziale & I & "." & estensionefilesequenziale, Form7.List2(v).List(I)
Else
    FTP1.UploadFile Form6.txtDir.Text & nomefilesequenziale & I & "." & estensionefilesequenziale, Form7.List2(v).List(I)
    'FTP1.UploadFile Form6.txtDir.Text & frmConfig.File.Text & I & "." & frmConfig.estensione.Text, Form6.List2(v).List(I)
End If

ProgressBar2.value = I
Me.Refresh

If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub
Next I

Form6.Adodc1.Recordset.MoveNext
FTP1.Disconnect

ProgressBar1.value = v
Me.Refresh

ProgressBar2.Max = Form7.List2(v + 1).ListCount
ProgressBar2.value = 0

Next v

Form6.Adodc1.Recordset.MoveFirst
MsgBox ("UPLOAD COMPLETATO CON SUCCESSO")
ocio = False
Me.Height = 3000
txtLabel.Visible = False

End Sub














Private Sub FTP1_GetError(Error As String, Func As String, ErrorNum As Long)
  FTP1.Disconnect
  MsgBox ("Errore rilevato. controllare che tutte le impostazioni siano corrette e riprovare" & " - " & Form6.Url.Text)
  ocio = True
  Me.Height = 3000
  txtLabel.Visible = False
  
 
End Sub




Private Sub saveini()

Dim lngResult As Long

'strFileName = App.Path & "\Projects\" & Nome_File_Salva & ".ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("Impossibile salvare", vbExclamation)
End If

End Sub




