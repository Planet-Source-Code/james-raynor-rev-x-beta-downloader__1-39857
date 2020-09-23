VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{8FF0514F-A9CD-4CA9-AB6E-31D3B9591CA0}#1.0#0"; "ftpocx.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{B31EC2AB-FEC0-11D4-9B23-0000B49F239E}#2.0#0"; "NmFileTypReg.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REV-X - Downloader by James Raynor"
   ClientHeight    =   3696
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8424
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3696
   ScaleWidth      =   8424
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton command3 
      Height          =   336
      Left            =   6240
      TabIndex        =   23
      Top             =   3000
      Width           =   2124
      _ExtentX        =   3747
      _ExtentY        =   593
      BTYPE           =   14
      TX              =   "DOWNLOAD FILE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      DataField       =   "Id"
      DataSource      =   "Adodc1"
      Height          =   336
      Left            =   2916
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   5292
      Visible         =   0   'False
      Width           =   828
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4344
      Top             =   5280
   End
   Begin VB.TextBox ext 
      Height          =   288
      Left            =   3264
      TabIndex        =   20
      Top             =   5304
      Visible         =   0   'False
      Width           =   588
   End
   Begin VB.TextBox dir 
      Height          =   288
      Left            =   1428
      TabIndex        =   19
      Top             =   5244
      Visible         =   0   'False
      Width           =   588
   End
   Begin VB.TextBox thumb 
      Height          =   288
      Left            =   1944
      TabIndex        =   18
      Top             =   5364
      Visible         =   0   'False
      Width           =   588
   End
   Begin VB.TextBox dirD 
      Height          =   300
      Left            =   216
      TabIndex        =   17
      Top             =   5292
      Visible         =   0   'False
      Width           =   408
   End
   Begin VB.TextBox fileD 
      Height          =   300
      Left            =   456
      TabIndex        =   16
      Top             =   5316
      Visible         =   0   'False
      Width           =   408
   End
   Begin VB.TextBox valor 
      Height          =   300
      Left            =   1440
      TabIndex        =   15
      Top             =   5256
      Visible         =   0   'False
      Width           =   408
   End
   Begin VB.TextBox txtTemp 
      Height          =   300
      Left            =   1020
      TabIndex        =   14
      Top             =   5292
      Visible         =   0   'False
      Width           =   408
   End
   Begin VB.TextBox dirdd 
      Height          =   300
      Left            =   792
      TabIndex        =   13
      Top             =   5256
      Visible         =   0   'False
      Width           =   408
   End
   Begin VB.TextBox txtSize 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   144
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3312
      Width           =   1488
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   504
      Left            =   1644
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "No description"
      Top             =   3060
      Width           =   4428
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   276
      Left            =   72
      TabIndex        =   6
      Top             =   2280
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   487
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   444
      Left            =   60
      TabIndex        =   3
      Top             =   24
      Width           =   8316
      Begin VB.Image Image3 
         Height          =   240
         Left            =   5412
         Picture         =   "Form1.frx":0326
         Stretch         =   -1  'True
         Top             =   156
         Width           =   2820
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   36
         Picture         =   "Form1.frx":3AB8
         Stretch         =   -1  'True
         Top             =   156
         Width           =   2820
      End
      Begin VB.Image Image1 
         Height          =   276
         Left            =   2892
         Picture         =   "Form1.frx":724A
         Stretch         =   -1  'True
         Top             =   144
         Width           =   2460
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   1000
      Left            =   60
      TabIndex        =   2
      Top             =   504
      Width           =   8316
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Left            =   7752
         Top             =   420
      End
      Begin VB.TextBox txtProgetto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3144
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "No project loaded"
         Top             =   468
         Width           =   4908
      End
      Begin prjChameleon.chameleonButton Command2 
         Height          =   372
         Left            =   120
         TabIndex        =   0
         Top             =   372
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   656
         BTYPE           =   4
         TX              =   "Open file..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":A9E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Project name:"
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
         Left            =   1740
         TabIndex        =   5
         Top             =   408
         Width           =   6408
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   672
      Left            =   60
      TabIndex        =   1
      Top             =   1512
      Width           =   8316
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7020
         Top             =   264
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
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Download in progress..."
         Top             =   264
         Visible         =   0   'False
         Width           =   3444
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   312
         Left            =   1620
         Top             =   228
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   550
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   0
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtProgress 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   5268
         TabIndex        =   12
         Top             =   300
         Width           =   2928
      End
      Begin prjChameleon.chameleonButton command1 
         Height          =   372
         Left            =   120
         TabIndex        =   26
         Top             =   204
         Width           =   372
         _ExtentX        =   656
         _ExtentY        =   656
         BTYPE           =   4
         TX              =   "?"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":AA00
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.CheckBox CheckBox1 
         Height          =   468
         Left            =   564
         TabIndex        =   22
         Top             =   168
         Width           =   1104
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1947;825"
         Value           =   "1"
         Caption         =   "Auto - RETRY"
         SpecialEffect   =   0
         FontHeight      =   156
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   276
      Left            =   72
      TabIndex        =   7
      Top             =   2640
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   487
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin ftpOCX.FTP FTP1 
      Left            =   4920
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      TransferType    =   2
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   2484
      Top             =   5244
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DefaultExt      =   "tpl"
      DialogTitle     =   "Scegli file da ricostruire..."
      Filter          =   "*.tpl"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3912
      Top             =   5172
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DefaultExt      =   "rvx"
      DialogTitle     =   "Apri file downloader.."
      Filter          =   "*.rvx"
   End
   Begin prjChameleon.chameleonButton command4 
      Height          =   252
      Left            =   6240
      TabIndex        =   24
      Top             =   3348
      Width           =   1008
      _ExtentX        =   1778
      _ExtentY        =   445
      BTYPE           =   5
      TX              =   "Build file..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":AA1C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btnCheck 
      Height          =   252
      Left            =   7368
      TabIndex        =   25
      Top             =   3348
      Width           =   1008
      _ExtentX        =   1778
      _ExtentY        =   445
      BTYPE           =   5
      TX              =   "Check file..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":AA38
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin NmRegisterFileType.NmFileTypeRegister nmFTR 
      Left            =   1068
      Top             =   5640
      _ExtentX        =   677
      _ExtentY        =   677
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " File description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   72
      TabIndex        =   8
      Top             =   3000
      Width           =   6060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Dim dirdestinazione As String

Dim totalsplit As Long
Dim isdownload
Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String
Dim strFileName
Dim riprova As Boolean
Dim checkriprova As Boolean
Dim maxriprova As Long
Dim perc As String
'Dim tempfilename2 As String

Dim opzione

Private Sub btnCheck_Click()

On Error Resume Next
CheckBox1.value = False
txtLabel.Visible = True
Me.Refresh

Form6.Adodc1.Recordset.MoveLast
tempor5 = Form6.Adodc1.Recordset.AbsolutePosition
tempor5 = tempor5 + 1

Form6.Adodc1.Recordset.MoveFirst


For v = 1 To tempor5 - 1

FTP1.Connect App.Title, Form6.Url.Text, Form6.porta.Text, Form6.Nome.Text, Form6.Passw.Text
    
I = Form7.List2(v).ListCount - 1

If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub

thumb.Text = nomefilesequenziale
Dir.Text = directorydacreare
ext.Text = estensionefilesequenziale

If okCrearedirectory = True Then
    FTP1.DownloadFile Form6.txtDir.Text & Dir.Text & "/" & thumb.Text & I & "." & ext.Text, "c:\temp.del"
Else
    FTP1.DownloadFile Form6.txtDir.Text & thumb.Text & I & "." & ext.Text, "c:\temp.del"
    
End If

If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub
Form6.Adodc1.Recordset.MoveNext
FTP1.Disconnect

DeleteFile ("c:\temp.del")
Me.Refresh

Next v
If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub

Form6.Adodc1.Recordset.MoveFirst
MsgBox ("No errors")
ocio = False
txtLabel.Visible = False
txtProgress.Visible = False

End Sub

Private Sub Command1_Click()
ExecuteFile App.Path & "\Downloader.PDF"
End Sub

Private Sub Command2_Click()

Dim tempfilename2 As String

On Error Resume Next

Me.Height = 2640
txtLabel.Visible = False
txtProgetto.Text = "No project loaded"


Form6.Show: Form6.Visible = False
Form7.Show: Form7.Visible = False
Form6.Adodc1.Recordset.MoveFirst
For v = 1 To 70
Form6.Adodc1.Recordset.MoveFirst
Form6.Adodc1.Recordset.Delete
'Adodc1.Refresh
Next v

For v = 1 To 49
Form7.List2(v).Clear
Next v

With CommonDialog1
    .Filter = "REV-X Projects (*.rvx)|*.rvx"
    .CancelError = False
    .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
    
    tempfilename2 = .FileName
End With

DecompressFile tempfilename2, "c:\temp.ini"

strFileName = "c:\temp.ini"

'REVX FIRMA
KeySection = "REVX"
KeyKey = "REVX"
loadini
If KeyValue <> 5 Then a = MsgBox("Cannot open file. Unknown format.", vbCritical) = vbOK: Exit Sub


'NOME PROGETTO
KeySection = "Progetto"
KeyKey = "Nome"
loadini
nomeprogetto = KeyValue


'INFORMAZIONI PROGETTO
KeySection = "Informazioni"
KeyKey = "Info"
loadini
descrizioneprogetto = KeyValue


'FileDaSplittareConPercorso
KeySection = "FileDaSplittareConPercorso"
KeyKey = "Info"
loadini
FileDaSplittareConPercorso = KeyValue


'FileDaSplittare
KeySection = "FileDaSplittare"
KeyKey = "Info"
loadini
FileDaSplittare = KeyValue


'FileDaSplittareSize
KeySection = "FileDaSplittareSize"
KeyKey = "Info"
loadini
FileDaSplittareSize = KeyValue
txtSize.Text = FileDaSplittareSize


'singlesplit
KeySection = "singlesplit"
KeyKey = "Info"
loadini
singlesplit = KeyValue


'directoryconglisplit
KeySection = "directoryconglisplit"
KeyKey = "Info"
loadini
directoryconglisplit = KeyValue


'nomefilesequenziale
KeySection = "nomefilesequenziale"
KeyKey = "Info"
loadini
nomefilesequenziale = KeyValue


'estensionefilesequenziale
KeySection = "estensionefilesequenziale"
KeyKey = "Info"
loadini
estensionefilesequenziale = KeyValue


'directorydacreare
KeySection = "directorydacreare"
KeyKey = "Info"
loadini
directorydacreare = KeyValue


'okCrearedirectory
KeySection = "okCrearedirectory"
KeyKey = "Info"
loadini
okCrearedirectory = KeyValue

'okCrearedirectory
KeySection = "falsohtml"
KeyKey = "Info"
loadini
falsohtml = KeyValue

'INFORMAZIONI QUANTIFTP
KeySection = "QUANTIFTP"
KeyKey = "Quanti"
loadini
tempor5 = KeyValue

On Error Resume Next

Adodc1.Recordset.MoveFirst
For v = 1 To 100
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Delete
Next v

'INFORMAZIONI FTP
For v = 1 To tempor5

Form6.Adodc1.Recordset.AddNew

KeySection = "FTP"
KeyKey = "Url" & v
loadini
Form6.Url.Text = KeyValue
KeyKey = "Dir" & v
loadini
Form6.txtDir.Text = KeyValue
KeyKey = "Nome" & v
loadini
Form6.Nome.Text = KeyValue
KeyKey = "Passw" & v
loadini
Form6.Passw.Text = KeyValue
KeyKey = "Porta" & v
loadini
Form6.porta.Text = KeyValue
KeyKey = "Numero" & v
loadini
Form6.Numero.Text = KeyValue

Form6.Adodc1.Recordset.Update

Next v

Adodc1.Recordset.MoveFirst
Form6.Numero.Text = 1


For v = 1 To 49

KeySection = "COUNT"
KeyKey = "Count" & v
loadini
tempor = KeyValue

KeySection = "LIST" & v
For I = 1 To tempor
KeyKey = "FileListed" & I
loadini
Form7.List2(v).AddItem KeyValue
Next I
Next v


Form2.txtDescrizione = descrizioneprogetto
Form2.txtProgetto = nomeprogetto
'Me.Visible = False
'Form10.Show
DeleteFile ("c:\temp.ini")

Height = 4050
txtProgetto.Text = nomeprogetto
Text2.Text = descrizioneprogetto
If FileDaSplittareSize < 1048576 Then txtSize.Text = FileDaSplittareSize & " Kb"
If FileDaSplittareSize >= 1048576 Then txtSize.Text = FileDaSplittareSize & " Kb"
ProgressBar1.value = 0
ProgressBar2.value = 0

End Sub






Private Sub Command3_Click()

On Error Resume Next
maxriprova = 0
'Dim dirdestinazione As String

If command3.Caption = "Cancel" Then isdownload = 1
If isdownload = 1 Then GoTo returntonormal

If isdownload = 0 Then Command2.Enabled = False: command4.Enabled = False: btnCheck.Enabled = False: command3.Caption = "Cancel"

riprovasub:
Me.Refresh
If maxriprova > 30 Then maxriprova = 0: GoTo returntonormal
totalsplit = 0
txtProgress.Visible = True: txtProgress.Text = "": totalsize = singlesplit
Form6.Adodc1.Refresh

Dim Security As SECURITY_ATTRIBUTES
ret& = CreateDirectory(App.Path & "\Temp_" & nomeprogetto, Security)
dirdestinazione = App.Path & "\Temp_" & nomeprogetto
dirdd.Text = dirdestinazione

On Error Resume Next

txtLabel.Visible = True
ProgressBar1.Max = Form6.Adodc1.Recordset.RecordCount
ProgressBar2.Max = Form7.List2(1).ListCount
ProgressBar1.value = 0
ProgressBar2.value = 0
Form6.Adodc1.Recordset.MoveFirst
Me.Height = 4050
Me.Refresh


Form6.Adodc1.Recordset.MoveLast
tempor5 = Form6.Adodc1.Recordset.AbsolutePosition
tempor5 = tempor5 + 1


Form6.Adodc1.Recordset.MoveFirst


For v = 1 To tempor5 - 1
    
If isdownload = 1 Then GoTo returntonormal

Randomize Timer: temporan = Rnd(10000): valor.Text = temporan
    
FTP1.Connect App.Title & valor.Text, Form6.Url.Text, Form6.porta.Text, Form6.Nome.Text, Form6.Passw.Text
If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub
    
If isdownload = 1 Then GoTo returntonormal
If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub
If riprova = True Then riprova = False: GoTo riprovasub
For I = 0 To Form7.List2(v).ListCount - 1
If isdownload = 1 Then GoTo returntonormal

'If frmConfig.txtdirectory.Text = True Then FTP1.UploadFile Text1.Text & Text2.Text & i & "." & Text3.Text, List2.List(i)
If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub

thumb.Text = nomefilesequenziale
Dir.Text = directorydacreare
ext.Text = estensionefilesequenziale

totalsize = singlesplit + totalsize
If totalsize > FileDaSplittareSize Then totalsize = FileDaSplittareSize
'totalsize = (singlesplit * I) * v
Dim timdsize As String
timdsize = totalsize & " di " & txtSize.Text
txtProgress.Text = timdsize
Me.Refresh

If FileExists(dirdd.Text & "\" & Form7.List2(v).List(I)) = True Then GoTo nexti

If isdownload = 1 Then GoTo returntonormal
If riprova = True Then riprova = False: GoTo riprovasub

If okCrearedirectory = True Then
    FTP1.DownloadFile Form6.txtDir.Text & Dir.Text & "/" & thumb.Text & I & "." & ext.Text, dirdd.Text & "\" & Form7.List2(v).List(I)
Else
    FTP1.DownloadFile Form6.txtDir.Text & thumb.Text & I & "." & ext.Text, dirdd.Text & "\" & Form7.List2(v).List(I)
    
    'FTP1.UploadFile Form6.txtDir.Text & frmConfig.File.Text & I & "." & frmConfig.estensione.Text, Form6.List2(v).List(I)
End If

nexti:
ProgressBar2.value = I
Me.Refresh
If isdownload = 1 Then GoTo returntonormal
If riprova = True Then riprova = False: GoTo riprovasub

If ocio = True Then ocio = False: txtLabel.Visible = False: Exit Sub
Next I

Form6.Adodc1.Recordset.MoveNext
FTP1.Disconnect

ProgressBar1.value = v
Me.Refresh

ProgressBar2.Max = Form7.List2(v + 1).ListCount
ProgressBar2.value = 0
If isdownload = 1 Then GoTo returntonormal
If riprova = True Then riprova = False: GoTo riprovasub
Next v

Form6.Adodc1.Recordset.MoveFirst
MsgBox ("DOWNLOAD COMPLETE")
ocio = False
Me.Height = 2640


Timer2.interval = 500: Timer2.Enabled = True
Exit Sub

returntonormal:

On Error Resume Next
If CheckBox1.value = True Then MsgBox "Please retry later...": CheckBox1.value = False
FTP1.Disconnect
isdownload = 0
ocio = False
Me.Height = 2640
txtLabel.Visible = False
txtProgetto.Text = "No project loaded"
Command2.Enabled = True: command4.Enabled = True: btnCheck.Enabled = True: command3.Caption = "Download": txtProgress.Visible = False

End Sub







Private Sub Command4_Click()
Dim tempdir As String
Dim tempfile As String

With CommonDialog2
    .Filter = "Templates files (*.tpl)|*.tpl"
    .DialogTitle = "Please select the TPL to reassemble.."
    .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
    
    tempfile = .FileName


End With

Dim temp30 As String
folder = BrowseFolder("Save where...", Me)
tempdir = folder
temp30 = tempdir & "\" & FileDaSplittare

Dim err_descr As String

    If Not ReassembleFile(tempfile, False, temp30) Then
        MsgBox err_descr
    Else
        MsgBox "File created."
    End If
   
End Sub





Private Sub Form_Terminate()
Unload Form1
Unload Form10
Unload Form6
Unload Form7
Unload form9
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form10
Unload Form6
Unload Form7
Unload form9
Unload Me
End Sub

Private Sub FTP1_GetError(Error As String, Func As String, ErrorNum As Long)
On Error Resume Next
If CheckBox1.value = False Then
  FTP1.Disconnect
  MsgBox ("ERROR. Please retry later")
  ocio = True
  Me.Height = 2640
  txtLabel.Visible = False
  txtProgetto.Text = "No project loaded"
  isdownload = 0
  Command2.Enabled = True: command4.Enabled = True: btnCheck.Enabled = True: command3.Caption = "Download": txtProgress.Visible = False
  CheckBox1.value = True
  Exit Sub
End If

If CheckBox1.value = True Then
  FTP1.Disconnect
  txtLabel.Text = "Auto-retry in 10 seconds..."
  Me.Refresh
  Wait 10
  txtLabel.Text = "Download in progress..."
  riprova = True
  Me.Refresh
  maxriprova = maxriprova + 1
End If
 
End Sub

Private Sub Form_Load()
riprova = False
checkriprova = True

isdownload = 0
Me.Height = 2640
If App.PrevInstance = True Then Unload Me
On Error Resume Next
'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "db.mdb" & ";Persist Security Info=False"
'Adodc1.Enabled = True

nomeprogetto = "New project"
descrizioneprogetto = "Project description"


perc = App.Path
If Right(perc, 1) <> "\" Then perc = perc + "\"

    nmFTR.ApplicationIconIndex = 0
    nmFTR.ApplicationToRun = perc + "RevX Downloader.exe"
    nmFTR.ExtFullDescription = "REV-X Downloader Beta file"
    nmFTR.ExtShortDescription = "Rev-X file"
    nmFTR.FileContentType = "noMime"
    nmFTR.FileExtension = ".rvx"
    
    nmFTR.RegisterNewFileType

End Sub




Private Sub Option1_Click()
opzione = 1
btnAvanti.Enabled = True
End Sub

Private Sub Option2_Click()
opzione = 2
btnAvanti.Enabled = True
End Sub




Private Sub loadini()

Dim lngResult As Long
'Dim strFileName
Dim strResult As String * 100
'strFileName = App.Path & "\Projects\" & Nome_File_Apri & ".rvs" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("Cannot open file. Invalid format.", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub


Private Sub Timer1_Timer()
Me.Refresh
End Sub


Private Sub Timer2_Timer()
On Error Resume Next

Timer2.Enabled = False

'folderr = BrowseFolder("Salva dove...", Me)
'tempdirr = folderr

'dirD.Text = dirdestinazione
'fileD.Text = FileDaSplittare
'txtTemp.Text = tempdirr

'Dim tempor As String
'Dim tempor2 As String
'tempor3 = txtTemp.Text & "\" & fileD.Text
'tempor2 = dirdd.Text & "\" & fileD.Text & ".tpl"

'tempsplit1 = tempor2
'tempsplit2 = tempor3

'Dim err_descr As String

'    If Not Split.ReassembleFile(tempsplit1, False, tempsplit2) Then
        'MsgBox err_descr
        'okdel = 1
    'Else
        'MsgBox "File creato correttamente"
    'End If

Dim tempdir As String
Dim tempfile As String


fileD.Text = FileDaSplittare
'tempfile = dirdd.Text & "\" & fileD.Text & ".tpl"

'With CommonDialog2
    '.Filter = "Files templates (*.tpl)|*.tpl"
    '.InitDir = dirdd.Text
    '.DialogTitle = "Seleziona il file TPL da riassemblare.."
    '.ShowOpen
        'If Len(.FileName) = 0 Then txtLabel.Visible = False: txtProgetto.Text = "Nessun progetto caricato": txtProgress.Visible = False: Command2.Enabled = True: Exit Sub

    'tempfile = .FileName

'End With
'Dim tempfile As String

SetCurrentDirectory dirdd.Text
tempfile = dirdd.Text & "\" & fileD.Text & ".tpl"

Dim temp30 As String
folder = BrowseFolder("Save the downloaded file where...", Me)
tempdir = folder
temp30 = tempdir & "\" & FileDaSplittare

Dim err_descr As String

    If Not ReassembleFile(tempfile, False, temp30) Then
        MsgBox err_descr
        okdel = 1
    Else
        MsgBox "File created."
    End If


If okdel = 1 Then okdel = 0: Exit Sub

Form6.Adodc1.Recordset.MoveLast
tempor5 = Form6.Adodc1.Recordset.AbsolutePosition
tempor5 = tempor5 + 1

Form6.Adodc1.Recordset.MoveFirst


On Error Resume Next

For v = 1 To tempor5 - 1
For I = 0 To Form7.List2(v).ListCount - 1
Kill (dirdd.Text & "\" & Form7.List2(v).List(I))
Next I
Form6.Adodc1.Recordset.MoveNext
Next v

txtLabel.Visible = False
txtProgetto.Text = "No project loaded"
Command2.Enabled = True: command4.Enabled = True: btnCheck.Enabled = True: command3.Caption = "Download": txtProgress.Visible = False
Me.Height = 2640
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False

If startfile = "" Then Exit Sub

Dim tempfilename2 As String

On Error Resume Next

Me.Height = 2640
txtLabel.Visible = False
txtProgetto.Text = "No project loaded"


Form6.Show: Form6.Visible = False
Form7.Show: Form7.Visible = False
Form6.Adodc1.Recordset.MoveFirst
For v = 1 To 70
Form6.Adodc1.Recordset.MoveFirst
Form6.Adodc1.Recordset.Delete
'Adodc1.Refresh
Next v

For v = 1 To 49
Form7.List2(v).Clear
Next v

tempfilename2 = startfile


DecompressFile tempfilename2, "c:\temp.ini"

strFileName = "c:\temp.ini"

'REVX FIRMA
KeySection = "REVX"
KeyKey = "REVX"
loadini
If KeyValue <> 5 Then a = MsgBox("Cannot open file. Invalid format.", vbCritical) = vbOK: Exit Sub

txtLabel.Text = "Load in progress..."
txtLabel.Visible = True
Me.Refresh

'NOME PROGETTO
KeySection = "Progetto"
KeyKey = "Nome"
loadini
nomeprogetto = KeyValue


'INFORMAZIONI PROGETTO
KeySection = "Informazioni"
KeyKey = "Info"
loadini
descrizioneprogetto = KeyValue


'FileDaSplittareConPercorso
KeySection = "FileDaSplittareConPercorso"
KeyKey = "Info"
loadini
FileDaSplittareConPercorso = KeyValue


'FileDaSplittare
KeySection = "FileDaSplittare"
KeyKey = "Info"
loadini
FileDaSplittare = KeyValue


'FileDaSplittareSize
KeySection = "FileDaSplittareSize"
KeyKey = "Info"
loadini
FileDaSplittareSize = KeyValue
txtSize.Text = FileDaSplittareSize


'singlesplit
KeySection = "singlesplit"
KeyKey = "Info"
loadini
singlesplit = KeyValue


'directoryconglisplit
KeySection = "directoryconglisplit"
KeyKey = "Info"
loadini
directoryconglisplit = KeyValue


'nomefilesequenziale
KeySection = "nomefilesequenziale"
KeyKey = "Info"
loadini
nomefilesequenziale = KeyValue


'estensionefilesequenziale
KeySection = "estensionefilesequenziale"
KeyKey = "Info"
loadini
estensionefilesequenziale = KeyValue


'directorydacreare
KeySection = "directorydacreare"
KeyKey = "Info"
loadini
directorydacreare = KeyValue


'okCrearedirectory
KeySection = "okCrearedirectory"
KeyKey = "Info"
loadini
okCrearedirectory = KeyValue

'okCrearedirectory
KeySection = "falsohtml"
KeyKey = "Info"
loadini
falsohtml = KeyValue

'INFORMAZIONI QUANTIFTP
KeySection = "QUANTIFTP"
KeyKey = "Quanti"
loadini
tempor5 = KeyValue

On Error Resume Next

Adodc1.Recordset.MoveFirst
For v = 1 To 100
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Delete
Next v

'INFORMAZIONI FTP
For v = 1 To tempor5

Form6.Adodc1.Recordset.AddNew

KeySection = "FTP"
KeyKey = "Url" & v
loadini
Form6.Url.Text = KeyValue
KeyKey = "Dir" & v
loadini
Form6.txtDir.Text = KeyValue
KeyKey = "Nome" & v
loadini
Form6.Nome.Text = KeyValue
KeyKey = "Passw" & v
loadini
Form6.Passw.Text = KeyValue
KeyKey = "Porta" & v
loadini
Form6.porta.Text = KeyValue
KeyKey = "Numero" & v
loadini
Form6.Numero.Text = KeyValue

Form6.Adodc1.Recordset.Update

Next v

Adodc1.Recordset.MoveFirst
Form6.Numero.Text = 1


For v = 1 To 49

KeySection = "COUNT"
KeyKey = "Count" & v
loadini
tempor = KeyValue

KeySection = "LIST" & v
For I = 1 To tempor
KeyKey = "FileListed" & I
loadini
Form7.List2(v).AddItem KeyValue
Next I
Next v


Form2.txtDescrizione = descrizioneprogetto
Form2.txtProgetto = nomeprogetto
'Me.Visible = False
'Form10.Show
DeleteFile ("c:\temp.ini")

Height = 4050
txtProgetto.Text = nomeprogetto
Text2.Text = descrizioneprogetto
If FileDaSplittareSize < 1048576 Then txtSize.Text = FileDaSplittareSize & " Kb"
If FileDaSplittareSize >= 1048576 Then txtSize.Text = FileDaSplittareSize & " Kb"
ProgressBar1.value = 0
ProgressBar2.value = 0

txtLabel.Text = "Download in progress..."
txtLabel.Visible = False
Me.Refresh


End Sub
