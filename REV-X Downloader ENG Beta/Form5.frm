VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{8FF0514F-A9CD-4CA9-AB6E-31D3B9591CA0}#1.0#0"; "ftpocx.ocx"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REV-X - Configuratore"
   ClientHeight    =   6240
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   8424
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8424
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Height          =   888
      Left            =   60
      TabIndex        =   15
      Top             =   4488
      Width           =   8328
      Begin VB.CommandButton btnSalva 
         Caption         =   "Salva"
         Height          =   276
         Left            =   6180
         TabIndex        =   7
         Top             =   528
         Width           =   2040
      End
      Begin VB.TextBox Dir 
         DataField       =   "Dir"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   3924
         TabIndex        =   20
         Top             =   168
         Width           =   4296
      End
      Begin VB.TextBox Passw 
         DataField       =   "Password"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   3096
         TabIndex        =   19
         Top             =   528
         Width           =   1608
      End
      Begin VB.TextBox Nome 
         DataField       =   "Nome"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   552
         TabIndex        =   18
         Top             =   528
         Width           =   1884
      End
      Begin VB.TextBox Url 
         DataField       =   "Url"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   552
         TabIndex        =   17
         Top             =   168
         Width           =   2928
      End
      Begin VB.TextBox porta 
         DataField       =   "Porta"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   5352
         TabIndex        =   16
         Top             =   528
         Width           =   612
      End
      Begin MSForms.Label Label13 
         Height          =   192
         Left            =   180
         TabIndex        =   8
         Top             =   192
         Width           =   444
         ForeColor       =   12582912
         BackColor       =   -2147483648
         Caption         =   "Url:"
         Size            =   "783;339"
         FontName        =   "TechnicBold"
         FontEffects     =   1073741825
         FontHeight      =   156
         FontCharSet     =   2
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   192
         Left            =   3600
         TabIndex        =   24
         Top             =   204
         Width           =   444
         ForeColor       =   12582912
         BackColor       =   -2147483648
         Caption         =   "Dir:"
         Size            =   "783;339"
         FontName        =   "TechnicBold"
         FontEffects     =   1073741825
         FontHeight      =   156
         FontCharSet     =   2
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label3 
         Height          =   192
         Left            =   60
         TabIndex        =   23
         Top             =   564
         Width           =   624
         ForeColor       =   12582912
         BackColor       =   -2147483648
         Caption         =   "Nome:"
         Size            =   "1101;339"
         FontName        =   "TechnicBold"
         FontEffects     =   1073741825
         FontHeight      =   156
         FontCharSet     =   2
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label4 
         Height          =   192
         Left            =   2520
         TabIndex        =   22
         Top             =   564
         Width           =   624
         ForeColor       =   12582912
         BackColor       =   -2147483648
         Caption         =   "Passw:"
         Size            =   "1101;339"
         FontName        =   "TechnicBold"
         FontEffects     =   1073741825
         FontHeight      =   156
         FontCharSet     =   2
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label5 
         Height          =   192
         Left            =   4788
         TabIndex        =   21
         Top             =   576
         Width           =   624
         ForeColor       =   12582912
         BackColor       =   -2147483648
         Caption         =   "Porta:"
         Size            =   "1101;339"
         FontName        =   "TechnicBold"
         FontEffects     =   1073741825
         FontHeight      =   156
         FontCharSet     =   2
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      Height          =   672
      Left            =   60
      TabIndex        =   3
      Top             =   5424
      Width           =   8316
      Begin ftpOCX.FTP FTP1 
         Left            =   1464
         Top             =   132
         _ExtentX        =   847
         _ExtentY        =   847
         Enabled         =   -1  'True
         ConnectionType  =   536870912
      End
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   2028
         Top             =   228
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
         Left            =   3144
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Attendere...."
         Top             =   252
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   312
         Left            =   2160
         Top             =   252
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\gianluca\Desktop\REVX\DB.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\gianluca\Desktop\REVX\DB.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "FTP"
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
      Begin VB.CommandButton btnHelp 
         Caption         =   "Help..."
         Height          =   372
         Left            =   96
         TabIndex        =   6
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnIndietro2 
         Caption         =   "<== Indietro"
         Height          =   372
         Left            =   5556
         TabIndex        =   5
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnAvanti2 
         Caption         =   "Avanti ==>"
         Height          =   372
         Left            =   6900
         TabIndex        =   4
         Top             =   216
         Width           =   1308
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3948
      Left            =   60
      TabIndex        =   2
      Top             =   504
      Width           =   8316
      Begin VB.CommandButton btnIndietro 
         Caption         =   "<<<"
         Height          =   312
         Left            =   7080
         TabIndex        =   14
         Top             =   3504
         Width           =   552
      End
      Begin VB.CommandButton btnAvanti 
         Caption         =   ">>>"
         Height          =   312
         Left            =   7704
         TabIndex        =   13
         Top             =   3504
         Width           =   552
      End
      Begin VB.CommandButton btnTest 
         Caption         =   "TEST"
         Height          =   312
         Left            =   7080
         TabIndex        =   12
         Top             =   3132
         Width           =   1164
      End
      Begin VB.CommandButton btnAccountMinus 
         Caption         =   "Elimina"
         Height          =   312
         Left            =   7116
         TabIndex        =   11
         Top             =   708
         Width           =   1140
      End
      Begin VB.CommandButton btnAccountPlus 
         Caption         =   "Aggiungi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   7104
         TabIndex        =   10
         Top             =   324
         Width           =   1140
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form5.frx":030A
         Height          =   3492
         Left            =   120
         TabIndex        =   9
         Top             =   324
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   6160
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   0   'False
         BackColor       =   16777215
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   16
         RowDividerStyle =   5
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "TechnicBold"
            Size            =   7.8
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Id"
            Caption         =   "Id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Url"
            Caption         =   "Url"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Nome"
            Caption         =   "Nome"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Password"
            Caption         =   "Password"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Dir"
            Caption         =   "Dir"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Numero"
            Caption         =   "Numero"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Porta"
            Caption         =   "Porta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   204,094
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2004,095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1716,095
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   204,094
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   504
            EndProperty
         EndProperty
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
         Caption         =   "Rev-X - Step 5: Configura gli account..."
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
         TabIndex        =   1
         Top             =   144
         Width           =   8184
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempor As Boolean

Private Sub btnAccountMinus_Click()
On Error Resume Next
Adodc1.Recordset.Delete

End Sub

Private Sub btnAccountPlus_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
btnAccountPlus.Visible = False
btnAccountMinus.Visible = False
'edit.Visible = False

End Sub

Private Sub btnAvanti_Click()
On Error Resume Next
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast

End Sub

Private Sub btnAvanti2_Click()
Me.Visible = False
Form6.Show
End Sub

Private Sub btnIndietro_Click()
On Error Resume Next
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst

End Sub

Private Sub btnIndietro2_Click()
Me.Visible = False
Form4.Show
End Sub

Private Sub btnSalva_Click()
On Error Resume Next
Adodc1.Recordset.Update
btnAccountPlus.Visible = True
btnAccountMinus.Visible = True
edit.Visible = True

End Sub

Private Sub btnTest_Click()
txtLabel.Visible = True
Me.Refresh

On Error Resume Next
FTP1.Connect App.Title, Url.Text, porta.Text, Nome.Text, Passw.Text
FTP1.UploadFile Dir.Text & "test.txt", App.Path & "\test.txt"
FTP1.DeleteSelection Dir.Text & "test.txt"
If tempor = True Then
    tempor = False
    Exit Sub
End If

MsgBox ("Errore. Controllare l'URL, il nome, la password ed il percorso")
FTP1.Disconnect
txtLabel.Visible = False

'On Error GoTo erroreftp

'Inet1.RemotePort = porta.Text
'Inet1.Url = Url.Text
'Inet1.UserName = Nome.Text
'Inet1.Password = Passw.Text
'Inet1.Execute , "CD " & Dir.Text
'Inet1.Execute , "CLOSE" ' Chiude la connessione.
'MsgBox "Ftp configurato correttamente", vbOKOnly
'txtLabel.Visible = False
'Exit Sub

End Sub

Private Sub Form_Load()

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "db.mdb" & ";Persist Security Info=False"
FTP1.ConnectionType = CONNECT_PASSIVE
FTP1.TransferType = TRANSFER_BINARY
Adodc1.Enabled = True

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
End Sub

Private Sub FTP1_Message(MsgNum As ftpOCX.MessageTypes)
  'If MsgNum = MCONNECTED Then
   ' DoList "*.*"
  If MsgNum = MUPLOADED Then
    tempor = True
    MsgBox "FTP OK. Verificare comunque la corretta digitazione del percorso."
    txtLabel.Visible = False
  End If
End Sub

