VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Aggiungi_Ftp 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aggiungi FTP"
   ClientHeight    =   2904
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7260
   Icon            =   "Aggiungi_Ftp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2904
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Aggiungi FTP"
      BeginProperty Font 
         Name            =   "TechnicBold"
         Size            =   7.8
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2748
      Left            =   36
      TabIndex        =   0
      Top             =   48
      Width           =   7080
      Begin VB.CommandButton Command1 
         Caption         =   "Chiudi"
         Height          =   312
         Left            =   6060
         TabIndex        =   12
         Top             =   2340
         Width           =   960
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
         Left            =   120
         TabIndex        =   9
         Top             =   2316
         Width           =   960
      End
      Begin VB.CommandButton btnIndietro 
         Caption         =   "<<<"
         Height          =   312
         Left            =   1260
         TabIndex        =   11
         Top             =   2316
         Width           =   552
      End
      Begin VB.CommandButton btnAvanti 
         Caption         =   ">>>"
         Height          =   312
         Left            =   1872
         TabIndex        =   10
         Top             =   2316
         Width           =   552
      End
      Begin VB.TextBox porta 
         DataField       =   "Porta"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   4512
         TabIndex        =   8
         Top             =   1848
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.TextBox Numero 
         Height          =   288
         Left            =   252
         TabIndex        =   7
         Top             =   1524
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox Id_Ftp 
         DataField       =   "Id"
         DataSource      =   "Adodc1"
         Height          =   336
         Left            =   3984
         TabIndex        =   6
         Text            =   "1"
         Top             =   1836
         Visible         =   0   'False
         Width           =   468
      End
      Begin VB.TextBox Url 
         DataField       =   "Url"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   3048
         TabIndex        =   5
         Top             =   1872
         Visible         =   0   'False
         Width           =   888
      End
      Begin VB.TextBox Nome 
         DataField       =   "Nome"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   2136
         TabIndex        =   4
         Top             =   1872
         Visible         =   0   'False
         Width           =   888
      End
      Begin VB.TextBox Passw 
         DataField       =   "Password"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   1212
         TabIndex        =   3
         Top             =   1860
         Visible         =   0   'False
         Width           =   888
      End
      Begin VB.TextBox Dir 
         DataField       =   "Dir"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   264
         TabIndex        =   2
         Top             =   1860
         Visible         =   0   'False
         Width           =   888
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   312
         Left            =   1176
         Top             =   1488
         Visible         =   0   'False
         Width           =   5832
         _ExtentX        =   10287
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
         Caption         =   "FTP Creati"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Aggiungi_Ftp.frx":030A
         Height          =   2004
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   3535
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
         ColumnCount     =   5
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   204,094
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2496,189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   996,095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   996,095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2148,094
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Aggiungi_Ftp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnAccountPlus_Click()
On Error Resume Next
Dim tempor As Long

Form6.Adodc1.Recordset.MoveLast
tempor = Form6.Adodc1.Recordset.AbsolutePosition
If tempor = "-1" Then tempor = 0
tempor = tempor + 1
Numero.Text = tempor

If Numero.Text = "49" Then
    MsgBox ("Limite massimo di FTP raggiunto")
    Exit Sub
End If


Form6.Adodc1.Recordset.AddNew
Form6.Url.Text = Url.Text
Form6.Passw.Text = Passw.Text
Form6.Nome.Text = Nome.Text
Form6.txtDir.Text = Dir.Text
Form6.Numero.Text = Numero.Text
Form6.porta.Text = porta.Text
Form6.Adodc1.Recordset.Update
End Sub

Private Sub btnAvanti_Click()
On Error Resume Next
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
End Sub

Private Sub btnIndietro_Click()
On Error Resume Next
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
End Sub

Private Sub chameleonButton2_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "db.mdb" & ";Persist Security Info=False"
Adodc1.Enabled = True
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub
