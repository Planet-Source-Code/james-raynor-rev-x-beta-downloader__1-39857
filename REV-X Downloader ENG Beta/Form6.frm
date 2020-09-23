VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   ClientHeight    =   72
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   72
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   72
   ScaleWidth      =   72
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   444
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   8316
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Rev-X - Step 6: Seleziona quali account utilizzare..."
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
   Begin VB.Frame Frame3 
      Height          =   3948
      Left            =   60
      TabIndex        =   4
      Top             =   504
      Width           =   8316
      Begin VB.CommandButton btnAggiungi 
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
         Left            =   2436
         TabIndex        =   18
         Top             =   3552
         Width           =   1320
      End
      Begin VB.CommandButton btnCancella 
         Caption         =   "Cancella"
         Height          =   312
         Left            =   864
         TabIndex        =   17
         Top             =   3552
         Width           =   1320
      End
      Begin VB.CommandButton btnAvanti 
         Caption         =   ">>>"
         Height          =   312
         Left            =   4044
         TabIndex        =   16
         Top             =   3552
         Width           =   552
      End
      Begin VB.CommandButton btnIndietro 
         Caption         =   "<<<"
         Height          =   312
         Left            =   168
         TabIndex        =   15
         Top             =   3552
         Width           =   552
      End
      Begin VB.TextBox Numero 
         DataField       =   "Numero"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   1008
         TabIndex        =   14
         Top             =   3000
         Visible         =   0   'False
         Width           =   888
      End
      Begin VB.TextBox Id_Ftp 
         DataField       =   "Id"
         DataSource      =   "Adodc1"
         Height          =   336
         Left            =   4008
         TabIndex        =   12
         Text            =   "1"
         Top             =   3252
         Visible         =   0   'False
         Width           =   468
      End
      Begin VB.TextBox Url 
         DataField       =   "Url"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   3072
         TabIndex        =   11
         Top             =   3288
         Visible         =   0   'False
         Width           =   888
      End
      Begin VB.TextBox Nome 
         DataField       =   "Nome"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   2160
         TabIndex        =   10
         Top             =   3288
         Visible         =   0   'False
         Width           =   888
      End
      Begin VB.TextBox Passw 
         DataField       =   "Password"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   1236
         TabIndex        =   9
         Top             =   3312
         Visible         =   0   'False
         Width           =   888
      End
      Begin VB.TextBox txtDir 
         DataField       =   "Dir"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   288
         TabIndex        =   8
         Top             =   3276
         Visible         =   0   'False
         Width           =   888
      End
      Begin VB.TextBox porta 
         DataField       =   "Porta"
         DataSource      =   "Adodc1"
         Height          =   264
         Left            =   312
         TabIndex        =   7
         Top             =   2988
         Visible         =   0   'False
         Width           =   612
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form6.frx":030A
         Height          =   3276
         Left            =   168
         TabIndex        =   13
         Top             =   216
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   5779
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
               ColumnWidth     =   1595,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   996,095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   996,095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   4199,811
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   672
      Left            =   60
      TabIndex        =   0
      Top             =   4524
      Width           =   8316
      Begin VB.CommandButton btnAvanti2 
         Caption         =   "Avanti ==>"
         Height          =   372
         Left            =   6900
         TabIndex        =   3
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnIndietro2 
         Caption         =   "<== Indietro"
         Height          =   372
         Left            =   5556
         TabIndex        =   2
         Top             =   216
         Width           =   1308
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "Help..."
         Height          =   372
         Left            =   96
         TabIndex        =   1
         Top             =   216
         Width           =   1308
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   312
         Left            =   1524
         Top             =   252
         Visible         =   0   'False
         Width           =   2088
         _ExtentX        =   3683
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
         Caption         =   "Lista Account"
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
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAggiungi_Click()
Aggiungi_Ftp.Show
End Sub

Private Sub btnAvanti_Click()
On Error Resume Next
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast

End Sub

Private Sub btnAvanti2_Click()
Me.Visible = False
Form7.Show
End Sub

Private Sub btnCancella_Click()
On Error Resume Next
Adodc1.Recordset.Delete
End Sub

Private Sub btnIndietro_Click()
On Error Resume Next
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst

End Sub

Private Sub btnIndietro2_Click()
Me.Visible = False
Form5.Show
End Sub

Private Sub Form_Load()
'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "db.mdb" & ";Persist Security Info=False"
'Adodc1.Enabled = True

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
