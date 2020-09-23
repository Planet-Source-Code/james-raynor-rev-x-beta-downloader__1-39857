VERSION 5.00
Begin VB.Form frmfinalsplit 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2292
   LinkTopic       =   "Form2"
   ScaleHeight     =   300
   ScaleWidth      =   2292
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Ricostruzione file....."
      Top             =   12
      Width           =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1116
      Top             =   792
   End
End
Attribute VB_Name = "frmfinalsplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = txtLabel.Width
Me.Height = txtLabel.Height
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim err_descr As String

'If Not ReassembleFile(dirD.Text & "\" & fileD.Text & ".tpl", False, txtTemp.Text & "\" & fileD.Text) Then
    If Not ReassembleFile(tempsplit1, False, tempsplit2) Then
        MsgBox err_descr
        okdel = 1
    Else
        MsgBox "File creato correttamente"
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

Unload Me
Form1.Show
End Sub


