VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   11076
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   18624
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   11076
   ScaleWidth      =   18624
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR"
      Height          =   732
      Left            =   1200
      TabIndex        =   7
      Top             =   7080
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "SAVE"
      Height          =   732
      Left            =   4080
      TabIndex        =   6
      Top             =   7080
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1692
      Left            =   11160
      Top             =   3480
      Width           =   2412
      _ExtentX        =   4255
      _ExtentY        =   2985
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
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=CRMS;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=CRMS;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "ACCUSED_CONTACT_T"
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
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   612
      Left            =   7200
      TabIndex        =   5
      Top             =   7080
      Width           =   1692
   End
   Begin VB.TextBox Text2 
      DataField       =   "PHONE_NO"
      DataSource      =   "Adodc1"
      Height          =   612
      Left            =   2520
      TabIndex        =   4
      Top             =   4440
      Width           =   2892
   End
   Begin VB.TextBox Text1 
      DataField       =   "ACCUSED_ID"
      DataSource      =   "Adodc1"
      Height          =   612
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   2292
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NO"
      ForeColor       =   &H8000000E&
      Height          =   732
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   3972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCUSED ID"
      ForeColor       =   &H8000000E&
      Height          =   1092
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   4332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCUSED CONTACT INFORMATION"
      ForeColor       =   &H8000000E&
      Height          =   1452
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   6492
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form7.Show
Form8.Hide
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Save

Adodc1.Recordset.Update

End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
End Sub

Private Sub Text2_Change()
If (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 124) Or (KeyAscii = 32 Or KeyAscii = 8) Or (KeyAscii = 64 And at = 0) Then
Exit Sub
ElseIf (KeyAscii = 64) Then
at = 1
Else
Beep
KeyAscii = 0
End If
End Sub
