VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ACCUSED 
   Caption         =   "ACCUSED"
   ClientHeight    =   10716
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   19128
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   10716
   ScaleWidth      =   19128
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "DATA REPORT"
      Height          =   492
      Left            =   1080
      TabIndex        =   26
      Top             =   9600
      Width           =   1572
   End
   Begin VB.TextBox Text8 
      DataField       =   "FIR_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   612
      Left            =   3240
      TabIndex        =   25
      Top             =   7920
      Width           =   2052
   End
   Begin VB.TextBox Text1 
      DataField       =   "ACCUSED_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   612
      Left            =   3240
      TabIndex        =   24
      Top             =   1080
      Width           =   3852
   End
   Begin VB.PictureBox Picture1 
      Height          =   3372
      Left            =   8280
      ScaleHeight     =   3324
      ScaleWidth      =   4644
      TabIndex        =   23
      Top             =   840
      Width           =   4692
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BACK"
      Height          =   492
      Left            =   960
      TabIndex        =   22
      Top             =   8880
      Width           =   1452
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SAVE"
      Height          =   492
      Left            =   7560
      TabIndex        =   21
      Top             =   8880
      Width           =   2172
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR"
      Height          =   492
      Left            =   3120
      TabIndex        =   20
      Top             =   8880
      Width           =   1212
   End
   Begin VB.TextBox Text7 
      DataField       =   "EMAIL"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   492
      Left            =   3240
      TabIndex        =   19
      Top             =   7200
      Width           =   3852
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   612
      Left            =   16920
      Top             =   5760
      Visible         =   0   'False
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   1080
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=acccc"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "acccc"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "ACCUSED_T"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   18120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BROWSE"
      Height          =   492
      Left            =   9720
      TabIndex        =   18
      Top             =   4320
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   492
      Left            =   4680
      TabIndex        =   17
      Top             =   8880
      Width           =   2292
   End
   Begin VB.OptionButton Option1 
      Caption         =   "FEMALE"
      Height          =   372
      Index           =   1
      Left            =   5400
      TabIndex        =   16
      Top             =   5160
      Width           =   1812
   End
   Begin VB.TextBox Text6 
      DataField       =   "NIC"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   492
      Left            =   3240
      TabIndex        =   15
      Top             =   6240
      Width           =   3852
   End
   Begin VB.TextBox Text5 
      DataField       =   "GENDER"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Left            =   7800
      TabIndex        =   14
      Top             =   5160
      Width           =   1332
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MALE"
      Height          =   372
      Index           =   0
      Left            =   3240
      TabIndex        =   13
      Top             =   5160
      Width           =   1812
   End
   Begin VB.TextBox Text4 
      DataField       =   "AGE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   492
      Left            =   3240
      TabIndex        =   12
      Top             =   3960
      Width           =   3852
   End
   Begin VB.TextBox Text3 
      DataField       =   "FATHER_NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   492
      Left            =   3240
      TabIndex        =   11
      Top             =   2760
      Width           =   3852
   End
   Begin VB.TextBox Text2 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   492
      Left            =   3240
      TabIndex        =   10
      Top             =   1920
      Width           =   3852
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "F.I.R ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   9
      Top             =   8040
      Width           =   1572
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "PHOTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10080
      TabIndex        =   8
      Top             =   360
      Width           =   1932
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   7
      Top             =   7320
      Width           =   2052
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "NIC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   1812
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   2652
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "AGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   1692
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FATHER'S NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2052
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2052
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCUSED ID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCUSED INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   7092
   End
End
Attribute VB_Name = "ACCUSED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
OFFICER.Show
ACCUSED.Hide
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command4_Click()
REPORT.Text8.Text = Text1.Text
'Adodc1.Recordset.AddNew
Adodc1.Recordset.Save
Adodc1.Recordset.Update
'Command4.Enabled = True
'Form6.Show
'ACCUSED.Hide
End Sub

Private Sub Command5_Click()
'Adodc1.Recordset.Delete

End Sub

Private Sub Command6_Click()
DR_Accused.Show
End Sub

Private Sub Form_Load()
'Check1(0).Value = False
'Check1(1).Value = False
'Text4.Text = ""
Adodc1.Recordset.AddNew
'Label5.Caption = Form7.Text2.Text
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
Text5.Text = "MALE"
ElseIf Option1(1).Value = True Then
Text5.Text = "FEMALE"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 124) Or (KeyAscii = 32 Or KeyAscii = 8) Or (KeyAscii = 64 And at = 0) Then
Exit Sub
ElseIf (KeyAscii = 64) Then
at = 1
Else
Beep
KeyAscii = 0
End If

End Sub

