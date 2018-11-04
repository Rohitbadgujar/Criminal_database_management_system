VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form OFFICER 
   Caption         =   "INVESTIGATION OFFICER"
   ClientHeight    =   9540
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   16716
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   9540
   ScaleWidth      =   16716
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "DATA REPORT"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7200
      Width           =   2772
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   14520
      Top             =   1440
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
      Width           =   2532
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Width           =   2532
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7200
      Width           =   2772
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   15120
      Top             =   7440
      Visible         =   0   'False
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   656
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
      Connect         =   "DSN=inv"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "inv"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "INVESTIGATION_OFFICER_T"
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
   Begin VB.TextBox Text3 
      DataField       =   "RANK"
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
      ForeColor       =   &H00C0C000&
      Height          =   612
      Left            =   2760
      TabIndex        =   6
      Top             =   4440
      Width           =   3252
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
      ForeColor       =   &H00C0C000&
      Height          =   612
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   3252
   End
   Begin VB.TextBox Text1 
      DataField       =   "OFFICER_ID"
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
      ForeColor       =   &H00C0C000&
      Height          =   612
      Left            =   2760
      TabIndex        =   4
      Top             =   1560
      Width           =   3252
   End
   Begin VB.Image Image2 
      Height          =   984
      Left            =   14160
      Picture         =   "Form9.frx":5EEC42
      Top             =   600
      Width           =   1584
   End
   Begin VB.Image Image1 
      Height          =   972
      Left            =   14160
      Picture         =   "Form9.frx":5F6B5C
      Top             =   600
      Width           =   1596
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RANK"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1092
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   2412
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1092
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   3132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OFFICIER ID"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   732
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2772
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INVESTIGATION OFFICIER"
      BeginProperty Font 
         Name            =   "Microsoft JhengHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1332
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   6972
   End
End
Attribute VB_Name = "OFFICER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Command1_Click()
'Unload Me
CASER.Show
OFFICER.Hide
End Sub

Private Sub Command2_Click()
'CASER.Label10.Caption = Text1.Text
CASER.Text6.Text = Text1.Text
REPORT.Text5.Text = Text1.Text
Adodc1.Recordset.Save
Adodc1.Recordset.Update
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Command4_Click()
DR_Officer.Show
End Sub

Private Sub Form_Load()
a = 1
Adodc1.Recordset.AddNew
End Sub

Private Sub Timer1_Timer()
If (Image1.Visible = True) Then
Image1.Visible = False
Image2.Visible = True
ElseIf (Image2.Visible = True) Then
Image2.Visible = False
Image1.Visible = True
End If




End Sub
