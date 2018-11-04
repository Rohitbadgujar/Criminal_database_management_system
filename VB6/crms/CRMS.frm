VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PETITIONER 
   Caption         =   "PETITIONER"
   ClientHeight    =   10908
   ClientLeft      =   1404
   ClientTop       =   1248
   ClientWidth     =   19272
   LinkTopic       =   "Form3"
   Picture         =   "CRMS.frx":0000
   ScaleHeight     =   10908
   ScaleWidth      =   19272
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   21240
      Top             =   12000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DATA REPORT"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   13560
      TabIndex        =   15
      Top             =   720
      Width           =   2292
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "CRMS.frx":153BE6
      Height          =   3612
      Left            =   960
      TabIndex        =   14
      Top             =   6840
      Width           =   12132
      _ExtentX        =   21400
      _ExtentY        =   6371
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   552
      Left            =   10680
      Top             =   7080
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   974
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
      Connect         =   "DSN=PETITIONER_T"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "PETITIONER_T"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "PETITIONER_T"
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command3 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   13560
      TabIndex        =   13
      Top             =   4440
      Width           =   2292
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   13560
      TabIndex        =   12
      Top             =   3360
      Width           =   2292
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   17640
      Top             =   0
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
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
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=CRMS;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=CRMS;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "PETITIONER_T"
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
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   13560
      MaskColor       =   &H0080FF80&
      TabIndex        =   11
      Top             =   2040
      Width           =   2292
   End
   Begin VB.TextBox Text5 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   6000
      TabIndex        =   10
      Top             =   2040
      Width           =   3972
   End
   Begin VB.TextBox Text4 
      DataField       =   "FATHER_NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   6000
      TabIndex        =   9
      Top             =   3240
      Width           =   3972
   End
   Begin VB.TextBox Text3 
      DataField       =   "ADDRESS"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   852
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4320
      Width           =   3972
   End
   Begin VB.TextBox Text2 
      DataField       =   "CONTACT"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   7
      Top             =   5880
      Width           =   3972
   End
   Begin VB.TextBox Text1 
      DataField       =   "NIC"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   3972
   End
   Begin VB.Image Image1 
      Height          =   972
      Left            =   20880
      Picture         =   "CRMS.frx":153BFB
      Top             =   11160
      Width           =   1596
   End
   Begin VB.Image Image2 
      Height          =   984
      Left            =   20880
      Picture         =   "CRMS.frx":15BACD
      Top             =   11160
      Width           =   1584
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT NUMBER :"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   972
      Left            =   0
      TabIndex        =   5
      Top             =   5880
      Width           =   4092
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   972
      Left            =   0
      TabIndex        =   4
      Top             =   4680
      Width           =   3012
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FATHER'S NAME"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   852
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   3972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   972
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   3012
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NIC"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GET INFORMATION ABOUT PETITIONER :"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   852
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   6852
   End
End
Attribute VB_Name = "PETITIONER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Integer
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""


End Sub
Private Sub Command2_Click()


FIR.Text10.Text = Text5.Text
FIR.Text7.Text = Text1.Text
REPORT.Text1.Text = Text1.Text
REPORT.Text2.Text = Text5.Text
'Adodc1.Recordset.MoveLast
Adodc1.Recordset.Save
Adodc1.Recordset.Update
End Sub


Private Sub Command3_Click()
VICTIM.Show
PETITIONER.Hide
End Sub

Private Sub Command4_Click()
'DataReport1.comman
'b = Val(Text1.Text)
'DataReport1.DataSource = "select * from petitioner_t where NIC=" & Text1.Text
DR_PEtitioner.Show
End Sub

Private Sub Form_Load()
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
