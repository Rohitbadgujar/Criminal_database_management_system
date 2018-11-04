VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form REPORT 
   Caption         =   "Form11"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3624
   LinkTopic       =   "Form11"
   Picture         =   "REPORT.frx":0000
   ScaleHeight     =   2340
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12840
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FINISH"
      Height          =   612
      Left            =   14400
      TabIndex        =   18
      Top             =   9600
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   612
      Left            =   16800
      TabIndex        =   17
      Top             =   9600
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1092
      Left            =   14160
      Top             =   1440
      Visible         =   0   'False
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   1926
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
      Connect         =   "DSN=REPORT"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "REPORT"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "REPORT_T"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text8 
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
      Height          =   492
      Left            =   5760
      TabIndex        =   15
      Top             =   9120
      Width           =   1452
   End
   Begin VB.TextBox Text7 
      DataField       =   "CASE_ID"
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
      Height          =   492
      Left            =   5760
      TabIndex        =   14
      Top             =   7920
      Width           =   1212
   End
   Begin VB.TextBox Text6 
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
      Height          =   492
      Left            =   5760
      TabIndex        =   13
      Top             =   6720
      Width           =   972
   End
   Begin VB.TextBox Text5 
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
      Height          =   492
      Left            =   5760
      TabIndex        =   12
      Top             =   5640
      Width           =   1932
   End
   Begin VB.TextBox Text4 
      DataField       =   "VICTIM_ADDRESS"
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
      Height          =   852
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4440
      Width           =   3372
   End
   Begin VB.TextBox Text3 
      DataField       =   "VICTIM_NAME"
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
      Height          =   492
      Left            =   5520
      TabIndex        =   10
      Top             =   3480
      Width           =   3372
   End
   Begin VB.TextBox Text2 
      DataField       =   "PETITIONER_NAME"
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
      Height          =   492
      Left            =   5520
      TabIndex        =   9
      Top             =   2400
      Width           =   3372
   End
   Begin VB.TextBox Text1 
      DataField       =   "PETITITIONER_NIC"
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
      Height          =   492
      Left            =   5520
      TabIndex        =   8
      Top             =   1440
      Width           =   1812
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE REPORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   732
      Left            =   4560
      TabIndex        =   16
      Top             =   120
      Width           =   6252
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCUSED_ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   732
      Left            =   600
      TabIndex        =   7
      Top             =   9240
      Width           =   2292
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE_ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   852
      Left            =   600
      TabIndex        =   6
      Top             =   8040
      Width           =   2292
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "FIR ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1092
      Left            =   600
      TabIndex        =   5
      Top             =   6960
      Width           =   1932
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "INVESTIGATION_OFFICER_IT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   852
      Left            =   600
      TabIndex        =   4
      Top             =   5640
      Width           =   3612
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "VICTIM_ADDRESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   852
      Left            =   600
      TabIndex        =   3
      Top             =   4440
      Width           =   2172
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "VICTIM_NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   732
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   2172
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PETITIONER_NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   612
      Left            =   600
      TabIndex        =   1
      Top             =   2520
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PETITIONER_NIC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   2172
   End
End
Attribute VB_Name = "REPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo cancel
CommonDialog1.PrinterDefault = True
CommonDialog1.Flags = 0&
CommonDialog1.ShowPrinter
cancel:
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

