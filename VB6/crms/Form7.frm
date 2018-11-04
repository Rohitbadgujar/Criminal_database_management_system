VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FIR 
   Caption         =   "FIR"
   ClientHeight    =   10224
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   18876
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   10224
   ScaleWidth      =   18876
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "DATA REPORT"
      Height          =   732
      Left            =   16080
      TabIndex        =   28
      Top             =   7440
      Width           =   2292
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00404000&
      Height          =   492
      Left            =   4800
      TabIndex        =   27
      Top             =   9720
      Width           =   1452
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   4800
      TabIndex        =   26
      Top             =   8040
      Width           =   1332
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   4800
      TabIndex        =   25
      Top             =   7080
      Width           =   1692
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PETITIONER_ID"
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
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   4800
      TabIndex        =   24
      Top             =   5880
      Width           =   2052
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "TIME_LODGED"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   4800
      TabIndex        =   23
      Top             =   3840
      Width           =   1212
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   4800
      TabIndex        =   22
      Top             =   960
      Width           =   2532
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      ItemData        =   "Form7.frx":2E6F90
      Left            =   6240
      List            =   "Form7.frx":2E6F9A
      TabIndex        =   21
      Top             =   3960
      Width           =   972
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "Form7.frx":2E6FA6
      Left            =   6240
      List            =   "Form7.frx":2E6FB0
      TabIndex        =   20
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808000&
      Caption         =   "SAVE"
      Height          =   852
      Left            =   13320
      MaskColor       =   &H00808000&
      TabIndex        =   19
      Top             =   8760
      Width           =   2292
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808000&
      Caption         =   "CLEAR"
      Height          =   852
      Left            =   10320
      MaskColor       =   &H00808000&
      TabIndex        =   18
      Top             =   8760
      Width           =   2532
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "NEXT"
      Height          =   852
      Left            =   16080
      MaskColor       =   &H00808000&
      TabIndex        =   17
      Top             =   8760
      Width           =   2412
   End
   Begin MSACAL.Calendar Calendar1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   2172
      Left            =   11400
      TabIndex        =   16
      Top             =   1560
      Width           =   5772
      _Version        =   524288
      _ExtentX        =   10181
      _ExtentY        =   3831
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2013
      Month           =   9
      Day             =   26
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00404000&
      Height          =   492
      Left            =   4800
      TabIndex        =   15
      Top             =   8880
      Width           =   2532
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   852
      Left            =   15240
      Top             =   5520
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   1503
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
      Connect         =   "DSN=FIR_T"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "FIR_T"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "FIR_T"
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
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      DataField       =   "DATE_LODGED"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "m/d/yy h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      ForeColor       =   &H00404000&
      Height          =   492
      Left            =   4800
      TabIndex        =   3
      Top             =   4920
      Width           =   1572
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      DataField       =   "INCIDENT_PLACE"
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
      ForeColor       =   &H00404000&
      Height          =   612
      Left            =   4800
      TabIndex        =   2
      Top             =   3000
      Width           =   2532
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      DataField       =   "INCIDENT_TIME"
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
      ForeColor       =   &H00404000&
      Height          =   492
      Left            =   4800
      TabIndex        =   1
      Top             =   2400
      Width           =   1092
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      DataField       =   "INCIDENT_DATE"
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
      ForeColor       =   &H00404000&
      Height          =   492
      Left            =   4800
      TabIndex        =   0
      Top             =   1680
      Width           =   2532
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "PETITIONER NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   612
      Left            =   120
      TabIndex        =   14
      Top             =   9840
      Width           =   2532
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   240
      TabIndex        =   13
      Top             =   9120
      Width           =   2412
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "VICTIME ADDRESS"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   120
      TabIndex        =   12
      Top             =   8280
      Width           =   2652
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "VICTIM NAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   1812
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PETITIONER ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   3132
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE LODGE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   2172
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time LODGE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   1452
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "INCIDENT PLACE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   3252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "INCIDENT TIME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "INCIDENT DATE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   3732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FIR ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1332
   End
End
Attribute VB_Name = "FIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d, c As Date
Dim a As Integer
Private Sub Calendar1_Click()
'If (a = 2) Then
c = DateSerial(Calendar1.Year, Calendar1.Month, Calendar1.Day)
Text3.Text = c

'ElseIf a = 3 Then
'd = DateSerial(Calendar1.Year, Calendar1.Month, Calendar1.Day)
'Text3.Text = d
'End If
End Sub

Private Sub Command1_Click()
ACCUSED.Show
FIR.Hide
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
'Text12.Text = ""
End Sub

Private Sub Command3_Click()
ACCUSED.Text8.Text = Text2.Text
CASER.Text5.Text = Text2.Text
CASER.Text4.Text = Text11.Text
REPORT.Text6.Text = Text2.Text
REPORT.Text7.Text = Text11.Text
Adodc1.Recordset.Save
Adodc1.Recordset.Update
'ACCUSED.Show
'FIR.Hide
End Sub

Private Sub Command4_Click()
'Adodc2.CommandType = adCmdText
'Adodc2.Recordset = "select to_char(date_lodge,'DD-MON-YY HH24:MI:SS') "" from dual"
DR_Fir.Show
End Sub

Private Sub Form_Load()

'Text7.Text = Label13.Caption
' Text8.Text = Label14.Caption
'Text9.Text = Label15.Caption
' Text10.Text = Label12.Caption

'Adodc2.CommandType = adCmdText
Adodc1.Recordset.AddNew
'Label12.Caption = Form3.Text5.Text
'Label13.Caption = Form3.Text1.Text

'Label15.Caption = Form4.Text2.Text
'Label16.Caption = ACCUSED.Text8.Text
End Sub

Private Sub Text1_Click()
Text1.Text = Now
End Sub

Private Sub Text3_Click()
a = 2
End Sub

Private Sub Text6_Click()
Text6.Text = Now
End Sub
