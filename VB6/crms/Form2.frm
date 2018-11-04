VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MENU 
   Caption         =   "MENU"
   ClientHeight    =   10344
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   18528
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10344
   ScaleWidth      =   18528
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H00808080&
      Caption         =   "NEW FIR"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   2040
      Width           =   2172
   End
   Begin VB.CommandButton Command6 
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
      Height          =   732
      Left            =   7680
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CASE ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   5400
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   21480
      Top             =   13320
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11760
      Top             =   5520
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   732
      Left            =   15120
      Top             =   8160
      Visible         =   0   'False
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1291
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "DSN=FI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "FI"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "select * from fir_t;"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   492
      Left            =   14880
      Top             =   6840
      Visible         =   0   'False
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   868
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "DSN=ACC"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ACC"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "select * from case_t;"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "Form2.frx":97F37
      Height          =   1452
      Left            =   1680
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   15252
      _ExtentX        =   26903
      _ExtentY        =   2561
      _Version        =   393216
      ForeColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   32
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
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form2.frx":97F4C
      Height          =   2412
      Left            =   1680
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   14292
      _ExtentX        =   25210
      _ExtentY        =   4255
      _Version        =   393216
      ForeColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   32
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
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "PETITION NIC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   2760
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":97F61
      Height          =   2172
      Left            =   1680
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   14652
      _ExtentX        =   25845
      _ExtentY        =   3831
      _Version        =   393216
      ForeColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   32
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
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
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
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   852
      Left            =   16560
      Top             =   8640
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1503
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "DSN=fg"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "fg"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "select * from petitioner_t"
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
      BackColor       =   &H00808080&
      Caption         =   "SEARCH"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   2052
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000D&
      Caption         =   "PETITIONER"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   13.8
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000D&
      Caption         =   "VICTIM"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   13.8
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   6600
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   3012
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "CONTINUE"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   22.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   3360
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   6252
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   22.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   9000
      Picture         =   "Form2.frx":97F76
      TabIndex        =   0
      Top             =   2040
      Width           =   2412
   End
   Begin VB.Image Image1 
      Height          =   972
      Left            =   21120
      Picture         =   "Form2.frx":12FEAD
      Top             =   12480
      Width           =   1596
   End
   Begin VB.Image Image2 
      Height          =   984
      Left            =   21120
      Picture         =   "Form2.frx":137D7F
      Top             =   12480
      Width           =   1584
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   28.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1212
      Left            =   18120
      TabIndex        =   12
      Top             =   2160
      Width           =   4812
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   732
      Left            =   19680
      TabIndex        =   11
      Top             =   1440
      Width           =   3852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NEW FIR AS :"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   852
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   11292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MENU "
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1692
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   6972
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, c As Integer
 Dim b As String

Private Sub Command1_Click()
DataGrid1.Visible = False
DataGrid2.Visible = False
DataGrid3.Visible = False
Command3.Visible = True
Command5.Visible = True
Command6.Visible = True
End Sub

Private Sub Command2_Click()

'VICTIM.Adodc1.Recordset.Close


If Option1.Value Then
PETITIONER.Show
MENU.Hide
Else
VICTIM.Show
MENU.Hide
End If
End Sub

Private Sub Command3_Click()
Command3.Visible = False
Command5.Visible = False
Command6.Visible = False
a = Val(InputBox("ENTER PETITIONER NIC NO."))
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from petitioner_t where NIC=" & a
Adodc1.Refresh
DataGrid1.Visible = True
DataGrid2.Visible = False
DataGrid3.Visible = False

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Command3.Visible = False
Command5.Visible = False
Command6.Visible = False
b = Val(InputBox("ENTER CASE ID"))
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select * from CASE_t where CASE_id=" & b
Adodc2.Refresh
DataGrid2.Visible = True
DataGrid1.Visible = False
DataGrid3.Visible = False
End Sub

Private Sub Command6_Click()
Command3.Visible = False
Command5.Visible = False
Command6.Visible = False
c = Val(InputBox("ENTER FIR ID"))
Adodc3.CommandType = adCmdText
Adodc3.RecordSource = "select * from fir_t where fir_id=" & c
Adodc3.Refresh
DataGrid3.Visible = True
DataGrid1.Visible = False
DataGrid2.Visible = False
End Sub

Private Sub Command8_Click()
Command3.Visible = False
Command5.Visible = False
Command6.Visible = False
Command1.Value = False
Command4.Visible = False
Option1.Visible = True
Option2.Visible = True
Command2.Visible = True
Command4.Visible = True
Label2.Visible = True
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Time
End Sub

Private Sub Timer2_Timer()
If (Image1.Visible = True) Then
Image1.Visible = False
Image2.Visible = True
ElseIf (Image2.Visible = True) Then
Image2.Visible = False
Image1.Visible = True
End If



End Sub
