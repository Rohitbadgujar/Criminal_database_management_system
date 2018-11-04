VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CASER 
   ClientHeight    =   10176
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   15156
   LinkTopic       =   "Form10"
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   27783.71
   ScaleMode       =   0  'User
   ScaleWidth      =   50611.43
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "DATA REPORT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   10800
      TabIndex        =   20
      Top             =   7200
      Width           =   2052
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FINISH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   672
      Left            =   10800
      TabIndex        =   19
      Top             =   6240
      Width           =   2052
   End
   Begin VB.TextBox Text6 
      DataField       =   "OFFICER_ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   732
      Left            =   2760
      TabIndex        =   18
      Top             =   8280
      Width           =   2772
   End
   Begin VB.TextBox Text5 
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
      ForeColor       =   &H00808000&
      Height          =   732
      Left            =   2880
      TabIndex        =   17
      Top             =   6960
      Width           =   1812
   End
   Begin VB.TextBox Text4 
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
      ForeColor       =   &H00808000&
      Height          =   612
      Left            =   2760
      TabIndex        =   16
      Top             =   1200
      Width           =   1332
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form10.frx":5EEC42
      Height          =   1212
      Left            =   -2160
      TabIndex        =   15
      Top             =   9120
      Width           =   16572
      _ExtentX        =   29231
      _ExtentY        =   2138
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   360
      Top             =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   10800
      TabIndex        =   14
      Top             =   5160
      Width           =   2052
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   10800
      TabIndex        =   13
      Top             =   4080
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   10800
      Picture         =   "Form10.frx":5EEC57
      TabIndex        =   12
      Top             =   3000
      Width           =   2052
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   492
      Left            =   8880
      Top             =   240
      Width           =   3012
      _ExtentX        =   5313
      _ExtentY        =   868
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
      Connect         =   "DSN=caas"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "caas"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "CASE_T"
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
      DataField       =   "SECTION_OF_LAW"
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
      ForeColor       =   &H00808000&
      Height          =   732
      Left            =   4560
      TabIndex        =   9
      Top             =   5400
      Width           =   3372
   End
   Begin VB.TextBox Text2 
      DataField       =   "CASE_DETAIL"
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
      ForeColor       =   &H00808000&
      Height          =   1692
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Width           =   5052
   End
   Begin VB.TextBox Text1 
      DataField       =   "CASE_STATUS"
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
      ForeColor       =   &H00808000&
      Height          =   612
      Left            =   7680
      TabIndex        =   5
      Top             =   2040
      Width           =   2652
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000013&
      Caption         =   "CLOSE"
      ForeColor       =   &H000000FF&
      Height          =   612
      Index           =   1
      Left            =   5400
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   2040
      Width           =   2052
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000013&
      Caption         =   "OPEN"
      ForeColor       =   &H000000FF&
      Height          =   612
      Index           =   0
      Left            =   3120
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   2040
      Width           =   2052
   End
   Begin VB.Image Image1 
      Height          =   972
      Left            =   0
      Picture         =   "Form10.frx":78A9CE
      Top             =   0
      Width           =   1596
   End
   Begin VB.Image Image2 
      Height          =   984
      Left            =   0
      Picture         =   "Form10.frx":7928A0
      Top             =   0
      Width           =   1584
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "OFFICER ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   0
      TabIndex        =   11
      Top             =   8280
      Width           =   2892
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "FIR ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   1332
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SECTION OF LAW"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   852
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   4092
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE DETAIL"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   2412
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE STATUS "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE ID"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2532
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CASE DETAIL"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1692
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   8772
   End
End
Attribute VB_Name = "CASER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Adodc1.Recordset.Save
'Adodc1.Recordset.Update

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
REPORT.Show
CASER.Hide
End Sub

Private Sub Command5_Click()
DR_case.Show
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
Text1.Text = "OPEN"
ElseIf Option1(1).Value = True Then
Text1.Text = "CLOSED"
End If
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
