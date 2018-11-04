VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   10200
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   19464
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   10200
   ScaleWidth      =   19464
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   1332
      Left            =   10080
      Top             =   6000
      Visible         =   0   'False
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   2350
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
      Connect         =   "DSN=ash"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ash"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "ACCUSED_T"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form6.frx":70CEC
      Height          =   3372
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   6852
      _ExtentX        =   12086
      _ExtentY        =   5948
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
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR"
      Height          =   612
      Left            =   12960
      TabIndex        =   6
      Top             =   3120
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   612
      Left            =   15120
      TabIndex        =   5
      Top             =   3120
      Width           =   1452
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1092
      Left            =   15960
      Top             =   360
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
      RecordSource    =   "ACCUSED_ADDRESS_T"
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
      Left            =   17040
      TabIndex        =   4
      Top             =   3120
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      DataField       =   "ADDRESS"
      DataSource      =   "Adodc1"
      Height          =   1452
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3240
      Width           =   4332
   End
   Begin VB.Label Label4 
      Height          =   732
      Left            =   2400
      TabIndex        =   8
      Top             =   1560
      Width           =   2052
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   972
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCUSED ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   732
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCUSED ADDRESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   732
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   4212
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form8.Show
Form6.Hide
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

Private Sub Text1_Change()
'Text1.Text = Form5.Text1.Text
End Sub

Private Sub Label4_Click()
Label4.Caption = Form5.Text1.Text
End Sub
