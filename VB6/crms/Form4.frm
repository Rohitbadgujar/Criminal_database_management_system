VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form VICTIM 
   Caption         =   "VICTIM"
   ClientHeight    =   12420
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   19668
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   12420
   ScaleWidth      =   19668
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "DATA REPORT"
      Height          =   612
      Left            =   10320
      TabIndex        =   13
      Top             =   2280
      Width           =   2412
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":4C261
      Height          =   3132
      Left            =   240
      TabIndex        =   12
      Top             =   7200
      Width           =   13452
      _ExtentX        =   23728
      _ExtentY        =   5525
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
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
      Caption         =   "NEXT"
      Height          =   732
      Left            =   10320
      TabIndex        =   11
      Top             =   5040
      Width           =   2412
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   732
      Left            =   10320
      TabIndex        =   10
      Top             =   4080
      Width           =   2412
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   732
      Left            =   10320
      TabIndex        =   9
      Top             =   3120
      Width           =   2412
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   13800
      Top             =   2280
      Visible         =   0   'False
      Width           =   3012
      _ExtentX        =   5313
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
      Connect         =   "DSN=VICTIM_T"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "VICTIM_T"
      OtherAttributes =   ""
      UserName        =   "CRMS"
      Password        =   "mat"
      RecordSource    =   "VICTIM_T"
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
   Begin VB.TextBox Text4 
      DataField       =   "CONTACT"
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
      ForeColor       =   &H00000000&
      Height          =   492
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   8
      Top             =   6000
      Width           =   5292
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
      ForeColor       =   &H00000000&
      Height          =   612
      Left            =   3360
      TabIndex        =   7
      Top             =   2760
      Width           =   5292
   End
   Begin VB.TextBox Text2 
      DataField       =   "ADDRESS"
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
      ForeColor       =   &H00000000&
      Height          =   492
      Left            =   3360
      TabIndex        =   6
      Top             =   4200
      Width           =   5292
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H00000000&
      Height          =   492
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   5292
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   360
      TabIndex        =   4
      Top             =   6120
      Width           =   2772
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FATHER'S NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   2772
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
      Height          =   612
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   2652
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2412
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VICTIM  INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   7080
      TabIndex        =   0
      Top             =   480
      Width           =   3972
   End
End
Attribute VB_Name = "VICTIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
FIR.Text8.Text = Text1.Text
FIR.Text9.Text = Text2.Text
REPORT.Text3.Text = Text1.Text
REPORT.Text4.Text = Text2.Text
Adodc1.Recordset.Save
Adodc1.Recordset.Update
'ADO.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub

Private Sub Command3_Click()

FIR.Show
VICTIM.Hide
End Sub

Private Sub Command4_Click()
DR_Victim.Show
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 124) Or (KeyAscii = 32 Or KeyAscii = 8) Or (KeyAscii = 64 And at = 0) Then
Exit Sub
ElseIf (KeyAscii = 64) Then
at = 1
Else
Beep
KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Then
Exit Sub
Else
Beep
Keyscii = 1
End If
End Sub

