VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form LOGIN 
   BackColor       =   &H0080C0FF&
   Caption         =   "LOGIN"
   ClientHeight    =   8268
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   16044
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4.87053e5
   ScaleMode       =   0  'User
   ScaleWidth      =   1.86102e6
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   21360
      Top             =   12000
   End
   Begin VB.PictureBox Picture1 
      Height          =   3852
      Left            =   7920
      Picture         =   "Form1.frx":697842
      ScaleHeight     =   3804
      ScaleWidth      =   2844
      TabIndex        =   8
      Top             =   2520
      Width           =   2892
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   600
      MaskColor       =   &H80000013&
      TabIndex        =   6
      Top             =   5640
      Width           =   3012
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   4440
      MaskColor       =   &H000000FF&
      TabIndex        =   7
      Top             =   5640
      Width           =   2412
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3720
      Width           =   3012
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4200
      TabIndex        =   4
      Top             =   2520
      Width           =   3012
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
      Height          =   372
      Left            =   840
      TabIndex        =   9
      Top             =   6840
      Width           =   612
      URL             =   $"Form1.frx":69BEDC
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1080
      _cy             =   656
   End
   Begin VB.Image Image1 
      Height          =   972
      Left            =   21000
      Picture         =   "Form1.frx":69BFA3
      Top             =   11160
      Width           =   1596
   End
   Begin VB.Image Image2 
      Height          =   984
      Left            =   21000
      Picture         =   "Form1.frx":6A3E75
      Top             =   11160
      Width           =   1584
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD :"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   3252
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME :"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   3252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN PAGE "
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1332
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   10812
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   255
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Username As String                       'Username Variable
Dim Password As String                       'Password Variable
 
 Dim Username1 As String                       'Username Variable
Dim Password1 As String                       'Password Variable
 
Private Sub Command2_Click()
 
Unload Me      'Closes the program
 
End Sub

Private Sub Command1_Click()
 
Username = "COM"              'What the Username Variable is
Password = "PUT"              'What the Password Variable is
 
Username1 = "POLICE"              'What the Username Variable is
Password1 = "POLICE"              'What the Password Variable is
 
'If Statement Starting Here
'It States that if the username is = to text box and password is = password text box, then it should login, else display error message!
If Username = Text1 And Password = Text2 Then
 
           MENU.Show    'Shows the Login Success Form
           LOGIN.Hide       'Hides the Login Screen
 
Else
 
'A Message box that displays you have entered the wrong username and password
           
MsgBox ("You have entered the wrong username and password")
 
End If


 
End Sub

Private Sub Form_Load()
'WindowsMediaPlayer1.Close

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
