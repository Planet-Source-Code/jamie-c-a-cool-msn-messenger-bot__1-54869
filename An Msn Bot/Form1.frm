VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   ForeColor       =   &H8000000C&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Check My Email!"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   4320
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   840
      Picture         =   "Form1.frx":0F4E
      ScaleHeight     =   855
      ScaleWidth      =   5655
      TabIndex        =   7
      Top             =   0
      Width           =   5655
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H000000FF&
      Caption         =   "Send Message"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1560
      Picture         =   "Form1.frx":A860
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1080
      Picture         =   "Form1.frx":ABA0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   720
      Picture         =   "Form1.frx":AEFF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   360
      Picture         =   "Form1.frx":B261
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":B593
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":B8C3
      Height          =   855
      Left            =   1200
      TabIndex        =   11
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Your Current Nickname Is:"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "My Status:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents MSN As Messenger
Attribute MSN.VB_VarHelpID = -1
Dim groups As IMessengerGroups
Dim group As IMessengerGroup
Dim contacts As IMessengerContacts
Dim contact As IMessengerContact
Dim window As IMessengerConversationWnd

Private Declare Function SetParent Lib "user32" (ByVal hWndChild&, ByVal hWndNewParent&) As Long

Private Sub Command1_Click()
Messenger.MyStatus = MISTATUS_AWAY
End Sub

Private Sub Command2_Click()
Messenger.Signout
Unload Me
End Sub

Private Sub Command3_Click()
Messenger.MyStatus = MISTATUS_ONLINE
End Sub

Private Sub Command4_Click()
Messenger.MyStatus = MISTATUS_BUSY
End Sub

Private Sub Command5_Click()
Messenger.MyStatus = MISTATUS_INVISIBLE
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub

Private Sub Command7_Click()
Messenger.OpenInbox
End Sub

Private Sub Command8_Click()
Messenger.FriendlyName = Text2.Text
End Sub

Private Sub Form_Load()
Text1.Text = Messenger.MyFriendlyName
End Sub

