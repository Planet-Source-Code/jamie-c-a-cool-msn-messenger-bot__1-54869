VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Send An Instant Message!"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3405
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Text            =   "Your Message"
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send!!!"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "User@hotmail.com"
      Top             =   2160
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "This Feature Is Unavailable Yet! Please See Our Site For Updates!"
End Sub
