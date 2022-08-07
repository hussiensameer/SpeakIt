VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Speak It"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpeak 
      Caption         =   "Speak It"
      Height          =   480
      Left            =   3113
      TabIndex        =   1
      Top             =   5610
      Width           =   1275
   End
   Begin VB.TextBox Text 
      Height          =   5430
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   7485
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSpeak_Click()
    SpeakString Text.Text
End Sub
