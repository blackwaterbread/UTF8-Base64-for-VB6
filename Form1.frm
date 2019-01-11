VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command2 
      Caption         =   "Decode"
      Height          =   585
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encode"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Text            =   "Output"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "Input"
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Text2.Text = Base64Encode(Text1.Text)

End Sub

Private Sub Command2_Click()

Text2.Text = Base64Decode(Text1.Text)

End Sub
