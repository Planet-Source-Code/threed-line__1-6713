VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test form"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin Project1.ThreeDLine ThreeDLine2 
      Height          =   4275
      Left            =   6720
      TabIndex        =   4
      Top             =   240
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   7541
      Orientation     =   1
   End
   Begin Project1.ThreeDLine ThreeDLine1 
      Height          =   45
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   79
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "3D Line Source code test form. This was written in VB 6 by Tom Garratt."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label Label 
      Caption         =   "Wow, now its vertical!"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "Look, Horizontal!"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

End Sub
