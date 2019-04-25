VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19605
   LinkTopic       =   "Form3"
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19600
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line4 
      BorderColor     =   &H0000FFFF&
      X1              =   2279.418
      X2              =   2279.418
      Y1              =   1440
      Y2              =   5280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      X1              =   1799.541
      X2              =   1799.541
      Y1              =   1440
      Y2              =   5280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   1319.663
      X2              =   1319.663
      Y1              =   1440
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   839.786
      X2              =   839.786
      Y1              =   1440
      Y2              =   5280
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9602
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End             'escÍË³ö

End Sub

