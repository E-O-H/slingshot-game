VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   DrawWidth       =   5
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Left            =   1680
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   0
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   18000
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   17400
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      Height          =   255
      Left            =   18960
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "0"
      Height          =   255
      Left            =   18720
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape bullet 
      BorderColor     =   &H8000000B&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   100
      Left            =   240
      Shape           =   3  'Circle
      Top             =   2640
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000B&
      X1              =   360
      X2              =   2280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      X1              =   360
      X2              =   2280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   3  'Circle
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   3  'Circle
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PN As Integer
Dim px(1000), py(1000), r(1000) As Long
Dim draw, shipdraw As Boolean
Dim s1x, s1y, s2x, s2y As Double
Dim player As Integer
Dim ang, x, y, vx, vy, ax, ay, length As Double


Private Sub Form_Load()
Timer1.Interval = 0
Timer2.Interval = 0
Label5.Caption = MT
Form3.Label5.Caption = MT
Timer2.Interval = 0
End Sub

Private Sub Form_initialize()     '这里的事件不可用Form_Load，因为form2在程序执行初期就装入内存（此时变量PNM为空），而非show方法执行时！！！
Randomize
PN = Int(Rnd * PNM + 1)
draw = True
For i = 1 To PN
r(i) = Int(Rnd * 3000)
px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
    For k = 1 To i - 1
    If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
    i = i - 1
    draw = False
    Exit For
    End If
    Next k
If draw = True Then
Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)
Else
draw = True
End If
Next i
Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000
Label1.Caption = 0
Label2.Caption = 0
Label3.Caption = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
Label4.Caption = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End             'esc退出
If KeyCode = 32 Then                 'space开始运行
 If Timer1.Interval = 0 Then
   If player = 0 Then
   
   ang = Atn((Line1.Y2 - Line1.Y1) / (Line1.X2 - Line1.X1))
   If Line1.X2 - Line1.X1 < 0 Then
   ang = ang + 3.141592653589
   End If

   bullet.Left = Shape1.Left + 200 + 302 * Cos(ang)
   bullet.Top = Shape1.Top + 200 + 302 * Sin(ang)
   bullet.Visible = True
   vx = (Line1.X2 - Line1.X1) / 30
   vy = (Line1.Y2 - Line1.Y1) / 30
   ElseIf player = 1 Then
   
   ang = Atn((Line2.Y2 - Line2.Y1) / (Line2.X2 - Line2.X1))
   If Line2.X2 - Line2.X1 < 0 Then
   ang = ang + 3.141592653589
   End If

   bullet.Left = Shape2.Left + 200 + 302 * Cos(ang)
   bullet.Top = Shape2.Top + 200 + 302 * Sin(ang)
   bullet.Visible = True
   vx = (Line2.X2 - Line2.X1) / 30
   vy = (Line2.Y2 - Line2.Y1) / 30                                     '初速度调节处
   End If
Timer1.Interval = 1
Timer2.Interval = 100
 End If
End If          'space开始运行
If KeyCode = 37 Then                 '方向键左
   If Timer1.Interval = 0 And player = 0 Then
   ang = Atn((Line1.Y2 - Line1.Y1) / (Line1.X2 - Line1.X1))
   If Line1.X2 - Line1.X1 < 0 Then
   ang = ang + 3.141592653589
   End If
        If Shift = 0 Then
        ang = ang - 0.17
        ElseIf Shift = 1 Then
        ang = ang - 0.5
        ElseIf Shift = 2 Then
        ang = ang - 0.034
        ElseIf Shift = 4 Then
        ang = ang - 0.0068
        ElseIf Shift = 6 Then
        ang = ang - 0.00136
        ElseIf Shift = 5 Then
        ang = ang - 0.000272
        End If
   length = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
   Line1.X2 = Line1.X1 + length * Cos(ang)
   Line1.Y2 = Line1.Y1 + length * Sin(ang)
   Label3.Caption = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
   ElseIf Timer1.Interval = 0 And player = 1 Then
      ang = Atn((Line2.Y2 - Line2.Y1) / (Line2.X2 - Line2.X1))
   If Line2.X2 - Line2.X1 < 0 Then
   ang = ang + 3.141592653589
   End If
        If Shift = 0 Then
        ang = ang - 0.17
        ElseIf Shift = 1 Then
        ang = ang - 0.5
        ElseIf Shift = 2 Then
        ang = ang - 0.034
        ElseIf Shift = 4 Then
        ang = ang - 0.0068
        ElseIf Shift = 6 Then
        ang = ang - 0.00136
        ElseIf Shift = 5 Then
        ang = ang - 0.000272
        End If
   length = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
   Line2.X2 = Line2.X1 + length * Cos(ang)
   Line2.Y2 = Line2.Y1 + length * Sin(ang)
   Label4.Caption = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
   End If
End If
If KeyCode = 39 Then                 '方向键右
   If Timer1.Interval = 0 And player = 0 Then
   ang = Atn((Line1.Y2 - Line1.Y1) / (Line1.X2 - Line1.X1))
     If Line1.X2 - Line1.X1 < 0 Then
     ang = ang + 3.141592653589
     End If
        If Shift = 0 Then
        ang = ang + 0.17
        ElseIf Shift = 1 Then
        ang = ang + 0.5
        ElseIf Shift = 2 Then
        ang = ang + 0.034
        ElseIf Shift = 4 Then
        ang = ang + 0.0068
        ElseIf Shift = 6 Then
        ang = ang + 0.00136
        ElseIf Shift = 5 Then
        ang = ang + 0.000272
        End If
   length = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
   Line1.X2 = Line1.X1 + length * Cos(ang)
   Line1.Y2 = Line1.Y1 + length * Sin(ang)
   Label3.Caption = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
   ElseIf Timer1.Interval = 0 And player = 1 Then
   ang = Atn((Line2.Y2 - Line2.Y1) / (Line2.X2 - Line2.X1))
     If Line2.X2 - Line2.X1 < 0 Then
     ang = ang + 3.141592653589
     End If
        If Shift = 0 Then
        ang = ang + 0.17
        ElseIf Shift = 1 Then
        ang = ang + 0.5
        ElseIf Shift = 2 Then
        ang = ang + 0.034
        ElseIf Shift = 4 Then
        ang = ang + 0.0068
        ElseIf Shift = 6 Then
        ang = ang + 0.00136
        ElseIf Shift = 5 Then
        ang = ang + 0.000272
        End If
   length = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
   Line2.X2 = Line2.X1 + length * Cos(ang)
   Line2.Y2 = Line2.Y1 + length * Sin(ang)
   Label4.Caption = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
   End If
End If
If KeyCode = 38 Then                        '方向键上
   If Timer1.Interval = 0 And player = 0 Then
   ang = Atn((Line1.Y2 - Line1.Y1) / (Line1.X2 - Line1.X1))
        If Line1.X2 - Line1.X1 < 0 Then
        ang = ang + 3.141592653589
        End If
   length = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
        If Shift = 0 Then
        length = length + 100
        ElseIf Shift = 1 Then
        length = length + 500
        ElseIf Shift = 2 Then
        length = length + 20
        ElseIf Shift = 4 Then
        length = length + 4
        ElseIf Shift = 6 Then
        length = length + 1
        End If
   Line1.X2 = Line1.X1 + length * Cos(ang)
   Line1.Y2 = Line1.Y1 + length * Sin(ang)
   Label3.Caption = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
   ElseIf Timer1.Interval = 0 And player = 1 Then
   ang = Atn((Line2.Y2 - Line2.Y1) / (Line2.X2 - Line2.X1))
        If Line2.X2 - Line2.X1 < 0 Then
        ang = ang + 3.141592653589
        End If
   length = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
        If Shift = 0 Then
        length = length + 100
        ElseIf Shift = 1 Then
        length = length + 500
        ElseIf Shift = 2 Then
        length = length + 20
        ElseIf Shift = 4 Then
        length = length + 4
        ElseIf Shift = 6 Then
        length = length + 1
        End If
   Line2.X2 = Line2.X1 + length * Cos(ang)
   Line2.Y2 = Line2.Y1 + length * Sin(ang)
   Label4.Caption = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
   End If
End If
If KeyCode = 40 Then                        '方向键下
   If Timer1.Interval = 0 And player = 0 Then
   ang = Atn((Line1.Y2 - Line1.Y1) / (Line1.X2 - Line1.X1))
        If Line1.X2 - Line1.X1 < 0 Then
        ang = ang + 3.141592653589
        End If
   length = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
        If Shift = 0 Then
        length = length - 100
        ElseIf Shift = 1 Then
        length = length - 500
        ElseIf Shift = 2 Then
        length = length - 20
        ElseIf Shift = 4 Then
        length = length - 4
        ElseIf Shift = 6 Then
        length = length - 1
        End If
   Line1.X2 = Line1.X1 + length * Cos(ang)
   Line1.Y2 = Line1.Y1 + length * Sin(ang)
   Label3.Caption = Sqr((Line1.X2 - Line1.X1) ^ 2 + (Line1.Y2 - Line1.Y1) ^ 2)
   ElseIf Timer1.Interval = 0 And player = 1 Then
   ang = Atn((Line2.Y2 - Line2.Y1) / (Line2.X2 - Line2.X1))
        If Line2.X2 - Line2.X1 < 0 Then
        ang = ang + 3.141592653589
        End If
   length = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
        If Shift = 0 Then
        length = length - 100
        ElseIf Shift = 1 Then
        length = length - 500
        ElseIf Shift = 2 Then
        length = length - 20
        ElseIf Shift = 4 Then
        length = length - 4
        ElseIf Shift = 6 Then
        length = length + -1
        End If
   Line2.X2 = Line2.X1 + length * Cos(ang)
   Line2.Y2 = Line2.Y1 + length * Sin(ang)
   Label4.Caption = Sqr((Line2.X2 - Line2.X1) ^ 2 + (Line2.Y2 - Line2.Y1) ^ 2)
   End If
End If

'新的快捷键操作在此处添加












If KeyCode = 8 Then                                'backspace重置
    Cls
    Form3.Cls
    Timer1.Interval = 0
    Timer2.Interval = 0
    Label5.Caption = MT
    Form3.Label5.Caption = MT
    bullet.Visible = False
    vx = 0
    vy = 0
    ax = 0
    ay = 0
    Form3.Hide
        If Label1.Caption < Label2.Caption Then
        player = 0
        ElseIf Label1.Caption > Label2.Caption Then
        player = 1
        Else
        player = 0
        End If
    Randomize
    PN = Int(Rnd * PNM + 1)
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 49 Then
    Cls
        Form3.Cls
    Randomize
    PN = 1
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为1
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 50 Then
    Cls
    Form3.Cls
    Randomize
    PN = 2
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为2
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 51 Then
    Cls
        Form3.Cls

    Randomize
    PN = 3
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为3
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 52 Then
    Cls
        Form3.Cls

    Randomize
    PN = 4
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为4
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 53 Then
    Cls
        Form3.Cls

    Randomize
    PN = 5
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为5
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 54 Then
    Cls
        Form3.Cls

    Randomize
    PN = 6
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为6
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 55 Then
    Cls
        Form3.Cls

    Randomize
    PN = 7
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为7
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 56 Then
    Cls
        Form3.Cls

    Randomize
    PN = 8
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为8
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 57 Then
    Cls
        Form3.Cls

    Randomize
    PN = 9
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为9
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If
If KeyCode = 48 Then
    Cls
        Form3.Cls

    Randomize
    PN = 10
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i                                '手动设星球数为10
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000

End If

End Sub


Private Sub Timer1_Timer()
x = bullet.Left + 50
y = bullet.Top + 50
ax = 0
ay = 0
For i = 1 To PN
ax = ax - (r(i) ^ 3 * (x - px(i)) / ((x - px(i)) ^ 2 + (y - py(i)) ^ 2) ^ ((1 - index) / 2)) / 500 * G
ay = ay - (r(i) ^ 3 * (y - py(i)) / ((x - px(i)) ^ 2 + (y - py(i)) ^ 2) ^ ((1 - index) / 2)) / 500 * G
Next i
vx = vx + ax
vy = vy + ay
x = x + vx
y = y + vy
If player = 0 Then
Form2.PSet (x, y), RGB(255, 0, 0)
ElseIf player = 1 Then
Form2.PSet (x, y), RGB(0, 0, 255)
End If
Form3.PSet (x / n + 9600 - 9600 / n, y / n + 6000 - 6000 / n), RGB(255, 255, 255)
bullet.Left = x - 50
bullet.Top = y - 50                                                                                '飞行代码
For i = 1 To PN                                               '碰撞星球
    If (x - px(i)) ^ 2 + (y - py(i)) ^ 2 < (r(i) + 50) ^ 2 Then
        If player = 0 Then
        player = 1
        ElseIf player = 1 Then
        player = 0
        End If
    bullet.Visible = False
    vx = 0
    vy = 0
    ax = 0
    ay = 0
    Timer1.Interval = 0
    Timer2.Interval = 0
    Label5.Caption = MT
    Form3.Label5.Caption = MT
    Exit For
    End If
Next i
If (x - s1x) ^ 2 + (y - s1y) ^ 2 < 88209 Then                 '碰撞飞船1
    Cls
    Form3.Cls
    Timer1.Interval = 0
    Timer2.Interval = 0
    Label5.Caption = MT
    Form3.Label5.Caption = MT
    bullet.Visible = False
    vx = 0
    vy = 0
    ax = 0
    ay = 0
    Randomize
    PN = Int(Rnd * PNM + 1)
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)

    Else
    draw = True
    End If
    Next i
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000
Label2.Caption = Label2.Caption + 1
        If Label1.Caption < Label2.Caption Then
        player = 0
        ElseIf Label1.Caption > Label2.Caption Then
        player = 1
        Else
        player = 0
        End If
End If
If (x - s2x) ^ 2 + (y - s2y) ^ 2 < 88209 Then                 '碰撞飞船2
    Cls
    Form3.Cls
    Timer1.Interval = 0
    Timer2.Interval = 0
    Label5.Caption = MT
    Form3.Label5.Caption = MT
    bullet.Visible = False
    vx = 0
    vy = 0
    ax = 0
    ay = 0
    Randomize
    PN = Int(Rnd * PNM + 1)
    draw = True
    For i = 1 To PN
    r(i) = Int(Rnd * 3000)
    px(i) = Int(Rnd * (19200 - 2 * r(i)) + r(i))
    py(i) = Int(Rnd * (12000 - 2 * r(i)) + r(i))
        For k = 1 To i - 1
        If (px(i) - px(k)) ^ 2 + (py(i) - py(k)) ^ 2 <= (r(i) + r(k)) ^ 2 Then
        i = i - 1
        draw = False
        Exit For
        End If
        Next k
    If draw = True Then
    Circle (px(i), py(i)), r(i), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Form3.Circle (px(i) / n + 9600 - 9600 / n, py(i) / n + 6000 - 6000 / n), r(i) / n, RGB(255, 255, 255)
    Else
    draw = True
    End If
    Next i
    Do
Shape1.Top = Int(Rnd * (12000 - 2 * Shape1.Width) + Shape1.Width)
Shape1.Left = Int(Rnd * (19200 - 2 * Shape1.Width) + Shape1.Width)
s1x = Shape1.Left + Shape1.Width / 2
s1y = Shape1.Top + Shape1.Height / 2
  For i = 1 To PN
  If (s1x - px(i)) ^ 2 + (s1y - py(i)) ^ 2 <= (r(i) + Shape1.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
Loop While shipdraw = False
Do
Shape2.Top = Int(Rnd * (12000 - 2 * Shape2.Width) + Shape2.Width)
Shape2.Left = Int(Rnd * (19200 - 2 * Shape2.Width) + Shape2.Width)
s2x = Shape2.Left + Shape2.Width / 2
s2y = Shape2.Top + Shape2.Height / 2
  For i = 1 To PN
  If (s2x - px(i)) ^ 2 + (s2y - py(i)) ^ 2 <= (r(i) + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  Exit For
  Else
  shipdraw = True
  End If
  Next i
  If (s2x - s1x) ^ 2 + (s2y - s1y) ^ 2 <= (Shape1.Width / 2 + Shape2.Width / 2) ^ 2 Then
  shipdraw = False
  End If
Loop While shipdraw = False                    '定位飞船
Line1.X1 = s1x
Line1.X2 = s1x + 1000
Line1.Y1 = s1y
Line1.Y2 = s1y + 1000
Line2.X1 = s2x
Line2.X2 = s2x + 1000
Line2.Y1 = s2y
Line2.Y2 = s2y + 1000
Label1.Caption = Label1.Caption + 1
        If Label1.Caption < Label2.Caption Then
        player = 0
        ElseIf Label1.Caption > Label2.Caption Then
        player = 1
        Else
        player = 0
        End If
End If
If x > 19200 Or x < 0 Or y > 12000 Or y < 0 Then          '越过小边界
Form3.Show
Else
Form3.Hide
End If

If x / n + 9600 - 9600 / n > 19200 Or x / n + 9600 - 9600 / n < 0 Or y / n + 6000 - 6000 / n > 12000 Or y / n + 6000 - 6000 / n < 0 Then '越过大边界
        If player = 0 Then
        player = 1
        ElseIf player = 1 Then
        player = 0
        End If
    bullet.Visible = False
    vx = 0
    vy = 0
    ax = 0
    ay = 0
    Timer1.Interval = 0
    Timer2.Interval = 0
    Label5.Caption = MT
    Form3.Label5.Caption = MT
    Form3.Hide
End If
If Timer2.Interval = 0 Then Form3.Hide
End Sub

Private Sub Timer2_Timer()
Label5.Caption = Label5.Caption - 1
Form3.Label5.Caption = Form2.Label5.Caption
If Label5.Caption = 0 Then '时间到
Form3.Hide
        If player = 0 Then
        player = 1
        ElseIf player = 1 Then
        player = 0
        End If
    bullet.Visible = False
    vx = 0
    vy = 0
    ax = 0
    ay = 0
    Timer1.Interval = 0
    Timer2.Interval = 0
    Label5.Caption = MT
    Form3.Label5.Caption = MT
End If
If Timer2.Interval = 0 Then Label5.Caption = MT
End Sub
