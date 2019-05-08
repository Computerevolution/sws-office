VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "SWS Office"
   ClientHeight    =   5295
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7185
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   400
      Left            =   960
      Top             =   2760
   End
   Begin VB.PictureBox Picture3 
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   0
      Width           =   255
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0C0C0&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3795
         ScaleWidth      =   195
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   2400
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   2520
      ScaleHeight     =   795
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   2880
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   7
         Top             =   0
         Width           =   135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   0
         TabIndex        =   4
         Text            =   "100"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "show size"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu 保存 
         Caption         =   "Save"
      End
      Begin VB.Menu open 
         Caption         =   "Open"
      End
   End
   Begin VB.Menu wholee 
      Caption         =   "Show"
      Begin VB.Menu sta 
         Caption         =   "Start show mode"
      End
      Begin VB.Menu showw 
         Caption         =   "show size"
         Begin VB.Menu rte 
            Caption         =   "bigger 50%"
         End
         Begin VB.Menu hnfh 
            Caption         =   "smaller 50%"
         End
         Begin VB.Menu enter 
            Caption         =   "user decide"
         End
      End
   End
   Begin VB.Menu in 
      Caption         =   "insert"
      Begin VB.Menu text 
         Caption         =   "text area"
      End
   End
   Begin VB.Menu mo 
      Caption         =   "advanced"
      Begin VB.Menu vc 
         Caption         =   "view code"
      End
   End
   Begin VB.Menu help 
      Caption         =   "help"
      Begin VB.Menu about 
         Caption         =   "about"
      End
      Begin VB.Menu hh 
         Caption         =   "help"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod(1000) As String
Dim ia, xx, yy, bia As Integer
Dim pa As String
Dim max, ooo As Long
Private Sub em_Click()

Dim ii As New Form1
ii.Show
ii.re pa
Me.Hide
End Sub

Private Sub about_Click()
MsgBox "SWS Office 0.12.2,From CR2019，GNU Software，Shaoqirui dev."
End Sub

Private Sub Command1_Click()
Dim i As Integer
i = 0
While i < ia
Label1(i).FontSize = Label1(i).FontSize * CInt(Text2.text) / 100
i = i + 1
Wend
Picture1.Visible = False
Text2.text = "100"

End Sub

Private Sub enter_Click()
Picture1.Visible = True
End Sub

Private Sub Form_Load()
ia = 1
End Sub

Private Sub hh_Click()
Form6.Show
End Sub

Private Sub hnfh_Click()
Dim i As Integer
i = 0
While i < ia
Label1(i).FontSize = Label1(i).FontSize / 2
i = i + 1
Wend
End Sub

Private Sub Label1_DblClick(Index As Integer)
Form5.Text1.text = Label1(Index)
Form5.Text2.text = Label1(Index).FontSize
Form5.Show
Form5.ad Index
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
xx = X
yy = Y
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Label1(Index).Left = Label1(Index).Left + X - xx
Label1(Index).Top = Label1(Index).Top + Y - yy
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx = X
yy = Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Picture1.Left = Picture1.Left + X - xx
Picture1.Top = Picture1.Top + Y - yy
End If
End Sub

Private Sub open_Click()
Form2.Show
End Sub

Public Function re(h As String)
Dim i As Integer
i = 1
While i < ia
Label1(i).Visible = False
Label1(i) = ""
i = i + 1
Wend
If ia > bia Then
bia = ia
End If
ia = 0
pa = h
nfilenum = FreeFile
Open h For Input As nfilenum
linel = 1
Do While Not EOF(nfilenum)
Line Input #nfilenum, snel
stext = stext & snel & vbCrLf
cod(linel) = snel
linel = linel + 1
Loop
Close nfilenum
Text1.text = stext
i = 0
While i < linel
If i = 1 Then
If cod(i) <> "s10" Then
MsgBox "此文件在更高版本的SWS Office或其他不受兼容的软件上被创建，文件打开后非常有可能出现不兼容的情况。"
End If
End If
If cod(i) = "l" Then
If ia >= bia Then
Load Label1(ia)
End If
i = i + 1
Label1(ia) = ""
While cod(i) <> "end!#"
Label1(ia) = Label1(ia) + vbCrLf + cod(i)
i = i + 1
Wend
i = i + 1
Label1(ia).Left = CInt(cod(i))
i = i + 1
Label1(ia).Top = CInt(cod(i))
i = i + 1
Label1(ia).FontSize = CInt(cod(i))
Label1(ia).Visible = True
ia = ia + 1
End If
i = i + 1
Wend
End Function

Private Sub 查看页面源代码_Click()

End Sub

Private Sub Picture2_Click()
Picture1.Visible = False
Text2.text = "100"
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
yy = Y

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Picture4.Top = Picture4.Top + Y - yy
Timer2.Enabled = True
End If
End Sub

Private Sub rte_Click()
Dim i As Integer
i = 0
While i < ia
Label1(i).FontSize = Label1(i).FontSize * 1.5
i = i + 1
Wend
End Sub

Private Sub sta_Click()
Form1.re pa
Form1.Show
End Sub

Private Sub text_Click()
If ia >= bia Then
Load Label1(ia)
End If
i = i + 1
Label1(ia) = "新文本框"
i = i + 1
Label1(ia).Left = 100
i = i + 1
Label1(ia).Top = 100
i = i + 1
Label1(ia).FontSize = 15
Label1(ia).Visible = True
ia = ia + 1
End Sub

Private Sub Timer1_Timer()
Dim i As Long
i = 0
max = 0
While i < ia
If Label1(i).Width + Label1(i).Top > max Then
max = Label1(i).Height + Label1(i).Top
End If
i = i + 1
Wend
If max > Form1.Height Then
Picture4.Height = Form4.Height / (max / Form4.Height)
Picture3.Height = Form4.Height * 2
End If
End Sub

Private Sub Timer2_Timer()
Dim i As Integer
i = 0
While i < ia
If Label1(i).Width + Label1(i).Top > max Then
Label1(i).Top = Label1(i).Top - ((ooo - Picture4.Top) / Form1.Height) * max
End If
i = i + 1
Wend
Timer2.Enabled = False
ooo = Picture4.Top
End Sub

Private Sub vc_Click()
Dim a As String
a = "s10" + vbCrLf
Dim i As Integer
i = 0
While i < ia
a = a + "l" + vbCrLf + Label1(i).Caption + vbCrLf + CStr(Label1(i).Left) + vbCrLf + CStr(Label1(i).Top) + vbCrLf + CStr(Label1(i).FontSize) + vbCrLf
i = i + 1
Wend
Form3.Text1.text = a
Form3.Show
End Sub

Private Sub 保存_Click()
Dim a As String
a = "s10" + vbCrLf
Dim i As Integer
i = 0
While i < ia
a = a + "l" + vbCrLf + Label1(i).Caption + vbCrLf + "end!#" + vbCrLf + CStr(Label1(i).Left) + vbCrLf + CStr(Label1(i).Top) + vbCrLf + CStr(Label1(i).FontSize) + vbCrLf
i = i + 1
Wend
Open pa For Output As #1
 Contents = a
 Print #1, Contents
 Close 1
End Sub
