VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "SWS Office编辑"
   ClientHeight    =   5295
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod(1000) As String
Dim ia, xx, yy, bia As Integer
Dim pa As String
Private Sub em_Click()

Dim ii As New Form1
ii.Show
ii.re pa
Me.Hide
End Sub

Private Sub about_Click()
MsgBox "SWS Office,计算机革命2019产物，开源软件，邵启瑞主导开发"
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
ia = 1
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

Private Sub hh_Click()
Form6.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Command1.Left = X
Command1.Top = Y
Command1.Visible = True
Else
Command1.Visible = False
End If
End Sub


Private Sub 查看页面源代码_Click()

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
a = a + "l" + vbCrLf + Label1(i).Caption + vbCrLf + CStr(Label1(i).Left) + vbCrLf + CStr(Label1(i).Top) + vbCrLf + CStr(Label1(i).FontSize) + vbCrLf
i = i + 1
Wend
Open pa For Output As #1
 Contents = a
 Print #1, Contents
 Close 1
End Sub

