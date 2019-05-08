VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "SWS Office Help"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form6"
   ScaleHeight     =   5730
   ScaleWidth      =   8550
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   8415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search Help"
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "enter the name of the thing that you have problem"
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim j As String
j = Text1.text
If InStr(1, j, "文本框") > 0 Then
Text2.text = "关于文本框" + vbCrLf + "sws office支持创建文本框，您可以通过 插入--文本框 创建一个文本框。" + vbCrLf + "sws office不赞成删除文本框，如果您真的想删除某个文本框，您可以清空文本框内文字，我们在保存时会对其进行回收处理"
End If
If InStr(1, j, "解码") > 0 Then
Text2.text = "关于解码" + vbCrLf + "sws office使用sw(sws window)解码技术，是世界上最快最小的office文件解码器。"
End If
If InStr(1, j, "保存") > 0 Then
Text2.text = "关于保存" + vbCrLf + "sws office使用sw(sws window)编码技术保存文档，是世界上最快最小的office文件编码器。" + vbCrLf + "如何保存文件:菜单栏--文件--保存"
End If
If InStr(1, j, "打开") > 0 Then
Text2.text = "关于打开" + vbCrLf + "sws office使用sw(sws window)解码技术，是世界上最快最小的office文件解码器。" + vbCrLf + "如何打开文件:菜单栏--文件--打开"
End If
If InStr(1, j, "ppt") > 0 Or InStr(1, j, "word") Or InStr(1, j, "excel") > 0 > 0 Then
Text2.text = "关于其它Office创建的文件" + vbCrLf + "sws office追求小巧，因此不支持速度慢，占用空间大的ppt文件，如果您想打开ppt文件推荐使用开源软件Libre Office"
End If
End Sub

