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
   StartUpPosition =   3  '����ȱʡ
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
If InStr(1, j, "�ı���") > 0 Then
Text2.text = "�����ı���" + vbCrLf + "sws office֧�ִ����ı���������ͨ�� ����--�ı��� ����һ���ı���" + vbCrLf + "sws office���޳�ɾ���ı�������������ɾ��ĳ���ı�������������ı��������֣������ڱ���ʱ�������л��մ���"
End If
If InStr(1, j, "����") > 0 Then
Text2.text = "���ڽ���" + vbCrLf + "sws officeʹ��sw(sws window)���뼼�����������������С��office�ļ���������"
End If
If InStr(1, j, "����") > 0 Then
Text2.text = "���ڱ���" + vbCrLf + "sws officeʹ��sw(sws window)���뼼�������ĵ����������������С��office�ļ���������" + vbCrLf + "��α����ļ�:�˵���--�ļ�--����"
End If
If InStr(1, j, "��") > 0 Then
Text2.text = "���ڴ�" + vbCrLf + "sws officeʹ��sw(sws window)���뼼�����������������С��office�ļ���������" + vbCrLf + "��δ��ļ�:�˵���--�ļ�--��"
End If
If InStr(1, j, "ppt") > 0 Or InStr(1, j, "word") Or InStr(1, j, "excel") > 0 > 0 Then
Text2.text = "��������Office�������ļ�" + vbCrLf + "sws office׷��С�ɣ���˲�֧���ٶ�����ռ�ÿռ���ppt�ļ�����������ppt�ļ��Ƽ�ʹ�ÿ�Դ���Libre Office"
End If
End Sub

