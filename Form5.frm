VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Please enter new value"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form5"
   ScaleHeight     =   1395
   ScaleWidth      =   5445
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "text size"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h As Integer
Private Sub Command1_Click()
Form4.Label1(h) = Text1.text
Form4.Label1(h).FontSize = CInt(Text2.text)
Me.Hide
End Sub

Public Function ad(a As Integer)
h = a
End Function
