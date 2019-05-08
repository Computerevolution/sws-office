VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Open"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3525
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   1920
      Pattern         =   "*.st"
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2610
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form4.re File1.Path + "\" + File1.FileName
Me.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

