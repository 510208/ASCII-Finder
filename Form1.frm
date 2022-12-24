VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ASCII碼查詢工具"
   ClientHeight    =   615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   4215
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "查詢按鍵ASCII(&A)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ASCIICode.Show
End Sub
