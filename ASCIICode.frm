VERSION 5.00
Begin VB.Form ASCIICode 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "ASCII查詢器"
   ClientHeight    =   1500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   1500
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFFF&
      Caption         =   "Keycode："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label dyj 
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Appearance      =   0  '平面
      BackColor       =   &H00C0E0FF&
      Caption         =   "對應鍵："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFC0&
      Caption         =   "ASCII："
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "ASCIICode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Label4.Caption = KeyCode
    Label9.Caption = Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    dyj.Caption = ""
    Label2.Caption = ""
    Label4.Caption = ""
    dyj.ForeColor = QBColor(0)
    Label2.Caption = KeyAscii
    If KeyAscii = Null Then
        dyj.ForeColor = QBColor(4)
        dyj.Caption = "找不到此ASCII"
    Else
        Select Case KeyAscii
            Case 9
                dyj.Caption = "Tab"
            Case 13
                dyj.Caption = "Enter"
            Case 32
                dyj.Caption = "Space(空白鍵)"
            Case Else
                dyj.Caption = Chr(KeyAscii)
        End Select
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Label4.Caption = KeyCode
End Sub

Private Sub Label9_Click()
End Sub
