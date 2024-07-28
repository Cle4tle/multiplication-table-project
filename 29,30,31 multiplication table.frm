VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   2835
   ClientTop       =   2325
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   14490
   Begin VB.CommandButton Cmd_x12 
      Caption         =   "x 12"
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x11 
      Caption         =   "x 11"
      Height          =   375
      Left            =   6120
      TabIndex        =   26
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x10 
      Caption         =   "x 10"
      Height          =   375
      Left            =   6120
      TabIndex        =   25
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x9 
      Caption         =   "x 9"
      Height          =   375
      Left            =   6120
      TabIndex        =   24
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x8 
      Caption         =   "x 8"
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x7 
      Caption         =   "x 7"
      Height          =   375
      Left            =   6120
      TabIndex        =   22
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x6 
      Caption         =   "x 6"
      Height          =   375
      Left            =   6120
      TabIndex        =   21
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x5 
      Caption         =   "x 5"
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x4 
      Caption         =   "x 4"
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x3 
      Caption         =   "x 3"
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x2 
      Caption         =   "x 2"
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Cmd_x1 
      Caption         =   "x 1"
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Cmd_Clr 
      Caption         =   "Clear"
      Height          =   975
      Left            =   3840
      TabIndex        =   14
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_multiply 
      Caption         =   "Multiply"
      Height          =   975
      Left            =   3840
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Txt_input 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      TabIndex        =   12
      Top             =   1320
      Width           =   4935
   End
   Begin VB.CommandButton Cmd_9 
      Caption         =   "9"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Cmd_8 
      Caption         =   "8"
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Cmd_7 
      Caption         =   "7"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Cmd_6 
      Caption         =   "6"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Cmd_5 
      Caption         =   "5"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Cmd_4 
      Caption         =   "4"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Cmd_backspace 
      Caption         =   "<"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Cmd_dec 
      Caption         =   "."
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Cmd_0 
      Caption         =   "0"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Cmd_3 
      Caption         =   "3"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Cmd_2 
      Caption         =   "2"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Cmd_1 
      Caption         =   "1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Lbl_output 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   6840
      TabIndex        =   29
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Lbl_number 
      Alignment       =   2  'Center
      Caption         =   "29, 30, 31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   28
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Lbl_title 
      Alignment       =   2  'Center
      Caption         =   "³Ë·¨±í"
      BeginProperty Font 
         Name            =   "¿¬Ìå"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   15
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ls(9) As Integer 'array for the values of the buttons
Dim i As Integer 'for loop variable
Function textoutinit() 'clears output textbox and adds =
    Lbl_output.Caption = "=" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " =" & vbNewLine & " ="
End Function
Private Sub Form_Load()
    textoutinit
    For i = 0 To 9 Step 1 'for loop to populate array
        ls(i) = i
        Print ls(i)
        Next
End Sub
Private Sub Cmd_backspace_Click()
    Txt_input.SetFocus
    Txt_input.SelStart = Len(Txt_input.Text)
    SendKeys ("{BACKSPACE}")
End Sub
Private Sub Cmd_Clr_Click()
    Txt_input.Text = "" 'clears textbox
    textoutinit
End Sub
Private Sub Cmd_0_Click()
    Txt_input.Text = Txt_input.Text + CStr(ls(0))
End Sub
Private Sub Cmd_1_Click()
    Txt_input.Text = Txt_input.Text + CStr(ls(1))
End Sub
Private Sub Cmd_2_click()
    Txt_input.Text = Txt_input.Text + CStr(ls(2))
End Sub
Private Sub Cmd_3_click()
    Txt_input.Text = Txt_input.Text + CStr(ls(3))
End Sub
Private Sub Cmd_4_click()
    Txt_input.Text = Txt_input.Text + CStr(ls(4))
End Sub
Private Sub Cmd_5_click()
    Txt_input.Text = Txt_input.Text + CStr(ls(5))
End Sub
Private Sub Cmd_6_click()
    Txt_input.Text = Txt_input.Text + CStr(ls(6))
End Sub
Private Sub Cmd_7_click()
    Txt_input.Text = Txt_input.Text + CStr(ls(7))
End Sub
Private Sub Cmd_8_click()
    Txt_input.Text = Txt_input.Text + CStr(ls(8))
End Sub
Private Sub Cmd_9_click()
    Txt_input.Text = Txt_input.Text + CStr(ls(9))
End Sub
Private Sub Cmd_dec_Click()
    Txt_input.Text = Txt_input.Text + "."
End Sub
Private Sub Txt_input_KeyPress(KeyAscii As Integer) 'rejects non-numeral inputs
    Select Case KeyAscii
            Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyDelete, _
            vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab, vbKeyBack
                If KeyAscii = 46 Then If InStr(1, Txt_input.Text, ".") Then KeyAscii = 0
            Case Else
                KeyAscii = 0
                Beep
        End Select
End Sub

