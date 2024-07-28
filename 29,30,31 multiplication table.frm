VERSION 5.00
Begin VB.Form Frm1 
   AutoRedraw      =   -1  'True
   Caption         =   "乘法表 29, 30, 31"
   ClientHeight    =   6810
   ClientLeft      =   2835
   ClientTop       =   2325
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   8175
   Begin VB.CommandButton Cmd_ce 
      Caption         =   "Clear Everything"
      Height          =   495
      Left            =   3720
      TabIndex        =   42
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Cmd_neg 
      Caption         =   "-"
      Height          =   495
      Left            =   2640
      TabIndex        =   41
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   11
      Left            =   5160
      TabIndex        =   40
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   10
      Left            =   5160
      TabIndex        =   39
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   9
      Left            =   5160
      TabIndex        =   38
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   8
      Left            =   5160
      TabIndex        =   37
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   7
      Left            =   5160
      TabIndex        =   36
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   6
      Left            =   5160
      TabIndex        =   35
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   5
      Left            =   5160
      TabIndex        =   34
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   33
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   32
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   2
      Left            =   5160
      TabIndex        =   31
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   30
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   9
      Left            =   2640
      TabIndex        =   29
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   8
      Left            =   1560
      TabIndex        =   28
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   7
      Left            =   480
      TabIndex        =   27
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   6
      Left            =   2640
      TabIndex        =   26
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   5
      Left            =   1560
      TabIndex        =   25
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   24
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   3
      Left            =   2640
      TabIndex        =   23
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   22
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   21
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Cmd_multnum 
      Caption         =   "x 1"
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Cmd_Clr 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Cmd_multiply 
      Caption         =   "Multiply"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Txt_input 
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton Cmd_backspace 
      Caption         =   "←"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Cmd_dec 
      Caption         =   "."
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Cmd_numpad 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   5880
      TabIndex        =   20
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   5880
      TabIndex        =   19
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5880
      TabIndex        =   18
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5880
      TabIndex        =   17
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5880
      TabIndex        =   16
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5880
      TabIndex        =   15
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5880
      TabIndex        =   14
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Lbl_output 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Liberation Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   9
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Lbl_number 
      Alignment       =   2  'Center
      Caption         =   "莫丰泽(29),沈俊达(30),戴于珅(31)"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   4935
   End
   Begin VB.Label Lbl_title 
      Alignment       =   2  'Center
      Caption         =   "乘法表"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "Frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ls(9) As Integer 'array for the values of the buttons
Dim i As Integer 'for loop var
Function textoutinit()
    For i = 0 To 11
        Lbl_output(i).FontSize = 19 'change lbl fontsize
        Lbl_output(i).Caption = "=" 'clear output
        Lbl_output(i).AutoSize = True 'set autosize
        Cmd_multnum(i).Caption = ("× " & i + 1) 'set multnum button caption
    Next
End Function
Private Sub Frm_Load()
    textoutinit
    For i = 0 To 9  'populate array and change numpad caption
        ls(i) = i
        Cmd_numpad(i).Caption = (i)
    Next
End Sub
Private Sub Cmd_backspace_Click()
    Txt_input.SetFocus
    Txt_input.SelStart = Len(Txt_input.Text)
    SendKeys ("{BACKSPACE}")
End Sub
Private Sub Cmd_Clr_Click()
    Txt_input.Text = "" 'clears textbox
End Sub
Private Sub Cmd_ce_Click()
    Txt_input.Text = ""
    textoutinit
End Sub
Private Sub Cmd_neg_Click()
    If InStr(1, Txt_input.Text, "-") = False Then
        Txt_input.Text = "-" + Txt_input.Text
    End If
End Sub
Private Sub Cmd_multiply_Click()
    For i = 0 To 11
        Lbl_output(i).Caption = "= " & Val(Txt_input.Text) * (i + 1)
    Next
End Sub
Private Sub Cmd_numpad_Click(Index As Integer)
    Txt_input.Text = Txt_input.Text + CStr(ls(Index))
End Sub
Private Sub Cmd_multnum_Click(Index As Integer)
    Lbl_output(Index).Caption = "= " & Val(Txt_input.Text) * (Index + 1)
End Sub
Private Sub Cmd_dec_Click()
    If InStr(1, Txt_input.Text, ".") = False Then
        Txt_input.Text = Txt_input.Text + "."
    End If
End Sub

Private Sub Txt_input_KeyPress(KeyAscii As Integer) 'rejects non-numeral inputs
    Select Case KeyAscii
            Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyDelete, vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab, vbKeyBack, 45
                If KeyAscii = 46 Then If InStr(1, Txt_input.Text, ".") Then KeyAscii = 0
                If KeyAscii = 45 Then If InStr(1, Txt_input.Text, "-") Then KeyAscii = 0
            Case Else
                KeyAscii = 0
                Beep
        End Select
End Sub
