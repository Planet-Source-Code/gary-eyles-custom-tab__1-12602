VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Perpetua"
      Size            =   27.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin CustomTabs.gTab gTab1 
      Height          =   1695
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Doneit As Boolean

Private Sub Command1_Click()
Dim fnt As New CLogFont
Dim rct As RECT
Dim rct2 As RECT

Cls

Set fnt.LOGFONT = Me.Font
            
'If Doneit Then
'    fnt.Rotation = 0
'Else
    fnt.Rotation = 90
'    Doneit = True
'End If


Dim ooFont As Long
ooFont = SelectObject(Me.hdc, fnt.Handle)

DrawText Me.hdc, "GARY", 4, rct2, DT_SINGLELINE Or DT_VCENTER Or DT_CALCRECT

rct.left = ScaleWidth / 2
rct.tOp = (ScaleHeight / 2) - rct2.Right
rct.Bottom = rct.tOp + rct2.Right * 2 + rct2.Bottom * 2
rct.Right = rct2.Bottom * 2 + rct.left
Line (rct.left, rct.tOp)-(rct.Right, rct.Bottom), QBColor(14), B
DrawText Me.hdc, "&GARY", 5, rct, DT_SINGLELINE Or DT_VCENTER

Line (ScaleWidth / 2, 0)-(ScaleWidth / 2, ScaleHeight), 0
Line (0, ScaleHeight / 2)-(ScaleWidth, ScaleHeight / 2), 0

SelectObject Me.hdc, ooFont
fnt.CleanUp
Set fnt = Nothing
End Sub

Private Sub Form_Load()
gTab1.InsertTab "&TEST ", "TESTING"
gTab1.InsertTab "T&EST2 ", "TESTING2"
End Sub

Private Sub Form_Resize()
'Command1_Click
gTab1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
