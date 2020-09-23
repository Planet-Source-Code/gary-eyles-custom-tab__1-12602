VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Custom tabs"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   StartUpPosition =   2  'CenterScreen
   Begin CustomTabs.vbalImageList gList 
      Left            =   2040
      Top             =   6000
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   5640
      Images          =   "Form1.frx":014A
      KeyCount        =   6
      Keys            =   "ÿÿÿÿÿ"
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete All"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Own text"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete tab"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add tab"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   1335
   End
   Begin CustomTabs.gTab gTab1 
      Height          =   5055
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8916
      Begin VB.Frame Frame1 
         Caption         =   "Information"
         Height          =   4575
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton Command10 
            Caption         =   "Test"
            Height          =   375
            Left            =   2880
            TabIndex        =   21
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Without"
            Height          =   375
            Left            =   2640
            TabIndex        =   8
            Top             =   2760
            Width           =   1455
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Icons"
            Height          =   375
            Left            =   2640
            TabIndex        =   7
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Hot tracking"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Bottom"
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   10
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Right"
            Height          =   375
            Index           =   3
            Left            =   2160
            TabIndex        =   11
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Style"
            Height          =   375
            Left            =   1200
            TabIndex        =   12
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Top"
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   16
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Left"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Rotatet text (left / right views only)"
            Height          =   495
            Left            =   360
            TabIndex        =   14
            Top             =   2280
            Width           =   2055
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Highlight selected tab"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Enabled (True)"
            Height          =   375
            Left            =   2640
            TabIndex        =   9
            Top             =   2040
            Width           =   1455
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Icon placement"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1080
            TabIndex        =   22
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Tab text ="
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1200
            TabIndex        =   18
            Top             =   600
            Width           =   645
         End
         Begin VB.Label Label4 
            Caption         =   "Tab index ="
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CustomTabs = "&Network|Ne&w|Add/&Remove|&Boot|Re&pair|Parano&ia|Mou&se|&General|E&xplorer|Desk&top|My &Computer"
Private Const CustomTabsTips = "Change Network settings|New settings|Add/Remove programs|Change different Boot settings|Repair various settings|Some useful options|Change Mouse settings|Various settings|Change explorer settings|Different desktop settings|My Computer settings"

Private Sub Check1_Click()
gTab1.tRotateText CBool(Check1.Value)
gTab1.RefreshTabs
End Sub

Private Sub Check2_Click()
gTab1.ButtonHighlight CBool(Check2.Value)
End Sub

Private Sub Check3_Click()
If CBool(Check3.Value) Then
    gTab1.IconPlacement True
Else
    gTab1.IconPlacement False
End If

gTab1.RefreshTabs
End Sub

Private Sub Check4_Click()
gTab1.HotTrack CBool(Check4.Value)
Frame1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command1_Click()
gTab1.DeleteTab gTab1.tTabIndex
DoEvents
gTab1.RefreshTabs
Frame1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command10_Click()
'gTab1.SetTooltipBkColor QBColor(14)
'gTab1.SetTooltipTextColor QBColor(4)

Form2.Show
End Sub

Private Sub Command2_Click()
Debug.Print "ADDING A NEW TAB"

gTab1.InsertTab "Tab " & Rnd
DoEvents
gTab1.RefreshTabs
Frame1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command3_Click()
Dim cTmp As Long
cTmp = gTab1.tTabIndex

gTab1.InsertTab InputBox("Enter text", "Enter text"), _
    InputBox("Enter tooltip text", "Enter text"), , cTmp + 1

gTab1.RefreshTabs

Frame1.Visible = False
Frame1.Visible = True

gTab1.tTabIndex cTmp
End Sub

Private Sub Command4_Click(Index As Integer)
If Index = 0 Then
    gTab1.ChangeStyle ttop, gTab1.GetStyleButton
ElseIf Index = 1 Then
    gTab1.ChangeStyle tbottom, gTab1.GetStyleButton
ElseIf Index = 2 Then
    gTab1.ChangeStyle tleft, gTab1.GetStyleButton
ElseIf Index = 3 Then
    gTab1.ChangeStyle tRight, gTab1.GetStyleButton
End If

DoEvents
gTab1.RefreshTabs
Frame1.Visible = False
Frame1.Visible = True

Form_Resize
End Sub

Private Sub Command5_Click()
If gTab1.GetStyleButton Then
    gTab1.ChangeStyle gTab1.GetStyle, False
Else
    gTab1.ChangeStyle gTab1.GetStyle, True
End If

Frame1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command6_Click()
gTab1.DeleteAllTabs
End Sub

Private Sub Command7_Click()
If gTab1.Enabled Then
    gTab1.Enabled False
    Command7.Caption = "Enabled (False)"
Else
    gTab1.Enabled True
    Command7.Caption = "Enabled (True)"
End If
End Sub

Private Sub Command8_Click()
Dim TmpFrm As Form

Set TmpFrm = New Form1
TmpFrm.Caption = "Without Images"
TmpFrm.Show
End Sub

Private Sub Command9_Click()
If gTab1.GetImagelist <> 0 Then
    gTab1.SetImageList
Else
    gTab1.SetImageList gList
End If
End Sub

Private Sub Form_Load()
'Setup up tabs like Tweak UI

Dim TmpFrm As Form
Dim ThereIsAnotherForm As Boolean

For Each TmpFrm In Forms
    If TmpFrm.hwnd <> Me.hwnd Then
        ThereIsAnotherForm = True
    End If
Next

If Not ThereIsAnotherForm Then
    gTab1.SetImageList gList
End If

Dim TmpStrings() As String
TmpStrings() = Split(CustomTabs, "|")
Dim TmpStringsTips() As String
TmpStringsTips() = Split(CustomTabsTips, "|")

Dim c As Integer
For c = UBound(TmpStrings) To LBound(TmpStrings) Step -1
        gTab1.InsertTab TmpStrings(c), TmpStringsTips(c), Int(Rnd * gList.ImageCount)
Next
End Sub

Private Sub Form_Resize()
On Error Resume Next

If gTab1.GetStyle = tbottom Then
    Command1.tOp = 5
    Command2.tOp = 5
    Command3.tOp = 5
    Command6.tOp = 5
    gTab1.Move 0, 5 * 2 + Command1.Height, ScaleWidth, ScaleHeight - Command1.Height - 10
Else
    Command1.tOp = ScaleHeight - Command1.Height - 5
    Command2.tOp = ScaleHeight - Command2.Height - 5
    Command3.tOp = ScaleHeight - Command3.Height - 5
    Command6.tOp = ScaleHeight - Command6.Height - 5
    gTab1.Move 0, 0, ScaleWidth, ScaleHeight - Command1.Height - 10
End If
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    gTab1.DoMenu
End If
End Sub

Private Sub gTab1_Resize()
On Error Resume Next
Frame1.left = gTab1.pLeft * 15 + 10 * 15
Frame1.tOp = gTab1.pTop * 15 + (10 * 15)
Frame1.Width = gTab1.pRight * 15 - (20 * 15) - gTab1.pLeft * 15
Frame1.Height = (gTab1.pBottom * 15) - Frame1.tOp - (10 * 15)
End Sub

Private Sub gTab1_gTabChange(tTabIndex As Long, tTabString As String)
Label3.Caption = tTabIndex

Dim Tmp As Long
Dim TmpS As String

Tmp = InStr(1, tTabString, "&", vbTextCompare)
If Tmp > 0 Then
    TmpS = Mid(tTabString, 1, Tmp - 1)
    TmpS = TmpS & Mid(tTabString, Tmp + 1, Len(tTabString) - Tmp)
End If

Label2.Caption = TmpS
End Sub

