VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "uscream@vip.hr"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin Project1.uSc_MultiOption uSc_MultiOption1 
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      CaptionAligning =   1
      Aligning        =   1
      Positioning     =   1
      Caption         =   "uScream waz here"
      OptionCount     =   4
      OptionSelected  =   3
      Horizontal      =   -1  'True
      BCOL            =   255
      FCOL            =   16777215
   End
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   3135
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   2775
         Begin VB.CommandButton Command8 
            Caption         =   "Change Caption Aligning"
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   2535
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Change Positioning"
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   2535
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Change Aligning"
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2775
         Begin VB.CommandButton Command1 
            Caption         =   "Selected?"
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "-"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "+"
            Height          =   375
            Left            =   720
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   2775
         Begin VB.CommandButton Command5 
            Caption         =   "Set caption"
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Set tooltip of selected option"
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   240
            Width           =   2535
         End
      End
      Begin Project1.uSc_MultiOption uSc_MultiOption3 
         Height          =   1575
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2778
         Appearance3D    =   1
         Border          =   1
         CaptionAligning =   2
         Aligning        =   1
         Positioning     =   1
         Caption         =   "Use buttons above to ""Do"" this control"
         OptionCount     =   4
         OptionSelected  =   2
         Horizontal      =   0   'False
         BCOL            =   13160664
         FCOL            =   0
      End
   End
   Begin VB.PictureBox picRulOut 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.VScrollBar scrollRul 
         Height          =   2295
         LargeChange     =   5000
         Left            =   2760
         SmallChange     =   500
         TabIndex        =   1
         Top             =   1680
         Width           =   255
      End
      Begin VB.PictureBox picRulIn 
         BackColor       =   &H80000005&
         Height          =   3135
         Left            =   0
         ScaleHeight     =   3075
         ScaleWidth      =   2955
         TabIndex        =   2
         Top             =   480
         Width           =   3015
         Begin Project1.uSc_MultiOption chkRulesExist_P 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            CaptionAligning =   2
            Aligning        =   1
            Positioning     =   1
            Caption         =   "uScream waz here"
            OptionCount     =   2
            OptionSelected  =   2
            Horizontal      =   -1  'True
            BCOL            =   -2147483643
            FCOL            =   -2147483640
         End
         Begin Project1.uSc_MultiOption chkRulesExist_M 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            CaptionAligning =   2
            Aligning        =   1
            Positioning     =   1
            Caption         =   "uScream waz here"
            OptionCount     =   3
            OptionSelected  =   1
            Horizontal      =   -1  'True
            BCOL            =   -2147483643
            FCOL            =   -2147483640
         End
         Begin VB.Line lineRules 
            X1              =   120
            X2              =   2400
            Y1              =   480
            Y2              =   480
         End
      End
   End
   Begin Project1.uSc_MultiOption uSc_MultiOption2 
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   4080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Border          =   1
      CaptionAligning =   1
      Aligning        =   2
      Positioning     =   3
      Caption         =   "uScream waz here"
      OptionCount     =   2
      OptionSelected  =   2
      Horizontal      =   -1  'True
      BCOL            =   16777215
      FCOL            =   0
   End
   Begin Project1.uSc_MultiOption uSc_MultiOption4 
      Height          =   735
      Left            =   240
      TabIndex        =   20
      Top             =   5040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      CaptionAligning =   1
      Aligning        =   1
      Positioning     =   4
      Caption         =   "uScream waz here"
      OptionCount     =   8
      OptionSelected  =   7
      Horizontal      =   -1  'True
      BCOL            =   16711680
      FCOL            =   16777215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox uSc_MultiOption3.OptionSelected
End Sub

Private Sub Command2_Click()
uSc_MultiOption3.OptionCount = uSc_MultiOption3.OptionCount - 1
End Sub

Private Sub Command3_Click()
uSc_MultiOption3.OptionCount = uSc_MultiOption3.OptionCount + 1
End Sub

Private Sub Command4_Click()
uSc_MultiOption3.SetOptionToolTip uSc_MultiOption3.OptionSelected, Text1.Text
End Sub

Private Sub Command5_Click()
uSc_MultiOption3.Caption = Text1.Text
End Sub

Private Sub Command6_Click()
Select Case uSc_MultiOption3.Aligning
Case 1: uSc_MultiOption3.Aligning = 2
Case 2: uSc_MultiOption3.Aligning = 1
End Select
End Sub

Private Sub Command7_Click()
Select Case uSc_MultiOption3.Positioning
Case 1: uSc_MultiOption3.Positioning = 2
Case 2: uSc_MultiOption3.Positioning = 3
Case 3: uSc_MultiOption3.Positioning = 4
Case 4: uSc_MultiOption3.Positioning = 1
End Select
End Sub

Private Sub Command8_Click()
Select Case uSc_MultiOption3.CaptionAligning
Case 1: uSc_MultiOption3.CaptionAligning = 2
Case 2: uSc_MultiOption3.CaptionAligning = 1
End Select
End Sub

Private Sub Form_Load()


chkRulesExist_M(1).Caption = "Caption 1-1"

For i = 2 To 20
    Load chkRulesExist_M(i)

    chkRulesExist_M(i).Top = chkRulesExist_M(i - 1).Top + chkRulesExist_M(i - 1).Height + 50

    chkRulesExist_M(i).Caption = "Caption 1-" & (i)
    
    chkRulesExist_M(i).Visible = True

Next



lineRules.Y1 = chkRulesExist_M(chkRulesExist_M.UBound).Top + chkRulesExist_M(chkRulesExist_M.UBound).Height + 50
lineRules.Y2 = lineRules.Y1
lineRules.X1 = 100
lineRules.X2 = picRulIn.Width - 130 - scrollRul.Width



chkRulesExist_P(1).Caption = "Caption 2-1"


chkRulesExist_P(1).Top = lineRules.Y1 + 100


For i = 2 To 10
    Load chkRulesExist_P(i)

    chkRulesExist_P(i).Top = chkRulesExist_P(i - 1).Top + chkRulesExist_P(i - 1).Height + 50

    chkRulesExist_P(i).Caption = "Caption 2-" & (i)
    
    chkRulesExist_P(i).Visible = True

Next

picRulIn.Left = 0
'picRulIn.Width = lblRules3.Width

picRulIn.Top = 0
picRulIn.Left = 0
picRulIn.Height = chkRulesExist_P(chkRulesExist_P.UBound).Top + chkRulesExist_P(chkRulesExist_P.UBound).Height + 50
picRulIn.Width = picRulOut.Width
picRulIn.BorderStyle = 0

scrollRul.Top = 0
scrollRul.Left = picRulOut.Width - scrollRul.Width - 60
scrollRul.Height = picRulOut.Height - 60

scrollRul.Value = 0
scrollRul.Max = picRulIn.Height - picRulOut.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
For i = 2 To 20
    Unload chkRulesExist_M(i)
Next
For i = 2 To 10
    Unload chkRulesExist_P(i)
Next
End Sub

Private Sub scrollRul_Change()
    picRulIn.Top = -scrollRul.Value
End Sub
