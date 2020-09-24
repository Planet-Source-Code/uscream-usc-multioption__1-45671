VERSION 5.00
Begin VB.UserControl uSc_MultiOption 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "USC_MU~1.ctx":0000
   Begin VB.OptionButton optOption 
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   220
   End
   Begin VB.Label lblLabel 
      Caption         =   "uScream waz here"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "uSc_MultiOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************
'*  Copyright (C) uScream 2003 - All Rights Reserved    *
'*                                                      *
'*  Contact: uscream@vip.hr                             *
'*                                                      *
'*  CHANGE HISTORY:                                     *
'*      22.05.2003. - v 1.0                             *
'*                                                      *
'*      27.05.2003. - v 1.01                            *
'*          * Added more positioning properties         *
'*              -Aligning = 1/2 instead Horizontal = t/f*
'*          * Now every option can have his own ToolTip *
'*          * Added Border & Appearance properties      *
'********************************************************

Public Enum typePositioning
    [Left_top] = 1
    [Left_Bottom] = 2
    [Right_top] = 3
    [Right_Bottom] = 4
End Enum
Private varPositioning As typePositioning

Public Enum typeAligning
    [Horizontal] = 1
    [Vertical] = 2
End Enum
Private varAligning As typeAligning
Private varCaptionAligning As typeAligning

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private intOptionCount As Integer
Private intOptionSelected As Integer
Private boolHorizontal As Boolean
Private BackC As Long
Private ForeC As Long

Private AllTips() As String

Private Function ConvertFromSystemColor(ByVal theColor As Long) As Long
    Call OleTranslateColor(theColor, 0, ConvertFromSystemColor)
End Function

Public Property Get Border() As Boolean
    Border = UserControl.BorderStyle
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    UserControl.BorderStyle() = IIf(New_Border, 1, 0)
    PropertyChanged "Border"
    Update_Ctrl
End Property

Public Property Get Apperarance3D() As Boolean
    Apperarance3D = UserControl.Appearance
End Property

Public Property Let Apperarance3D(ByVal New_Apperarance3D As Boolean)
    UserControl.Appearance() = IIf(New_Apperarance3D, 1, 0)
    PropertyChanged "Apperarance3D"
    Update_Ctrl
End Property


Public Property Get BackColor() As OLE_COLOR
BackColor = BackC
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
BackC = theCol
If Not Ambient.UserMode Then BackO = theCol
PropertyChanged "BCOL"
Update_Ctrl
End Property


Public Property Get ForeColor() As OLE_COLOR
ForeColor = ForeC
End Property
Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
ForeC = theCol
If Not Ambient.UserMode Then ForeO = theCol
PropertyChanged "FCOL"
Update_Ctrl
End Property

Public Property Get Positioning() As typePositioning
    Positioning = varPositioning
End Property

Public Property Let Positioning(ByVal New_Positioning As typePositioning)
    'If New_OptionCount < 1 Then New_Horizontal = 1
    varPositioning = New_Positioning
    PropertyChanged "Positioning"
    Update_Ctrl
End Property


Public Property Get Aligning() As typeAligning
    Aligning = varAligning
End Property

Public Property Let Aligning(ByVal New_Aligning As typeAligning)
    'If New_OptionCount < 1 Then New_Horizontal = 1
    varAligning = New_Aligning
    PropertyChanged "Aligning"
    Update_Ctrl
End Property

Public Property Get CaptionAligning() As typeAligning
    CaptionAligning = varCaptionAligning
End Property

Public Property Let CaptionAligning(ByVal New_CaptionAligning As typeAligning)
    'If New_OptionCount < 1 Then New_Horizontal = 1
    varCaptionAligning = New_CaptionAligning
    PropertyChanged "CaptionAligning"
    Update_Ctrl
End Property



Public Property Get OptionSelected() As Integer
    OptionSelected = intOptionSelected
End Property

Public Property Let OptionSelected(ByVal New_OptionSelected As Integer)
    If New_OptionSelected < 1 Then New_OptionSelected = 1
    If New_OptionSelected > intOptionCount Then New_OptionSelected = intOptionCount
    
    intOptionSelected = New_OptionSelected
    PropertyChanged "OptionSelected"
    Update_Ctrl
End Property

Public Property Get OptionCount() As Integer
    OptionCount = intOptionCount
End Property

Public Property Let OptionCount(ByVal New_OptionCount As Integer)
    If New_OptionCount < 1 Then New_OptionCount = 1
    If New_OptionCount < intOptionSelected Then Let OptionSelected = New_OptionCount
    
    intOptionCount = New_OptionCount
    PropertyChanged "OptionCount"
    Update_Ctrl
End Property

Public Property Get Caption() As String
    Caption = lblLabel.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblLabel.Caption() = New_Caption
    PropertyChanged "Caption"
    Update_Ctrl
End Property


Private Sub optOption_Click(index As Integer)
If ((optOption(index).Value = True) And (intOptionSelected <> index)) Then
    OptionSelected = index
    optOption(index).SetFocus
End If
End Sub


Private Sub UserControl_InitProperties()
varCaptionAligning = 1
varAligning = 1
varPositioning = 1
boolHorizontal = True
intOptionCount = 3
intOptionSelected = 1
BackC = GetSysColor(COLOR_BTNFACE)
ForeC = GetSysColor(COLOR_BTNTEXT)
Update_Ctrl
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Appearance = PropBag.ReadProperty("Appearance3D", 0)
    UserControl.BorderStyle = PropBag.ReadProperty("Border", 0)
    varCaptionAligning = PropBag.ReadProperty("CaptionAligning", 1)
    varAligning = PropBag.ReadProperty("Aligning", 1)
    varPositioning = PropBag.ReadProperty("Positioning", 1)
    lblLabel.Caption = PropBag.ReadProperty("Caption", lblLabel.Caption)
    intOptionCount = PropBag.ReadProperty("OptionCount", intOptionCount)
    intOptionSelected = PropBag.ReadProperty("OptionSelected", intOptionSelected)
    boolHorizontal = PropBag.ReadProperty("Horizontal", boolHorizontal)
    BackC = PropBag.ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
    ForeC = PropBag.ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))
Update_Ctrl
End Sub


Private Sub UserControl_Resize()
    Update_Position
    Update_Label
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance3D", UserControl.Appearance, 0)
    Call PropBag.WriteProperty("Border", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("CaptionAligning", varCaptionAligning)
    Call PropBag.WriteProperty("Aligning", varAligning)
    Call PropBag.WriteProperty("Positioning", varPositioning)
    Call PropBag.WriteProperty("Caption", lblLabel.Caption)
    Call PropBag.WriteProperty("OptionCount", intOptionCount)
    Call PropBag.WriteProperty("OptionSelected", intOptionSelected)
    Call PropBag.WriteProperty("Horizontal", boolHorizontal)
    Call PropBag.WriteProperty("BCOL", BackC)
    Call PropBag.WriteProperty("FCOL", ForeC)
End Sub


Private Sub Update_Ctrl()
    Update_OptionCount
    Update_OptionSelected
    Update_Position
    Update_Label
    Update_Colors
End Sub

Private Sub Update_Position()
optOption(1).Top = 0
optOption(1).Left = 0

For i = 2 To intOptionCount
    
    Select Case varAligning
    Case Horizontal:
        optOption(i).Left = optOption(i - 1).Left + optOption(i - 1).Width + 10 'Adjusted for XP
        optOption(i).Top = optOption(i - 1).Top
    Case Vertical:
        optOption(i).Top = optOption(i - 1).Top + optOption(i - 1).Height + 10
        optOption(i).Left = optOption(i - 1).Left
    End Select
Next


For i = 1 To intOptionCount
    Select Case varPositioning
    Case Left_top, Left_Bottom:

    Case Right_Bottom, Right_top:
        Select Case varAligning
        Case Horizontal:
            optOption(i).Left = optOption(i).Left + (UserControl.Width - ((optOption(1).Width * intOptionCount) + 10))
        Case Vertical:
            optOption(i).Left = UserControl.Width - optOption(1).Width
        End Select
    End Select
    
    Select Case varPositioning
    Case Left_top, Right_top:
        
    Case Left_Bottom, Right_Bottom:
        Select Case varAligning
        Case Horizontal:
            optOption(i).Top = UserControl.Height - optOption(1).Height
        Case Vertical:
            optOption(i).Top = optOption(i).Top + (UserControl.Height - ((optOption(1).Height * intOptionCount) + 10))
        End Select
    End Select
    
Next



ReDim Preserve AllTips(intOptionCount)
For i = 1 To intOptionCount
    optOption(i).ToolTipText = AllTips(i)
    optOption(i).Visible = True
Next

End Sub


Private Sub Update_Label()


Select Case varCaptionAligning
Case Horizontal:
    lblLabel.Width = UserControl.Width
    lblLabel.Left = 0
    Select Case varPositioning
        Case Right_Bottom, Left_Bottom:
            If optOption(optOption.lBound).Top > 0 Then
                lblLabel.Height = optOption(optOption.lBound).Top
            Else
                lblLabel.Height = 0
            End If
            lblLabel.Top = 0
        Case Right_top, Left_top:
            If UserControl.Height - optOption(optOption.UBound).Top - optOption(optOption.UBound).Height > 0 Then
                lblLabel.Height = UserControl.Height - optOption(optOption.UBound).Top - optOption(optOption.UBound).Height
            Else
                lblLabel.Height = 0
            End If
            lblLabel.Top = optOption(optOption.UBound).Top + optOption(optOption.UBound).Height
    End Select
    
Case Vertical:
    lblLabel.Height = UserControl.Height
    lblLabel.Top = 0
    Select Case varPositioning
        Case Left_top, Left_Bottom:
            If UserControl.Width - optOption(optOption.UBound).Left - optOption(optOption.UBound).Width > 0 Then
                lblLabel.Width = UserControl.Width - optOption(optOption.UBound).Left - optOption(optOption.UBound).Width
            Else
                lblLabel.Width = 0
            End If
            lblLabel.Left = optOption(optOption.UBound).Left + optOption(optOption.UBound).Width
        Case Right_top, Right_Bottom:
            If optOption(optOption.lBound).Left > 0 Then
                lblLabel.Width = optOption(optOption.lBound).Left
            Else
                lblLabel.Width = 0
            End If
            lblLabel.Left = 0
    End Select
End Select


End Sub

Private Sub Update_OptionCount()

    For i = optOption.UBound To optOption.lBound + 1 Step -1
        If IsObject(optOption(i)) Then Unload optOption(i)
    Next

    For i = 2 To intOptionCount
        Load optOption(i)
    Next

End Sub

Private Sub Update_OptionSelected()
    If Not (intOptionSelected < optOption.lBound Or intOptionSelected > optOption.UBound) Then
        optOption(intOptionSelected).Value = True
    End If
End Sub

Private Sub Update_Colors()

    UserControl.BackColor = ConvertFromSystemColor(BackC)
    lblLabel.BackColor = UserControl.BackColor

    UserControl.ForeColor = ConvertFromSystemColor(ForeC)
    lblLabel.ForeColor = UserControl.ForeColor
    

For i = optOption.lBound To optOption.UBound
    optOption(i).BackColor = UserControl.BackColor
    optOption(i).ForeColor = UserControl.ForeColor
Next
End Sub

Public Function SetOptionToolTip(index As Integer, ToolTip As String) As Byte
    If index <= optOption.UBound And index >= optOption.lBound Then
        AllTips(index) = ToolTip
        Update_Ctrl
        SetOptionToolTip = 1
    Else
        SetOptionToolTip = 0
    End If
End Function
