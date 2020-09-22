VERSION 5.00
Begin VB.UserControl MultiView 
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   ScaleHeight     =   3885
   ScaleWidth      =   5070
   ToolboxBitmap   =   "DynData.ctx":0000
   Begin VB.PictureBox picViewWindow 
      Height          =   3435
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      Begin VB.PictureBox picBothScrActive 
         Height          =   200
         Left            =   4200
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   1
         Top             =   3120
         Width           =   200
      End
      Begin VB.HScrollBar hscrViewWindow 
         Height          =   200
         LargeChange     =   20
         Left            =   0
         SmallChange     =   10
         TabIndex        =   2
         Top             =   3120
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.VScrollBar vscrViewWindow 
         Height          =   3000
         LargeChange     =   20
         Left            =   3960
         Max             =   -100
         SmallChange     =   10
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.PictureBox picArea 
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   4245
         TabIndex        =   4
         Top             =   0
         Width           =   4310
         Begin VB.ComboBox cboItem 
            Height          =   315
            Index           =   0
            Left            =   3600
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.TextBox txtItem 
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Line lnLower 
            BorderColor     =   &H80000014&
            Index           =   0
            Visible         =   0   'False
            X1              =   840
            X2              =   1680
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line lnUpper 
            BorderColor     =   &H80000010&
            Index           =   0
            Visible         =   0   'False
            X1              =   840
            X2              =   1680
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label lblItem 
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "MultiView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum tType
    tNone = 1
    tNumeric = 2
    tText = 3
    tCombo = 4
End Enum
Public Enum tRelation
    tSibling = 1
    tChild = 2
End Enum
Public Enum mInputType
    mvNumeric = 1
    mvString = 2
End Enum
Public Enum mInputQty
    mvSingle = 1
    mvMulti = 2
End Enum
Public Event ComboActivate(iIndex As Integer)
Public Event TextActivate(iIndex As Integer)

Private iLblVSpace As Integer
Private iYOrigin As Integer
Private iNewControl As Integer
Private iNewLine As Integer
Private iLineOffset As Integer
Private iTrueIndex() As Integer
Private iVScrollSize As Integer
Private iHScrollSize As Integer
Private iLabelWidth As Integer
Private iTextBoxWidth As Integer
Private iMaxWidth As Integer
'Private bShowCombo As Boolean
Private iNewCBO As Integer
Private iInputType As mInputType
Private iInputQty As mInputQty
Private mProperties As Object

Public Property Get Text(Optional Index As Variant)
Dim txtBox As TextBox
Dim bFoundIndex As Boolean

    If IsMissing(Index) Then Index = 1
    If IsNumeric(Index) Then
        If Index > 0 Then
            If Index > iNewControl Then
                Err.Raise 381, , "Invalid property-array index"
            Else
                Text = txtItem(Index).Text
            End If
        Else
            Err.Raise 381, , "Invalid property-array index"
        End If
    Else
        For Each txtBox In txtItem
            If LCase(Index) = LCase(txtBox.Tag) Then
                Text = txtBox.Text
                bFoundIndex = True
                Exit For
            End If
        Next
        If Not bFoundIndex Then
            Err.Raise 381, , "Invalid property-array index"
        End If
    End If

End Property

Public Property Get Count()

    Count = iNewControl

End Property

Public Property Get InputQty() As mInputQty

    InputQty = iInputQty

End Property

Public Property Let InputQty(mvInputQty As mInputQty)

    iInputQty = mvInputQty

End Property

Public Property Get InputType() As mInputType

    InputType = iInputType

End Property

Public Property Let InputType(mvInputType As mInputType)

    iInputType = mvInputType

End Property

Public Sub AddComboItem(sString As String, iData As Integer)

    cboItem(iNewCBO).AddItem sString
    cboItem(iNewCBO).ItemData(cboItem(iNewCBO).ListCount - 1) = iData
    If txtItem(cboItem(iNewCBO).Tag) = iData Then
        cboItem(iNewCBO).ListIndex = cboItem(iNewCBO).ListCount - 1
        cboItem(iNewCBO) = cboItem(iNewCBO).List(cboItem(iNewCBO).ListIndex)
    End If
    

End Sub

Private Sub cboItem_Change(Index As Integer)
    
    If Not cboItem(Index).ListIndex Then
        txtItem(cboItem(Index).Tag) = cboItem(Index).ItemData(cboItem(Index).ListIndex)
    End If

End Sub

Private Sub cboItem_Click(Index As Integer)

    If Not cboItem(Index).ListIndex Then
        txtItem(cboItem(Index).Tag) = cboItem(Index).ItemData(cboItem(Index).ListIndex)
    End If

End Sub

Private Sub cboItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub

Private Sub cboItem_KeyPress(Index As Integer, KeyAscii As Integer)

    KeyAscii = 0

End Sub

Private Sub cboItem_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub

Private Sub txtItem_Change(Index As Integer)

    If txtItem(Index).Text = "" Then
        Call txtItem_Validate(Index, False)
    End If

End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
        Case 13
            Call txtItem_Validate(Index, False)
        'Case Else
            'MsgBox KeyAscii
    End Select
    

End Sub

Private Sub txtItem_Validate(Index As Integer, Cancel As Boolean)
Dim bTest As Boolean
Dim i As Integer

    For i = 1 To iNewControl
        If txtItem(i).Text <> "" Then
            If (IsNumeric(txtItem(i).Text) And iInputType = mvNumeric) Or (Not (IsNumeric(txtItem(i).Text)) And iInputType = mvString And txtItem(i).Text <> "") Then
                If bTest And iInputQty = mvSingle Then
                    MsgBox "You have too many boxes filled."
                    txtItem(Index).Text = ""
                    Exit Sub
                End If
                bTest = True
            End If
        End If
    Next i
    
'    If lblItem(Index).Tag = True Then
'        If IsNumeric(txtItem(Index)) Then
'            If iLastCBO <> Index Then
'                cboItem(0).Clear
'                'Load Item Data for combo box
'                RaiseEvent ComboActivate(iTrueIndex(Index))
'            End If
'            iLastCBO = Index
'            cboItem(0).Top = txtItem(Index).Top
'            cboItem(0).Left = (txtItem(Index).Left + txtItem(Index).Width) + 120
'            cboItem(0).Visible = True
'            picArea.Width = (cboItem(0).Left + cboItem(0).Width) + 60
'            CalculateSize
'            cboItem(0).SetFocus
'        Else
'            If iLastCBO = Index Then
'                cboItem(0).Visible = False
'            End If
'        End If
'    End If

End Sub

Private Sub UserControl_Initialize()
    
    iLblVSpace = 400
    iYOrigin = 120
    iInputType = mvNumeric
    iInputQty = mvSingle
    
    picViewWindow.Left = 0
    picViewWindow.Top = 0
    picViewWindow.Height = UserControl.Height
    picViewWindow.Width = UserControl.Width
    
    
    picArea.BorderStyle = 0     'Flat, No 3D
    picBothScrActive.BorderStyle = 0
    picBothScrActive.Visible = False
    vscrViewWindow.LargeChange = iLblVSpace / 2
    picArea.Height = 0
    If (txtItem(0).Top + txtItem(0).Height) > picArea.Height Then
        picArea.Height = txtItem(iNewControl).Top + iLblVSpace
        ResizeViewWindow
    End If
    picArea.Width = (txtItem(0).Width + txtItem(0).Left) + 60
    lnLower(0).x1 = 0
    lnLower(0).X2 = picArea.Width
    lnUpper(0).x1 = lnLower(0).x1
    lnUpper(0).X2 = lnLower(0).X2
    CalculateSize

End Sub

Private Sub UserControl_Resize()
    
    picViewWindow.Height = UserControl.Height
    picViewWindow.Width = UserControl.Width

End Sub

Private Sub ResizeViewWindow()

    CalculateSize
    If picViewWindow.Width > 150 Then
        vscrViewWindow.Left = picViewWindow.Width - 260        '(vscrViewWindow.Width + 60)
        picBothScrActive.Left = vscrViewWindow.Left
    End If
    If picViewWindow.Height > 120 Then
        hscrViewWindow.Top = picViewWindow.Height - 250
        picBothScrActive.Top = hscrViewWindow.Top
    End If
    If picViewWindow.Width > 250 Then
        If picBothScrActive.Visible Then
            hscrViewWindow.Width = picViewWindow.Width - 250
        Else
            hscrViewWindow.Width = picViewWindow.Width - 50
        End If
    End If
    If picViewWindow.Height > 250 Then
        If picBothScrActive.Visible Then
            vscrViewWindow.Height = picViewWindow.Height - 250
        Else
            vscrViewWindow.Height = picViewWindow.Height - 50
        End If
    End If

End Sub

Private Sub CalculateSize()

    'This checks the Vertical Area vs. the View Window
    If picViewWindow.Height < picArea.Height Then
        'Area Is Greater than Window
        vscrViewWindow.Visible = True
    ElseIf hscrViewWindow.Visible And hscrViewWindow.Top < (txtItem(iNewControl).Top + txtItem(iNewControl).Height) Then
        'Area Is Greater than Window
        vscrViewWindow.Visible = True
    Else
        'Area is equal or smaller
        vscrViewWindow.Visible = False
        picArea.Top = 0
        vscrViewWindow.Value = 0
    End If

    'This checks the Horizontal Area vs. the rightmost portion
    'of the items in the area
    If vscrViewWindow.Visible And vscrViewWindow.Left < (lnLower(0).X2) Then
        'Area Is Greater than Window
        hscrViewWindow.Visible = True
    ElseIf picViewWindow.Width < (lnLower(0).X2 + 60) Then
        'Area Is Greater than Window
        hscrViewWindow.Visible = True
    Else
        'Area is equal or smaller
        hscrViewWindow.Visible = False
        picArea.Left = 0
        hscrViewWindow.Value = 0
    End If
    
    
    iVScrollSize = picArea.Height - picViewWindow.Height
    iHScrollSize = picArea.Width - picViewWindow.Width
    If vscrViewWindow.Visible And hscrViewWindow.Visible Then
        iVScrollSize = iVScrollSize + 250
        iHScrollSize = iHScrollSize + 250
        picBothScrActive.Visible = True
    Else
        picBothScrActive.Visible = False
    End If
    If iVScrollSize > 5 Then
        vscrViewWindow.Max = iVScrollSize * -1
        If Abs(iVScrollSize) < 300 Then
            vscrViewWindow.LargeChange = iVScrollSize
        Else
            vscrViewWindow.LargeChange = iVScrollSize / 2
        End If
        vscrViewWindow.SmallChange = iVScrollSize / 18
    End If
    If iHScrollSize > 5 Then
        hscrViewWindow.Max = iHScrollSize * -1
        hscrViewWindow.LargeChange = iHScrollSize / 6
        hscrViewWindow.SmallChange = iHScrollSize / 18
    End If

End Sub

Private Sub Remove(iRemoveItem As Integer)

    If iRemoveItem > 0 Then
        If iRemoveItem > iNewControl Then
            Err.Raise 381, , "Invalid property-array index"
        Else
            Unload lblItem(iNewControl)
            Unload txtItem(iNewControl)
            iNewControl = iNewControl - 1
        End If
    Else
        If iRemoveItem = -1 Then
            Do While iNewControl > 0
                Unload lblItem(iNewControl)
                Unload txtItem(iNewControl)
                iNewControl = iNewControl - 1
            Loop
            Do While iNewLine > 0
                Unload lnLower(iNewLine)
                Unload lnUpper(iNewLine)
                iNewLine = iNewLine - 1
            Loop
            Do While iNewCBO > 0
                Unload cboItem(iNewCBO)
                iNewCBO = iNewCBO - 1
            Loop
        Else
            Err.Raise 381, , "Invalid property-array index"
        End If
    End If
    If picArea.Height > picViewWindow.Height Then
        If (txtItem(iNewControl).Top + txtItem(iNewControl).Height) < picArea.Height Then
            picArea.Height = txtItem(iNewControl).Top + iLblVSpace
            ResizeViewWindow
        End If
    End If
    If picArea.Width < picViewWindow.Width Then
        ResizeViewWindow
    End If

End Sub

Private Sub hscrViewWindow_Change()

    picArea.Left = hscrViewWindow.Value

End Sub

Private Sub hscrViewWindow_Scroll()

    picArea.Left = hscrViewWindow.Value

End Sub

Private Sub picViewWindow_Resize()

    ResizeViewWindow

End Sub

Private Sub vscrViewWindow_Change()

    picArea.Top = vscrViewWindow.Value

End Sub

Private Sub vscrViewWindow_Scroll()

    picArea.Top = vscrViewWindow.Value

End Sub

Public Sub Clear()

    Remove (-1)
    'bShowCombo = False
    cboItem(0).Visible = False

End Sub

Public Sub Delete(Index As Integer)

    Remove (Index)

End Sub
Private Sub New_Add(vRelative As Variant, tRelationship As tRelation, tInputType As tType, sCaption As String, sValue As Variant)

    'bShowCombo = bCombo
    iNewControl = iNewControl + 1
    
    'Load Lable
    Load lblItem(iNewControl)
    lblItem(iNewControl).Left = lblItem(iNewControl - 1).Left
    lblItem(iNewControl).Top = lblItem(iNewControl - 1).Top + iLblVSpace
    lblItem(iNewControl).Caption = sCaption
    
    'Load Text Box
    Load txtItem(iNewControl)
    txtItem(iNewControl).Left = txtItem(iNewControl - 1).Left
    txtItem(iNewControl).Top = txtItem(iNewControl - 1).Top + iLblVSpace
    If (txtItem(iNewControl).Top + txtItem(iNewControl).Height) > picArea.Height Then
        picArea.Height = txtItem(iNewControl).Top + iLblVSpace
        ResizeViewWindow
    End If
    txtItem(iNewControl).Text = sValue
    'txtItem(iNewControl).Tag = bCombo
    
    'Show Items
    If iNewControl = 1 Then
        lblItem(iNewControl).Top = iYOrigin
        txtItem(iNewControl).Top = iYOrigin
    End If
    lblItem(iNewControl).Visible = True
    txtItem(iNewControl).Visible = True
    
    DoEvents


End Sub


Public Sub AddSeparator()
Dim iOffset As Integer

    iNewLine = iNewLine + 1
    'iTrueIndex(iNewControl) = iTrueIndex(iNewControl) + 1
    Load lnLower(iNewLine)
    Load lnUpper(iNewLine)

    iLineOffset = iLineOffset + 120
    iOffset = (lblItem(iNewControl).Top + lblItem(iNewControl).Height) + iLineOffset
    lnUpper(iNewLine).y1 = iOffset
    lnUpper(iNewLine).Y2 = iOffset
    iLineOffset = iLineOffset + 15
    iOffset = (lblItem(iNewControl).Top + lblItem(iNewControl).Height) + iLineOffset
    lnLower(iNewLine).y1 = iOffset
    lnLower(iNewLine).Y2 = iOffset
    
    If iOffset > picArea.Height Then
        picArea.Height = iOffset + iLblVSpace
        ResizeViewWindow
    End If
    
    lnUpper(iNewLine).Visible = True
    lnLower(iNewLine).Visible = True
    DoEvents

End Sub

Public Sub Add(ByVal vKey As Variant, ByVal sCaption As String, ByVal sValue As String, Optional bCombo As Boolean)
Dim l As Long

    'bShowCombo = bCombo
    If IsEmpty(bCombo) Then bCombo = False
    iNewControl = iNewControl + 1
    
    ReDim Preserve iTrueIndex(iNewControl)
    iTrueIndex(iNewControl) = iTrueIndex(iNewControl - 1) + 1
    If iLineOffset > 0 Then iTrueIndex(iNewControl) = iTrueIndex(iNewControl) + 1
    'Load Lable
    Load lblItem(iNewControl)
    lblItem(iNewControl).Left = lblItem(iNewControl - 1).Left
    lblItem(iNewControl).Top = lblItem(iNewControl - 1).Top + iLblVSpace + iLineOffset
    lblItem(iNewControl).Caption = sCaption
    lblItem(iNewControl).Tag = bCombo
    
    'Load Text Box
    Load txtItem(iNewControl)
    txtItem(iNewControl).Left = txtItem(iNewControl - 1).Left
    txtItem(iNewControl).Top = txtItem(iNewControl - 1).Top + iLblVSpace + iLineOffset
    If (txtItem(iNewControl).Top + txtItem(iNewControl).Height) > picArea.Height Then
        picArea.Height = txtItem(iNewControl).Top + iLblVSpace
        ResizeViewWindow
    End If
    txtItem(iNewControl).Text = sValue
    txtItem(iNewControl).Tag = vKey
    If bCombo Then
        'Load Combo Box
        iNewCBO = iNewCBO + 1
        Load cboItem(iNewCBO)
        cboItem(iNewCBO).Top = txtItem(iNewControl).Top
        cboItem(iNewCBO).Left = txtItem(iNewControl).Left
        cboItem(iNewCBO).Width = txtItem(iNewControl).Width + 1000
        cboItem(iNewCBO).Tag = iNewControl
        cboItem(iNewCBO).Text = "< Select One >"
        'txtItem(iNewControl).Width = cboItem(iNewControl).Width
        picArea.Width = (cboItem(iNewCBO).Left + cboItem(iNewCBO).Width) + 60
        For l = 0 To iNewLine
            lnLower(l).X2 = picArea.Width
            lnUpper(l).X2 = lnLower(l).X2
        Next
        CalculateSize
        'cboItem(iNewControl).SetFocus
    End If
    
    'Show Items
    iLineOffset = 0
    If iNewControl = 1 Then
        lblItem(iNewControl).Top = iYOrigin
        txtItem(iNewControl).Top = iYOrigin
        If bCombo Then
            cboItem(iNewCBO).Top = iYOrigin
        End If
    End If
    lblItem(iNewControl).Visible = True
    txtItem(iNewControl).Visible = True
    If bCombo Then
        cboItem(iNewCBO).Visible = True
        txtItem(iNewControl).Visible = False
    End If
    
    DoEvents


End Sub

Public Sub LoadViewer(clsProperties As Object)
Dim mProperty As Object
Dim iGroupNum As Integer
Dim iTempGroupNum As Integer
Dim iTempGroupComboNum As Integer
Dim iValueToLoad As Integer
Dim bLoadValue As Boolean

    Clear
    Set mProperties = clsProperties
    'get the property values collection from the processor
    For Each mProperty In mProperties
        'add the item
        If mProperty.Group > 99 Then
            'Must be a combo box item here
            iTempGroupNum = Int(mProperty.Group / 100)
            iTempGroupComboNum = mProperty.Group - (iTempGroupNum * 100)
            'add separator bar here if it is needed...
            If iTempGroupNum > iGroupNum And iGroupNum > 0 Then
                'add separator
                AddSeparator
            End If
            If iTempGroupComboNum = 0 Then
                'add the combo box here
                Add mProperty.Key, mProperty.Name, mProperty.Value, True
            Else
                'add the list items here
                AddComboItem mProperty.Name, mProperty.Value
            End If
        Else
            iTempGroupNum = mProperty.Group
            'add separator bar here if it is needed...
            If iTempGroupNum > iGroupNum And iGroupNum > 0 Then
                'add separator
                AddSeparator
            End If
            Add mProperty.Key, mProperty.Name, mProperty.Value
        End If
        iGroupNum = iTempGroupNum
    Next

End Sub

Public Sub SaveChanges()
Dim i As Integer

    For i = 1 To iNewControl
        mProperties(txtItem(i).Tag).Value = txtItem(i).Text
    Next i

End Sub
