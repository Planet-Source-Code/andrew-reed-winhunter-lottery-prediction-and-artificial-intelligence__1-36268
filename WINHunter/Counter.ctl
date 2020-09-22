VERSION 5.00
Begin VB.UserControl Counter 
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   FillStyle       =   0  'Solid
   ScaleHeight     =   825
   ScaleWidth      =   1380
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   780
      Top             =   150
   End
   Begin VB.Image imgCounter 
      Height          =   345
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum Counters
    Counter0 = 101
    Counter1 = 102
    Counter2 = 103
    Counter3 = 104
    Counter4 = 105
    Counter5 = 106
    Counter6 = 107
    Counter7 = 108
    Counter8 = 109
    Counter9 = 110
End Enum

'Default Property Values:
Const m_def_TimerEnabled = 0
Const m_def_Digits = 0
Const m_def_Value = 0

'Property Variables:
Dim m_TimerEnabled  As Boolean
Dim m_Value         As Long
Dim m_Digits        As Integer

Private Sub Timer1_Timer()
    If Len(Str(m_Value + 1)) > Len(Str(m_Digits)) Then
        'm_Digits = Len(LTrim(m_Value + 1)) - 1
        Me.Digits = Len(LTrim(m_Value + 1))
    End If
    Value = m_Value + 1
End Sub

Private Sub UserControl_Initialize()
    
    Value = 0
    m_Digits = 0
    imgCounter(0).Move Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
    
End Sub

Private Sub UserControl_Paint()
    UserControl.Line (0, 0)-(New_Width * Screen.TwipsPerPixelX, 0), QBColor(8)
    UserControl.Line (0, 0)-(0, 24 * Screen.TwipsPerPixelY), QBColor(8)
    UserControl.Line ((New_Width + 1) * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-((New_Width + 1) * Screen.TwipsPerPixelX, 24 * Screen.TwipsPerPixelY), RGB(255, 255, 255) 'QBColor(1)
    UserControl.Line (Screen.TwipsPerPixelX, 24 * Screen.TwipsPerPixelY)-((New_Width + 1) * Screen.TwipsPerPixelX, 24 * Screen.TwipsPerPixelY), RGB(255, 255, 255) 'QBColor(1)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = (New_Width + 2) * Screen.TwipsPerPixelX
    UserControl.Height = 25 * Screen.TwipsPerPixelY
End Sub

Private Function New_Width() As Long

    New_Width = (((14 * m_Digits) - Fix(m_Digits / 0.99)) + 14)

End Function

Public Property Get Digits() As Integer
    Digits = m_Digits + 1
End Property

Public Property Let Digits(ByVal Digit_Count As Integer)
    
    If Digit_Count = 0 Then Exit Property
    If Digit_Count > 10 Then Exit Property
    Digit_Count = Digit_Count - 1
    Redraw_Digits Digit_Count
    
End Property

Private Sub Redraw_Digits(ByVal Digit_Qty As Integer)
Dim iShift As Integer

    Do While m_Digits < Digit_Qty
        m_Digits = m_Digits + 1
        iShift = New_Width - 13
        Load imgCounter(m_Digits)
        imgCounter(m_Digits).Visible = True
        imgCounter(m_Digits).Move iShift * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
    Loop
    Do While m_Digits > Digit_Qty
        m_Digits = m_Digits - 1
        imgCounter(m_Digits).Visible = False
        Unload imgCounter(m_Digits)
    Loop
    UserControl.Width = (New_Width + 2) * Screen.TwipsPerPixelX
    

End Sub

Public Property Get Value() As Long
Attribute Value.VB_Description = "Value being displayed"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
Dim tNumber     As String
Dim vNumbers    As Variant
Dim tString     As String
Dim i           As Integer

    m_Value = New_Value
    PropertyChanged "Value"
    'Update the display
    For i = 0 To m_Digits
        If i = 0 Then
            tString = "0"
        Else
            tString = "0\," & tString
        End If
    Next i
    tNumber = Format$(m_Value, tString)
    vNumbers = Split(tNumber, ",")
    For i = 0 To m_Digits
        Do While CLng(vNumbers(i)) > 100
            vNumbers(i) = vNumbers(i) / 10
        Loop
        Do While CLng(vNumbers(i)) > 9.999
            vNumbers(i) = vNumbers(i) - 10
        Loop
        LoadNumber i, CInt(vNumbers(i))
    Next i
    
End Property

Private Sub UserControl_InitProperties()
'Initialize Properties for User Control
    Value = m_def_Value
    m_Digits = m_def_Digits
    Redraw_Digits 0
    m_TimerEnabled = m_def_TimerEnabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'Load property values from storage
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Redraw_Digits PropBag.ReadProperty("Digits", m_def_Digits)
    m_TimerEnabled = PropBag.ReadProperty("TimerEnabled", m_def_TimerEnabled)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Write property values to storage
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Digits", m_Digits, m_def_Digits)
    Call PropBag.WriteProperty("TimerEnabled", m_TimerEnabled, m_def_TimerEnabled)
End Sub

Private Sub LoadNumber(ImageIndex As Integer, Number As Integer)
Dim iCounter As Integer

    Select Case Number
    Case 0
        iCounter = Counter0
    Case 1
        iCounter = Counter1
    Case 2
        iCounter = Counter2
    Case 3
        iCounter = Counter3
    Case 4
        iCounter = Counter4
    Case 5
        iCounter = Counter5
    Case 6
        iCounter = Counter6
    Case 7
        iCounter = Counter7
    Case 8
        iCounter = Counter8
    Case 9
        iCounter = Counter9
    Case Else
        Exit Sub
    End Select
    imgCounter(ImageIndex).Picture = LoadResPicture(iCounter, vbResBitmap)
End Sub

Public Property Get TimerEnabled() As Boolean
Attribute TimerEnabled.VB_Description = "If true starts a timer"
    TimerEnabled = m_TimerEnabled
End Property

Public Property Let TimerEnabled(ByVal New_TimerEnabled As Boolean)
    m_TimerEnabled = New_TimerEnabled
    PropertyChanged "TimerEnabled"
    Timer1.Enabled = m_TimerEnabled
End Property

