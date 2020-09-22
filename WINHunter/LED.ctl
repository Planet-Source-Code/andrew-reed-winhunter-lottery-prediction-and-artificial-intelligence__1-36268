VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl LED 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   ScaleHeight     =   2190
   ScaleWidth      =   1470
   ToolboxBitmap   =   "LED.ctx":0000
   Begin VB.Timer Blinker 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   960
      Top             =   1440
   End
   Begin PicClip.PictureClip pcLEDS 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   3387
      _Version        =   393216
      Rows            =   8
      Cols            =   3
      Picture         =   "LED.ctx":0312
   End
End
Attribute VB_Name = "LED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum LEDState
    LEDOn = 0
    LEDOff = 1
    LEDBlink = -1
End Enum

Public Enum LEDColor
    LEDRed = 0
    LEDGreen = 3
    LEDYellow = 6
    LEDBlue = 9
End Enum

Public Enum LEDShape
    LEDRound = 0
    LEDSquare = 12
End Enum
'Default Property Values:
Const m_def_State = 0
Const m_def_Color = 0
Const m_def_Shape = 0

Private mState As LEDState
Private mBlinkState As Boolean
Private mColor As LEDColor
Private mShape As LEDShape
Private mEnabled As Boolean
Private mCellNum As Long


Private Sub Redraw_LED()

    'the LED is not blinking
    If Not mState Or Not mEnabled Then
        If mEnabled Then
            'Get the color, which returns the initial "ON" state color
            mCellNum = mColor + mState + mShape
        Else
            mCellNum = mColor + 2 + mShape
        End If
        UserControl.MaskPicture = pcLEDS.GraphicCell(mCellNum)
        UserControl.Picture = pcLEDS.GraphicCell(mCellNum)
    End If

End Sub

Private Sub Blinker_Timer()
'Dim CellNum As Long

    'the LED is blinking
    If mState And mEnabled Then
        'Get the color, which returns the initial "ON" state color
        mCellNum = mColor + Abs(mBlinkState) + mShape
        UserControl.MaskPicture = pcLEDS.GraphicCell(mCellNum)
        UserControl.Picture = pcLEDS.GraphicCell(mCellNum)
        mBlinkState = Not mBlinkState
    End If


End Sub

Private Sub UserControl_Initialize()

    UserControl.Width = pcLEDS.CellWidth * 13
    UserControl.Height = pcLEDS.CellHeight * 13
    UserControl.MaskColor = RGB(255, 0, 255)
    'UserControl.MaskPicture = pcLEDS.GraphicCell(0)

End Sub

Private Sub UserControl_InitProperties()
'Initialize Properties for User Control
    mState = m_def_State
    mColor = m_def_Color
    mShape = m_def_Shape
    mEnabled = True
    Redraw_LED
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'Load property values from storage
    mState = PropBag.ReadProperty("State", m_def_State)
    mColor = PropBag.ReadProperty("Color", m_def_Color)
    mShape = PropBag.ReadProperty("Shape", m_def_Shape)
    mEnabled = PropBag.ReadProperty("Enabled", True)
    If Not mState = LEDBlink Then
        'set the initial blink state
        mBlinkState = Not mState
    End If
End Sub

Private Sub UserControl_Resize()

    UserControl.Width = pcLEDS.CellHeight * 13
    UserControl.Height = pcLEDS.CellHeight * 13
    Exit Sub

End Sub

Private Sub UserControl_Show()

    Redraw_LED

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Write property values to storage
    Call PropBag.WriteProperty("State", mState, m_def_State)
    Call PropBag.WriteProperty("Color", mColor, m_def_Color)
    Call PropBag.WriteProperty("Shape", mShape, m_def_Shape)
    Call PropBag.WriteProperty("Enabled", mEnabled, True)
End Sub


Public Property Let Enabled(ByVal bEnabled As Boolean)
    mEnabled = bEnabled
    Redraw_LED
End Property
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let State(ByVal eState As LEDState)
    mState = eState
    If Not mState = LEDBlink Then
        'set the initial blink state
        mBlinkState = Not mState
        Blinker.Enabled = False
        Redraw_LED
    Else
        Blinker.Enabled = True
    End If
End Property
Public Property Get State() As LEDState
    State = mState
End Property

Public Property Let Shape(ByVal eShape As LEDShape)
    mShape = eShape
    Redraw_LED
End Property
Public Property Get Shape() As LEDShape
    Shape = mShape
End Property

Public Property Let Color(ByVal eColor As LEDColor)
    mColor = eColor
    Redraw_LED
End Property
Public Property Get Color() As LEDColor
    Color = mColor
End Property

