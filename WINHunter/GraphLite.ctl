VERSION 5.00
Begin VB.UserControl GraphLite 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "GraphLite.ctx":0000
End
Attribute VB_Name = "GraphLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_VerticalTickInterval = 1
Const m_def_HorizontalTickFrequency = 1
Const m_def_Title = ""
Const m_def_PlotPoints = 0
Const m_def_DisplayLegend = 0
Const m_def_ChartType = 0
Const m_def_LowScale = 0
Const m_def_HighScale = 100

'Property Variables:
Dim m_TitleFont As Font
Dim m_VerticalTickInterval As Variant
Dim m_HorizontalTickFrequency As Variant
Dim m_Title As String
Dim m_PlotPoints As Boolean
Dim m_DisplayLegend As Boolean
Dim m_LowScale As Double
Dim m_HighScale As Double
Dim m_Columns As Integer
Enum ChartTypes
   Bar = 0
   Line = 1
End Enum
Dim m_ChartType As ChartTypes

'internal data storage
Dim PlotData() As Variant
Dim PlotColors(15) As Long
Dim Legends(15) As String
Dim ChartWidth As Long
Dim ChartHeight As Long
Dim TitleOffset As Long
Private mCol As Collection

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
   ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   UserControl.ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"
End Property

Public Property Get HighScale() As Double
   HighScale = m_HighScale
End Property

Public Property Let HighScale(ByVal New_HighScale As Double)
   m_HighScale = New_HighScale
   PropertyChanged "HighScale"
End Property
Public Property Get vLowScale() As Double
   vLowScale = m_LowScale
End Property

Public Property Let vLowScale(ByVal New_LowScale As Double)
   m_LowScale = New_LowScale
   PropertyChanged "vLowScale"
End Property

Public Property Get GotData() As Boolean

    If mCol.Count > 0 Then
        GotData = True
    Else
        GotData = False
    End If

End Property

Public Property Get Columns() As Integer
   Columns = m_Columns
End Property

Public Property Let Columns(ByVal New_Columns As Integer)
   m_Columns = New_Columns
   PropertyChanged "Columns"
End Property
Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   UserControl.Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
   Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set UserControl.Font = New_Font
   PropertyChanged "Font"
End Property

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()

    Set mCol = New Collection

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get ChartType() As ChartTypes
   ChartType = m_ChartType
End Property

Public Property Let ChartType(ByVal New_ChartType As ChartTypes)
   m_ChartType = New_ChartType
   PropertyChanged "ChartType"
End Property

Public Sub SetSeriesOptions(Series As Integer, Optional PlotColor As Long, Optional Legend As String)
If Series < 0 Or Series > 15 Then
   MsgBox "Too many data series -- limit is 15"
   Exit Sub
End If
If Len(Legend) > 0 Then
   Legends(Series) = Legend
End If
If Not IsMissing(PlotColor) Then
   PlotColors(Series) = PlotColor
End If

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
Dim i As Integer
   Set Font = Ambient.Font
   m_ChartType = m_def_ChartType
   m_DisplayLegend = m_def_DisplayLegend
   m_PlotPoints = m_def_PlotPoints
   m_Title = m_def_Title
   
   For i = 0 To 15
      PlotColors(i) = -1
   Next i
   
   m_VerticalTickInterval = m_def_VerticalTickInterval
   m_HorizontalTickFrequency = m_def_HorizontalTickFrequency
   Set m_TitleFont = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
   UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
   UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
   Set Font = PropBag.ReadProperty("Font", Ambient.Font)
   UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
   UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
   m_ChartType = PropBag.ReadProperty("ChartType", m_def_ChartType)
   m_DisplayLegend = PropBag.ReadProperty("DisplayLegend", m_def_DisplayLegend)
   m_PlotPoints = PropBag.ReadProperty("PlotPoints", m_def_PlotPoints)
   m_Title = PropBag.ReadProperty("Title", m_def_Title)
   m_VerticalTickInterval = PropBag.ReadProperty("VerticalTickInterval", m_def_VerticalTickInterval)
   m_HorizontalTickFrequency = PropBag.ReadProperty("HorizontalTickFrequency", m_def_HorizontalTickFrequency)
   Set m_TitleFont = PropBag.ReadProperty("TitleFont", Ambient.Font)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
   Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
   Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
   Call PropBag.WriteProperty("Font", Font, Ambient.Font)
   Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
   Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
   Call PropBag.WriteProperty("ChartType", m_ChartType, m_def_ChartType)
   Call PropBag.WriteProperty("DisplayLegend", m_DisplayLegend, m_def_DisplayLegend)
   Call PropBag.WriteProperty("PlotPoints", m_PlotPoints, m_def_PlotPoints)
   Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
   Call PropBag.WriteProperty("VerticalTickInterval", m_VerticalTickInterval, m_def_VerticalTickInterval)
   Call PropBag.WriteProperty("HorizontalTickFrequency", m_HorizontalTickFrequency, m_def_HorizontalTickFrequency)
   Call PropBag.WriteProperty("TitleFont", m_TitleFont, Ambient.Font)
End Sub

Public Function Refresh() As Variant

UserControl.Cls
ChartWidth = UserControl.ScaleWidth - 120
ChartHeight = UserControl.ScaleHeight - 120
TitleOffset = 0
If Len(m_Title) > 0 Then
   DrawTitle
End If
If m_DisplayLegend Then
   DrawLegend
End If

PlotChart

End Function
Public Property Get DisplayLegend() As Boolean
   DisplayLegend = m_DisplayLegend
End Property

Public Property Let DisplayLegend(ByVal New_DisplayLegend As Boolean)
   m_DisplayLegend = New_DisplayLegend
   PropertyChanged "DisplayLegend"
End Property

Public Sub RegisterData(SeriesNumber As Integer, RegData As Variant, ValData As Variant)
'ReDim PlotData(UBound(RegData), UBound(RegData, 2)) As Variant
'Dim n As Integer
'Dim i As Integer
Dim CurrentIndex As Integer
On Error GoTo RegisterError

CurrentIndex = (SeriesNumber * 2)
mCol.Add RegData, "ASeries" & SeriesNumber
mCol.Add ValData, "BSeries" & SeriesNumber

'RegData() = x point
'ValData() = y point

'For n = 0 To UBound(RegData)
'   For i = 0 To UBound(RegData, 2)
'      PlotData(n, i) = RegData(n, i)
'      If n > 0 Then 'a data series
'         If PlotColors(n - 1) = -1 Then PlotColors(n - 1) = QBColor(n - 1)
'      End If
'   Next
'Next
Exit Sub
RegisterError:

Select Case Err.Number
    Case 457
        mCol.Remove (CurrentIndex)
        mCol.Remove (CurrentIndex - 1)
        Resume
    Case Else
        MsgBox "error in graph"
End Select


End Sub

Public Property Get PlotPoints() As Boolean
   PlotPoints = m_PlotPoints
End Property

Public Property Let PlotPoints(ByVal New_PlotPoints As Boolean)
   m_PlotPoints = New_PlotPoints
   PropertyChanged "PlotPoints"
End Property

Public Property Get Title() As String
   Title = m_Title
End Property

Public Property Let Title(ByVal New_Title As String)
   m_Title = New_Title
   PropertyChanged "Title"
End Property

Private Sub DrawTitle()
Dim f As Font

Set f = UserControl.Font
Set UserControl.Font = m_TitleFont
UserControl.CurrentX = (UserControl.ScaleWidth - UserControl.TextWidth(m_Title)) / 2
UserControl.Print m_Title
TitleOffset = UserControl.CurrentY
Set UserControl.Font = f

End Sub
Private Sub DrawLegend()
Dim LegendLeft As Long
Dim n As Integer, y As Long

On Error GoTo LegendError:

LegendLeft = UserControl.ScaleWidth
If m_DisplayLegend Then
   For n = 0 To UBound(PlotData) - 1
      If LegendLeft > (UserControl.ScaleWidth - (UserControl.TextWidth(Legends(n)) + 180)) Then
         LegendLeft = UserControl.ScaleWidth - (UserControl.TextWidth(Legends(n)) + 180)
      End If
   Next n
   ChartWidth = LegendLeft - 180
   UserControl.CurrentY = (UserControl.ScaleHeight - _
      (UserControl.TextHeight("X") + 120) * (UBound(PlotData) + 2)) / 2
   For n = 0 To UBound(PlotData) - 1
      y = UserControl.CurrentY + UserControl.TextHeight("X")
      UserControl.Line (LegendLeft, y + ((UserControl.TextHeight(Legends(n)) - 60) / 2))-Step(60, 60), PlotColors(n), BF
      UserControl.CurrentX = LegendLeft + 120
      UserControl.CurrentY = y
      UserControl.Print Legends(n)
   Next n
End If

LegendExit:

   Exit Sub
      
LegendError:
   If Err = 9 Then
      'fail silently, not initialized yet
   Else
      Err.Raise 32007, "GraphLiteProject.GraphLite", _
         "Error " & Err & " plotting legend: " & Error$(Err)
   End If
   Resume LegendExit
   
End Sub
Private Sub PlotChart()
Dim Columns As Long
Dim n As Long, i As Long, d As Double
Dim x As Long, y As Long
Dim x1 As Long, y1 As Long
Dim PlotTop As Long, PlotBottom As Long, PlotLeft As Long, PlotRight As Long
Dim TickString As String
Dim LowTick As Double
Dim BarWidth As Integer
Dim RegData As Variant
Dim ValData As Variant

On Error GoTo PlotError

'determine horizontal extent
'Columns = UBound(PlotData, 2) + 1
Columns = m_Columns



'adjust vertical scale if necessary
'For n = 1 To UBound(PlotData)
'   For i = 0 To UBound(PlotData, 2)
'      If m_HighScale < PlotData(n, i) Then m_HighScale = PlotData(n, i)
'      If m_LowScale > PlotData(n, i) Then m_LowScale = PlotData(n, i)
'   Next i
'Next n

'define plot area
PlotLeft = 120 'may be overridden later
PlotRight = ChartWidth
PlotTop = TitleOffset
PlotBottom = UserControl.ScaleHeight - (UserControl.TextHeight("X") * 2)

'determine vertical tick scale
If m_LowScale / m_VerticalTickInterval = Int(m_LowScale / m_VerticalTickInterval) Then
   LowTick = m_LowScale
Else
   LowTick = Int(m_LowScale / m_VerticalTickInterval) * m_VerticalTickInterval
End If

'determine left spacing
'check vertical captions
For d = LowTick To HighScale Step m_VerticalTickInterval
   If PlotLeft < (UserControl.TextWidth(Format$(d)) + 120) Then
      PlotLeft = (UserControl.TextWidth(Format$(d)) + 120)
   End If
Next d
'check caption for first horizontal tick
'If PlotLeft < (UserControl.TextWidth(PlotData(0, 0)) / 2) + 60 Then
'   PlotLeft = (UserControl.TextWidth(PlotData(0, 0)) / 2) + 60
'End If

'draw row ticks
For d = LowTick To HighScale Step m_VerticalTickInterval
   y = PlotBottom - (PlotBottom - PlotTop) * ((d - LowTick) / (m_HighScale - LowTick))
   UserControl.Line (PlotLeft, y)-Step(60, 0)
   UserControl.CurrentX = PlotLeft - (UserControl.TextWidth(Format$(d)) + 60)
   UserControl.CurrentY = y - (UserControl.TextHeight("X") * 0.5)
   UserControl.Print Format$(d)
Next d

'draw plot box
UserControl.Line (PlotLeft, PlotTop)-(PlotRight, PlotBottom), , B

'draw column ticks and captions
For i = 1 To Columns Step m_HorizontalTickFrequency
   x = PlotLeft + (((PlotRight - PlotLeft) / (Columns)) * i)
   UserControl.Line (x, PlotBottom)-Step(0, -60)
   'UserControl.CurrentX = x - (UserControl.TextWidth(PlotData(0, i)) / 2)
   UserControl.CurrentX = x - (UserControl.TextWidth(i) / 2)
   UserControl.CurrentY = PlotBottom + (UserControl.TextHeight("X") * 0.5)
   UserControl.Print i
Next i
'base barwidth on series and points
'If m_ChartType = Bar Then
'   BarWidth = (PlotRight - (PlotLeft + 60)) / (Columns * UBound(PlotData)) - 30
'   If BarWidth <= 15 Then BarWidth = 30
'End If


'plot graph
For i = 1 To mCol.Count Step 2
    RegData = mCol(i)
    ValData = mCol(i + 1)
    For n = 1 To UBound(ValData)
        'determine coordinates
        x = PlotLeft + (((PlotRight - PlotLeft) / (Columns)) * RegData(n)) - 15
        If ValData(n) = LowTick Then
           y = PlotBottom - 15
        Else
           y = (PlotBottom - ((ValData(n) - LowTick) / (m_HighScale - LowTick) _
              * (PlotBottom - PlotTop))) - 15
        End If
        Select Case m_ChartType
            'Case Bar
            '   'adjust x for series
            '   x = PlotLeft + (((PlotRight - PlotLeft) / Columns) * i) - 15
            '   'x = x + 30 + ((n - 1) * (BarWidth + 30))
            '   x = x + 30 + ((n) * (BarWidth + 30))
            '   UserControl.Line (x, y)-(x + BarWidth, PlotBottom - 15), PlotColors((i / 2) - 1), BF
            Case Line
               'draw data point
               If m_PlotPoints Then
                  UserControl.Line (x, y)-Step(30, 30), PlotColors((i / 2) - 1), BF
                  'UserControl.Circle Step(-15, -15), 60, PlotColors((i / 2) - 1)
               End If
               'draw data graph
               If n <> 1 Then
                  UserControl.Line (x + 15, y + 15)-(x1 + 15, y1 + 15), PlotColors((i / 2) - 1)
               End If
               x1 = x
               y1 = y
        End Select
    Next n
Next i
'        'plot graph
'        For n = 1 To UBound(PlotData)
'           For i = 0 To UBound(PlotData, 2)
'              'determine coordinates
'              x = PlotLeft + (((PlotRight - PlotLeft) / (Columns - 1)) * i) - 15
'              If PlotData(n, i) = LowTick Then
'                 y = PlotBottom - 15
'              Else
'                 y = (PlotBottom - ((PlotData(n, i) - LowTick) / (m_HighScale - LowTick) _
'                    * (PlotBottom - PlotTop))) - 15
'              End If
'              Select Case m_ChartType
'              Case Bar
'                 'adjust x for series
'                 x = PlotLeft + (((PlotRight - PlotLeft) / Columns) * i) - 15
'                 x = x + 30 + ((n - 1) * (BarWidth + 30))
'                 UserControl.Line (x, y)-(x + BarWidth, PlotBottom - 15), PlotColors(n - 1), BF
'              Case Line
'                 'draw data point
'                 If m_PlotPoints Then
'                    UserControl.Line (x, y)-Step(30, 30), PlotColors(n - 1), BF
'                 End If
'                 'draw data graph
'                 If i <> 0 Then
'                    UserControl.Line (x + 15, y + 15)-(x1 + 15, y1 + 15), PlotColors(n - 1)
'                 End If
'                 x1 = x
'                 y1 = y
'              End Select
'           Next i
'        Next n

PlotExit:

   Exit Sub
      
PlotError:
   If Err = 9 Then
    Resume
      'fail silently, not initialized yet
   Else
      Err.Raise 32007, "GraphLiteProject.GraphLite", _
         "Error " & Err & " plotting graph: " & Error$(Err)
   End If
   Resume PlotExit
      
End Sub
Public Property Get VerticalTickInterval() As Variant
   VerticalTickInterval = m_VerticalTickInterval
End Property

Public Property Let VerticalTickInterval(ByVal New_VerticalTickInterval As Variant)
   m_VerticalTickInterval = New_VerticalTickInterval
   PropertyChanged "VerticalTickInterval"
End Property

Public Property Get HorizontalTickFrequency() As Variant
   HorizontalTickFrequency = m_HorizontalTickFrequency
End Property

Public Property Let HorizontalTickFrequency(ByVal New_HorizontalTickFrequency As Variant)
   m_HorizontalTickFrequency = New_HorizontalTickFrequency
   PropertyChanged "HorizontalTickFrequency"
End Property

Public Property Get TitleFont() As Font
   Set TitleFont = m_TitleFont
End Property

Public Property Set TitleFont(ByVal New_TitleFont As Font)
   Set m_TitleFont = New_TitleFont
   PropertyChanged "TitleFont"
End Property

