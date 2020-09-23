VERSION 5.00
Begin VB.UserControl dProgressBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   HasDC           =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   1260
   ToolboxBitmap   =   "dProgressBar.ctx":0000
End
Attribute VB_Name = "dProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2

Enum pBarBorder
    None = 0
    Fixed = 1
    FixedFlat = 2
End Enum

Private pBar_Min As Long
Private pBar_Max As Long
Private pBar_Value As Long
Private pBar_Color As OLE_COLOR
Private pBarBStyle As pBarBorder
Private m_ShowValue As Boolean

Event ProgressChange()
'Event Declarations:
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

Private Sub RenderPBar()
On Error Resume Next
Dim NumStr As String
Dim pValue As Long

    With UserControl
        .Cls
        'Processbar Value
        pValue = (pBar_Value - pBar_Min)
        'Draw the Processbar
        If (pValue > 0) Then
            UserControl.Line (0, 0)-(pValue, .ScaleHeight), TranslateColor(BarColor), BF
        End If
        
        'This display the Precent value
        If (m_ShowValue) Then
            NumStr = pBar_Value & "%"
            .FontBold = True
            'Center the display value
            .CurrentX = (.ScaleWidth - .TextWidth(NumStr)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(NumStr)) \ 2
            'Print the display value
            UserControl.Print NumStr
        End If
        
    End With
    
    RaiseEvent ProgressChange
    pValue = 0
    NumStr = vbNullString
    
End Sub

Private Sub FlatBorder(ByVal hwnd As Long, MakeControlFlat As Boolean)
Dim TFlat As Long
    'This little sub is used to give a fixed border style a Flat Style
    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    If MakeControlFlat Then
        TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    Else
        TFlat = TFlat And Not WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    End If
    
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Sub

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Sub UserControl_InitProperties()
    'Init default Props
    Min = 0
    Max = 100
    Value = 50
    BarColor = vbHighlight
    ShowValue = True
    BorderStyle = FixedFlat
    ForeColor = &HFFFFFF
End Sub

Private Sub UserControl_Paint()
    'Redraw the processbar
    Call RenderPBar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Min", Min, 0)
    Call PropBag.WriteProperty("Max", Max, 100)
    Call PropBag.WriteProperty("Value", Value, 50)
    Call PropBag.WriteProperty("BarColor", BarColor, vbHighlight)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", BorderStyle, 2)
    Call PropBag.WriteProperty("ShowValue", ShowValue, True)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Min = PropBag.ReadProperty("Min", 0)
    Max = PropBag.ReadProperty("Max", 100)
    Value = PropBag.ReadProperty("Value", 50)
    BarColor = PropBag.ReadProperty("BarColor", vbHighlight)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 2)
    ShowValue = PropBag.ReadProperty("ShowValue", True)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.ScaleHeight = (pBar_Max - pBar_Min)
    UserControl.ScaleWidth = (pBar_Max - pBar_Min)
    Call UserControl_Paint
End Sub

Public Property Get Min() As Long
    Min = pBar_Min
End Property

Public Property Let Min(ByVal NewMin As Long)
    'Check that we are within the correct rages
    If (NewMin >= pBar_Max) Then NewMin = (Max - 1)
    If (NewMin < 0) Then NewMin = 0
    pBar_Min = NewMin
    If (Value < pBar_Min) Then Value = pBar_Min
    UserControl_Resize
    PropertyChanged "Min"
End Property

Public Property Get Max() As Long
    Max = pBar_Max
End Property

Public Property Let Max(ByVal NewMax As Long)
    'Check that we are within the correct rages
    If (NewMax < 1) Then NewMax = 1
    If (NewMax <= pBar_Min) Then NewMax = pBar_Min + 1
    pBar_Max = NewMax
    If (Value > pBar_Max) Then Value = NewMax
    'Redraw the processbar
    UserControl_Resize
    PropertyChanged "Max"
End Property

Public Property Get Value() As Long
    Value = pBar_Value
End Property

Public Property Let Value(ByVal NewValue As Long)
    'Check that we are within the correct rages
    If (NewValue > pBar_Max) Then NewValue = pBar_Max
    If (NewValue < pBar_Min) Then NewValue = pBar_Min
    
    pBar_Value = NewValue
    'Redraw the processbar
    Call UserControl_Resize
    PropertyChanged "Value"
End Property

Public Property Get BarColor() As OLE_COLOR
    BarColor = pBar_Color
End Property

Public Property Let BarColor(ByVal NewBarColor As OLE_COLOR)
    pBar_Color = NewBarColor
    'Redraw the processbar
    Call UserControl_Resize
    PropertyChanged "BarColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    'Redraw the processbar
    Call UserControl_Resize
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As pBarBorder
    BorderStyle = pBarBStyle
End Property

Public Property Let BorderStyle(ByVal NewBorder As pBarBorder)
    pBarBStyle = NewBorder
    'Turn off the flat border
    FlatBorder UserControl.hwnd, False
    'Turn on the standred thick border
    If (pBarBStyle = Fixed) Then
        UserControl.BorderStyle = 1
    End If
    'Turns on the new flat border
    If (pBarBStyle = FixedFlat) Then
        UserControl.BorderStyle = 1
        FlatBorder UserControl.hwnd, True
    End If
    'Turn off all bordering
    If (pBarBStyle = None) Then
        UserControl.BorderStyle = 0
    End If
    'Redraw the processbar
    Call UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

Public Property Get ShowValue() As Boolean
    ShowValue = m_ShowValue
End Property

Public Property Let ShowValue(ByVal vNewValue As Boolean)
    m_ShowValue = vNewValue
    'Redraw the processbar
    Call UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    'Redraw the processbar
    Call UserControl_Resize
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

