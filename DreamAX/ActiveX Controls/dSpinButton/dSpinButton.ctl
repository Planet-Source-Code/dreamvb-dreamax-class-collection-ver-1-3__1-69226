VERSION 5.00
Begin VB.UserControl dSpinButton 
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   ScaleHeight     =   64
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
   ToolboxBitmap   =   "dSpinButton.ctx":0000
   Begin VB.Timer Tmr1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   1080
   End
   Begin VB.CommandButton cmddown 
      Height          =   495
      Left            =   0
      Picture         =   "dSpinButton.ctx":00FA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   435
   End
   Begin VB.CommandButton cmdup 
      Height          =   495
      Left            =   0
      Picture         =   "dSpinButton.ctx":0137
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "dSpinButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Value As Long
Private m_Max As Long

Private m_UpDownState As Integer

Event Change()

Private Sub UpDownCheck(Button As Integer, State As Integer, mEnabled As Boolean)
    'Only update if Mouse button is Left.
    If (Button <> vbLeftButton) Then
        Exit Sub
    Else
        m_UpDownState = State '0 = up,1 = down
        'Enable of disable timer.
        Tmr1.Enabled = mEnabled
    End If
End Sub

Private Sub cmddown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpDownCheck Button, 1, True
End Sub

Private Sub cmddown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpDownCheck Button, 1, False
End Sub

Private Sub cmdup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpDownCheck Button, 0, True
End Sub

Private Sub cmdup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpDownCheck 1, 0, False
End Sub

Private Sub Tmr1_Timer()
    
    If (m_UpDownState = 0) Then
        'Move Value Up.
        Value = (Value + 1)
    Else
        'Move Value Down.
        Value = (Value - 1)
    End If
    
    'Check that we are in the correct range.
    If (Value <= 0) Then Value = 0
    If (Value >= Max) Then Value = Max
    
    'Update Event
    RaiseEvent Change
    'Allow other sys tasks to process.
    DoEvents
    
End Sub

Private Sub UserControl_InitProperties()
    Value = 0
    Max = 100
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", 0)
    Max = PropBag.ReadProperty("Max", 100)
    cmdup.Enabled = PropBag.ReadProperty("Enabled", True)
    cmddown.Enabled = PropBag.ReadProperty("Enabled", True)
    cmdup.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    cmddown.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", Value, 0)
    Call PropBag.WriteProperty("Max", Max, 100)
    Call PropBag.WriteProperty("Enabled", cmdup.Enabled, True)
    Call PropBag.WriteProperty("Enabled", cmddown.Enabled, True)
    Call PropBag.WriteProperty("BackColor", cmdup.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", cmddown.BackColor, &H8000000F)
End Sub

Private Sub UserControl_Resize()
    'Resize the buttons
    With UserControl
        'Set Widths
        .cmdup.Width = .ScaleWidth
        .cmddown.Width = .ScaleWidth
        'Set Button2 Top
        .cmddown.Top = (.ScaleHeight \ 2)
        'Set Heights
        .cmdup.Height = (.ScaleHeight \ 2)
        .cmddown.Height = (.ScaleHeight \ 2)
    End With
End Sub

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Long)
    m_Value = NewValue
    PropertyChanged "Value"
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal NewMax As Long)
    m_Max = NewMax
    PropertyChanged "Max"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = cmdup.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    cmdup.Enabled() = New_Enabled
    cmddown.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = cmdup.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    cmdup.BackColor() = New_BackColor
    cmddown.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

