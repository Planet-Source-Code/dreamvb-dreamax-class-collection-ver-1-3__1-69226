VERSION 5.00
Begin VB.UserControl dFlash 
   CanGetFocus     =   0   'False
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   330
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dFlash.ctx":0000
   ScaleHeight     =   330
   ScaleWidth      =   330
   ToolboxBitmap   =   "dFlash.ctx":018A
   Begin VB.Timer Tmr1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "dFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private m_FrmObj As Form

Public Sub FormObject(FrmObj As Form)
    Set m_FrmObj = FrmObj
End Sub

Public Sub FlashWindowwStart()
    Tmr1.Enabled = True
End Sub

Public Sub FlashWindowStop()
    Tmr1.Enabled = False
End Sub

Private Sub Tmr1_Timer()
    FlashWindow m_FrmObj.hwnd, 1
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 330, 330
End Sub

Private Sub UserControl_Terminate()
    If (Tmr1.Enabled) Then Tmr1.Enabled = False
End Sub

Public Property Get Busy() As Boolean
    Busy = (Tmr1.Enabled)
End Property

Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
    Interval = Tmr1.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    Tmr1.Interval() = New_Interval
    PropertyChanged "Interval"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Tmr1.Interval = PropBag.ReadProperty("Interval", 300)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Interval", Tmr1.Interval, 300)
End Sub

