VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dMouseHover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long

Private WithEvents obj_Button As CommandButton
Attribute obj_Button.VB_VarHelpID = -1
Private WithEvents obj_PicBox As PictureBox
Attribute obj_PicBox.VB_VarHelpID = -1

Event MouseEnter()
Event MouseLeave()

Private Sub xMouse(Obj As Object, X As Single, Y As Single)
    If (X < 0) Or (X > Obj.Width) _
    Or (Y < 0) Or (Y > Obj.Height) Then
        ReleaseCapture
        RaiseEvent MouseLeave
    ElseIf GetCapture() <> Obj.Hwnd Then
        RaiseEvent MouseEnter
        SetCapture Obj.Hwnd
    End If
End Sub

Private Sub obj_Button_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call xMouse(obj_Button, X, Y)
End Sub
'
Private Sub obj_PicBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call xMouse(obj_PicBox, X, Y)
End Sub

Public Sub AttachControl(dControl As Object)
    Select Case TypeName(dControl)
        Case "CommandButton"
            Set obj_Button = dControl
        Case "PictureBox"
            Set obj_PicBox = dControl
    End Select
End Sub
