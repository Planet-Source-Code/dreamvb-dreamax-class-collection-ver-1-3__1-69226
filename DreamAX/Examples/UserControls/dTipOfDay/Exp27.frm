VERSION 5.00
Begin VB.Form FrmExp27 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip Of The Day Active X -Example"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1575
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show Tips"
      Height          =   375
      Left            =   255
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin Exp27.dTipOfDay dTipOfDay1 
      Left            =   3030
      Top             =   210
      _extentx        =   794
      _extenty        =   714
      dialogcaption   =   "Tips"
      tipimg          =   "Exp27.frx":0000
      outlineborder   =   8438015
      panelcolor      =   33023
      headerfont      =   "Exp27.frx":0124
      tipfont         =   "Exp27.frx":0154
   End
End
Attribute VB_Name = "FrmExp27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mShowAgian As Boolean

Private Sub Command2_Click()
    mShowAgian = dTipOfDay1.ShowValue
    If (mShowAgian) Then
        dTipOfDay1.ShowDialog
    End If
    
    MsgBox mShowAgian
End Sub

Private Sub Command3_Click()
    dTipOfDay1.ShowValue = True
End Sub

Private Sub cmdExit_Click()
    Unload FrmExp27
End Sub

Private Sub cmdShow_Click()
    mShowAgian = dTipOfDay1.ShowValue
    'You can use the top line and store the value of mShowAgian in a File or Regedit
    'ex If (Not mShowAgian) Then Exit Sub 'This will not show the dialog agian.
    
    'Show the Dialog
    Call dTipOfDay1.ShowDialog
    
    If (dTipOfDay1.ShowValue) Then
        MsgBox "You selected to show the dialog agian."
    Else
        MsgBox "You selected not to show the dialog agian."
    End If
End Sub

Private Sub Form_Load()
Dim MyTips As Collection
    'We add some tips useing a collection this makes things easyer
    Set MyTips = New Collection
    MyTips.Add "You can assign your own Colors, Fonts, Icon to the tips dialog"
    MyTips.Add "You can use a collection object to load your tips"
    MyTips.Add "That you can disable your tips by checking the check box"
    MyTips.Add "That you can set custom captions on the buttons"
    MyTips.Add "That i can't think of anything else here to write :)"
    'Assign the Tips
    dTipOfDay1.TipStrings = MyTips
    Set MyTips = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmExp27 = Nothing
End Sub
