VERSION 5.00
Begin VB.Form FrmExp22 
   Caption         =   "Properties List - ActiveX Example"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5010
   StartUpPosition =   1  'CenterOwner
   Begin Exp22.dPropList dPropList1 
      Height          =   2550
      Left            =   195
      TabIndex        =   2
      Top             =   105
      Width           =   3450
      _extentx        =   6085
      _extenty        =   4498
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00000000&
      Caption         =   "E&xit"
      Height          =   350
      Left            =   3720
      TabIndex        =   1
      Top             =   555
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetSel 
      Caption         =   "Get Selected"
      Height          =   350
      Left            =   3720
      TabIndex        =   0
      Top             =   135
      Width           =   1215
   End
End
Attribute VB_Name = "FrmExp22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_SelKey As String
Private m_SelType As CrtType

Private Sub SetupList()
Dim cLst1 As New Collection
Dim cLst2 As New Collection
    
    cLst1.Add "True"
    cLst1.Add "False"
    
    cLst2.Add "Soild"
    cLst2.Add "Dotted"
    cLst2.Add "Dashed"
    cLst2.Add "Dash-Dot"
    cLst2.Add "Transparent"
    'Clear properties box
    dPropList1.Clear
    'Add some property items
    dPropList1.AddProp "Name", "A", tTextbox
    dPropList1.AddProp "BackColor", "B", tTextbox
    dPropList1.AddProp "ForeColor", "C", tTextbox
    dPropList1.AddProp "Caption", "D", tTextbox
    dPropList1.AddProp "Enabled", "E", tComboBox
    dPropList1.AddProp "BorderStyle", "F", tComboBox
    dPropList1.AddProp "Height", "G", tTextbox
    dPropList1.AddProp "Width", "h", tTextbox
    dPropList1.AddProp "Style", "I", tComboBox
    dPropList1.AddProp "BorderWidth", "J", tTextbox
    'Set default property data
    
    dPropList1.SetPropItemValue "A", "ShpButton", tTextbox
    dPropList1.SetPropItemValue "B", "#000000", tTextbox
    dPropList1.SetPropItemValue "C", "#FFF000", tTextbox
    dPropList1.SetPropItemValue "D", "E&xit", tTextbox
    dPropList1.SetPropItemValue "E", cLst1, tComboBox
    dPropList1.SetPropItemValue "F", cLst2, tComboBox
    dPropList1.SetPropItemValue "G", 350, tTextbox
    dPropList1.SetPropItemValue "H", 1155, tTextbox
    dPropList1.SetPropItemValue "J", 1, tTextbox
    Set cLst1 = Nothing
    cLst1.Add "XP"
    cLst1.Add "Default"
    cLst1.Add "Office97"
    '
    dPropList1.SetPropItemValue "I", cLst1, tComboBox
    
    Set cLst1 = Nothing
    Set cLst2 = Nothing
End Sub

Private Sub cmdExit_Click()
    dPropList1.Clear
    Unload FrmExp22
End Sub

Private Sub cmdGetSel_Click()
    MsgBox dPropList1.GetPropItemValue(m_SelKey, m_SelType), vbInformation
End Sub

Private Sub dPropList1_PropClick(sKey As String, ItemProp As CrtType)
    m_SelKey = sKey
    m_SelType = ItemProp
End Sub

Private Sub Form_Load()
    Call SetupList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dPropList1.Clear
    Set FrmExp22 = Nothing
End Sub
