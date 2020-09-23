VERSION 5.00
Begin VB.UserControl dTipOfDay 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dTipOfDay.ctx":0000
   ScaleHeight     =   405
   ScaleWidth      =   450
   ToolboxBitmap   =   "dTipOfDay.ctx":0173
End
Attribute VB_Name = "dTipOfDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Tips As Collection
Private m_DlgCaption As String
Private m_Img As IPictureDisp
Private m_BColor As OLE_COLOR
Private m_PanelColor As OLE_COLOR
Private m_TipArea As OLE_COLOR

Private m_LnBStyle As BorderStyleConstants
Private m_NextBut As String
Private m_PrevBut As String
Private m_CloseBut As String
Private m_ShowValue As Boolean

Private m_HeaderCaption As String
Private m_HeadLnColor As OLE_COLOR
Private m_HeadFont As Font
Private m_HeadColor As OLE_COLOR

Private m_TipFont As Font
Private m_TipColor As OLE_COLOR
Private m_CheckBoxCap As String

Private Const m_defcaption As String = "Tip of the day"
Private Const m_defHeadCap As String = "Did you know"

Public Sub ShowDialog()
    With frmTips
        .m_BorderColor = m_BColor
        .pSidePanel.BackColor = m_PanelColor
        .pTipHolder.BackColor = m_TipArea
        .lnTop.BorderColor = m_HeadLnColor
        .lnTop.BorderStyle = m_LnBStyle
        .ImgIco = m_Img
        .Caption = m_DlgCaption
        .cmdButton(0).Caption = m_NextBut
        .cmdButton(1).Caption = m_PrevBut
        .cmdButton(2).Caption = m_CloseBut
        Set .lblHeader.Font = m_HeadFont
        .lblHeader.ForeColor = m_HeadColor
        .lblHeader.Caption = m_HeaderCaption
        Set .lblDesc.Font = m_TipFont
        .lblDesc.ForeColor = m_TipColor
        .chkShow.Caption = m_CheckBoxCap
        'Send the tip collection along
        Set .TipCollection = m_Tips
        .chkShow.Value = Abs(m_ShowValue)
        .Show vbModal, UserControl.Parent
        m_ShowValue = .ChkVal
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_HeaderCaption = m_defHeadCap
    m_DlgCaption = m_defcaption
    m_CheckBoxCap = "Don't show tips at start up."
    m_NextBut = "&Next"
    m_PrevBut = "&Previous"
    m_CloseBut = "&Close"
    m_BColor = vb3DShadow
    m_PanelColor = vb3DShadow
    m_TipArea = vbWhite
    m_HeadLnColor = vb3DShadow
    m_LnBStyle = vbBSSolid
    m_HeadColor = vbBlack
    m_TipColor = vbBlack
    m_ShowValue = True
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 450, 405
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_HeaderCaption = PropBag.ReadProperty("HeaderCaption", m_defHeadCap)
    m_DlgCaption = PropBag.ReadProperty("DialogCaption", m_defcaption)
    Set m_Img = PropBag.ReadProperty("TipImg", Nothing)
    m_BColor = PropBag.ReadProperty("OutLineBorder", vb3DShadow)
    m_PanelColor = PropBag.ReadProperty("PanelColor", vb3DShadow)
    m_TipArea = PropBag.ReadProperty("TipAreaColor", vbWhite)
    m_HeadLnColor = PropBag.ReadProperty("HeaderLineColor", vb3DShadow)
    m_LnBStyle = PropBag.ReadProperty("HeaderLineStyle", vbBSSolid)
    m_NextBut = PropBag.ReadProperty("NextButtonCaption", "&Next")
    m_PrevBut = PropBag.ReadProperty("PreviousButtonCaption", "&Previous")
    m_CloseBut = PropBag.ReadProperty("CloseButtonCaption", "&Close")
    m_ShowValue = PropBag.ReadProperty("ShowValue", True)
    Set m_HeadFont = PropBag.ReadProperty("HeaderFont", Ambient.Font)
    m_HeadColor = PropBag.ReadProperty("HeaderColor", vbBlack)
    Set m_TipFont = PropBag.ReadProperty("TipFont", Ambient.Font)
    m_TipColor = PropBag.ReadProperty("TipColor", vbBlack)
    m_CheckBoxCap = PropBag.ReadProperty("CheckBoxCaption", "Don't show tips at start up.")
End Sub

Private Sub UserControl_Terminate()
    Set m_Tips = Nothing
    Set frmTips = Nothing
    Set m_Img = Nothing
    Set m_HeadFont = Nothing
    Set m_TipFont = Nothing
    '
    m_DlgCaption = ""
    m_NextBut = ""
    m_PrevBut = ""
    m_CloseBut = ""
    m_HeaderCaption = ""
    m_CheckBoxCap = ""
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("HeaderCaption", m_HeaderCaption, m_defHeadCap)
    Call PropBag.WriteProperty("DialogCaption", m_DlgCaption, m_defcaption)
    Call PropBag.WriteProperty("TipImg", m_Img, Nothing)
    Call PropBag.WriteProperty("OutLineBorder", m_BColor, vb3DShadow)
    Call PropBag.WriteProperty("PanelColor", m_PanelColor, vb3DShadow)
    Call PropBag.WriteProperty("TipAreaColor", m_TipArea, vbWhite)
    Call PropBag.WriteProperty("HeaderLineColor", m_HeadLnColor, vb3DShadow)
    Call PropBag.WriteProperty("HeaderLineStyle", m_LnBStyle, vbBSSolid)
    Call PropBag.WriteProperty("NextButtonCaption", m_NextBut, "&Next")
    Call PropBag.WriteProperty("PreviousButtonCaption", m_PrevBut, "&Previous")
    Call PropBag.WriteProperty("CloseButtonCaption", m_CloseBut, "&Close")
    Call PropBag.WriteProperty("ShowValue", m_ShowValue, True)
    Call PropBag.WriteProperty("HeaderFont", m_HeadFont, Ambient.Font)
    Call PropBag.WriteProperty("HeaderColor", m_HeadColor, vbBlack)
    Call PropBag.WriteProperty("TipFont", m_TipFont, Ambient.Font)
    Call PropBag.WriteProperty("TipColor", m_TipColor, vbBlack)
    Call PropBag.WriteProperty("CheckBoxCaption", m_CheckBoxCap, "Don't show tips at start up.")
End Sub

Public Property Let TipStrings(ByVal vNewStrCol As Collection)
    Set m_Tips = vNewStrCol
End Property

Public Property Get HeaderCaption() As String
    HeaderCaption = m_HeaderCaption
End Property

Public Property Let HeaderCaption(ByVal vNewHCaption As String)
    m_HeaderCaption = vNewHCaption
    PropertyChanged "HeaderCaption"
End Property

Public Property Get DialogCaption() As String
    DialogCaption = m_DlgCaption
End Property

Public Property Let DialogCaption(ByVal vNewCaption As String)
    m_DlgCaption = vNewCaption
    PropertyChanged "DialogCaption"
End Property

Public Property Get TipImg() As Picture
    Set TipImg = m_Img
End Property

Public Property Set TipImg(ByVal New_TipImg As Picture)
    Set m_Img = New_TipImg
    PropertyChanged "TipImg"
End Property

Public Property Get OutLineBorder() As OLE_COLOR
    OutLineBorder = m_BColor
End Property

Public Property Let OutLineBorder(ByVal vNewColor As OLE_COLOR)
    m_BColor = vNewColor
    PropertyChanged "OutLineBorder"
End Property

Public Property Get PanelColor() As OLE_COLOR
    PanelColor = m_PanelColor
End Property

Public Property Let PanelColor(ByVal vNewPColor As OLE_COLOR)
    m_PanelColor = vNewPColor
    PropertyChanged "PanelColor"
End Property

Public Property Get TipAreaColor() As OLE_COLOR
   TipAreaColor = m_TipArea
End Property

Public Property Let TipAreaColor(ByVal vNewValue As OLE_COLOR)
    m_TipArea = vNewValue
    PropertyChanged "TipAreaColor"
End Property

Public Property Get HeaderLineColor() As OLE_COLOR
    HeaderLineColor = m_HeadLnColor
End Property

Public Property Let HeaderLineColor(ByVal vNewLnColor As OLE_COLOR)
    m_HeadLnColor = vNewLnColor
    PropertyChanged "HeaderLineColor"
End Property

Public Property Get HeaderLineStyle() As BorderStyleConstants
    HeaderLineStyle = m_LnBStyle
End Property

Public Property Let HeaderLineStyle(ByVal vNewBorderStyle As BorderStyleConstants)
    m_LnBStyle = vNewBorderStyle
    PropertyChanged "HeaderLineStyle"
End Property

Public Property Get NextButtonCaption() As String
   NextButtonCaption = m_NextBut
End Property

Public Property Let NextButtonCaption(ByVal vNewValue As String)
    m_NextBut = vNewValue
    PropertyChanged "NextButtonCaption"
End Property

Public Property Get PreviousButtonCaption() As String
    PreviousButtonCaption = m_PrevBut
End Property

Public Property Let PreviousButtonCaption(ByVal vNewValue As String)
    m_PrevBut = vNewValue
    PropertyChanged "PreviousButtonCaption"
End Property

Public Property Get CloseButtonCaption() As String
    CloseButtonCaption = m_CloseBut
End Property

Public Property Let CloseButtonCaption(ByVal vNewValue As String)
    m_CloseBut = vNewValue
    PropertyChanged "CloseButtonCaption"
End Property

Public Property Get ShowValue() As Boolean
   ShowValue = m_ShowValue
End Property

Public Property Let ShowValue(ByVal vNewValue As Boolean)
    m_ShowValue = vNewValue
    PropertyChanged "ShowValue"
End Property

Public Property Get HeaderFont() As Font
Attribute HeaderFont.VB_Description = "Returns a Font object."
    Set HeaderFont = m_HeadFont
End Property

Public Property Set HeaderFont(ByVal New_HeaderFont As Font)
    Set m_HeadFont = New_HeaderFont
    PropertyChanged "HeaderFont"
End Property

Public Property Get HeaderColor() As OLE_COLOR
    HeaderColor = m_HeadColor
End Property

Public Property Let HeaderColor(ByVal vNewValue As OLE_COLOR)
    m_HeadColor = vNewValue
    PropertyChanged "HeaderColor"
End Property

Public Property Get TipFont() As Font
Attribute TipFont.VB_Description = "Returns a Font object."
    Set TipFont = m_TipFont
End Property

Public Property Set TipFont(ByVal New_TipFont As Font)
    Set m_TipFont = New_TipFont
    PropertyChanged "TipFont"
End Property

Public Property Get TipColor() As OLE_COLOR
    TipColor = m_TipColor
End Property

Public Property Let TipColor(ByVal New_TipColor As OLE_COLOR)
    m_TipColor = New_TipColor
    PropertyChanged "TipColor"
End Property

Public Property Get CheckBoxCaption() As String
    CheckBoxCaption = m_CheckBoxCap
End Property

Public Property Let CheckBoxCaption(ByVal New_Caption As String)
    m_CheckBoxCap = New_Caption
    PropertyChanged "CheckBoxCaption"
End Property
