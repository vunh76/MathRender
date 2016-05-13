VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBracket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsBox
Private m_Content As clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_Parent As clsBox
Private m_OpenBracketWidth As Long
Private m_CloseBracketWidth As Long
Private m_nFontSize As Long
Private m_bRequired As Boolean
Public Property Get Content() As clsBox
  Set Content = m_Content
End Property

Public Property Set Content(ByVal vNewValue As clsBox)
  Set m_Content = vNewValue
  Set m_Content.Parent = Me
  m_Content.FontSize = m_nFontSize
End Property

Private Sub Class_Initialize()
  m_bRequired = True
End Sub

Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
  clsBox_ClassName = "bracket"
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim b As New clsBracket
  Set b.Content = m_Content
  Set clsBox_Copy = b
End Property

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hdc As Long)
  Const PI = 3.1415926
  Dim nx As Long, ny As Long, ny1 As Long
  Dim lf As LOGFONT
  Dim hNewFont As Long
  Dim hOldFont As Long
  Dim ha As Long, hd As Long
  Const CHAR1 = &H59
  Const CHAR2 = &H5B
  Const CHAR3 = &H5A
  Const CHAR4 = &H5C
  If m_Content.ClassName = "text" Or m_Content.Height < 1.5 * modHelper.TextHeight("(", hdc) Then 'Brackets is the same line as text formular
    modHelper.DrawText x, y, "(", hdc
    m_Content.Draw x + m_OpenBracketWidth, y, hdc
    modHelper.DrawText x + m_Content.Width + m_OpenBracketWidth, y, ")", hdc
  Else
       
    m_Content.Draw x + m_OpenBracketWidth, y, hdc
    
    hOldFont = GetCurrentObject(hdc, OBJ_FONT)
    GetObject hOldFont, Len(lf), lf
    lf.lfCharSet = SYMBOL_CHARSET
    lf.lfFaceName = MATH_EXTENSION_FONT & Chr(0)
    hNewFont = CreateFontIndirect(lf)
    hOldFont = SelectObject(hdc, hNewFont)
    ha = modHelper.TextAscent(hdc)
    
    modHelper.DrawText x, y - ha \ 2, Chr(CHAR1), hdc
    modHelper.DrawText x + m_Width - m_OpenBracketWidth, y - ha \ 2, Chr(CHAR3), hdc
    nx = x + m_OpenBracketWidth \ 2
    ny = y + ha \ 2
    ny1 = y + m_Height - ha \ 2
    modHelper.DrawLine nx, ny, nx, ny1, hdc
    
    modHelper.DrawText x, y + m_Height - ha, Chr(CHAR2), hdc
    modHelper.DrawText x + m_Width - m_OpenBracketWidth, y + m_Height - ha, Chr(CHAR4), hdc
    nx = x + m_Width - m_OpenBracketWidth \ 2 - 1
    ny = y + ha \ 2
    ny1 = y + m_Height - ha \ 2
    modHelper.DrawLine nx, ny, nx, ny1, hdc
    
    SelectObject hdc, hOldFont
    DeleteObject hNewFont
  End If
End Sub

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  m_nFontSize = RHS
  If Not m_Content Is Nothing Then
    m_Content.FontSize = m_nFontSize
  End If
End Property

Private Property Get clsBox_FontSize() As Long
  clsBox_FontSize = m_nFontSize
End Property

Private Property Get clsBox_Height() As Long
  clsBox_Height = m_Height
End Property

Private Sub clsBox_Layout(hdc As Long)
  Dim lf As LOGFONT
  Dim hNewFont As Long
  Dim hOldFont As Long
  m_Content.Layout hdc
  If m_Content.ClassName = "text" Or m_Content.Height < 1.5 * modHelper.TextHeight("(", hdc) Then 'Brackets is the same line as text formular
    m_OpenBracketWidth = modHelper.TextWidth("(", hdc)
  Else
    hOldFont = GetCurrentObject(hdc, OBJ_FONT)
    GetObject hOldFont, Len(lf), lf
    lf.lfCharSet = SYMBOL_CHARSET
    lf.lfFaceName = MATH_EXTENSION_FONT & Chr(0)
    hNewFont = CreateFontIndirect(lf)
    hOldFont = SelectObject(hdc, hNewFont)
    m_OpenBracketWidth = modHelper.TextWidth(Chr(&H59), hdc)
    SelectObject hdc, hOldFont
    DeleteObject hNewFont
  End If
  m_CloseBracketWidth = m_OpenBracketWidth
  m_Height = m_Content.Height
  m_Width = m_Content.Width + m_OpenBracketWidth + m_CloseBracketWidth
  m_Ascent = m_Content.Ascent
End Sub

Private Property Set clsBox_Parent(ByVal RHS As clsBox)
  Set m_Parent = RHS
End Property

Private Property Get clsBox_Parent() As clsBox
  Set clsBox_Parent = m_Parent
End Property

Private Property Get clsBox_Width() As Long
  clsBox_Width = m_Width
End Property

Public Property Get AutoRemove() As Boolean
  AutoRemove = m_bRequired
End Property

Public Property Let AutoRemove(ByVal vNewValue As Boolean)
  m_bRequired = vNewValue
End Property
