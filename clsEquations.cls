VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_nFontSize As Long
Private m_Parent As clsBox
Private m_Equations() As clsBox
Private m_nNumberEquation As Long
Private m_nLineSpace As Long

Public Sub CreateEquation(ByVal n As Long)
  Dim i As Long
  For i = 1 To m_nNumberEquation
    Set m_Equations(i) = Nothing
  Next
  Erase m_Equations
  m_nNumberEquation = n
  ReDim m_Equations(1 To n)
End Sub
Public Property Get Equation(ByVal nIndex As Long) As clsBox
  If nIndex <= m_nNumberEquation Then
    Set Equation = m_Equations(nIndex)
  End If
End Property
Public Property Set Equation(ByVal nIndex As Long, EQ As clsBox)
  If nIndex <= m_nNumberEquation Then
    Set m_Equations(nIndex) = EQ
  End If
End Property
Private Property Get clsBox_Copy() As clsBox
  Dim EQ As New clsEquations
  Dim i As Long
  EQ.CreateEquation m_nNumberEquation
  For i = 1 To m_nNumberEquation
    Set EQ.Equation(i) = m_Equations(i)
  Next
  Set clsBox_Copy = EQ
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
  Dim i As Long
  Dim nx As Long, ny As Long
  Dim lf As LOGFONT, hNewFont As Long, hOldFont As Long
  Dim t1 As Long, t2 As Long
  Const CHAR1 = &H59
  Const CHAR2 = &H5D
  Const CHAR3 = &H5B
  ny = y + m_nLineSpace \ 2
  For i = 1 To m_nNumberEquation
    nx = x + m_nLineSpace
    m_Equations(i).Draw nx, ny, hDC
    ny = ny + m_Equations(i).Height + m_nLineSpace
  Next
  lf = modHelper.GetLogFont(modHelper.GetFont(hDC))
  lf.lfCharSet = SYMBOL_CHARSET
  lf.lfFaceName = "Lucida Bright Math Extension" & Chr(0)
  hNewFont = CreateFontIndirect(lf)
  hOldFont = modHelper.SetFont(hNewFont, hDC)
  
  modHelper.DrawText x, y, Chr(CHAR1), hDC
  
  t1 = modHelper.TextWidth(Chr(CHAR2), hDC)
  ny = y + m_Height \ 2 - modHelper.TextAscent(hDC) \ 2
  modHelper.DrawLine x + t1 \ 2, y + modHelper.TextAscent(hDC), x + t1 \ 2, ny, hDC
  
  nx = x
  modHelper.DrawText nx, ny, Chr(CHAR2), hDC
  
  ny = ny + modHelper.TextAscent(hDC)
  modHelper.DrawLine x + t1 \ 2, ny, x + t1 \ 2, y + m_Height - (modHelper.TextHeight(Chr(CHAR3), hDC) - modHelper.TextAscent(hDC) \ 2), hDC
  
  modHelper.DrawText x, y + m_Height - modHelper.TextHeight(Chr(CHAR3), hDC), Chr(CHAR3), hDC
  
  modHelper.SetFont hOldFont, hDC
  DeleteObject hNewFont
End Sub

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  Dim i As Long
  m_nFontSize = RHS
  For i = 1 To m_nNumberEquation
    If Not m_Equations(i) Is Nothing Then
      m_Equations(i).FontSize = m_nFontSize
    End If
  Next
End Property

Private Sub clsBox_Layout(hDC As Long)
  Dim i As Long
  m_Width = 0
  m_Height = 0
  m_Ascent = 0
  m_nLineSpace = modHelper.TextHeight("X", hDC)
  For i = 1 To m_nNumberEquation
    m_Equations(i).Layout hDC
    If m_Width < m_Equations(i).Width Then
      m_Width = m_Equations(i).Width
    End If
    m_Height = m_Height + m_Equations(i).Height + m_nLineSpace
  Next
  m_Width = m_Width + m_nLineSpace
  m_Ascent = m_Height / 2 + modHelper.TextDecent(hDC)
End Sub

Private Property Set clsBox_Parent(ByVal RHS As clsBox)
  Set m_Parent = RHS
End Property

Private Property Get clsBox_Parent() As clsBox
  clsBox_Parent = m_Parent
End Property

Private Property Get clsBox_Width() As Long
  clsBox_Width = m_Width
End Property
Private Property Get clsBox_FontSize() As Long
  clsBox_FontSize = m_nFontSize
End Property

Private Property Get clsBox_Height() As Long
  clsBox_Height = m_Height
End Property

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property
Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
  clsBox_ClassName = "equation"
End Property

