VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_Parent As clsBox
Private m_Numerator As clsBox
Private m_Denominator As clsBox
Private m_nFontSize As Long

Private Sub Class_Initialize()
  Set m_Numerator = Nothing
  Set m_Denominator = Nothing
  Set m_Parent = Nothing
End Sub

Private Sub Class_Terminate()
  Set m_Numerator = Nothing
  Set m_Denominator = Nothing
  Set m_Parent = Nothing
End Sub

Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
   clsBox_ClassName = "fraction"
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim f As New clsFraction
  Set f.Numerator = m_Numerator
  Set f.Denominator = m_Denominator
  Set clsBox_Copy = f
End Property

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hdc As Long)
  Dim nx As Long
  Dim ny As Long
  Dim ls As Long
  'ls = modHelper.TextDecent(hdc)
  ls = modHelper.TextInternalLeading(hdc)
  nx = x + (m_Width - m_Numerator.Width) / 2
  ny = y
  m_Numerator.Draw nx, ny, hdc
  ny = ny + m_Numerator.Height + (ls / 3) * 2
  
  modHelper.DrawLine x, ny, x + m_Width, ny, hdc
  
  ny = ny + ls / 3
  nx = x + (m_Width - m_Denominator.Width) / 2
  m_Denominator.Draw nx, ny, hdc
End Sub

Public Property Get Numerator() As clsBox
  Set Numerator = m_Numerator
End Property

Public Property Set Numerator(ByVal vNewValue As clsBox)
  Dim b As clsBracket
  If vNewValue.ClassName = "bracket" Then
    Set b = vNewValue
    Set m_Numerator = b.Content
  Else
    Set m_Numerator = vNewValue
  End If
  Set m_Numerator.Parent = Me
  m_Numerator.FontSize = m_nFontSize
End Property

Public Property Get Denominator() As clsBox
  Set Denominator = m_Denominator
End Property

Public Property Set Denominator(ByVal vNewValue As clsBox)
  Dim b As clsBracket
  If vNewValue.ClassName = "bracket" Then
    Set b = vNewValue
    Set m_Denominator = b.Content
  Else
    Set m_Denominator = vNewValue
  End If
  Set m_Denominator.Parent = Me
  m_Denominator.FontSize = m_nFontSize
End Property

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  m_nFontSize = RHS
  If Not m_Numerator Is Nothing Then
    m_Numerator.FontSize = m_nFontSize
  End If
  If Not m_Denominator Is Nothing Then
    m_Denominator.FontSize = m_nFontSize
  End If
End Property

Private Property Get clsBox_FontSize() As Long
  clsBox_FontSize = m_nFontSize
End Property

Private Property Get clsBox_Height() As Long
  clsBox_Height = m_Height
End Property

Private Sub clsBox_Layout(hdc As Long)
  Dim ls As Long
  m_Numerator.Layout hdc
  m_Denominator.Layout hdc
  'ls = modHelper.TextDecent(hdc)
  ls = modHelper.TextInternalLeading(hdc)
  If m_Numerator.Width > m_Denominator.Width Then
    m_Width = m_Numerator.Width
  Else
    m_Width = m_Denominator.Width
  End If
  m_Height = m_Numerator.Height + m_Denominator.Height + ls
     
  Dim hNewFont As Long
  Dim hOldFont As Long
  Dim lf As LOGFONT
  hOldFont = modHelper.GetFont(hdc)
  lf = modHelper.GetLogFont(hOldFont)
  lf.lfCharSet = SYMBOL_CHARSET
  lf.lfFaceName = MATH_SYMBOL_FONT & Chr(0)
  lf.lfHeight = m_nFontSize
  hNewFont = CreateFontIndirect(lf)
  modHelper.SetFont hNewFont, hdc
  m_Ascent = m_Numerator.Height + (ls / 3) * 2 + modHelper.TextAscent(hdc) - modHelper.TextHeight(Chr(PLUS_SIGN), hdc) \ 2
  modHelper.SetFont hOldFont, hdc
  DeleteObject hNewFont
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

