VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class implemention for SUM, PROD, INT, UNI function
Option Explicit
Implements clsBox
Private m_Parent As clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_SignWidth As Long
Private m_SignAscent As Long
Private m_SignDecent As Long
Private m_strSign As String
Private m_nFontSize As Long
Private m_Content As clsBox
Private m_Upper As clsBox
Private m_Lower As clsBox
Public Property Get Content() As clsBox
  Set Content = m_Content
End Property

Public Property Set Content(ByVal vNewValue As clsBox)
  Set m_Content = vNewValue
End Property

Private Sub Class_Initialize()
  Set m_Parent = Nothing
  Set m_Content = Nothing
  Set m_Upper = Nothing
  Set m_Lower = Nothing
End Sub

Private Sub Class_Terminate()
  Set m_Parent = Nothing
  Set m_Content = Nothing
  Set m_Upper = Nothing
  Set m_Lower = Nothing
End Sub
Private Property Get clsBox_Copy() As clsBox
  Dim s As New clsISPU
  Set s.Content = m_Content
  Set s.Lower = m_Lower
  Set s.Upper = m_Upper
  s.Sign = m_strSign
  Set clsBox_Copy = s
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hdc As Long)
  Dim X1 As Long, Y1 As Long
  Dim n As Long, h As Long
  Dim h1 As Long, h2 As Long, w As Long
  Dim hOldFont As Long, hNewFont As Long
  Dim lf As LOGFONT
  Dim cs As Long
  cs = modHelper.TextInternalLeading(hdc)
  
  If Not m_Upper Is Nothing Then
    X1 = x + (m_SignWidth - m_Upper.Width) / 2
    Y1 = y + m_Ascent - m_Upper.Height - m_SignAscent
    m_Upper.Draw X1, Y1, hdc
    h1 = m_Upper.Height
  Else
    h1 = 0
  End If
  
  hOldFont = modHelper.GetFont(hdc)
  lf = modHelper.GetLogFont(hOldFont)
  lf.lfCharSet = SYMBOL_CHARSET
  lf.lfFaceName = MATH_SYMBOL_FONT & Chr(0)
  lf.lfHeight = m_nFontSize * 2
  hNewFont = CreateFontIndirect(lf)
  modHelper.SetFont hNewFont, hdc
  w = modHelper.TextWidth(m_strSign, hdc)
  
  X1 = x + (m_SignWidth - w) / 2
  Y1 = y + m_Ascent - m_SignAscent
  modHelper.DrawText X1, Y1, m_strSign, hdc
  modHelper.SetFont hOldFont, hdc
  
  If Not m_Lower Is Nothing Then
    X1 = x + (m_SignWidth - m_Lower.Width) / 2
    Y1 = y + m_Ascent + m_SignDecent
    m_Lower.Draw X1, Y1, hdc
  End If
  
  X1 = x + m_SignWidth + cs
  Y1 = y + m_Ascent - m_Content.Ascent
  m_Content.Draw X1, Y1, hdc
  
  DeleteObject hNewFont
End Sub

Private Sub clsBox_Layout(hdc As Long)
  Dim h1 As Long, h2 As Long, w1 As Long, w2 As Long, h3 As Long
  Dim hOldFont As Long, hNewFont As Long
  Dim lf As LOGFONT
  Dim h As Long, w As Long, n As Long
  Dim ac1 As Long, ac2 As Long, dc1 As Long, dc2 As Long
  
  Dim cs As Long
  'cs = modHelper.TextDecent(hdc) / 2
  cs = modHelper.TextInternalLeading(hdc)
  m_Content.Layout hdc
  ac2 = m_Content.Ascent
  dc2 = m_Content.Descent
  If Not m_Upper Is Nothing Then
    m_Upper.Layout hdc
    h1 = m_Upper.Height
    w1 = m_Upper.Width
    ac1 = m_Upper.Height
  Else
    h1 = 0
    w1 = 0
    ac1 = 0
  End If
  
  If Not m_Lower Is Nothing Then
    m_Lower.Layout hdc
    h2 = m_Lower.Height
    w2 = m_Lower.Width
    dc1 = m_Lower.Height
  Else
    h2 = 0
    w2 = 0
    dc1 = 0
  End If
  
  Dim extra As Long
  extra = modHelper.GetFreeExtra(m_Content.FontSize, hdc)
  
  hOldFont = modHelper.GetFont(hdc)
  lf = modHelper.GetLogFont(hOldFont)
  lf.lfCharSet = SYMBOL_CHARSET
  lf.lfFaceName = MATH_SYMBOL_FONT & Chr(0)
  lf.lfHeight = m_nFontSize * 2
  hNewFont = CreateFontIndirect(lf)
  modHelper.SetFont hNewFont, hdc
  w = modHelper.TextWidth(m_strSign, hdc)
  h3 = modHelper.TextHeight(m_strSign, hdc)
  m_SignAscent = modHelper.TextHeight(m_strSign, hdc) / 2 + extra
  m_SignDecent = modHelper.TextHeight(m_strSign, hdc) - m_SignAscent
  ac1 = ac1 + m_SignAscent
  dc1 = dc1 + modHelper.TextDecent(hdc)
  modHelper.SetFont hOldFont, hdc
  If ac1 > ac2 Then
    m_Ascent = ac1
  Else
    m_Ascent = ac2
  End If
  If dc1 > dc2 Then
    m_Height = m_Ascent + dc1
  Else
    m_Height = m_Ascent + dc2
  End If
  
  If w1 > w2 Then
    m_Width = w1
  Else
    m_Width = w2
  End If
  If m_Width < w Then
    m_Width = w
  End If
  m_SignWidth = m_Width
  m_Width = m_Width + cs + m_Content.Width
  DeleteObject hNewFont
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
  If Asc(m_strSign) = SUM_SIGN Then
    clsBox_ClassName = "sum"
  ElseIf Asc(m_strSign) = INT_SIGN Then
    clsBox_ClassName = "integral"
  ElseIf Asc(m_strSign) = PROD_SIGN Then
    clsBox_ClassName = "product"
  ElseIf Asc(m_strSign) = UNI_SIGN Then
    clsBox_ClassName = "union"
  End If
End Property

Public Property Get Upper() As clsBox
  Set Upper = m_Upper
End Property

Public Property Set Upper(ByVal vNewValue As clsBox)
  Set m_Upper = vNewValue
  If Not m_Upper Is Nothing Then
    m_Upper.FontSize = m_nFontSize * 4 / 5
    Set m_Upper.Parent = Me
  End If
End Property

Public Property Get Lower() As clsBox
  Set Lower = m_Lower
End Property

Public Property Set Lower(ByVal vNewValue As clsBox)
  Set m_Lower = vNewValue
  If Not m_Lower Is Nothing Then
    m_Lower.FontSize = m_nFontSize * 4 / 5
    Set m_Lower.Parent = Me
  End If
End Property


Private Property Let clsBox_FontSize(ByVal RHS As Long)
  m_nFontSize = RHS
  If Not m_Content Is Nothing Then
    m_Content.FontSize = m_nFontSize
  End If
  If Not m_Upper Is Nothing Then
    m_Upper.FontSize = m_nFontSize * 4 / 5
  End If
  If Not m_Lower Is Nothing Then
    m_Lower.FontSize = m_nFontSize * 4 / 5
  End If
End Property

Public Property Get Sign() As String
  Sign = m_strSign
End Property

Public Property Let Sign(ByVal str As String)
  m_strSign = str
End Property
