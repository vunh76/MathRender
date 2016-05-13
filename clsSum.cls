VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsBox
Private m_Parent As clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_SMW As Long
Private m_SMH As Long
Private m_SMA As Long
Private m_LeftW As Long
Private m_strSign As String
Private m_nFontSize As Long
Private m_Content As clsBox
Private m_Upper As clsBox
Private m_Lower As clsBox

Private Sub Class_Initialize()
  Set m_Parent = Nothing
  Set m_Content = Nothing
  Set m_Upper = Nothing
  Set m_Lower = Nothing
  m_strSign = Chr(SUM_SIGN)
End Sub

Private Sub Class_Terminate()
  Set m_Parent = Nothing
  Set m_Content = Nothing
  Set m_Upper = Nothing
  Set m_Lower = Nothing
End Sub

Private Property Get clsBox_Copy() As clsBox
  Dim s As New clsSum
  Set s.Content = m_Content
  Set s.Lower = m_Lower
  Set s.Upper = m_Upper
  Set clsBox_Copy = s
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hdc As Long)
  Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, h1 As Long, n As Long
  
  n = modHelper.TextInternalLeading(hdc)
  If Not m_Upper Is Nothing Then
    X1 = x + (m_LeftW - m_Upper.Width) / 2
    Y1 = y + m_Ascent - m_SMA
    m_Upper.Draw X1, Y1, hdc
    h1 = m_Upper.Height + n
  Else
    h1 = 0
  End If
  
  X1 = x + (m_LeftW - m_SMW) / 2
  Y1 = y + m_Ascent - m_SMA + h1
  X2 = X1 + m_SMW
  Y2 = Y1
  modHelper.DrawLine X1, Y1, X2, Y2, hdc
  X2 = X1 + m_SMW * 3 / 4
  Y2 = Y1 + m_SMH / 2
  modHelper.DrawLine X1, Y1, X2, Y2, hdc
  Y1 = Y1 + m_SMH
  modHelper.DrawLine X1, Y1, X2, Y2, hdc
  X2 = X1 + m_SMW
  Y2 = Y1
  modHelper.DrawLine X1, Y1, X2, Y2, hdc
  
  If Not m_Lower Is Nothing Then
    X1 = x + (m_LeftW - m_Lower.Width) / 2
    Y1 = y + m_Ascent - m_SMA + h1 + m_SMH + n
    m_Lower.Draw X1, Y1, hdc
  End If
  
  X1 = x + m_LeftW + 1
  Y1 = y + m_Ascent - m_Content.Ascent
  m_Content.Draw X1, Y1, hdc
End Sub

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

Private Sub clsBox_Layout(hdc As Long)
  Dim h1 As Long, h2 As Long, w1 As Long, w2 As Long
  Dim smD As Long, cD As Long, n As Long
  Dim hOldFont As Long, hNewFont As Long
  Dim lf As LOGFONT
  n = modHelper.TextInternalLeading(hdc)
  m_Content.Layout hdc
  If Not m_Upper Is Nothing Then
    m_Upper.Layout hdc
    h1 = m_Upper.Height + n
    w1 = m_Upper.Width
  Else
    h1 = 0
    w1 = 0
  End If
  If Not m_Lower Is Nothing Then
    m_Lower.Layout hdc
    h2 = m_Lower.Height + n
    w2 = m_Lower.Width
  Else
    h2 = 0
    w2 = 0
  End If
  'Calculation size of sigma sign
  m_SMW = modHelper.TextWidth("X", hdc)
  m_SMH = modHelper.TextAscent(hdc) * 2 / 3
  hOldFont = modHelper.GetFont(hdc)
  lf = modHelper.GetLogFont(hOldFont)
  lf.lfHeight = lf.lfHeight * 4 / 5
  hNewFont = CreateFontIndirect(lf)
  modHelper.SetFont hNewFont, hdc
  'm_SMW = m_SMW + modHelper.TextWidth("x", pic1, pic2)
  m_SMH = m_SMH + modHelper.TextHeight("x", hdc)
  m_SMW = m_SMH * 4 / 5
  modHelper.SetFont hOldFont, hdc
  
  'End sigma sign
  'Calculation m_Width
  If w1 > w2 Then
    m_Width = w1
  Else
    m_Width = w2
  End If
  If m_Width < m_SMW Then
    m_Width = m_SMW
  End If
  m_LeftW = m_Width
  m_Width = m_Width + m_Content.Width + 1
  'End m_Width
  'Calculation m_Height, m_Ascent
  
  m_SMA = h1 + m_SMH \ 2 + modHelper.OperatorWidth(m_nFontSize, hdc) \ 2 'Sigma ascent
  smD = m_SMH + h1 + h2 - m_SMA 'Sigma descent
  If m_SMA > m_Content.Ascent Then
    m_Ascent = m_SMA
  Else
    m_Ascent = m_Content.Ascent
  End If
  cD = m_Content.Height - m_Content.Ascent
  If cD > smD Then
    m_Height = m_Ascent + cD
  Else
    m_Height = m_Ascent + smD
  End If
  DeleteObject hNewFont
  'DeleteObject hOldFont
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
  clsBox_ClassName = "sum"
End Property

Public Property Get Content() As clsBox
  Set Content = m_Content
End Property

Public Property Set Content(ByVal vNewValue As clsBox)
  Set m_Content = vNewValue
  m_Content.FontSize = m_nFontSize
  Set m_Content.Parent = Me
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

Public Property Get Sign() As String
  Sign = m_strSign
End Property

Public Property Let Sign(ByVal str As String)
  m_strSign = str
End Property
