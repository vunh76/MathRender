VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOver"
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
Private m_nFontSize As Long
Private m_OverType As OVERTYPE
Public Property Get OverStyle() As OVERTYPE
  OverStyle = m_OverType
End Property
Public Property Let OverStyle(ByVal vNewValue As OVERTYPE)
  m_OverType = vNewValue
End Property
Private Sub Class_Initialize()
  Set m_Content = Nothing
  Set m_Parent = Nothing
  m_OverType = OVER_NORMAL
End Sub

Private Sub Class_Terminate()
  Set m_Content = Nothing
  Set m_Parent = Nothing
  m_OverType = OVER_NORMAL
End Sub

Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
  clsBox_ClassName = "over"
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim o As New clsOver
  Set o.Content = m_Content
  o.OverStyle = m_OverType
  Set clsBox_Copy = o
End Property

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
  Dim nx As Long, ny As Long, d As Long
  d = modHelper.TextInternalLeading(hDC)
  If m_OverType = OVER_NORMAL Then
    m_Content.Draw x, y, hDC
    modHelper.DrawLine x, y, x + m_Width - 1, y, hDC
  ElseIf m_OverType = OVER_VECTOR Then
    m_Content.Draw x, y + 2 * d, hDC
    modHelper.DrawLine x, y + d, x + m_Width - 1, y + d, hDC
    modHelper.DrawLine x + m_Width - 1, y + d, x + m_Width - d, y, hDC
    modHelper.DrawLine x + m_Width - 1, y + d, x + m_Width - d, y + 2 * d, hDC
  ElseIf m_OverType = OVER_ANGLE Then
    m_Content.Draw x, y + 2 * d, hDC
    modHelper.DrawLine x + 1, y + 2 * d, x + m_Width \ 2, y, hDC
    modHelper.DrawLine x + m_Width \ 2, y, x + m_Width - d, y + 2 * d, hDC
  ElseIf m_OverType = OVER_ABS Then
    m_Content.Draw x + 2, y, hDC
    modHelper.DrawLine x, y, x, y + m_Height, hDC
    modHelper.DrawLine x + m_Width - 1, y, x + m_Width - 1, y + m_Height, hDC
  Else
    m_Content.Draw x, y + 2 * d, hDC
    Arc hDC, x, y, x + m_Width - 1, y + m_Height, x + m_Width - 1, y + 2 * d, x, y + 2 * d
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
  If Not m_Content Is Nothing Then
    m_Content.FontSize = m_nFontSize
  End If
End Property

Private Sub clsBox_Layout(hDC As Long)
  Dim k As Long
  m_Content.Layout hDC
  m_Width = m_Content.Width
  m_Height = m_Content.Height
  m_Ascent = m_Content.Ascent
  If m_OverType = OVER_VECTOR Or m_OverType = OVER_ANGLE Or m_OverType = OVER_CHORD Then
    k = modHelper.TextInternalLeading(hDC)
    m_Height = m_Height + 2 * k
    m_Ascent = m_Ascent + 2 * k
  ElseIf m_OverType = OVER_ABS Then
    m_Width = m_Width + 4
  End If
End Sub

Private Property Set clsBox_Parent(ByVal RHS As clsBox)
  Set m_Parent = RHS
End Property

Private Property Get clsBox_Parent() As clsBox
  Set clsBox_Parent = m_Parent
End Property

Public Property Get Content() As clsBox
  Set Content = m_Content
End Property

Public Property Set Content(ByVal vNewValue As clsBox)
  Set m_Content = vNewValue
  Set m_Content.Parent = Me
  m_Content.FontSize = m_nFontSize
End Property

Private Property Get clsBox_Width() As Long
  clsBox_Width = m_Width
End Property
