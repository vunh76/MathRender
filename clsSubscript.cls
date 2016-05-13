VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubscript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsBox
Private m_Parent As clsBox
Private m_Base As clsBox
Private m_Subscript As clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_nFontSize As Long

Private Sub Class_Initialize()
  Set m_Parent = Nothing
  Set m_Base = Nothing
  Set m_Subscript = Nothing
End Sub

Private Sub Class_Terminate()
  Set m_Parent = Nothing
  Set m_Base = Nothing
  Set m_Subscript = Nothing
End Sub

Public Property Get Base() As clsBox
  Set Base = m_Base
End Property

Public Property Set Base(ByVal vNewValue As clsBox)
  Set m_Base = vNewValue
  m_Base.FontSize = m_nFontSize
  Set m_Base.Parent = Me
End Property

Public Property Get Subscript() As clsBox
  Set Subscript = m_Subscript
End Property

Public Property Set Subscript(ByVal vNewValue As clsBox)
  Dim b As clsBracket
  If vNewValue.ClassName = "bracket" Then
    Set b = vNewValue
    Set m_Subscript = b.Content
  Else
    Set m_Subscript = vNewValue
  End If
  m_Subscript.FontSize = m_nFontSize * 4 / 5
  Set m_Subscript.Parent = Me
End Property

Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
  clsBox_ClassName = "subscript"
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim s As New clsSubscript
  Set s.Base = m_Base
  Set s.Subscript = m_Subscript
  Set clsBox_Copy = s
End Property

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
  m_Base.Draw x, y, hDC
  x = x + m_Base.Width + 1
  y = y + (m_Height - m_Subscript.Height)
  m_Subscript.Draw x, y, hDC
End Sub

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  m_nFontSize = RHS
  If Not m_Base Is Nothing Then
    m_Base.FontSize = m_nFontSize
  End If
  If Not m_Subscript Is Nothing Then
    m_Subscript.FontSize = m_nFontSize * 4 / 5
  End If
End Property

Private Property Get clsBox_FontSize() As Long
  clsBox_FontSize = m_nFontSize
End Property

Private Property Get clsBox_Height() As Long
  clsBox_Height = m_Height
End Property

Private Sub clsBox_Layout(hDC As Long)
  Dim l As Long
  m_Base.Layout hDC
  m_Subscript.Layout hDC
  l = modHelper.TextAscent(hDC)
  m_Width = m_Base.Width + m_Subscript.Width + 1
  m_Height = m_Base.Ascent + m_Subscript.Height - l / 3
  m_Ascent = m_Base.Ascent
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
