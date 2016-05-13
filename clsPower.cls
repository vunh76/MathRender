VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsBox
Private m_Left As clsBox
Private m_Right As clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_Parent As clsBox
Private m_nFontSize As Long

Private Sub Class_Initialize()
  Set m_Left = Nothing
  Set m_Right = Nothing
  Set m_Parent = Nothing
End Sub

Private Sub Class_Terminate()
  Set m_Left = Nothing
  Set m_Right = Nothing
  Set m_Parent = Nothing
End Sub

Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
  clsBox_ClassName = "power"
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim p As New clsPower
  Set p.Left = m_Left
  Set p.Right = m_Right
  Set clsBox_Copy = p
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
  Dim l As Long
  Dim nx As Long, ny As Long
  l = modHelper.TextAscent(hDC)
  nx = x
  ny = y + m_Right.Ascent - l / 3
  m_Left.Draw nx, ny, hDC
  nx = x + m_Left.Width + 1
  m_Right.Draw nx, y, hDC
End Sub

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  m_nFontSize = RHS
  If Not Right Is Nothing Then
    Right.FontSize = m_nFontSize * 4 / 5
  End If
  If Not Left Is Nothing Then
    Left.FontSize = m_nFontSize
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
  m_Left.Layout hDC
  m_Right.Layout hDC
  m_Width = m_Left.Width + m_Right.Width + 1
  l = modHelper.TextAscent(hDC)
  m_Ascent = m_Left.Ascent + m_Right.Ascent - l / 3
  m_Height = m_Left.Height + m_Right.Ascent - l / 3
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

Public Property Get Left() As clsBox
  Set Left = m_Left
End Property

Public Property Set Left(ByVal vNewValue As clsBox)
  Set m_Left = vNewValue
  Set m_Left.Parent = Me
  m_Left.FontSize = m_nFontSize
End Property

Public Property Get Right() As clsBox
  Set Right = m_Right
End Property

Public Property Set Right(ByVal vNewValue As clsBox)
  Dim b As clsBracket
  If vNewValue.ClassName = "bracket" Then
    Set b = vNewValue
    Set m_Right = b.Content
  Else
    Set m_Right = vNewValue
  End If
  Set m_Right.Parent = Me
  m_Right.FontSize = m_nFontSize * 4 / 5
End Property

