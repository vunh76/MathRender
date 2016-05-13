VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRoot"
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
Private m_nFontSize As Long
Private m_Content As clsBox
Private m_Base As clsBox
Private m_RadicalWidth As Long

Private Sub Class_Initialize()
  Set m_Parent = Nothing
  Set m_Content = Nothing
End Sub

Private Sub Class_Terminate()
  Set m_Parent = Nothing
  Set m_Content = Nothing
End Sub
Public Property Get Base() As clsBox
    Set Base = m_Base
End Property

Public Property Set Base(vNewValue As clsBox)
    Set m_Base = vNewValue
    If Not m_Base Is Nothing Then
        m_Base.FontSize = 4 / 5 * m_nFontSize
    End If
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim r As New clsRoot
  Set r.Content = m_Content
  Set r.Base = m_Base
  Set clsBox_Copy = r
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
  Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, n As Long
  Dim m As Long
  m = modHelper.TextInternalLeading(hDC)
  n = m_RadicalWidth
  If Not m_Base Is Nothing Then
    X1 = x
    If m_Base.Height > m_Content.Height \ 2 Then
        Y1 = y
    Else
        Y1 = y + m_Content.Height \ 2 - m_Base.Height
    End If
    'm_Base.Draw X1, Y1
  End If
  X1 = x + n * 2 \ 3
  Y1 = y + m_Height
  X2 = x + n
  Y2 = y + m
  modHelper.DrawLine X1, Y1, X2, Y2, hDC
  X1 = X2
  Y1 = Y2
  X2 = X1 + m_Content.Width
  modHelper.DrawLine X1, Y1, X2, Y2, hDC
  
  X1 = x + n * 2 \ 3
  Y1 = y + m_Height
  X2 = x + (X1 - x) \ 3
  Y2 = y + m + (m_Height - m) \ 2
  modHelper.DrawLine X1, Y1, X2, Y2, hDC
  X1 = x
  Y1 = Y2
  modHelper.DrawLine X1, Y1, X2, Y2, hDC
  x = x + n
  y = y + modHelper.TextInternalLeading(hDC)
  m_Content.Draw x, y, hDC
End Sub

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  m_nFontSize = RHS
  If Not m_Content Is Nothing Then
    m_Content.FontSize = m_nFontSize
  End If
  If Not m_Base Is Nothing Then
    m_Base.FontSize = 4 / 5 * m_nFontSize
  End If
End Property

Private Sub clsBox_Layout(hDC As Long)
  Dim l As Long, nBW As Long, nBH As Long
  m_Content.Layout hDC
  If Not m_Base Is Nothing Then
    m_Base.Layout hDC
    nBW = m_Base.Width
    nBH = m_Base.Height
  Else
    nBW = 0
    nBH = 0
  End If
  l = modHelper.TextInternalLeading(hDC)
  If nBH <= m_Content.Height \ 2 Then
     m_Height = m_Content.Height + l
  Else
     m_Height = m_Content.Height + l + nBH - m_Content.Height \ 2
  End If
  If m_Content.Width > m_Content.Height \ 2 Then
    m_RadicalWidth = m_Content.Height \ 2
  Else
    m_RadicalWidth = m_Content.Width
  End If
  If nBW \ 2 < m_RadicalWidth Then
     m_Width = m_Content.Width + m_RadicalWidth + nBW \ 2
  Else
     m_Width = m_Content.Width + m_RadicalWidth + nBW - m_RadicalWidth
  End If
  If nBH <= m_Content.Height \ 2 Then
    m_Ascent = m_Content.Ascent + l
  Else
    m_Ascent = m_Content.Ascent + l + nBH - m_Content.Height \ 2
  End If
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
  clsBox_ClassName = "root"
End Property

Public Property Get Content() As clsBox
  Set Content = m_Content
End Property

Public Property Set Content(ByVal vNewValue As clsBox)
  Set m_Content = vNewValue
  m_Content.FontSize = m_nFontSize
  Set m_Content.Parent = Me
End Property
