VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConcat"
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
Private m_nElement As Long
Private m_Parts() As clsBox
Private m_nFontSize As Long

Private Sub Class_Initialize()
  m_nElement = 0
  Set m_Parent = Nothing
End Sub

Private Sub Class_Terminate()
  Dim i As Long
  For i = 1 To m_nElement
    Set m_Parts(i) = Nothing
  Next
  Erase m_Parts
End Sub

Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
  clsBox_ClassName = "concat"
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim c As New clsConcat, i As Long
  c.NumberElement = m_nElement
  For i = 1 To m_nElement
    Set c.Part(i) = m_Parts(i)
  Next
  Set clsBox_Copy = c
End Property

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hdc As Long)
  Dim i As Long
  Dim nX As Long, nY As Long
  Dim cs As Long
  'cs = modHelper.TextDecent(hdc) / 2
  cs = 2
  nX = x
  nY = y
  For i = 1 To m_nElement
    nY = y + m_Ascent - m_Parts(i).Ascent
    m_Parts(i).Draw nX, nY, hdc
    nX = nX + m_Parts(i).Width + cs
  Next
End Sub

Public Property Get NumberElement() As Long
  NumberElement = m_nElement
End Property

Public Property Let NumberElement(ByVal vNewValue As Long)
  m_nElement = vNewValue
  ReDim m_Parts(1 To m_nElement)
End Property

Public Property Get Part(ByVal nIndex As Long) As clsBox
  If nIndex > 0 And nIndex <= m_nElement Then
    Set Part = m_Parts(nIndex)
  Else
    Set Part = Null
  End If
End Property

Public Property Set Part(ByVal nIndex As Long, ByVal vNewValue As clsBox)
  If nIndex > 0 And nIndex <= m_nElement Then
    Set m_Parts(nIndex) = vNewValue
    Set m_Parts(nIndex).Parent = Me
    m_Parts(nIndex).FontSize = m_nFontSize
  End If
End Property

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  Dim i As Long
  m_nFontSize = RHS
  For i = 1 To m_nElement
    If Not m_Parts(i) Is Nothing Then
      m_Parts(i).FontSize = m_nFontSize
    End If
  Next
End Property

Private Property Get clsBox_FontSize() As Long
  clsBox_FontSize = m_nFontSize
End Property

Private Property Get clsBox_Height() As Long
  clsBox_Height = m_Height
End Property

Private Sub clsBox_Layout(hdc As Long)
  Dim i As Long
  Dim nMax As Long
  Dim cs As Long
  'cs = modHelper.TextDecent(hdc) / 2
  cs = 2
  For i = 1 To m_nElement
    m_Parts(i).Layout hdc
  Next
  
  nMax = 0
  For i = 1 To m_nElement
    If m_Parts(i).Height > nMax Then nMax = m_Parts(i).Height
  Next
  m_Height = nMax
  
  nMax = 0
  For i = 1 To m_nElement
    If i > 1 Then nMax = nMax + cs
    nMax = nMax + m_Parts(i).Width
  Next
  m_Width = nMax
  
  nMax = 0
  For i = 1 To m_nElement
    If m_Parts(i).Ascent > nMax Then nMax = m_Parts(i).Ascent
  Next
  m_Ascent = nMax
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


