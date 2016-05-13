VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_Base As clsBox
Private m_Sub As clsBox
Private m_Sup As clsBox
Private m_Parent As clsBox
Private m_nFontSize As clsBox
Public Property Get Base() As clsBox
  Set Base = m_Base
End Property
Public Property Set Base(box As clsBox)
  Set m_Base = box
End Property
Public Property Get Subscript() As clsBox
  Set Subscript = m_Sub
End Property
Public Property Get Supscript() As clsBox
  Set Supscript = m_Sup
End Property
Public Property Set Supscript(box As clsBox)
  Set m_Sup = box
End Property
Public Property Set Subscript(box As clsBox)
  Set m_Sub = box
End Property

Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
  clsBox_ClassName = "subsup"
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim c As clsSubSup
  Set c.Base = m_Base
  Set c.Subscript = m_Sub
  Set c.Supscript = m_Sup
  Set clsBox_Copy = c
End Property

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
End Sub

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  Dim i As Long
  m_nFontSize = RHS
  If Not m_Base Is Nothing Then
    m_Base.FontSize = m_nFontSize
  End If
  If Not m_Sub Is Nothing Then
    m_Sub.FontSize = 3 / 4 * m_nFontSize
  End If
  If Not m_Sup Is Nothing Then
    m_Sup.FontSize = 3 / 4 * m_nFontSize
  End If
End Property

Private Property Get clsBox_FontSize() As Long
  clsBox_FontSize = m_nFontSize
End Property

Private Property Get clsBox_Height() As Long
  clsBox_Height = m_Height
End Property

Private Sub clsBox_Layout(hDC As Long)
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



