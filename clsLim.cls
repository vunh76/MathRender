VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements clsBox
Private m_Content As clsBox
Private m_Under As clsBox
Private m_Parent As clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_nFontSize As Long
Public Property Get Content() As clsBox
  Set Content = m_Content
End Property

Public Property Set Content(vNewValue As clsBox)
    If Not vNewValue Is Nothing Then
        Set m_Content = vNewValue
    End If
End Property
Public Property Get Under() As clsBox
    Set Under = m_Under
End Property
Public Property Set Under(vNewValue As clsBox)
    If Not vNewValue Is Nothing Then
        Set m_Under = vNewValue
        m_Under.FontSize = m_nFontSize * 4 / 5
    End If
End Property
Private Sub Class_Initialize()
  Set m_Content = Nothing
  Set m_Under = Nothing
End Sub

Private Sub Class_Terminate()
  Set m_Content = Nothing
  Set m_Under = Nothing
End Sub

Private Property Get clsBox_Ascent() As Long
    clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
    clsBox_ClassName = "lim"
End Property

Private Property Get clsBox_Copy() As clsBox
    Dim l As New clsLim
    Set l.Content = m_Content
    Set l.Under = m_Under
    Set clsBox_Copy = l
End Property

Private Property Get clsBox_Descent() As Long
    clsBox_Descent = m_Height - m_Ascent
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
    Dim nLimH As Long, nLimW As Long
    Dim nX As Long, nY As Long
    Dim nUW As Long
    nlimA = modHelper.TextAscent(hDC)
    nLimW = modHelper.TextWidth("lim", hDC)
    If Not m_Under Is Nothing Then
        nUW = m_Under.Width
    Else
        nUW = 0
    End If
    If nUW > nLimW Then
        nX = x + (nUW - nLimW) \ 2
    Else
        nX = x
    End If
    nY = y + m_Ascent - nlimA
    modHelper.DrawText nX, nY, "lim", hDC
    If nUW > nLimW Then
        nX = x + nUW
    Else
        nX = x + nLimW
    End If
    nY = y + m_Ascent - m_Content.Ascent
    m_Content.Draw nX, nY, hDC
    If Not m_Under Is Nothing Then
        If nUW > nLimW Then
            nX = x
        Else
            nX = x + (nlimx - nUW) \ 2
        End If
        nY = y + m_Ascent
        m_Under.Draw nX, nY, hDC
    End If
End Sub

Private Property Get clsBox_FontSize() As Long
    clsBox_FontSize = m_nFontSize
End Property

Private Property Let clsBox_FontSize(ByVal RHS As Long)
    m_nFontSize = RHS
    If Not m_Under Is Nothing Then
        m_Under.FontSize = m_nFontSize * 4 / 5
    End If
    If Not m_Content Is Nothing Then
        m_Content.FontSize = m_nFontSize
    End If
End Property

Private Property Get clsBox_Height() As Long
    clsBox_Height = m_Height
End Property

Private Sub clsBox_Layout(hDC As Long)
    Dim nLimW As Long
    Dim nCW As Long, nCH As Long, nCA As Long
    Dim nUW As Long, nUH As Long, nUA As Long
    m_Content.Layout hDC
    nLimW = modHelper.TextWidth("lim", hDC)
    nCW = m_Content.Width
    nCH = m_Content.Height
    nCA = m_Content.Ascent
    If Not m_Under Is Nothing Then
        m_Under.Layout hDC
        nUW = m_Under.Width
        nUH = m_Under.Height
        nUA = m_Under.Ascent
    Else
        nUW = 0
        nUA = 0
        nUH = 0
    End If
    If nLimW > nUW Then
        m_Width = nLimW + nCW
    Else
        m_Width = nUW + nCW
    End If
    m_Ascent = nCA
    If nCH - nCA > nUH Then
        m_Height = nCH
    Else
        m_Height = nCA + nUH
    End If
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
