VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsBox
Private Enum ENTITYTYPE
  ENT_NORMAL
  ENT_MATH
End Enum
Private Type ENTITIES
  str As String
  nEntityType As ENTITYTYPE
  nHeight As Long
  nAscent As Long
  nWidth As Long
  nFontIndex As Long
End Type
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_InternalLeading As Long
Private m_Parent As clsBox
Private m_strContent As String
Private m_nFontSize As Long
Private m_Entities() As ENTITIES
Private m_nEntity As Long

Private Sub Class_Initialize()
  m_strContent = ""
  Set m_Parent = Nothing
End Sub

Private Sub Class_Terminate()
  m_strContent = ""
  Set m_Parent = Nothing
End Sub

Private Property Get clsBox_Ascent() As Long
  clsBox_Ascent = m_Ascent
End Property

Private Property Get clsBox_ClassName() As String
  clsBox_ClassName = "text"
End Property

Private Property Get clsBox_Copy() As clsBox
  Dim t As New clsText
  t.Text = m_strContent
  Set clsBox_Copy = t
End Property

Private Property Get clsBox_Descent() As Long
  clsBox_Descent = m_Height - m_Ascent
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
  Dim nx As Long, ny As Long, dw As Long
  Dim hOldFont As Long
  Dim hCharFont As Long
  Dim bChanged As Boolean
  Dim hOpFont As Long
  Dim lf As LOGFONT
  Dim st As String
  Dim i As Long
  Dim mo As MATH_OP
  hOldFont = modHelper.GetFont(hDC)
  lf = modHelper.GetLogFont(hOldFont)
   
  lf.lfItalic = False
  lf.lfCharSet = SYMBOL_CHARSET
  lf.lfFaceName = MATH_SYMBOL_FONT & Chr(0)
  hOpFont = CreateFontIndirect(lf)
  lf.lfFaceName = MATH_CHAR_FONT & Chr(0)
  hCharFont = CreateFontIndirect(lf)
  
  nx = x
  For i = 1 To m_nEntity
    st = ""
    bChanged = False
    If m_Entities(i).nEntityType = ENT_NORMAL Then
      mo = modHelper.FindOPCode(m_Entities(i).str)
      If mo.nCode > 0 Then
        st = Chr(mo.nCode)
      End If
      If mo.nFontIndex = 2 Then
        hOldFont = SelectObject(hDC, hOpFont)
        bChanged = True
      ElseIf mo.nFontIndex = 1 Then
        hOldFont = SelectObject(hDC, hCharFont)
        bChanged = True
      End If
    ElseIf m_Entities(i).nEntityType = ENT_MATH Then
      If m_Entities(i).nFontIndex = 1 Then
        hOldFont = SelectObject(hDC, hCharFont)
        bChanged = True
      ElseIf m_Entities(i).nFontIndex = 2 Then
        hOldFont = SelectObject(hDC, hOpFont)
        bChanged = True
      End If
    End If
    If st = "" Then
      st = m_Entities(i).str
    End If
    ny = y + m_Ascent - m_Entities(i).nAscent
    modHelper.DrawText nx, ny, st, hDC
    If m_Entities(i).str = "<>" Then
      modHelper.DrawLine nx + m_Entities(i).nWidth / 3, ny + m_Entities(i).nHeight * 3 / 4, nx + m_Entities(i).nWidth * 2 / 3, ny + m_Entities(i).nHeight / 4, hDC
    End If
    nx = nx + m_Entities(i).nWidth
    If bChanged Then
      modHelper.SetFont hOldFont, hDC
    End If
  Next
  
  DeleteObject hOpFont
  DeleteObject hCharFont
End Sub

Public Property Get Text() As String
  Text = m_strContent
End Property

Public Property Let Text(ByVal vNewValue As String)
  m_strContent = vNewValue
  ParseString
End Property

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  m_nFontSize = RHS
End Property

Private Property Get clsBox_FontSize() As Long
  clsBox_FontSize = m_nFontSize
End Property

Private Property Get clsBox_Height() As Long
  clsBox_Height = m_Height
End Property

Private Sub clsBox_Layout(hDC As Long)
  Dim hOpFont As Long
  Dim hCharFont As Long
  Dim hTextFont As Long
  Dim hOldFont As Long
  Dim lf As LOGFONT
  Dim i As Long
  Dim maxHeight As Long
  Dim maxAscent As Long
  Dim bChanged As Boolean
  Dim st As String
  Dim mo As MATH_OP
  hOldFont = GetCurrentObject(hDC, OBJ_FONT)
  GetObject hOldFont, Len(lf), lf
 
  lf.lfItalic = False
  lf.lfCharSet = SYMBOL_CHARSET
  lf.lfFaceName = MATH_SYMBOL_FONT & Chr(0)
  hOpFont = CreateFontIndirect(lf)
  lf.lfFaceName = MATH_CHAR_FONT & Chr(0)
  hCharFont = CreateFontIndirect(lf)
  m_Width = 0
  m_Height = 0
  m_Ascent = 0
  For i = 1 To m_nEntity
    bChanged = False
    st = ""
    If m_Entities(i).nEntityType = ENT_NORMAL Then
      mo = modHelper.FindOPCode(m_Entities(i).str)
      If mo.nCode > 0 Then
        st = Chr(mo.nCode)
      End If
      If mo.nFontIndex = 2 Then
        hOldFont = SelectObject(hDC, hOpFont)
        bChanged = True
      ElseIf mo.nFontIndex = 1 Then
        hOldFont = SelectObject(hDC, hCharFont)
        bChanged = True
      End If
    ElseIf m_Entities(i).nEntityType = ENT_MATH Then
      bChanged = True
      If m_Entities(i).nFontIndex = 1 Then
        hOldFont = SelectObject(hDC, hCharFont)
      ElseIf m_Entities(i).nFontIndex = 2 Then
        hOldFont = SelectObject(hDC, hOpFont)
      End If
    End If
    If st = "" Then
      st = m_Entities(i).str
    End If
    m_Entities(i).nWidth = modHelper.TextWidth(st, hDC)
    m_Width = m_Width + m_Entities(i).nWidth
    m_Entities(i).nHeight = modHelper.TextHeight(m_Entities(i).str, hDC)
    If maxHeight < m_Entities(i).nHeight Then
      maxHeight = m_Entities(i).nHeight
    End If
    m_Entities(i).nAscent = modHelper.TextAscent(hDC)
    If maxAscent < m_Entities(i).nAscent Then
      maxAscent = m_Entities(i).nAscent
    End If
    If bChanged Then
      SelectObject hDC, hOldFont
    End If
  Next
  m_Height = maxHeight
  m_Ascent = maxAscent
  DeleteObject hOpFont
  DeleteObject hCharFont
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

Public Function isOperator(ByVal str As String) As Boolean
  isOperator = InStr("+-*=", str) > 0 Or str = "<=" Or str = ">=" Or str = "+-" Or str = "-+" Or str = "<>"
End Function

Private Sub ParseString()
  Dim i As Long
  Dim st As String, ss As String, ch As String
  Dim bFound As Boolean, k As Long
  Dim ms As MATH_SYMBOL
  m_nEntity = 0
  Erase m_Entities
  i = 1
  st = ""
  ss = ""
  Do While i <= Len(m_strContent)
    ch = Mid(m_strContent, i, 1)
    If ch = "&" Then
      bFound = True
      If Len(ss) > 0 Then
        st = st & ss
      End If
      ss = "&"
    ElseIf ch = ";" Then
      If bFound And Len(ss) > 1 Then
        ss = ss & ch
        ms = modHelper.FindCharCode(ss)
        k = ms.nCode
        'k = modHelper.FindCharCode(ss)
        If k > 0 Then
          If Len(st) > 0 Then
            m_nEntity = m_nEntity + 1
            ReDim Preserve m_Entities(1 To m_nEntity)
            m_Entities(m_nEntity).nEntityType = ENT_NORMAL
            m_Entities(m_nEntity).str = st
            st = ""
          End If
          m_nEntity = m_nEntity + 1
          ReDim Preserve m_Entities(1 To m_nEntity)
          m_Entities(m_nEntity).nEntityType = ENT_MATH
          m_Entities(m_nEntity).str = Chr(k)
          m_Entities(m_nEntity).nFontIndex = ms.nFontIndex
          ss = ""
          bFound = False
        Else
          st = st & ss
          ss = ""
        End If
      Else
        If Len(ss) > 0 Then
          st = st & ss
        End If
        st = st & ch
      End If
    ElseIf bFound Then
      ss = ss & ch
    Else
      st = st & ch
    End If
    i = i + 1
    If i > Len(m_strContent) Then Exit Do
  Loop
  st = st & ss
  If st <> "" Then
    m_nEntity = m_nEntity + 1
    ReDim Preserve m_Entities(1 To m_nEntity)
    m_Entities(m_nEntity).nEntityType = ENT_NORMAL
    m_Entities(m_nEntity).str = st
  End If
End Sub
