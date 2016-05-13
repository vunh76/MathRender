Attribute VB_Name = "modRender"
'Module modRender
'Implements string parser to find math and text elements
Option Explicit

Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal x As Long, ByVal y As Long, _
                                             ByVal nWidth As Long, ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Enum ELEMENT_TYPE
  E_SPACE
  E_TEXT
  E_MATH
  E_NEWLINE
End Enum
Private Type ELEMENT_INFO
  strText As String
  nType As ELEMENT_TYPE
  lngLeft As Long
  lngTop As Long
  f As clsBox
End Type
Private m_nElement As Long
Private m_Elements() As ELEMENT_INFO
Private nLineHeight As Long
Private nMarginTop As Long
Private nMarginLeft As Long
Private nLineSpace As Long
Private strOpenTag As String    'Ignore
Private strCloseTag As String   'Ignore
Private nBottom As Long

Private Function isAlphaDigit(ByVal ch As String)
  Const ALPHASET = "0123456789abcdefghijklmnopqrstuvwxyz"
  ch = LCase(ch)
  isAlphaDigit = InStr(ALPHASET, ch) > 0
End Function
Private Sub Layout(picViewPort As PictureBox, picCanvas As PictureBox, Optional ByVal bRedraw As Boolean = False)
  Dim i As Long, j As Long, nStart As Long, nEnd As Long
  Dim x As Long, y As Long
  Dim arr As Variant
  Dim st As String
  Dim nMaxAscent As Long
  Dim nMaxHeight As Long
  Dim tm As TEXTMETRIC
  Dim nWidth As Long
  Dim nElementWidth As Long
  Dim bHasNewLine As Boolean
  
  If m_nElement = 0 Then
    picViewPort.Cls
    picCanvas.Cls
    vscScroll.Enabled = False
    Exit Sub
  End If
  picCanvas.Cls
  
  nMarginLeft = 10
  nMarginTop = 10
  nLineHeight = picCanvas.TextHeight("X")
  GetTextMetrics picCanvas.hdc, tm
  nLineSpace = 2
  
  If Not bRedraw Then
    offset = vscScroll.Value * nLineHeight
  End If
  
  x = nMarginLeft
  y = nMarginTop
  i = 1
  If m_nElement > 0 And (Not bRedraw) Then 'Bypass all invisible elements. Above edge of canvas
    Do While m_Elements(i).lngBottom < offset
      i = i + 1
      If i > m_nElement Then Exit Do
    Loop
    If i <= m_nElement Then
      y = m_Elements(i).lngTop - offset
      x = m_Elements(i).lngLeft
    Else
      Exit Sub
    End If
  End If
 
  Do While i <= m_nElement
    'Skip invisible elements, below edge of canvas
    If (Not bRedraw) And y > picCanvas.ScaleHeight Then Exit Do
    nWidth = nMarginLeft
    nMaxAscent = tm.tmAscent
    nMaxHeight = tm.tmHeight
    nStart = i
    nEnd = 0
    Do While nWidth < picCanvas.ScaleWidth
      If m_Elements(i).nType = E_SPACE Or m_Elements(i).nType = E_TEXT Then
        nElementWidth = picCanvas.TextWidth(m_Elements(i).strText)
      ElseIf m_Elements(i).nType = E_MATH Then
        m_Elements(i).f.Layout picCanvas
        nElementWidth = m_Elements(i).f.Width
      Else
        i = i + 1
        nEnd = i
        Exit Do
      End If
      If (nWidth + nElementWidth) < picCanvas.ScaleWidth Then
        nWidth = nWidth + nElementWidth
        If m_Elements(i).nType = E_MATH Then
          If m_Elements(i).f.Ascent > nMaxAscent Then
            nMaxAscent = m_Elements(i).f.Ascent
          End If
          If m_Elements(i).f.Height > nMaxHeight Then
            nMaxHeight = m_Elements(i).f.Height
          End If
        End If
      Else
        nEnd = i
        Exit Do
      End If
      i = i + 1
      If i > m_nElement Then
        nEnd = i
        Exit Do
      End If
    Loop
    nEnd = nEnd - 1
    If nEnd >= nStart Then
      bHasNewLine = False
      If m_Elements(nStart).nType = E_SPACE Then
        j = nStart + 1 'Skip space if it is beginning of line
      Else
        j = nStart
      End If
      Do While j <= nEnd
        If bRedraw Then
          m_Elements(i).lngTop = y
          m_Elements(i).lngLeft = x
        End If
        If m_Elements(j).nType = E_TEXT Or m_Elements(j).nType = E_SPACE Then
          picCanvas.CurrentX = x
          picCanvas.CurrentY = y + nMaxAscent - tm.tmAscent
          picCanvas.Print m_Elements(j).strText
          x = x + picCanvas.TextWidth(m_Elements(j).strText)
        ElseIf m_Elements(j).nType = E_MATH Then
          m_Elements(j).f.Draw x, y + nMaxAscent - m_Elements(j).f.Ascent, picCanvas
          x = x + m_Elements(j).f.Width
        Else 'New line
          x = nMarginLeft
          y = y + nMaxHeight + nLineSpace
          bHasNewLine = True
        End If
        j = j + 1
      Loop
      If Not bHasNewLine Then
        x = nMarginLeft
        y = y + nMaxHeight + nLineSpace
      End If
    End If
  Loop
End Sub
Public Function Render(ByVal strText As String, picCanvas As PictureBox)
  Dim i As Long, j As Long, ch As String, st As String, ok As Boolean
  Dim box As clsBox
  m_nElement = 0
  Erase m_Elements
  i = 1
  Do While i <= Len(strText)
    ch = Mid(strText, i, 1)
    st = ""
    Do While ch = " " Or ch = vbTab
      st = st & ch
      i = i + 1
      ch = Mid(strText, i, 1)
    Loop
    If Len(st) > 0 Then
      If m_nElement = 0 Then
        IncreaseArray
      ElseIf m_Elements(m_nElement).nType <> E_SPACE Then
        IncreaseArray
      End If
      m_Elements(m_nElement).nType = E_SPACE
      m_Elements(m_nElement).strText = st 'Space only
    End If
    
    If i > Len(strText) Then Exit Do
    If ch = vbCr Then
      IncreaseArray
      m_Elements(m_nElement).nType = E_NEWLINE
      i = i + 1
      If i <= Len(strText) Then 'Remove vbLf character
        If Mid(strText, i, 1) = vbLf Then i = i + 1
      End If
    Else
      st = ""
      Do
        If ch <> "{" Then
          st = st & ch
          i = i + 1
          If i > Len(strText) Then Exit Do
          ch = Mid(strText, i, 1)
        Else
          Exit Do
        End If
      Loop Until ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf
      'Automatic detect math formular
      ok = False
      If Right(st, 1) = "." Or Right(st, 1) = "," Or Right(st, 1) = ":" Or Right(st, 1) = ";" Then
        If Len(st) > 1 Then 'Trim last character if it is a deliminate
          st = Left(st, Len(st) - 1)
          i = i - 1
          ok = True
        End If
      End If
      If (Left(st, 1) = "[" And Right(st, 1) = "]") Or _
          InStr(st, "^") > 0 Or InStr(st, "/") > 0 Or _
          InStr(st, "+") > 0 Or InStr(st, "-") > 0 Or _
          InStr(st, "*") > 0 Or InStr(st, "_") > 0 Then
        On Error Resume Next
        Set box = Nothing
        Set box = ParseExpression(st) 'Try to check if it is a correct formular
        Err.Clear
        If Not box Is Nothing Then
          IncreaseArray
          m_Elements(m_nElement).nType = E_MATH
          m_Elements(m_nElement).strText = st
          Set m_Elements(m_nElement).f = box.Copy
          st = ""
        ElseIf ok Then
          st = st & Mid(strText, i, 1)
          i = i + 1
          ok = False
        End If
      End If
      If st <> "" Then 'Not math formular
        If ok Then
          st = st & Mid(strText, i, 1)
          i = i + 1
        End If
        If st <> "" Then
          If m_nElement = 0 Then
            IncreaseArray
          ElseIf m_Elements(m_nElement).nType <> E_TEXT Then
            IncreaseArray
          End If
          m_Elements(m_nElement).strText = m_Elements(m_nElement).strText & st
        End If
      End If
      If ch = "{" Then
        j = i + 1
        Do While j <= Len(strText)
          ch = Mid(strText, j, 1)
          If ch = "}" Or ch = "{" Then Exit Do
          j = j + 1
        Loop
        st = Mid(strText, i + 1, j - i - 1)
        If st <> "" Then
            If ch = "}" And st <> "" Then 'Math markup found
              On Error Resume Next
              Set box = Nothing
              Set box = ParseExpression(st)
              Err.Clear
              If Not box Is Nothing Then
                IncreaseArray
                m_Elements(m_nElement).nType = E_MATH
                m_Elements(m_nElement).strText = st
                Set m_Elements(m_nElement).f = box.Copy
                i = j + 1
                st = ""
              End If
            End If
            If st <> "" Then
              If m_nElement = 0 Then
                IncreaseArray
              ElseIf m_Elements(m_nElement).nType <> E_TEXT Then
                IncreaseArray
              End If
              m_Elements(m_nElement).strText = m_Elements(m_nElement).strText & Mid(strText, i, 1)
              i = i + 1
            End If
        Else
          If m_nElement = 0 Then
            IncreaseArray
          ElseIf m_Elements(m_nElement).nType <> E_TEXT Then
            IncreaseArray
          End If
          m_Elements(m_nElement).strText = m_Elements(m_nElement).strText & Mid(strText, i, 1)
          i = i + 1
        End If
      End If
      If ch = " " Or ch = vbTab Then
        IncreaseArray
        m_Elements(m_nElement).nType = E_SPACE
      End If
    End If
  Loop
  Layout picCanvas
'  Dim x As Long, y As Long
'  x = 10
'  y = 10
'  picCanvas.CurrentX = x
'  For i = 1 To m_nElement
'    If m_Elements(i).nType = E_TEXT Then
'      picCanvas.CurrentY = y
'      picCanvas.CurrentX = x
'      picCanvas.Print m_Elements(i).strText
'      y = y + picCanvas.TextHeight("X")
'    ElseIf m_Elements(i).nType = E_MATH Then
'      m_Elements(i).f.Layout picCanvas
'      m_Elements(i).f.Draw x, y, picCanvas
'      y = y + m_Elements(i).f.Height
'    Else
'      picCanvas.CurrentY = y
'      picCanvas.CurrentX = x
'      picCanvas.Print "______________"
'      y = y + picCanvas.TextHeight("X")
'    End If
'  Next
End Function

Private Sub IncreaseArray()
    m_nElement = m_nElement + 1
    ReDim Preserve m_Elements(1 To m_nElement)
    m_Elements(m_nElement).nType = E_TEXT
End Sub


