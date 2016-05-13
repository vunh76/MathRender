VERSION 5.00
Begin VB.UserControl ctlMathDraw 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ScaleHeight     =   3600
   ScaleWidth      =   5925
   ToolboxBitmap   =   "ctlMathDraw.ctx":0000
   Begin VB.PictureBox picViewPort 
      Height          =   1335
      Left            =   3240
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.VScrollBar vscScroll 
      Enabled         =   0   'False
      Height          =   3255
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "ctlMathDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************
'ctlMathDraw: Implementation a render engine
'Render engine was get from code of HTMLLabel. Copyright © 2001 Woodbury Associates.
'****************************************************************
Option Explicit

Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
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
  lngBottom As Long
  f As clsBox
End Type
Private m_CurrentBox As Long
Private m_nElement As Long
Private m_Elements() As ELEMENT_INFO
Private nLineHeight As Long
Private nMarginTop As Long
Private nMarginLeft As Long
Private nLineSpace As Long
Private strOpenTag As String    'Ignore
Private strCloseTag As String   'Ignore
Private nBottom As Long
Private strText As String
Private m_nFontSize As Long
Private m_bAllowEdit As Boolean
Public Event SaveCompleted()
Public Event ContentChanged()
Public Property Get AllowEdit() As Boolean
    AllowEdit = m_bAllowEdit
End Property
Public Property Let AllowEdit(ByVal vNewValue As Boolean)
    m_bAllowEdit = vNewValue
End Property
'******************************************************************
'Process event Double_Click on picViewport
'Just select a formula that bellows pointer and shows a dialog to edit
'******************************************************************
Private Sub picViewPort_DblClick()
  Dim offset As Long
  If Not m_bAllowEdit Then Exit Sub
  If m_CurrentBox > 0 Then
    frmEditor.strFormular = m_Elements(m_CurrentBox).strText
    frmEditor.Show vbModal
    If frmEditor.bState Then
       If m_Elements(m_CurrentBox).strText <> frmEditor.strFormular Then
         m_Elements(m_CurrentBox).strText = frmEditor.strFormular
         Set m_Elements(m_CurrentBox).f = modHelper.ParseExpression(frmEditor.strFormular)
         offset = vscScroll.Value
         Render True
         RaiseEvent ContentChanged
         If offset > 0 Then
            On Error GoTo ErrHdl
            vscScroll.Value = offset
         End If
       End If
    End If
  End If
  Exit Sub
ErrHdl:
  Exit Sub
End Sub
'******************************************************************
'Turn on mouse hook to get wheeling ability
'******************************************************************
Private Sub picViewPort_GotFocus()
  modWheelHook.WheelHook picViewPort
End Sub
'*****************************************************************
'Turn off mouse hook
'*****************************************************************
Private Sub picViewPort_LostFocus()
  modWheelHook.WheelUnHook
End Sub
'******************************************************************
'Process mouse down event.
'Select a formula under pointer. Underline it by a blue line
'******************************************************************
Private Sub picViewPort_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim offset As Long
    Dim i As Long, nX As Long, nY As Long
    If Not m_bAllowEdit Then Exit Sub
    If Button <> 1 Then Exit Sub
    If m_nElement = 0 Then Exit Sub
    offset = vscScroll.Value * nLineHeight
    For i = 1 To m_nElement
      If m_Elements(i).nType = E_MATH And (Not m_Elements(i).f Is Nothing) Then
        If m_Elements(i).lngBottom >= offset Then
            If m_Elements(i).lngTop > offset + picViewPort.ScaleHeight Then Exit Sub
            If m_Elements(i).lngLeft <= x And (m_Elements(i).lngLeft + m_Elements(i).f.Width) >= x Then
                If (m_Elements(i).lngTop - offset) <= y And (m_Elements(i).lngBottom - offset) >= y Then
                  If i = m_CurrentBox Then Exit Sub
                  HilightBox m_CurrentBox
                  HilightBox i
                  m_CurrentBox = i
                  Exit Sub
                End If
            End If
        End If
      End If
    Next
End Sub
'*****************************************************************************
'Underline a formula
'Input:
'       nBox: The formula that will be underlined
'*****************************************************************************
Private Sub HilightBox(ByVal nBox As Long)
    Dim offset As Long
    Dim oldDrawMode As Long
    Dim oldDrawStyle As Long
    Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
    offset = vscScroll.Value * nLineHeight
    If nBox > 0 And nBox <= m_nElement Then
      If m_Elements(nBox).lngBottom < offset Or m_Elements(nBox).lngTop - offset > picViewPort.ScaleHeight Then Exit Sub
      oldDrawMode = picViewPort.DrawMode
      picViewPort.DrawMode = vbNotXorPen
      oldDrawStyle = picViewPort.DrawStyle
      picViewPort.DrawStyle = vbDot
      X1 = m_Elements(nBox).lngLeft
      X2 = m_Elements(nBox).lngLeft + m_Elements(nBox).f.Width
      Y1 = m_Elements(nBox).lngBottom - offset
      Y2 = Y1
      picViewPort.Line (X1, Y1)-(X2, Y2), RGB(0, 0, 255)
      picViewPort.DrawMode = oldDrawMode
      picViewPort.DrawStyle = oldDrawStyle
    End If
End Sub
'*******************************************************************************
'Process initialize event
'*******************************************************************************
Private Sub UserControl_Initialize()
  picCanvas.Top = 0
  picCanvas.Left = 0
  picViewPort.Top = 0
  picViewPort.Left = 0
  vscScroll.Top = 0
  modHelper.Initialize
  Set modWheelHook.Scroll = vscScroll
End Sub
'*********************************************************************************
'Process resize event
'*********************************************************************************
Private Sub UserControl_Resize()
  On Error Resume Next
  If UserControl.Parent.WindowState <> vbMinimized And Height > 360 Then
    vscScroll.Left = Width - vscScroll.Width
    vscScroll.Height = Height
    picCanvas.Width = vscScroll.Left
    picCanvas.Height = vscScroll.Height
    picViewPort.Width = picCanvas.Width
    picViewPort.Height = picCanvas.Height
  End If
  If UserControl.Ambient.UserMode Then
    If m_nElement > 0 Then
      Render True
    End If
  End If
  picViewPort_Paint
End Sub
'***************************************************************************
'Return TRUE if ch is in ALPHASET
'Input:
'       ch: string to check
'***************************************************************************
Private Function isAlphaDigit(ByVal ch As String) As Boolean
  Const ALPHASET = "0123456789abcdefghijklmnopqrstuvwxyz"
  ch = LCase(ch)
  isAlphaDigit = InStr(ALPHASET, ch) > 0
End Function
'***************************************************************************
'Redraw on picture box. Base on code and ideal from HTMLLabel. Copyright © 2001 Woodbury Associates.
'Input:
'       bRedraw: TRUE if we need to re-calculate element's dimentions
'                FALSE if we need to re-display only
'***************************************************************************
Private Sub Render(Optional ByVal bRedraw As Boolean = False)
On Error GoTo RenderErrHandler
  Dim i As Long, j As Long, nStart As Long, nEnd As Long
  Dim x As Long, y As Long
  Dim st As String
  Dim nMaxAscent As Long
  Dim nMaxHeight As Long
  Dim nMaxDescent As Long
  Dim tm As TEXTMETRIC
  Dim nWidth As Long
  Dim nElementWidth As Long, offset As Long
  Dim bHasNewLine As Boolean
    
  If m_nElement = 0 Then
    picViewPort.Cls
    picCanvas.Cls
    vscScroll.Enabled = False
    Exit Sub
  End If
  picCanvas.Cls
  
  m_nFontSize = picCanvas.font.Size
    
  nLineHeight = picCanvas.TextHeight("X")
  GetTextMetrics picCanvas.hDC, tm
  
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
    nMaxDescent = nMaxHeight - nMaxAscent
    nStart = i
    nEnd = -1
    Do While nWidth < picCanvas.ScaleWidth
      If m_Elements(i).nType = E_SPACE Or m_Elements(i).nType = E_TEXT Then
        nElementWidth = picCanvas.TextWidth(m_Elements(i).strText)
      ElseIf m_Elements(i).nType = E_MATH Then
        If bRedraw Then
          m_Elements(i).f.FontSize = modHelper.GetFontSize(m_nFontSize, picCanvas.hDC)
          m_Elements(i).f.Layout picCanvas.hDC
        End If
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
          'If m_Elements(i).f.Height > nMaxHeight Then
          '  nMaxHeight = m_Elements(i).f.Height
          'End If
          If m_Elements(i).f.Descent > nMaxDescent Then
            nMaxDescent = m_Elements(i).f.Descent
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
    nMaxHeight = nMaxAscent + nMaxDescent
    If nEnd >= nStart Then
      bHasNewLine = False
      If m_Elements(nStart).nType = E_SPACE Then
        j = nStart + 1 'Skip space if it is beginning of line
      Else
        j = nStart
      End If
      Do While j <= nEnd
        If nEnd = (nStart + 1) Then
          If j = nStart And m_Elements(j).nType = E_MATH And m_Elements(nEnd).nType = E_NEWLINE Then
            x = (picCanvas.ScaleWidth - m_Elements(j).f.Width) / 2
          End If
          If m_Elements(nStart).nType = E_SPACE And m_Elements(nEnd).nType = E_MATH And j = nEnd Then
            x = (picCanvas.ScaleWidth - m_Elements(j).f.Width) / 2
          End If
        End If
        If nEnd = nStart And m_Elements(j).nType = E_MATH Then
          x = (picCanvas.ScaleWidth - m_Elements(j).f.Width) / 2
        End If
        If bRedraw Then
          m_Elements(j).lngTop = y
          m_Elements(j).lngLeft = x
          m_Elements(j).lngBottom = y + nMaxHeight
        End If
        If m_Elements(j).nType = E_TEXT Or m_Elements(j).nType = E_SPACE Then
          picCanvas.CurrentX = x
          picCanvas.CurrentY = y + nMaxAscent - tm.tmAscent
          picCanvas.Print m_Elements(j).strText
          x = x + picCanvas.TextWidth(m_Elements(j).strText)
        ElseIf m_Elements(j).nType = E_MATH Then
          'Set picCanvas.font = fFont
          m_Elements(j).f.Draw x, y + nMaxAscent - m_Elements(j).f.Ascent, picCanvas.hDC
          'Set picCanvas.font = tFont
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
    ElseIf nEnd >= 0 Then
      'Width of 1 element larger than canvas width
      j = nEnd + 1
      If m_Elements(j).nType = E_MATH Then
          nMaxHeight = nMaxAscent + nMaxDescent
          If m_Elements(j).f.Ascent > nMaxAscent Then
            nMaxAscent = m_Elements(j).f.Ascent
          End If
          'If m_Elements(j).f.Height > nMaxHeight Then
          '  nMaxHeight = m_Elements(j).f.Height
          'End If
          If m_Elements(j).f.Descent > nMaxDescent Then
            nMaxDescent = m_Elements(j).f.Descent
          End If
      End If
      nMaxHeight = nMaxAscent + nMaxDescent
      If bRedraw Then
          m_Elements(j).lngTop = y
          m_Elements(j).lngLeft = x
          m_Elements(j).lngBottom = y + nMaxHeight
      End If
      If m_Elements(j).nType = E_TEXT Or m_Elements(j).nType = E_SPACE Then
          picCanvas.CurrentX = x
          picCanvas.CurrentY = y + nMaxAscent - tm.tmAscent
          picCanvas.Print m_Elements(j).strText
      ElseIf m_Elements(j).nType = E_MATH Then
          'Set picCanvas.font = fFont
          m_Elements(j).f.Draw x, y + nMaxAscent - m_Elements(j).f.Ascent, picCanvas
          'Set picCanvas.font = tFont
          x = x + m_Elements(j).f.Width
      End If
      If j < m_nElement Then
        If m_Elements(j + 1).nType <> E_NEWLINE Then
          x = nMarginLeft
          y = y + nMaxHeight + nLineSpace
        End If
      End If
      i = i + 1
    End If
  Loop
  
  If bRedraw Then
    nBottom = y + 5
    'Calculation scrollbar's values
    Dim nLinePerPage As Long, nNumOfLine As Long
    If m_nElement > 0 Then
      nLinePerPage = picViewPort.ScaleHeight \ nLineHeight
      nNumOfLine = nBottom \ nLineHeight
      If nNumOfLine <= nLinePerPage Then
        vscScroll.Max = 0
      Else
        vscScroll.Max = nNumOfLine - nLinePerPage 'Substraction first page
      End If
    Else
      vscScroll.Max = 0
    End If
    vscScroll.Value = 0
    If vscScroll.Max = 0 Then
      vscScroll.Enabled = False
    Else
      vscScroll.Enabled = True
      vscScroll.LargeChange = nLinePerPage
      vscScroll.SmallChange = 1
    End If
  End If
  picViewPort_Paint
  Exit Sub
RenderErrHandler:
  Err.Clear
End Sub
'*************************************************************************
'Parse input string to build element array
'If an element is kind formula, it will call expression parser next
'*************************************************************************
Private Sub Parse()
'On Error GoTo ParseErrHandler
  Dim i As Long, j As Long, ch As String, st As String, ok As Boolean
  Dim box As clsBox
  m_nElement = 0
  m_CurrentBox = 0
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
          Set m_Elements(m_nElement).f = box
          'New
          'm_Elements(m_nElement).f.FontSize = m_nFontSize
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
          If Left(st, 1) = "#" And Right(st, 1) = "#" And Len(st) >= 5 Then
            If IsDate(Mid(st, 2, Len(st) - 2)) Then
              st = Mid(st, 2, Len(st) - 2)
              m_Elements(m_nElement).strText = m_Elements(m_nElement).strText & st
            Else
              m_Elements(m_nElement).strText = m_Elements(m_nElement).strText & st
            End If
          Else
            m_Elements(m_nElement).strText = m_Elements(m_nElement).strText & st
          End If
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
                Set m_Elements(m_nElement).f = box
                'New
                'm_Elements(m_nElement).f.FontSize = m_nFontSize
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
  Exit Sub
ParseErrHandler:
  Err.Clear
End Sub
'*********************************************************************************
'Resize array
'*********************************************************************************
Private Sub IncreaseArray()
    m_nElement = m_nElement + 1
    ReDim Preserve m_Elements(1 To m_nElement)
    m_Elements(m_nElement).nType = E_TEXT
End Sub
'*********************************************************************************
'Process Paint event
'*********************************************************************************
Private Sub picViewPort_Paint()
  If Not UserControl.Ambient.UserMode Then
    picCanvas.Cls
    picCanvas.CurrentX = nMarginLeft
    picCanvas.CurrentY = nMarginTop
    picCanvas.Print "Mathematical Formula Drawing Control"
    picCanvas.CurrentX = nMarginLeft
    picCanvas.CurrentY = nMarginTop + picCanvas.TextHeight("X") + nLineSpace
    picCanvas.Print "Copyright 2003 by School@net Technology Company, Ltd."
  End If
  BitBlt picViewPort.hDC, 0, 0, picViewPort.ScaleWidth, picViewPort.ScaleHeight, _
                          picCanvas.hDC, 0, 0, SRCCOPY
  HilightBox m_CurrentBox
End Sub
'********************************************************************************
'Process scroll event of scroll bar
'********************************************************************************
Private Sub vscScroll_Change()
  On Error Resume Next
  If UserControl.Ambient.UserMode And m_nElement > 0 Then
    If m_nElement > 0 Then
      Render False
      picViewPort_Paint
      picViewPort.SetFocus
    End If
  End If
End Sub

Private Sub vscScroll_GotFocus()
  picViewPort.SetFocus
End Sub
'***********************************************************************************
'Process key down event
'***********************************************************************************
Private Sub picViewPort_KeyDown(KeyCode As Integer, Shift As Integer)
  If vscScroll.Enabled = False Then Exit Sub
  Select Case KeyCode
    Case vbKeyUp
      If vscScroll.Value > vscScroll.Min Then
        vscScroll.Value = vscScroll.Value - vscScroll.SmallChange
      End If
    Case vbKeyDown
      If vscScroll.Value < vscScroll.Max Then
        vscScroll.Value = vscScroll.Value + vscScroll.SmallChange
      End If
    Case vbKeyPageUp
      If vscScroll.Value > vscScroll.Min Then
        If vscScroll.Value - vscScroll.LargeChange >= vscScroll.Min Then
          vscScroll.Value = vscScroll.Value - vscScroll.LargeChange
        Else
          vscScroll.Value = vscScroll.Min
        End If
      End If
    Case vbKeyPageDown
      If vscScroll.Value < vscScroll.Max Then
        If vscScroll.Value + vscScroll.LargeChange <= vscScroll.Max Then
          vscScroll.Value = vscScroll.Value + vscScroll.LargeChange
        Else
          vscScroll.Value = vscScroll.Max
        End If
      End If
    Case vbKeyHome
      If (Shift And vbCtrlMask) > 0 Then
        vscScroll.Value = vscScroll.Min
      End If
    Case vbKeyEnd
      If (Shift And vbCtrlMask) > 0 Then
        vscScroll.Value = vscScroll.Max
      End If
  End Select
End Sub
'***************************************************************************
'Interface implemented
'***************************************************************************
Public Property Get FontName() As String
  On Error Resume Next
  FontName = picCanvas.FontName
End Property

Public Property Let FontName(ByVal vNewValue As String)
  On Error Resume Next
  picCanvas.FontName = vNewValue
  Redraw
End Property

Public Property Get FontSize() As Long
  On Error Resume Next
  FontSize = picCanvas.FontSize
End Property

Public Property Let FontSize(ByVal vNewValue As Long)
  On Error Resume Next
  picCanvas.FontSize = vNewValue
  Redraw
End Property

Public Property Get font() As StdFont
  Set font = picCanvas.font
End Property

Public Property Set font(ByVal vNewValue As StdFont)
  On Error Resume Next
  Set picCanvas.font = vNewValue
  Redraw
End Property

Public Property Get FontBold() As Boolean
  On Error Resume Next
  FontBold = picCanvas.FontBold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
  On Error Resume Next
  picCanvas.FontBold = vNewValue
  Redraw
End Property

Public Property Get FontItalic() As Boolean
  On Error Resume Next
  FontItalic = picCanvas.FontItalic
End Property

Public Property Let FontItalic(ByVal vNewValue As Boolean)
  On Error Resume Next
  picCanvas.FontItalic = vNewValue
  Redraw
End Property

Public Property Get FontStrikethru() As Boolean
  On Error Resume Next
  FontStrikethru = picCanvas.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal vNewValue As Boolean)
  On Error Resume Next
  picCanvas.FontStrikethru = vNewValue
  Redraw
End Property

Public Property Get FontUnderline() As Boolean
  On Error Resume Next
  FontUnderline = picCanvas.FontUnderline
End Property

Public Property Let FontUnderline(ByVal vNewValue As Boolean)
  On Error Resume Next
  picCanvas.FontUnderline = vNewValue
  Redraw
End Property

Public Property Get ForeColor() As OLE_COLOR
  On Error Resume Next
  ForeColor = picCanvas.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
  On Error Resume Next
  picCanvas.ForeColor = vNewValue
  Redraw
End Property

Public Property Get BackColor() As OLE_COLOR
  On Error Resume Next
  BackColor = picCanvas.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
  On Error Resume Next
  picCanvas.BackColor = vNewValue
  Redraw
End Property
'********************************************************************************
'Rebuild input string from element array
'********************************************************************************
Public Property Get Text() As String
  Dim i As Long
  strText = ""
  For i = 1 To m_nElement
    If m_Elements(i).nType = E_MATH Then
        strText = strText & "{" & m_Elements(i).strText & "}"
    ElseIf m_Elements(i).nType = E_NEWLINE Then
        strText = strText & vbNewLine
    Else
        strText = strText & m_Elements(i).strText
    End If
  Next
  Text = strText
End Property
'*****************************************************************************
'Set string input
'*****************************************************************************
Public Property Let Text(ByVal vNewValue As String)
  If strText <> vNewValue Then
    strText = vNewValue
  Else
    Exit Property
  End If
  If Extender.Visible Then
    UserControl.MousePointer = vbHourglass
    vscScroll.Enabled = False
  End If
  Parse
  If Extender.Visible Then
    UserControl.MousePointer = vbDefault
  End If
End Property

Public Property Get LeftMargin() As Long
  LeftMargin = nMarginLeft
End Property

Public Property Let LeftMargin(ByVal vNewValue As Long)
  On Error Resume Next
  nMarginLeft = vNewValue
  Redraw
End Property

Public Property Get TopMargin() As Long
  TopMargin = nMarginTop
End Property

Public Property Let TopMargin(ByVal vNewValue As Long)
  On Error Resume Next
  nMarginTop = vNewValue
  Redraw
End Property

Public Property Get LineSpace() As Long
  LineSpace = nLineSpace
End Property

Public Property Let LineSpace(ByVal vNewValue As Long)
  On Error Resume Next
  nLineSpace = vNewValue
  Redraw
End Property
'*********************************************************************
'Redraw control when changing its properties
'**********************************************************************
Private Sub Redraw()
  If UserControl.Ambient.UserMode Then
    If m_nElement > 0 Then
      Render False
    End If
  End If
  picViewPort_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "AllowEdit", m_bAllowEdit, True
  PropBag.WriteProperty "Text", strText, ""
  PropBag.WriteProperty "BackColor", picCanvas.BackColor, vbWhite
  PropBag.WriteProperty "ForeColor", picCanvas.ForeColor, vbBlack
  'PropBag.WriteProperty "Font", picCanvas.Font
  PropBag.WriteProperty "FontName", picCanvas.FontName, ".VnTime"
  PropBag.WriteProperty "FontSize", picCanvas.FontSize, 12
  PropBag.WriteProperty "FontBold", picCanvas.FontBold, False
  PropBag.WriteProperty "FontItalic", picCanvas.FontItalic, False
  PropBag.WriteProperty "FontUnderline", picCanvas.FontUnderline, False
  PropBag.WriteProperty "FontStrikethru", picCanvas.FontStrikethru, False
  PropBag.WriteProperty "LeftMargin", nMarginLeft, 5
  PropBag.WriteProperty "TopMargin", nMarginTop, 5
  PropBag.WriteProperty "LineSpace", nLineSpace, 5
  'PropBag.WriteProperty "MathFontSize", m_nFontSize, 10
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_bAllowEdit = PropBag.ReadProperty("AllowEdit", True)
  strText = PropBag.ReadProperty("Text", "")
  nMarginLeft = PropBag.ReadProperty("LeftMargin", 5)
  nMarginTop = PropBag.ReadProperty("TopMargin", 5)
  nLineSpace = PropBag.ReadProperty("LineSpace", 5)
  picCanvas.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
  picCanvas.ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
  'Set picCanvas.Font = PropBag.ReadProperty("Font")
  picCanvas.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
  picCanvas.FontName = PropBag.ReadProperty("FontName", ".VnTime")
  picCanvas.FontSize = PropBag.ReadProperty("FontSize", 12)
  picCanvas.FontBold = PropBag.ReadProperty("FontBold", False)
  picCanvas.FontItalic = PropBag.ReadProperty("FontItalic", False)
  picCanvas.FontUnderline = PropBag.ReadProperty("FontUnderline", False)
  picCanvas.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
  'm_nFontSize = PropBag.ReadProperty("MathFontSize", 10)
End Sub
'****************************************************************************
'Call this to render input string on screen (picture box)
'****************************************************************************
Public Sub RenderOnScreen()
  If m_nElement > 0 Then
      Render True
  End If
End Sub
'*********************************************************************************
'Convert input string to RTF. But it has not completed yet
'Input:
'       strFileName: file name to export
'*********************************************************************************
Public Sub SaveRtf(ByVal strFileName)
  Dim f As Long
  f = FreeFile
  Open strFileName For Binary As f
  GetRtf f
  Close f
  RaiseEvent SaveCompleted
End Sub
'*******************************************************************************
'Convert all element to RTF. Called from SaveRTF
'*******************************************************************************
Public Sub GetRtf(hFile As Long)
  Dim st As String
  Dim i As Long
  Dim nF As Long
  On Error GoTo ErrHdl
  nF = picCanvas.FontSize
  st = "{\rtf1\ansi\ansicpg1252\uc0\deff0{\fonttbl"
  st = st & "{\f0\fswiss\fcharset0\fprq2 Schoolnet Sans Serif;}}"
  st = st & "{\colortbl;\red0\green0\blue0;\red255\green255\blue255;}"
  st = st & "\deftab1134\paperw11906\paperh16837\margl1134\margt1134\margr1134\margb1134"
  st = st & "\pard\plain\f0\fs20 "
  Put hFile, , st
  For i = 1 To m_nElement
    If m_Elements(i).nType = E_NEWLINE Then
      Put hFile, , "\par "
    ElseIf m_Elements(i).nType = E_TEXT Or m_Elements(i).nType = E_SPACE Then
      Text2Rtf m_Elements(i).strText, hFile
    Else
      m_Elements(i).f.FontSize = modHelper.GetFontSize(nF, picCanvas.hDC)
      If Formula2Rtf(m_Elements(i).f, hFile) = False Then
        Exit Sub
      End If
    End If
  Next
  Put hFile, , "\par }"
  Exit Sub
ErrHdl:
  MsgBox Err.Description & vbNewLine & "Module: GetRtf Property", vbCritical Or vbOKOnly, App.Title
End Sub
'***********************************************************************************
'Convert a formula to RTF
'Input:
'       f: The formular need to be converted
'       hFile: Handle of file to write out
'***********************************************************************************
Private Function Formula2Rtf(f As clsBox, hFile As Long) As Boolean
  Dim hMetaDC As Long
  Dim rc As RECT
  Dim rcMeta As RECT
  Dim hMetaFile As Long
  Dim hOldFont As Long
  Dim nBufferSize As Long
  Dim Buffer() As Byte
  Dim i As Long
  Dim st As String, ss As String
  Dim nAsc As Long
  Dim imgw As Long, imgh As Long, wgoal As Long, hgoal As Long
  Dim hNewPen As Long
  Dim hOldPen As Long
  Dim hDCref As Long
  Dim count As Long
  On Error GoTo ErrHdl
  
  hDCref = CreateCompatibleDC(ByVal 0)
  SetTextAlign hDCref, TA_LEFT Or TA_TOP
  hNewPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0))
  hOldFont = modHelper.GetFont(picCanvas.hDC)
  hOldPen = SelectObject(hDCref, hNewPen)
  SelectObject hDCref, hOldFont
  f.Layout hDCref
  nAsc = (f.Height - f.Ascent) * 1.5
  rc.Left = 0
  rc.Top = 0
  rc.Bottom = f.Height
  rc.Right = f.Width
  
  rcMeta.Top = 0
  rcMeta.Left = 0
  rcMeta.Right = MulDiv(rc.Right * 100, GetDeviceCaps(hDCref, HORZSIZE), GetDeviceCaps(hDCref, HORZRES))
  rcMeta.Bottom = MulDiv(rc.Bottom * 100, GetDeviceCaps(hDCref, VERTSIZE), GetDeviceCaps(hDCref, VERTRES))
  
  imgw = f.Width / GetDeviceCaps(hDCref, LOGPIXELSX) * 2540
  imgh = f.Height / GetDeviceCaps(hDCref, LOGPIXELSY) * 2540
  wgoal = f.Width / GetDeviceCaps(hDCref, LOGPIXELSX) * 1440
  hgoal = f.Height / GetDeviceCaps(hDCref, LOGPIXELSY) * 1440
  
  hMetaDC = CreateEnhMetaFile(0, vbNullString, rcMeta, "Formula" & Chr(0) & "Image" & Chr(0) & Chr(0))
   
  modHelper.SetFont hOldFont, hMetaDC
  hOldPen = SelectObject(hMetaDC, hNewPen)
  SetMapMode hMetaDC, MM_TEXT
  SetBkMode hMetaDC, TRANSPARENT
  SetTextAlign hMetaDC, TA_LEFT Or TA_TOP
  f.Draw 0, 0, hMetaDC
  
  hMetaFile = CloseEnhMetaFile(hMetaDC)
  
  nBufferSize = GetWinMetaFileBits(hMetaFile, 0, ByVal 0, MM_ANISOTROPIC, hDCref)
  ReDim Buffer(1 To nBufferSize) As Byte
  GetWinMetaFileBits hMetaFile, nBufferSize, Buffer(1), MM_ANISOTROPIC, hDCref
  st = "\plain\f0\fs20\dn" & CStr(nAsc) & " {\pict\wmetafile8\picw" & CStr(imgw) & "\pich" & CStr(imgh) & "\picwgoal" & CStr(wgoal) & "\pichgoal" & CStr(hgoal) & "\picscalex100\picscaley100 "
  count = Len(st)
  For i = 1 To nBufferSize
    ss = LCase(Hex(Buffer(i)))
    If Len(ss) = 1 Then ss = "0" & ss
    st = st & ss
    count = count + 2
    If count > 2048 And hFile <> 0 Then
      Put hFile, , st
      st = ""
      count = 0
    End If
  Next
  st = st & "}\plain\f0\fs20 "
  If hFile <> 0 Then
    Put hFile, , st
  End If
  DeleteEnhMetaFile hMetaFile
  DeleteObject hNewPen
  DeleteDC hDCref
  Formula2Rtf = True
  Exit Function
ErrHdl:
  MsgBox Err.Description & vbNewLine & "Module: Formula2Rtf", vbCritical Or vbOKOnly, App.Title
  Formula2Rtf = False
  CloseEnhMetaFile hMetaDC
  DeleteEnhMetaFile hMetaFile
  DeleteObject hNewPen
  DeleteDC hDCref
End Function
'******************************************************************************
'Convert a string to RTF
'Input:
'       str: The string need to be converted
'******************************************************************************
Private Function Text2Rtf(ByVal str As String, hFile As Long) As Boolean
  Dim i As Long
  Dim st As String, ch As String
  Dim count As Long
  If Trim(str) = "" Then
    Text2Rtf = True
    Exit Function
  End If
  st = ""
  count = 0
  For i = 1 To Len(str)
    ch = Mid(str, i, 1)
    If Asc(ch) > 127 Then
      st = st & "\'" & LCase(CStr(Hex(Asc(ch))))
      count = count + 3
    Else
      st = st & ch
      count = count + 1
    End If
    If count > 2048 And hFile <> 0 Then
      Put hFile, , st
      st = ""
      count = 0
    End If
  Next
  st = st & " "
  If st <> "" And hFile <> 0 Then
    Put hFile, , st
  End If
  Text2Rtf = True
End Function
'*****************************************************************************
'Render on VSPrinter. Semilar with render on screen but with printing ability
'Input:
'   x, y: where to render out. Those are byref arguments. Its return last position
'         after rendering
'   vps: VSPrinter object
'*****************************************************************************
Public Sub RenderOnVSP(x As Long, y As Long, vps As Object)
On Error GoTo RenderErrHandler
  Dim i As Long, j As Long, nStart As Long, nEnd As Long
  Dim st As String
  Dim nMaxAscent As Long
  Dim nMaxHeight As Long
  Dim nMaxDescent As Long
  Dim nWidth As Long
  Dim nElementWidth As Long, offset As Long
  Dim bHasNewLine As Boolean
  Dim sX As Double, sY As Double
  Dim nCharAscent As Long
  Parse
  Dim nClientWidth As Long
  If vps.Error > 0 Then
    MsgBox "Printer is not ready or error when accessing printer", vbCritical Or vbOKOnly, "Mathematics Equation Drawing"
    Exit Sub
  End If
  If m_nElement = 0 Then
    Exit Sub
  End If
  m_nFontSize = vps.font.Size
  sX = vps.TwipsPerPixelX
  sY = vps.TwipsPerPixelY

  nLineHeight = modHelper.TextHeight("X", vps.hDC)
  nCharAscent = modHelper.TextAscent(vps.hDC)
  i = 1
  nClientWidth = vps.PageWidth - vps.MarginRight
  Do While i <= m_nElement
    nWidth = x
    nMaxAscent = nCharAscent
    nMaxHeight = modHelper.TextHeight("X", vps.hDC)
    nMaxDescent = nMaxHeight - nMaxAscent
    nStart = i
    nEnd = -1
    Do While nWidth <= nClientWidth
      If m_Elements(i).nType = E_SPACE Or m_Elements(i).nType = E_TEXT Then
        nElementWidth = vps.TextWidth(m_Elements(i).strText)
      ElseIf m_Elements(i).nType = E_MATH Then
        'Set vps.font = fFont
        m_Elements(i).f.FontSize = modHelper.GetFontSize(m_nFontSize, vps.hDC)
        m_Elements(i).f.Layout vps.hDC
        'Set vps.font = tFont
        nElementWidth = m_Elements(i).f.Width * sX
      Else
        i = i + 1
        nEnd = i
        Exit Do
      End If
      If (nWidth + nElementWidth) <= nClientWidth Then
        nWidth = nWidth + nElementWidth
        If m_Elements(i).nType = E_MATH Then
          If m_Elements(i).f.Ascent > nMaxAscent Then
            nMaxAscent = m_Elements(i).f.Ascent
          End If
          'If m_Elements(i).f.Height > nMaxHeight Then
          '  nMaxHeight = m_Elements(i).f.Height
          'End If
          If m_Elements(i).f.Descent > nMaxDescent Then
            nMaxDescent = m_Elements(i).f.Descent
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
    nMaxHeight = nMaxAscent + nMaxDescent
    If nEnd >= nStart Then
      bHasNewLine = False
      If m_Elements(nStart).nType = E_SPACE Then
        j = nStart + 1 'Skip space if it is beginning of line
      Else
        j = nStart
      End If
      If (y + nMaxHeight * sY) > (vps.PageHeight - vps.MarginBottom) Then
        vps.NewPage
        y = vps.MarginTop
      End If
      Do While j <= nEnd
        If nEnd = (nStart + 1) Then
          If j = nStart And m_Elements(j).nType = E_MATH And m_Elements(nEnd).nType = E_NEWLINE Then
            x = (nClientWidth - m_Elements(j).f.Width * sX) / 2
          End If
          If m_Elements(nStart).nType = E_SPACE And m_Elements(nEnd).nType = E_MATH And j = nEnd Then
            x = (nClientWidth - m_Elements(j).f.Width * sX) / 2
          End If
        End If
        If nEnd = nStart And m_Elements(j).nType = E_MATH Then
          x = (nClientWidth - m_Elements(j).f.Width * sX) / 2
        End If

        If m_Elements(j).nType = E_TEXT Or m_Elements(j).nType = E_SPACE Then
          vps.CurrentX = x
          vps.CurrentY = y + (nMaxAscent - nCharAscent) * sX
          vps.Text = m_Elements(j).strText
          x = x + vps.TextWidth(m_Elements(j).strText)
        ElseIf m_Elements(j).nType = E_MATH Then
          'Set vps.font = fFont
          m_Elements(j).f.Draw x / sX, y / sY + nMaxAscent - m_Elements(j).f.Ascent, vps.hDC
          'Set vps.font = tFont
          x = x + m_Elements(j).f.Width * sX
        Else 'New line
          x = vps.MarginLeft
          y = y + nMaxHeight * sY + vps.LineSpacing
          bHasNewLine = True
        End If
        j = j + 1
      Loop
      If Not bHasNewLine Then
        x = vps.MarginLeft
        y = y + nMaxHeight * sY + vps.LineSpacing
      End If
    ElseIf nEnd >= 0 Then
      'Width of 1 element is larger than canvas width
      j = nEnd + 1
      If m_Elements(j).nType = E_MATH Then
          If m_Elements(j).f.Ascent > nMaxAscent Then
            nMaxAscent = m_Elements(j).f.Ascent
          End If
          'If m_Elements(j).f.Height > nMaxHeight Then
          '  nMaxHeight = m_Elements(j).f.Height
          'End If
          If m_Elements(i).f.Descent > nMaxDescent Then
            nMaxDescent = m_Elements(i).f.Descent
          End If
      End If
      nMaxHeight = nMaxAscent + nMaxDescent
      If m_Elements(j).nType = E_TEXT Or m_Elements(j).nType = E_SPACE Then
          vps.CurrentX = x
          vps.CurrentY = y + (nMaxAscent - nCharAscent) * sX
          vps.Text = m_Elements(j).strText
      ElseIf m_Elements(j).nType = E_MATH Then
          'Set vps.font = fFont
          m_Elements(j).f.Draw x / sX, y / sY + nMaxAscent - m_Elements(j).f.Ascent, vps.hDC
          'Set vps.font = tFont
          x = x + m_Elements(j).f.Width * sX
      End If
      If j < m_nElement Then
        If m_Elements(j + 1).nType <> E_NEWLINE Then
          x = vps.MarginLeft
          y = y + nMaxHeight * sY + vps.LineSpacing
        End If
      End If
      i = i + 1
    End If
  Loop
  Exit Sub
RenderErrHandler:
  Err.Clear
End Sub

