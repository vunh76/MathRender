VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsBox
Private m_nRows As Long
Private m_nCols As Long
Private m_RowHeight() As Long
Private m_RowAscent() As Long
Private m_ColWidth() As Long
Private m_Matrix() As clsBox
Private m_Width As Long
Private m_Height As Long
Private m_Ascent As Long
Private m_nFontSize As Long
Private m_Parent As clsBox
Private m_LineSpace As Long
Private m_ColSpace As Long
Private m_bDet As Boolean
Private m_bBracket As Boolean
Public Property Get IsBracket() As Boolean
  IsBracket = m_bBracket
End Property
Public Property Let IsBracket(ByVal b As Boolean)
  m_bBracket = b
End Property
Public Property Get IsDet() As Boolean
  IsDet = m_bDet
End Property
Public Property Let IsDet(ByVal vNewValue As Boolean)
  m_bDet = vNewValue
End Property
Public Sub CreateMatrix(ByVal m As Long, n As Long)
  Dim i As Long
  Dim j As Long
  For i = 1 To m_nRows
    For j = 1 To m_nCols
      Set m_Matrix(i, j) = Nothing
    Next
  Next
  Erase m_Matrix
  ReDim m_Matrix(1 To m, 1 To n)
  ReDim m_RowHeight(1 To m)
  ReDim m_RowAscent(1 To m)
  ReDim m_ColWidth(1 To n)
  m_nRows = m
  m_nCols = n
End Sub
Property Get Element(ByVal i As Long, ByVal j As Long) As clsBox
  If i <= m_nRows And j <= m_nCols Then
    Set Element = m_Matrix(i, j)
  Else
    Set Element = Nothing
  End If
End Property
Property Set Element(ByVal i As Long, ByVal j As Long, vNewValue As clsBox)
  Set m_Matrix(i, j) = vNewValue
  Set m_Matrix(i, j).Parent = Me
  m_Matrix(i, j).FontSize = m_nFontSize
End Property

Private Sub Class_Initialize()
  m_bBracket = True
  m_bDet = False
End Sub

Private Property Get clsBox_Copy() As clsBox
  Dim m As New clsMatrix
  Dim i As Long, j As Long
  m.CreateMatrix m_nRows, m_nCols
  m.IsDet = m_bDet
  m.IsBracket = m_bBracket
  For i = 1 To m_nRows
    For j = 1 To m_nCols
      Set m.Element(i, j) = m_Matrix(i, j)
    Next
  Next
  Set clsBox_Copy = m
End Property

Private Sub clsBox_Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
  Dim nX As Long, nY As Long, dy As Long, dx As Long
  Dim i As Long, j As Long
  If m_bBracket Then
    modHelper.DrawLine x, y, x, y + m_Height - 1, hDC
    If Not m_bDet Then
      modHelper.DrawLine x, y, x + m_ColSpace, y, hDC
      modHelper.DrawLine x, y + m_Height - 1, x + m_ColSpace, y + m_Height - 1, hDC
    End If
  End If
  dy = y + m_LineSpace / 2
  For i = 1 To m_nRows
    dx = x + m_ColSpace
    For j = 1 To m_nCols
      nX = dx + (m_ColWidth(j) - m_Matrix(i, j).Width) \ 2
      nY = dy + m_RowAscent(i) - m_Matrix(i, j).Ascent
      m_Matrix(i, j).Draw nX, nY, hDC
      dx = dx + m_ColWidth(j) + m_ColSpace
    Next
    dy = dy + m_RowHeight(i) + m_LineSpace
  Next
  If m_bBracket Then
    modHelper.DrawLine x + m_Width - 1, y, x + m_Width - 1, y + m_Height - 1, hDC
    If Not m_bDet Then
      modHelper.DrawLine x + m_Width - m_ColSpace, y, x + m_Width - 1, y, hDC
      modHelper.DrawLine x + m_Width - m_ColSpace, y + m_Height - 1, x + m_Width - 1, y + m_Height - 1, hDC
    End If
  End If
End Sub

Private Property Let clsBox_FontSize(ByVal RHS As Long)
  Dim i As Long, j As Long
  m_nFontSize = RHS
  For i = 1 To m_nRows
    For j = 1 To m_nCols
      If Not m_Matrix(i, j) Is Nothing Then
        m_Matrix(i, j).FontSize = m_nFontSize
      End If
    Next
  Next
End Property

Private Sub clsBox_Layout(hDC As Long)
  Dim i As Long
  Dim j As Long
  Dim maxHeight As Long
  Dim maxWidth As Long
  For j = 1 To m_nCols
    m_ColWidth(j) = 0
  Next
  m_ColSpace = modHelper.TextWidth("X", hDC)
  m_LineSpace = m_ColSpace
  maxHeight = 0
  maxWidth = 0
  For i = 1 To m_nRows
    m_RowHeight(i) = 0
    m_RowAscent(i) = 0
    For j = 1 To m_nCols
      m_Matrix(i, j).Layout hDC
      If m_Matrix(i, j).Height > m_RowHeight(i) Then
        m_RowHeight(i) = m_Matrix(i, j).Height
      End If
      If m_Matrix(i, j).Ascent > m_RowAscent(i) Then
        m_RowAscent(i) = m_Matrix(i, j).Ascent
      End If
      If m_ColWidth(j) < m_Matrix(i, j).Width Then
        m_ColWidth(j) = m_Matrix(i, j).Width
      End If
    Next
    maxHeight = maxHeight + m_RowHeight(i) + m_LineSpace
  Next
  maxWidth = 2 * m_ColSpace + m_ColWidth(1)
  For j = 2 To m_nCols
    maxWidth = maxWidth + m_ColSpace + m_ColWidth(j)
  Next
  m_Height = maxHeight
  m_Width = maxWidth
  m_Ascent = m_Height / 2 - modHelper.TextDecent(hDC)
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
  clsBox_ClassName = "matrix"
End Property
