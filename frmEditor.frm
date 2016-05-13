VERSION 5.00
Begin VB.Form frmEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equation Editor"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Schoolnet Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   45
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   4
      Top             =   0
      Width           =   6615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "§ãng"
      BeginProperty Font 
         Name            =   "Schoolnet Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "CËp nhËt"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Schoolnet Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Xem tr­íc"
      BeginProperty Font 
         Name            =   "Schoolnet Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtContent 
      BeginProperty Font 
         Name            =   "Schoolnet Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2400
      Width           =   6615
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bState As Boolean
Public strFormular As String
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If RenderFormular Then
    bState = True
    strFormular = txtContent.Text
    Unload Me
    Exit Sub
  End If
End Sub

Private Sub cmdPreview_Click()
    RenderFormular
End Sub

Private Sub Form_Load()
    bState = False
    txtContent.Text = strFormular
    txtContent.SelStart = Len(strFormular)
    RenderFormular
End Sub

Private Function RenderFormular() As Boolean
  Dim f As clsBox
  Dim x As Long, y As Long
  Dim st As String
  On Error GoTo ErrHdl
  picPreview.Cls
  Set f = modHelper.ParseExpression(txtContent.Text)
  If Not f Is Nothing Then
    f.FontSize = modHelper.GetFontSize(picPreview.FontSize, picPreview.hDC)
    f.Layout picPreview.hDC
    If f.Width >= picPreview.ScaleWidth Then
        x = 0
    Else
        x = (picPreview.ScaleWidth - f.Width) \ 2
    End If
    
    If f.Height >= picPreview.ScaleHeight Then
        y = 0
    Else
        y = (picPreview.ScaleHeight - f.Height) \ 2
    End If
    f.Draw x, y, picPreview.hDC
    Set f = Nothing
    RenderFormular = True
  Else
    GoTo ErrHdl
  End If
  Exit Function
ErrHdl:
  st = "Lçi có ph¸p. Xem l¹i c«ng thøc"
  x = (picPreview.ScaleWidth - modHelper.TextWidth(st, picPreview.hDC)) \ 2
  y = picPreview.ScaleHeight \ 2 - modHelper.TextAscent(picPreview.hDC)
  picPreview.CurrentX = x
  picPreview.CurrentY = y
  picPreview.Print st
  RenderFormular = False
End Function

