VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module clsBox
'Abstract base class for implementation other boxs
Option Explicit
Public Parent As clsBox
'Render formular on a VSPrinter or PictureBox
Public Sub Draw(ByVal x As Long, ByVal y As Long, hDC As Long)
  
End Sub
'Type of box
Public Property Get ClassName() As String

End Property
'Width of box
Public Property Get Width() As Long

End Property
'Height of box
Public Property Get Height() As Long

End Property
'Calculation all dimention of box
Public Sub Layout(hDC As Long)

End Sub
'Ascent of box
Public Property Get Ascent() As Long

End Property
'Descent of box
Public Property Get Descent() As Long

End Property
'Create another copy instance of box object
Public Property Get Copy() As clsBox

End Property

Public Property Get FontSize() As Long

End Property

Public Property Let FontSize(ByVal vNewValue As Long)

End Property

