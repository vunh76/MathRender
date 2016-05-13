Attribute VB_Name = "modHelper"
'Module modHelper. Base on clsExpressionParse.
'Sorry. I do not remember the author
'Implementation expression parser
'Grammar:
'MathExp    : Expression ['=' Expression]*
'Expression : Term ['+'|'-' Term]*
'Term       : Factor ['*'|'/' Fractor]*
'Fractor    : Element ['^' | '_' Element]*
'Element    : '(' Expression ')' | Atom
'Atom       : Symbol
'Symbol     : Leter[Letter]*
'Letter     : A-Z|0-9|'['']'',''.'
Option Explicit
Public Const MATH_SYMBOL_FONT = "Lucida Bright Math Symbol"
Public Const MATH_CHAR_FONT = "Lucida Bright Math Italic"
Public Const MATH_EXTENSION_FONT = "Lucida Bright Math Extension"
Public Const INT_SIGN = &H97
Public Const SUM_SIGN = &HAA
Public Const PROD_SIGN = &HA9
Public Const UNI_SIGN = &H7E
Public Const PLUS_SIGN = &H21
Private Const GENERIC_SYNTAX_ERR_MSG = "Syntax Error"
Private Const MAX_SYMBOLS = 83
Private Const MAX_MATHOP = 11
Private Enum ParserErrors
    PERR_FIRST = vbObjectError + 513
    PERR_SYNTAX_ERROR = PERR_FIRST
    PERR_DIVISION_BY_ZERO
    PERR_CLOSING_PARENTHESES_EXPECTED
    PERR_OVER_OPERATOR_EXPECTED
    PERR_LAST = PERR_OVER_OPERATOR_EXPECTED
End Enum
Private Enum ParserTokens
    TOK_UNKNOWN
    TOK_FIRST
    TOK_PM = TOK_FIRST
    TOK_ADD
    TOK_MP
    TOK_SUBTRACT
    TOK_EQUAL
    TOK_DIFF
    TOK_LEQ
    TOK_LESS
    TOK_GEQ
    TOK_GRE
    TOK_AND
    TOK_or
    TOK_MULTIPLY
    TOK_DIVIDE
    TOK_POWER
    TOK_SUBSCRIPT
    TOK_OPEN_PARENTHESES
    TOK_CLOSE_PARENTHESES
    TOK_LIST_SEPERATOR
    TOK_LAST = TOK_LIST_SEPERATOR
End Enum
Public Enum OVERTYPE
    OVER_NORMAL
    OVER_VECTOR
    OVER_ANGLE
    OVER_CHORD
    OVER_ABS
End Enum
Public Type MATH_SYMBOL
  strEntity As String
  nCode As Long
  nFontIndex As Long
End Type
Public Type MATH_OP
  strOP As String
  nCode As Long
  nFontIndex As Long
End Type
Private mMathSymbols() As MATH_SYMBOL
Private mMathOPs() As MATH_OP
'Private mMathCodes() As Long
Private mTokenSymbols() As String
Private mExpression As String
Private mPosition As Long
Private mLastTokenLength As Long
Private mProjectName As String
Public Function ParseExpression(ByVal strExpression As String) As clsBox
    'On Error GoTo ParseExpression_ErrHandler
    Dim Value As New clsBox
    Dim OtherValue As clsBox
    Dim CurrToken As ParserTokens
    mExpression = strExpression
    mPosition = 1
    SkipSpaces
    Set Value = Expression
    If mPosition <= Len(mExpression) Then
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
    End If
    Set ParseExpression = Value
    Exit Function
ParseExpression_ErrHandler:
    Set ParseExpression = Nothing
    SetErrSource "ParseExpression"
    Err.Raise Err.Number
End Function
Private Function Expression() As clsBox
    'On Error GoTo ParseExp_ErrHandler
    Dim Value As clsBox
    Dim OtherValue As clsBox
    Dim CurrToken As ParserTokens
    Set Value = Term
    Do While mPosition <= Len(mExpression)
        CurrToken = GetToken
        If CurrToken = TOK_PM Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text("+-"), OtherValue)
        ElseIf CurrToken = TOK_MP Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text("-+"), OtherValue)
        ElseIf CurrToken = TOK_ADD Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text("+"), OtherValue)
        ElseIf CurrToken = TOK_SUBTRACT Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text("-"), OtherValue)
        ElseIf CurrToken = TOK_EQUAL Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text("="), OtherValue)
        ElseIf CurrToken = TOK_DIFF Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text("<>"), OtherValue)
        ElseIf CurrToken = TOK_LESS Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text("<"), OtherValue)
        ElseIf CurrToken = TOK_GRE Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text(">"), OtherValue)
        ElseIf CurrToken = TOK_LEQ Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text("<="), OtherValue)
        ElseIf CurrToken = TOK_GEQ Then
            SkipLastToken
            Set OtherValue = Term
            Set Value = Concat3(Value, Text(">="), OtherValue)
        ElseIf CurrToken = TOK_UNKNOWN Then
            Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
        Else
            Set Expression = Value
            Exit Function
        End If
    Loop
    Set Expression = Value
    Exit Function
ParseExp_ErrHandler:
    SetErrSource "ParseExp"
    Err.Raise Err.Number
End Function
Private Function Term() As clsBox
    'On Error GoTo ParseTerm_ErrHandler
    Dim Value As clsBox
    Dim OtherValue As clsBox
    Dim CurrToken As ParserTokens
    Set Value = Factor
    Do While mPosition <= Len(mExpression)
        CurrToken = GetToken
        If CurrToken = TOK_MULTIPLY Then
            SkipLastToken
            Set OtherValue = Factor
            Set Value = Concat3(Value, Text("*"), OtherValue)
        ElseIf CurrToken = TOK_DIVIDE Then
            SkipLastToken
            Set OtherValue = Factor
            Set Value = Fraction(Value, OtherValue)
        ElseIf CurrToken = TOK_AND Then
            SkipLastToken
            Set OtherValue = Factor
            Set Value = Concat3(Value, Text("&and;"), OtherValue)
        ElseIf CurrToken = TOK_or Then
            SkipLastToken
            Set OtherValue = Factor
            Set Value = Concat3(Value, Text("&or;"), OtherValue)
        ElseIf CurrToken = TOK_UNKNOWN Then
            Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
        Else
            Set Term = Value
            Exit Function
        End If
    Loop
    Set Term = Value
    Exit Function
ParseTerm_ErrHandler:
    SetErrSource "ParseTerm"
    Err.Raise Err.Number
End Function
Private Function Factor() As clsBox
    'On Error GoTo ParseFactor_ErrHandler
    Dim Value As clsBox
    Dim OtherValue As clsBox
    Dim BracketTmp As clsBracket
    Dim CurrToken As ParserTokens
    Set Value = Element
    Do While mPosition <= Len(mExpression)
        CurrToken = GetToken
        If CurrToken = TOK_POWER Then
            SkipLastToken
            Set OtherValue = Element
            'Remove bracket if OtherValue is a super script
            If OtherValue.ClassName = "bracket" Then
              Set BracketTmp = OtherValue
              Set Value = Power(Value, BracketTmp.Content)
            Else
              Set Value = Power(Value, OtherValue)
            End If
        ElseIf CurrToken = TOK_SUBSCRIPT Then
            SkipLastToken
            Set OtherValue = Element
            'Remove bracket if OtherValue is a subscript script
            If OtherValue.ClassName = "bracket" Then
              Set BracketTmp = OtherValue
              Set Value = Subscript(Value, BracketTmp.Content)
            Else
              Set Value = Subscript(Value, OtherValue)
            End If
        Else
            Set Factor = Value
            Exit Function
        End If
    Loop
    Set Factor = Value
    Exit Function
ParseFactor_ErrHandler:
    SetErrSource "ParseFactor"
    Err.Raise Err.Number
End Function
Private Function Element() As clsBox
    'On Error GoTo ParseElement_ErrHandler
    Dim Sign As String
    Dim CurrToken As ParserTokens
    Dim Value As clsBox
    Sign = ""
    CurrToken = GetToken
    If CurrToken = TOK_SUBTRACT Then
        Sign = "-"
        SkipLastToken
    ElseIf CurrToken = TOK_ADD Then
        SkipLastToken
    End If
    CurrToken = GetToken
    If CurrToken = TOK_OPEN_PARENTHESES Then
        SkipLastToken
        Set Value = Bracket(Expression)
        CurrToken = GetToken
        If CurrToken = TOK_CLOSE_PARENTHESES Then
            SkipLastToken
        Else
            Err.Raise PERR_CLOSING_PARENTHESES_EXPECTED, , "')' Expected"
        End If
    Else
        Set Value = Atom
    End If
    If Sign = "" Then
      Set Element = Value
    Else
      Set Element = Concat2(Text("-"), Value)
    End If
    Exit Function
ParseElement_ErrHandler:
    SetErrSource "Element"
    Err.Raise Err.Number
End Function
Private Function Atom() As clsBox
    'On Error GoTo ParseAtom_ErrHandler
    Dim CurrPosition As Long
    Dim CurrToken As ParserTokens
    Dim Value As clsBox
    Dim st As String
    Dim i, j, k As Long
    Dim ArgumentValue As clsBox
    Dim ArgList() As clsBox
    Dim nArg As Long
    If mPosition > Len(mExpression) Then
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
    End If
    CurrPosition = mPosition
    If IsAlpha(Mid(mExpression, CurrPosition, 1)) Then
        CurrPosition = CurrPosition + 1
        Do While IsAlpha(Mid(mExpression, CurrPosition, 1))
            CurrPosition = CurrPosition + 1
        Loop
        st = Mid(mExpression, mPosition, CurrPosition - mPosition)
        mPosition = CurrPosition
        SkipSpaces
        CurrToken = GetToken
        If CurrToken = TOK_OPEN_PARENTHESES Then
           SkipLastToken
           nArg = 0
           Do While CurrPosition <= Len(mExpression)
              Set ArgumentValue = Expression
              nArg = nArg + 1
              ReDim Preserve ArgList(1 To nArg)
              Set ArgList(nArg) = ArgumentValue
              CurrToken = GetToken
              If CurrToken = TOK_CLOSE_PARENTHESES Then
                SkipLastToken
                Exit Do
              ElseIf CurrToken = TOK_LIST_SEPERATOR Then
                SkipLastToken
              Else
                Err.Raise PERR_CLOSING_PARENTHESES_EXPECTED
              End If
           Loop
           Set Value = ParseFunction(st, ArgList, nArg)
        Else
            Set Value = Text(st)
            mPosition = CurrPosition
            SkipSpaces
        End If
    Else
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
    End If
    Set Atom = Value
    Exit Function
ParseAtom_ErrHandler:
    SetErrSource "ParseAtom"
    Err.Raise Err.Number
End Function
'Get current tocken and set length of that tocken
Private Function GetToken() As ParserTokens
    Dim CurrToken As ParserTokens
    Dim i As ParserTokens
    If mPosition > Len(mExpression) Then
        GetToken = TOK_UNKNOWN
        Exit Function
    End If
    CurrToken = TOK_UNKNOWN
    mLastTokenLength = 0
    For i = TOK_FIRST To TOK_LAST
        If Mid(mExpression, mPosition, Len(mTokenSymbols(i))) = mTokenSymbols(i) Then
            CurrToken = i
            mLastTokenLength = Len(mTokenSymbols(i))
            Exit For
        End If
    Next
    GetToken = CurrToken
End Function
Private Function ParseFunction(ByVal strTockenName As String, ArgList() As clsBox, ByVal nArg As Long) As clsBox
  Dim v As clsBox
  Dim v1 As clsBox
  Dim i As Long
  Dim t1 As clsText
  Dim t2 As clsText
  Select Case UCase(strTockenName)
    Case "SQRT":
      Set v = Root(ArgList(1))
    Case "SUM"
      If nArg = 1 Then
        Set v = Sum(Nothing, Nothing, ArgList(1))
      ElseIf nArg = 2 Then
        Set v = Sum(ArgList(2), Nothing, ArgList(1))
      ElseIf nArg = 3 Then
        Set v = Sum(ArgList(2), ArgList(3), ArgList(1))
      Else
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
      End If
    Case "PROD"
      If nArg = 1 Then
        Set v = Prod(Nothing, Nothing, ArgList(1))
      ElseIf nArg = 2 Then
        Set v = Prod(ArgList(2), Nothing, ArgList(1))
      ElseIf nArg = 3 Then
        Set v = Prod(ArgList(2), ArgList(3), ArgList(1))
      Else
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
      End If
    Case "UNI"
      If nArg = 1 Then
        Set v = Uni(Nothing, Nothing, ArgList(1))
      ElseIf nArg = 2 Then
        Set v = Uni(ArgList(2), Nothing, ArgList(1))
      ElseIf nArg = 3 Then
        Set v = Uni(ArgList(2), ArgList(3), ArgList(1))
      Else
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
      End If
    Case "INT":
      If nArg = 2 Then
        Set v = InInt(Nothing, Nothing, Concat2(ArgList(1), ArgList(2)))
      ElseIf nArg = 4 Then
        Set v = InInt(ArgList(3), ArgList(4), Concat2(ArgList(1), ArgList(2)))
      Else
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
      End If
    Case "OVER"
      Set v = Over(ArgList(1))
    Case "VECTOR"
      Set v = Vector(ArgList(1))
    Case "ABS"
      Set v = ABSFunc(ArgList(1))
    Case "ANG"
      Set v = Ang(ArgList(1))
    Case "ARC"
      Set v = Arc(ArgList(1))
    Case "COMB"
      Set v = ArgList(1)
      For i = 2 To nArg
        Set v = Concat2(v, ArgList(i))
      Next
    Case "MATRIX", "DET", "NMATRIX"
      If nArg > 2 Then
        If ArgList(1).ClassName = "text" And ArgList(2).ClassName = "text" Then
          Set t1 = ArgList(1)
          Set t2 = ArgList(2)
          If UCase(strTockenName) = "DET" Then
            Set v = Matrix(nArg, CLng(t1.Text), CLng(t2.Text), ArgList, True)
          ElseIf UCase(strTockenName) = "NMATRIX" Then
            Set v = Matrix(nArg, CLng(t1.Text), CLng(t2.Text), ArgList, False, False)
          Else
            Set v = Matrix(nArg, CLng(t1.Text), CLng(t2.Text), ArgList)
          End If
        Else
          Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
        End If
      Else
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
      End If
    Case "EQU"
      If nArg > 1 Then
        If ArgList(1).ClassName = "text" Then
          Set t1 = ArgList(1)
          Set v = Equation(nArg, CLng(t1.Text), ArgList)
        Else
          Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
        End If
      Else
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
      End If
    Case "LIM"
      If nArg >= 2 Then
          Set v = Lim(ArgList(1), ArgList(2))
      ElseIf nArg = 1 Then
          Set v = Lim(ArgList(1), Nothing)
      Else
          Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
      End If
    Case Else
      Set v1 = ArgList(1)
      For i = 2 To nArg
        Set v1 = Concat3(v1, Text(","), ArgList(i))
      Next
      Set v = Concat2(Text(strTockenName), Bracket(v1))
    End Select
    Set ParseFunction = v
End Function
'Bypassing current tocken
Private Sub SkipLastToken()
    mPosition = mPosition + mLastTokenLength
    SkipSpaces
End Sub
'Bypassin all spaces
Private Sub SkipSpaces()
    Do While mPosition <= Len(mExpression) And _
             (Mid(mExpression, mPosition, 1) = " " Or _
              Mid(mExpression, mPosition, 1) = vbTab Or Mid(mExpression, mPosition, 1) = vbCr Or Mid(mExpression, mPosition, 1) = vbLf)
        mPosition = mPosition + 1
    Loop
End Sub
Private Function IsAlpha(ByVal ch As String)
  Const ALPHASET = "&;:.[]!0123456789abcdefghijklmnopqrstuvwxyz"
  ch = LCase(ch)
  IsAlpha = ch <> "" And InStr(ALPHASET, ch) > 0
End Function
Private Function GetProjectName() As String
On Error Resume Next
    Err.Raise 1, , " "
    GetProjectName = Err.Source
    Err.Clear
End Function
Private Sub SetErrSource(Name As String)
    If Err.Source = mProjectName Then
        Err.Source = Name
    Else
        Err.Source = Name & "->" & Err.Source
    End If
End Sub
Public Property Get LastErrorPosition() As Long
    LastErrorPosition = mPosition
End Property
Public Sub Initialize()
    Dim i As Long
    ReDim mTokenSymbols(TOK_FIRST To TOK_LAST)
    mTokenSymbols(TOK_PM) = "+-"
    mTokenSymbols(TOK_ADD) = "+"
    mTokenSymbols(TOK_MP) = "-+"
    mTokenSymbols(TOK_SUBTRACT) = "-"
    mTokenSymbols(TOK_MULTIPLY) = "*"
    mTokenSymbols(TOK_EQUAL) = "="
    mTokenSymbols(TOK_DIVIDE) = "/"
    mTokenSymbols(TOK_POWER) = "^"
    mTokenSymbols(TOK_SUBSCRIPT) = "_"
    mTokenSymbols(TOK_OPEN_PARENTHESES) = "("
    mTokenSymbols(TOK_CLOSE_PARENTHESES) = ")"
    mTokenSymbols(TOK_LIST_SEPERATOR) = ","
    mTokenSymbols(TOK_DIFF) = "<>"
    mTokenSymbols(TOK_LESS) = "<"
    mTokenSymbols(TOK_LEQ) = "<="
    mTokenSymbols(TOK_GRE) = ">"
    mTokenSymbols(TOK_GEQ) = ">="
    mTokenSymbols(TOK_AND) = "&and;"
    mTokenSymbols(TOK_or) = "&or;"
    ReDim mMathSymbols(1 To MAX_SYMBOLS)
    For i = 1 To MAX_SYMBOLS
      mMathSymbols(i).nFontIndex = 1
    Next
    mMathSymbols(1).strEntity = "&alpha;"
    mMathSymbols(1).nCode = &H2C
        
    mMathSymbols(2).strEntity = "&beta;"
    mMathSymbols(2).nCode = &H2D
    
    mMathSymbols(3).strEntity = "&chi;"
    mMathSymbols(3).nCode = &H41
    
    mMathSymbols(4).strEntity = "&delta;"
    mMathSymbols(4).nCode = &H2F
    
    mMathSymbols(5).strEntity = "&epsilon;"
    mMathSymbols(5).nCode = &H30
    
    mMathSymbols(6).strEntity = "&phi;"
    mMathSymbols(6).nCode = &H40
    
    mMathSymbols(7).strEntity = "&phiv;"
    mMathSymbols(7).nCode = &H49
    
    mMathSymbols(8).strEntity = "&gamma;"
    mMathSymbols(8).nCode = &H2E
    
    mMathSymbols(9).strEntity = "&eta;"
    mMathSymbols(9).nCode = &H32
    
    mMathSymbols(10).strEntity = "&kappa;"
    mMathSymbols(10).nCode = &H35
    
    mMathSymbols(11).strEntity = "&lamda;"
    mMathSymbols(11).nCode = &H36
        
    mMathSymbols(12).strEntity = "&mu;"
    mMathSymbols(12).nCode = &H37
    
    mMathSymbols(13).strEntity = "&nu;"
    mMathSymbols(13).nCode = &H38
    
    mMathSymbols(14).strEntity = "&pi;"
    mMathSymbols(14).nCode = &H3B
    
    mMathSymbols(15).strEntity = "&piv;"
    mMathSymbols(15).nCode = &H46
    
    mMathSymbols(16).strEntity = "&theta;"
    mMathSymbols(16).nCode = &H33
    
    mMathSymbols(17).strEntity = "&rho;"
    mMathSymbols(17).nCode = &H3C
    
    mMathSymbols(18).strEntity = "&sigma;"
    mMathSymbols(18).nCode = &H3D
    
    mMathSymbols(19).strEntity = "&finalsigma;"
    mMathSymbols(19).nCode = &H48
    
    mMathSymbols(20).strEntity = "&tau;"
    mMathSymbols(20).nCode = &H3E
    
    mMathSymbols(21).strEntity = "&upsilon;"
    mMathSymbols(21).nCode = &H3F
        
    mMathSymbols(22).strEntity = "&omega;"
    mMathSymbols(22).nCode = &H43
    
    mMathSymbols(23).strEntity = "&xi;"
    mMathSymbols(23).nCode = &H39
    
    mMathSymbols(24).strEntity = "&pxi;"
    mMathSymbols(24).nCode = &H42
    
    mMathSymbols(25).strEntity = "&zeta;"
    mMathSymbols(25).nCode = &H31
    
    mMathSymbols(26).strEntity = "&omicron;"
    mMathSymbols(26).nCode = &H3A
    
    mMathSymbols(27).strEntity = "&Alpha;"
    mMathSymbols(27).nCode = &H63
    
    mMathSymbols(28).strEntity = "&Beta;"
    mMathSymbols(28).nCode = &H64
    
    mMathSymbols(29).strEntity = "&Chi;"
    mMathSymbols(29).nCode = &H7A
    
    mMathSymbols(30).strEntity = "&Delta;"
    mMathSymbols(30).nCode = &H22
    
    mMathSymbols(31).strEntity = "&Epsilon;"
    mMathSymbols(31).nCode = &H67
        
    mMathSymbols(32).strEntity = "&Phi;"
    mMathSymbols(32).nCode = &H29
    
    mMathSymbols(33).strEntity = "&Gamma;"
    mMathSymbols(33).nCode = &H21
    
    mMathSymbols(34).strEntity = "&Eta;"
    mMathSymbols(34).nCode = &H6A
    
    mMathSymbols(35).strEntity = "&Iota;"
    mMathSymbols(35).nCode = &H6B
    
    mMathSymbols(36).strEntity = "&Kappa;"
    mMathSymbols(36).nCode = &H6D
    
    mMathSymbols(37).strEntity = "&Lamda;"
    mMathSymbols(37).nCode = &H24
    
    mMathSymbols(38).strEntity = "&Mu;"
    mMathSymbols(38).nCode = &H6F
    
    mMathSymbols(39).strEntity = "&Nu;"
    mMathSymbols(39).nCode = &H70
    
    mMathSymbols(40).strEntity = "&Omicron;"
    mMathSymbols(40).nCode = &H71
    
    mMathSymbols(41).strEntity = "&Pi;"
    mMathSymbols(41).nCode = &H26
        
    mMathSymbols(42).strEntity = "&Theta;"
    mMathSymbols(42).nCode = &H23
    
    mMathSymbols(43).strEntity = "&Rho;"
    mMathSymbols(43).nCode = &H72
    
    mMathSymbols(44).strEntity = "&Sigma;"
    mMathSymbols(44).nCode = &H27
    
    mMathSymbols(45).strEntity = "&Tau;"
    mMathSymbols(45).nCode = &H76
    
    mMathSymbols(46).strEntity = "&Upsilon;"
    mMathSymbols(46).nCode = &H28
    
    mMathSymbols(47).strEntity = "&Omega;"
    mMathSymbols(47).nCode = &H2B
    
    mMathSymbols(48).strEntity = "&Xi;"
    mMathSymbols(48).nCode = &H25
    
    mMathSymbols(49).strEntity = "&Psi;"
    mMathSymbols(49).nCode = &H2A
    
    mMathSymbols(50).strEntity = "&Zeta;"
    mMathSymbols(50).nCode = &H7C
    
    mMathSymbols(51).strEntity = "&infinity;"
    mMathSymbols(51).nCode = &H54
    For i = 51 To 83
        mMathSymbols(i).nFontIndex = 2
    Next
    
    mMathSymbols(52).strEntity = "&and;"
    mMathSymbols(52).nCode = &H82
        
    mMathSymbols(53).strEntity = "&or;"
    mMathSymbols(53).nCode = &H83
        
    mMathSymbols(54).strEntity = "&larrow;"
    mMathSymbols(54).nCode = &H43
        
    mMathSymbols(55).strEntity = "&rarrow;"
    mMathSymbols(55).nCode = &H44
        
    mMathSymbols(56).strEntity = "&uarrow;"
    mMathSymbols(56).nCode = &H45
        
    mMathSymbols(57).strEntity = "&darrow;"
    mMathSymbols(57).nCode = &H46
        
    mMathSymbols(58).strEntity = "&lrarrow;"
    mMathSymbols(58).nCode = &H47
        
    mMathSymbols(59).strEntity = "&larrowd;"
    mMathSymbols(59).nCode = &H4B
        
    mMathSymbols(60).strEntity = "&rarrowd;"
    mMathSymbols(60).nCode = &H4C
        
    mMathSymbols(61).strEntity = "&uarrowd;"
    mMathSymbols(61).nCode = &H4D
        
    mMathSymbols(62).strEntity = "&darrowd;"
    mMathSymbols(62).nCode = &H4E
        
    mMathSymbols(63).strEntity = "&lrarrowd;"
    mMathSymbols(63).nCode = &H4F
        
    mMathSymbols(64).strEntity = "&urarrow;"
    mMathSymbols(64).nCode = &H48
        
    mMathSymbols(65).strEntity = "&drarrow;"
    mMathSymbols(65).nCode = &H49
        
    mMathSymbols(66).strEntity = "&perp;"
    mMathSymbols(66).nCode = &H62
        
    mMathSymbols(67).strEntity = "&parallel;"
    mMathSymbols(67).nCode = &H8F
        
    mMathSymbols(68).strEntity = "&empty;"
    mMathSymbols(68).nCode = &H5E
        
    mMathSymbols(69).strEntity = "&in;"
    mMathSymbols(69).nCode = &H55
        
    mMathSymbols(70).strEntity = "&any;"
    mMathSymbols(70).nCode = &H5B
        
    mMathSymbols(71).strEntity = "&exist;"
    mMathSymbols(71).nCode = &H5C
        
    mMathSymbols(72).strEntity = "&subset;"
    mMathSymbols(72).nCode = &H3D
    
    mMathSymbols(73).strEntity = "&supset;"
    mMathSymbols(73).nCode = &H3E
    
    mMathSymbols(74).strEntity = "&vdots;"
    mMathSymbols(74).nCode = &HB4
    
    mMathSymbols(75).strEntity = "&lor;"
    mMathSymbols(75).nCode = &H83
   
    mMathSymbols(76).strEntity = "&land;"
    mMathSymbols(76).nCode = &H82
    
    mMathSymbols(77).strEntity = "&approx;"
    mMathSymbols(77).nCode = &H3C
    
    mMathSymbols(78).strEntity = "&equiv;"
    mMathSymbols(78).nCode = &H34
    
    mMathSymbols(79).strEntity = "&propto;"
    mMathSymbols(79).nCode = &H52
    
    mMathSymbols(80).strEntity = "&sim;"
    mMathSymbols(80).nCode = &H3B
    
    mMathSymbols(81).strEntity = "&ll;"
    mMathSymbols(81).nCode = &H3F
    
    mMathSymbols(82).strEntity = "&gg;"
    mMathSymbols(82).nCode = &H40
    
    mMathSymbols(83).strEntity = "&simeq;"
    mMathSymbols(83).nCode = &H4A
    
    ReDim mMathOPs(1 To MAX_MATHOP)
    For i = 1 To MAX_MATHOP
      mMathOPs(i).nFontIndex = 2
    Next
    mMathOPs(1).strOP = "+"
    mMathOPs(1).nCode = &H21
    mMathOPs(2).strOP = "="
    mMathOPs(2).nCode = &H22
    mMathOPs(3).strOP = "-"
    mMathOPs(3).nCode = &H23
    mMathOPs(4).strOP = "*"
    mMathOPs(4).nCode = &H25
    
    mMathOPs(5).strOP = "<="
    mMathOPs(5).nCode = &HD4
    mMathOPs(5).nFontIndex = 1
    
    mMathOPs(6).strOP = ">="
    mMathOPs(6).nCode = &HD5
    mMathOPs(6).nFontIndex = 1
    
    mMathOPs(7).strOP = "+-"
    mMathOPs(7).nCode = &H29
    mMathOPs(8).strOP = "-+"
    mMathOPs(8).nCode = &H2A

    mMathOPs(9).strOP = "<"
    mMathOPs(9).nCode = &H5E
    mMathOPs(9).nFontIndex = 1

    mMathOPs(10).strOP = ">"
    mMathOPs(10).nCode = &H60
    mMathOPs(10).nFontIndex = 1

    mMathOPs(11).strOP = "<>"
    mMathOPs(11).nCode = &H22

    mProjectName = GetProjectName
End Sub
'***************************************************************************
'Helper function to create a concatation with 2 factors
'***************************************************************************
Private Function Concat2(ByVal p1 As clsBox, ByVal p2 As clsBox) As clsConcat
  Dim c As New clsConcat
  c.NumberElement = 2
  Set c.Part(1) = p1
  Set c.Part(2) = p2
  Set Concat2 = c
End Function
'****************************************************************************
'Helper function to create a concatation with 3 factors
'****************************************************************************
Private Function Concat3(ByVal p1 As clsBox, ByVal p2 As clsBox, ByVal p3 As clsBox) As clsConcat
  Dim c As New clsConcat
  c.NumberElement = 3
  Set c.Part(1) = p1
  Set c.Part(2) = p2
  Set c.Part(3) = p3
  Set Concat3 = c
End Function
'*****************************************************************************
'Helper function to create a concatation with 4 factors
'*****************************************************************************
Private Function Concat4(ByVal p1 As clsBox, ByVal p2 As clsBox, ByVal p3 As clsBox, ByVal p4 As clsBox) As clsConcat
  Dim c As New clsConcat
  c.NumberElement = 4
  Set c.Part(1) = p1
  Set c.Part(2) = p2
  Set c.Part(3) = p3
  Set c.Part(4) = p4
  Set Concat4 = c
End Function
'******************************************************************************
'Helper function to create an atom factor(text object)
'******************************************************************************
Private Function Text(ByVal st As String) As clsText
  Dim t As New clsText
  t.Text = st
  Set Text = t
End Function
'******************************************************************************
'Helper function to create a fraction
'******************************************************************************
Private Function Fraction(ByVal nu As clsBox, ByVal de As clsBox) As clsFraction
  Dim f As New clsFraction
  Set f.Numerator = nu
  Set f.Denominator = de
  Set Fraction = f
End Function
'*******************************************************************************
'Helper function to create an expression width brackets
'*******************************************************************************
Private Function Bracket(ByVal c As clsBox) As clsBracket
  Dim b As New clsBracket
  Set b.Content = c
  Set Bracket = b
End Function
'*********************************************************************************
'Helper function to create a over operator
'*********************************************************************************
Private Function Over(ByVal c As clsBox) As clsOver
  Dim o As New clsOver
  Set o.Content = c
  Set Over = o
End Function
'*********************************************************************************
'Helper function to create a power operator
'*********************************************************************************
Private Function Power(ByVal l As clsBox, ByVal r As clsBox) As clsPower
  Dim p As New clsPower
  Set p.Left = l
  Set p.Right = r
  Set Power = p
End Function
'***********************************************************************************
'Helper function to create a subscript expression
'***********************************************************************************
Private Function Subscript(ByVal b As clsBox, ByVal s As clsBox) As clsSubscript
  Dim ss As New clsSubscript
  Set ss.Base = b
  Set ss.Subscript = s
  Set Subscript = ss
End Function
'***********************************************************************************
'Helper function to create a root expression
'***********************************************************************************
Private Function Root(ByVal c As clsBox) As clsRoot
  Dim r As New clsRoot
  Set r.Content = c
  Set Root = r
End Function
'***********************************************************************************
'Helper function to create a sum expression
'***********************************************************************************
Private Function Sum(ByVal b As clsBox, ByVal t As clsBox, ByVal c As clsBox) As clsISPU
  Dim s As New clsISPU
  Set s.Upper = t
  Set s.Lower = b
  Set s.Content = c
  s.Sign = Chr(SUM_SIGN)
  Set Sum = s
End Function
'***********************************************************************************
'Helper function to create a Integral expression
'***********************************************************************************
Private Function InInt(ByVal b As clsBox, ByVal t As clsBox, ByVal c As clsBox) As clsISPU
  Dim i As New clsISPU
  Set i.Upper = t
  Set i.Lower = b
  Set i.Content = c
  i.Sign = Chr(INT_SIGN)
  Set InInt = i
End Function
'***********************************************************************************
'Helper function to create a product expression
'***********************************************************************************
Private Function Prod(ByVal b As clsBox, ByVal t As clsBox, ByVal c As clsBox) As clsISPU
  Dim i As New clsISPU
  Set i.Upper = t
  Set i.Lower = b
  Set i.Content = c
  i.Sign = Chr(PROD_SIGN)
  Set Prod = i
End Function
'***********************************************************************************
'Helper function to create a union expression
'***********************************************************************************
Private Function Uni(ByVal b As clsBox, ByVal t As clsBox, ByVal c As clsBox) As clsISPU
  Dim i As New clsISPU
  Set i.Upper = t
  Set i.Lower = b
  Set i.Content = c
  i.Sign = Chr(UNI_SIGN)
  Set Uni = i
End Function
'***********************************************************************************
'Helper function to create a matrix expression
'***********************************************************************************
Private Function Matrix(ByVal nMax As Long, ByVal m As Long, ByVal n As Long, ElmList() As clsBox, Optional ByVal bDet As Boolean = False, Optional ByVal bBracket As Boolean = True) As clsBox
  Dim mx As New clsMatrix
  Dim i As Long, j As Long, k As Long
  mx.CreateMatrix m, n
  mx.IsDet = bDet
  mx.IsBracket = bBracket
  For i = 1 To m
    For j = 1 To n
      k = (i - 1) * n + j + 2
      If k > nMax Then
        Set mx.Element(i, j) = Text(".")
      Else
        Set mx.Element(i, j) = ElmList((i - 1) * n + j + 2)
      End If
    Next
  Next
  Set Matrix = mx
End Function
'***********************************************************************************
'Helper function to create an equation expression
'***********************************************************************************
Private Function Equation(ByVal nMax As Long, ByVal n As Long, EQ() As clsBox) As clsEquations
  Dim e As New clsEquations
  Dim i As Long
  e.CreateEquation n
  For i = 1 To n
    If i + 1 > nMax Then
      Set e.Equation(i) = Text(" ")
    Else
      Set e.Equation(i) = EQ(i + 1)
    End If
  Next
  Set Equation = e
End Function
'***********************************************************************************
'Helper function to create a vector expression
'***********************************************************************************
Private Function Vector(box As clsBox) As clsOver
  Dim o As New clsOver
  o.OverStyle = OVER_VECTOR
  Set o.Content = box
  Set Vector = o
End Function
'***********************************************************************************
'Helper function to create a abs expression
'***********************************************************************************
Private Function ABSFunc(box As clsBox) As clsOver
  Dim o As New clsOver
  o.OverStyle = OVER_ABS
  Set o.Content = box
  Set ABSFunc = o
End Function
'***********************************************************************************
'Helper function to create a angle sign
'***********************************************************************************
Private Function Ang(box As clsBox) As clsOver
  Dim o As New clsOver
  o.OverStyle = OVER_ANGLE
  Set o.Content = box
  Set Ang = o
End Function
'***********************************************************************************
'Helper function to create a arc sign
'***********************************************************************************
Private Function Arc(box As clsBox) As clsOver
  Dim o As New clsOver
  o.OverStyle = OVER_CHORD
  Set o.Content = box
  Set Arc = o
End Function
'***********************************************************************************
'Helper function to create a limited expression
'***********************************************************************************
Private Function Lim(box1 As clsBox, box2 As clsBox) As clsLim
  Dim l As New clsLim
  Set l.Content = box1
  Set l.Under = box2
  Set Lim = l
End Function
'***********************************************************************************
'Helper function to draw a line
'Input:
'       X1, X2, Y2, Y2: Where to draw
'       hDC: DC to draw
'***********************************************************************************
Public Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hDC As Long)
  Dim pt As POINTAPI
  MoveToEx hDC, X1, Y1, pt
  LineTo hDC, X2, Y2
End Sub
'***********************************************************************************
'Helper function to draw a circle
'Input:
'       X, X: center of circle
'       xs, ys, xe, ye: start and end position
'       hDC: DC to draw
'***********************************************************************************

Public Sub DrawCircle(ByVal x As Long, ByVal y As Long, ByVal r As Long, ByVal xs As Long, ByVal ys As Long, ByVal xe As Long, ye As Long, hDC As Long)
  'Arc hDC, x - r, y - r, x + r, y + r, xs, ys, xe, ye
End Sub
'***********************************************************************************
'Helper function to draw a line of text
'Input:
'       X, Y: Where to draw
'       st: what to draw
'       hDC: DC to draw
'***********************************************************************************

Public Sub DrawText(ByVal x As Long, ByVal y As Long, ByVal str As String, ByVal hDC As Long)
  TextOut hDC, x, y, str, Len(str)
End Sub
'***********************************************************************************
'Helper function to get ascent of text
'Input:
'       hDC: DC to draw
'***********************************************************************************
Public Function TextAscent(hDC As Long) As Long
  Dim tm As TEXTMETRIC
  GetTextMetrics hDC, tm
  TextAscent = tm.tmAscent
End Function
'***********************************************************************************
'Helper function to get descent of text
'Input:
'       hDC: DC to draw
'***********************************************************************************
Public Function TextDecent(hDC As Long) As Long
  Dim tm As TEXTMETRIC
  GetTextMetrics hDC, tm
  TextDecent = tm.tmDescent
End Function
'***********************************************************************************
'Helper function to get internal leading
'Input:
'       hDC: DC to draw
'***********************************************************************************
Public Function TextInternalLeading(hDC As Long) As Long
  Dim tm As TEXTMETRIC
  GetTextMetrics hDC, tm
  TextInternalLeading = tm.tmInternalLeading
End Function
'***********************************************************************************
'Helper function to get text width
'Input:
'       str: What to get
'       hDC: DC to draw
'***********************************************************************************
Public Function TextWidth(ByVal str As String, hDC As Long) As Long
  Dim sz As Size
  GetTextExtentPoint hDC, str, Len(str), sz
  TextWidth = sz.cx
End Function
'***********************************************************************************
'Helper function to get text height
'Input:
'       str: What to get
'       hDC: DC to draw
'***********************************************************************************
Public Function TextHeight(ByVal str As String, hDC As Long) As Long
  Dim sz As Size
  GetTextExtentPoint hDC, str, Len(str), sz
  TextHeight = sz.cy
End Function
'***********************************************************************************
'Helper function to get font from DC
'Input:
'       hDC: DC to draw
'***********************************************************************************
Public Function GetFont(hDC As Long) As Long
   Dim hFont As Long
   hFont = GetCurrentObject(hDC, OBJ_FONT)
   GetFont = hFont
End Function
'***********************************************************************************
'Helper function to LOGFONT from hFont
'Input:
'       hFont:
'***********************************************************************************
Public Function GetLogFont(hFont As Long) As LOGFONT
   Dim lf As LOGFONT
   GetObject hFont, Len(lf), lf
   GetLogFont = lf
End Function
'***********************************************************************************
'Helper function to set font for DC
'Input:
'       hFont: Font to set
'       hDC: DC handle
'***********************************************************************************
Public Function SetFont(hFont As Long, hDC As Long) As Long
   Dim hOldFont As Long
   hOldFont = SelectObject(hDC, hFont)
   SetFont = hOldFont
End Function
'***********************************************************************************
'Helper function to get width of plus sign
'Input:
'       FontSize:
'       hDC:
'***********************************************************************************
Public Function OperatorWidth(ByVal FontSize As Long, hDC As Long) As Long
  Dim hOldFont As Long
  Dim hFont As Long
  Dim lf As LOGFONT
  hOldFont = GetFont(hDC)
  lf = GetLogFont(hOldFont)
  lf.lfHeight = FontSize
  hFont = CreateFontIndirect(lf)
  hOldFont = SetFont(hFont, hDC)
  OperatorWidth = modHelper.TextWidth("+", hDC)
  SelectObject hDC, hOldFont
  DeleteObject hFont
End Function
'***********************************************************************************
'Helper function to create font size from point
'Input:
'       nPoint:
'       hDC: DC to draw
'***********************************************************************************
Public Function GetFontSize(ByVal nPoint As Long, hDC As Long)
  GetFontSize = -MulDiv(nPoint, GetDeviceCaps(hDC, LOGPIXELSY), 72)
End Function

Public Function GetFreeExtra(ByVal nFontSize As Long, hDC As Long) As Long
  Dim hOldFont As Long
  Dim hNewFont As Long
  Dim lf As LOGFONT
  lf = GetLogFont(GetFont(hDC))
  lf.lfHeight = nFontSize
  lf.lfCharSet = SYMBOL_CHARSET
  lf.lfFaceName = MATH_SYMBOL_FONT & Chr(0)
  hNewFont = CreateFontIndirect(lf)
  hOldFont = SetFont(hNewFont, hDC)
  GetFreeExtra = TextAscent(hDC) - TextHeight("X", hDC) / 2
  SelectObject hDC, hOldFont
  DeleteObject hNewFont
End Function

Public Function FindCharCode(ByVal str As String) As MATH_SYMBOL
  Dim i As Long
  Dim ms As MATH_SYMBOL
  ms.nCode = 0
  ms.strEntity = ""
  ms.nFontIndex = 0
  For i = 1 To MAX_SYMBOLS
    If mMathSymbols(i).strEntity = str Then
       FindCharCode = mMathSymbols(i)
       Exit Function
    End If
  Next
  FindCharCode = ms
End Function

Public Function FindOPCode(ByVal str As String) As MATH_OP
  Dim i As Long
  Dim mo As MATH_OP
  mo.nCode = 0
  mo.nFontIndex = 0
  mo.strOP = ""
  For i = 1 To MAX_MATHOP
    If mMathOPs(i).strOP = str Then
      FindOPCode = mMathOPs(i)
      Exit Function
    End If
  Next
  FindOPCode = mo
End Function
