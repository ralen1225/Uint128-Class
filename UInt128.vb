Option Explicit On
Option Strict On
' Module:   128-bit Unsigned Integer Class
' Author:   James Merrill

Imports IP_SummaryV2

Public Class UInt128
  ' High- and Low-order QWords
  Private _hi As ULong
  Private _lo As ULong

  'Access to _hi and _lo
  Public Property Hi As ULong
    Get
      Return _hi
    End Get
    Set(value As ULong)
      _hi = value
    End Set
  End Property
  Public Property Lo As ULong
    Get
      Return _lo
    End Get
    Set(value As ULong)
      _lo = value
    End Set
  End Property

  ' Default constructor
  Public Sub New()
    Hi = 0
    Lo = 0
  End Sub
  ' Combine high-order and low-order parts
  Public Sub New(ByVal argHi As ULong, ByVal argLo As ULong)
    ' save values internally
    Hi = argHi
    Lo = argLo
  End Sub
  ' Copy values from existing UInt128
  Public Sub New(ByVal arg128 As UInt128)
    ' Take values from arg and save internally
    Hi = arg128.Hi
    Lo = arg128.Lo
  End Sub

  Private Property Bit(Position As Integer) As Integer
    Get
      If position > 63 Then
        Return CInt((Hi And 1UL << (position - 64)) >> (position - 64))
      Else
        Return CInt((Lo And 1UL << position) >> position)
      End If
    End Get
    Set(value As Integer)
      If position > 63 Then
        Hi = Hi Or (1UL << (position - 64))
      Else
        Lo = Lo Or (1UL << position)
      End If
    End Set
  End Property

  ' Widen all Int/UInt types to UInt128. Reduces the number of overloads required for operators
  Public Shared Widening Operator CType(ByVal argULng As ULong) As UInt128
    Return New UInt128(0, argULng)
  End Operator

  ' Bit-wise And operator
  Public Shared Operator And(ByVal argLeft As UInt128, ByVal argRight As UInt128) As UInt128
    ' Perform And on each section pair and return result
    Return New UInt128(argLeft.Hi And argRight.Hi, argLeft.Lo And argRight.Lo)
  End Operator

  ' Bit-wise Or operator
  Public Shared Operator Or(ByVal argLeft As UInt128, ByVal argRight As UInt128) As UInt128
    ' Perform Xor on each section pair and return result
    Return New UInt128(argLeft.Hi Or argRight.Hi, argLeft.Lo Or argRight.Lo)
  End Operator

  ' Bit-wise Not operator
  Public Shared Operator Not(ByVal argLeft As UInt128) As UInt128
    ' Perform Not on each section and return result
    Return New UInt128(Not argLeft.Hi, Not argLeft.Lo)
  End Operator

  ' Bit-wise Xor operator
  Public Shared Operator Xor(ByVal argLeft As UInt128, ByVal argRight As UInt128) As UInt128
    ' Perform Xor on each section pair and return result
    Return New UInt128(argLeft.Hi Xor argRight.Hi, argLeft.Lo Xor argRight.Lo)
  End Operator

  ' Return Additive inverse of arguement
  Public Shared Operator -(ByVal arg128 As UInt128) As UInt128
    Return (Not arg128) + 1
  End Operator

  ' Return Multiplicative inverse of the given number using unary + operator
  ' This is not the traditional use of unary +
  ' Number MUST be odd (shift right until it is odd and record how many shifts it took
  ' to get there).
  ' This will make use of the fact that the * operator (above) drops overflow values.
  Public Shared Operator +(ByVal arg128 As UInt128) As UInt128
    Dim uxlTest, uxlNext As UInt128
    uxlNext = arg128
    uxlTest = uxlNext * arg128
    Do Until uxlTest = 1
      uxlNext = uxlNext * (2 - uxlTest)
      uxlTest = uxlTest * arg128
    Loop
    Return uxlNext
  End Operator

  ' Subtraction - Take additive inverse and pass to addition
  ' Uses the fact that addition drops overflow
  Public Shared Operator -(ByVal argLeft As UInt128, ByVal argRight As UInt128) As UInt128
    Return argLeft + (-argRight)
  End Operator

  ' Addition - Overflow is dropped
  Public Shared Operator +(ByVal argLeft As UInt128, ByVal argRight As UInt128) As UInt128
    Dim uxlResult As UInt128 = 0
    Dim blnCarry As Boolean = False
    If argLeft.Lo > ULong.MaxValue - argRight.Lo Then
      uxlResult.Lo = argLeft.Lo - (ULong.MaxValue - argRight.Lo + 1UL)
      blnCarry = True
    Else
      uxlResult.Lo = argLeft.Lo + argRight.Lo
    End If
    If argLeft.Hi > ULong.MaxValue - argRight.Hi Then
      uxlResult.Hi = argLeft.Hi - (ULong.MaxValue - argRight.Hi + 1UL)
    Else
      uxlResult.Hi = argLeft.Hi + argRight.Hi
    End If
    uxlResult.Hi += CULng(IIf(blnCarry, 1, 0))
    Return uxlResult
  End Operator

  ' Multiply - Overflow is dropped
  Public Shared Operator *(ByVal argLeft As UInt128, ByVal argRight As UInt128) As UInt128
    ' Split into 4 parts
    Dim intLeftParts() As UInteger = {CUInt(argLeft.Lo And &HFFFFFFFFUL), CUInt(argLeft.Lo >> 32), CUInt(argLeft.Hi And &HFFFFFFFFUL), CUInt(argLeft.Hi >> 32)}
    Dim intRightParts() As UInteger = {CUInt(argRight.Lo And &HFFFFFFFFUL), CUInt(argRight.Lo >> 32), CUInt(argRight.Hi And &HFFFFFFFFUL), CUInt(argRight.Hi >> 32)}
    ' Result registers - Use 8 to avoid runtime errors
    Dim lngResults(7) As ULong
    For i = 0 To 3 ' Cycle through Right arg parts
      For j = 0 To 3 ' Cycle through Left arg parts
        lngResults(i + j) += intRightParts(i) * intLeftParts(j)
        For k = i + j To 6 ' Move overflow into next one up
          lngResults(k + 1) += lngResults(k) >> 32
          lngResults(k) = lngResults(k) And &HFFFFFFFFUL
        Next ' k
      Next ' j 
    Next ' i 
    ' Put result together and return it - Overflow is dropped
    Return New UInt128(lngResults(3) << 32 Or lngResults(2), lngResults(1) << 32 Or lngResults(0))
  End Operator

  ' Division - Shift algorithm
  ' Adapted from https://en.wikipedia.org/wiki/Division_algorithm
  Public Shared Operator \(ByVal argLeft As UInt128, ByVal argRight As UInt128) As UInt128
    If argRight = 0 OrElse argRight > argLeft Then Return 0
    Dim uxlQuotient As UInt128 = 0
    Dim uxlRemainder As UInt128 = 0
    Dim intBits As Integer = argLeft.BitsUsed
    For i = intBits - 1 To 0 Step -1
      uxlRemainder = uxlRemainder << 1
      uxlRemainder.Bit(0) = argLeft.Bit(i)
      If uxlRemainder >= argRight Then
        uxlRemainder -= argRight
        uxlQuotient.Bit(i) = 1
      End If
    Next
    Return uxlQuotient
  End Operator

  Public Function BitsUsed() As Integer
    Dim lngTest As ULong
    Dim blnHi As Boolean = True
    Dim intReturn As Integer = 0
    If Hi = 0 Then
      lngTest = Lo
      blnHi = False
    Else
      lngTest = Hi
    End If
    Do Until intReturn > 63 OrElse lngTest < (1UL << intReturn)
      intReturn += 1
    Loop
    If blnHi Then intReturn += 64
    Return intReturn
  End Function

  ' Mod operator
  Public Shared Operator Mod(ByVal argLeft As UInt128, ByVal argRight As UInt128) As UInt128
    ' a - (b * (a \ b))
    Return argLeft - ((argLeft \ argRight) * argLeft)
  End Operator

  ' Exponent
  Public Shared Operator ^(ArgLeft As UInt128, argRight As Integer) As UInt128
    Dim uxlExponent As UInt128 = 1
    For i = 1 To argRight
      uxlExponent *= ArgLeft
    Next
    Return uxlExponent
  End Operator

  ' Shift right
  Public Shared Operator >>(ByVal argLeft As UInt128, ByVal argRight As Integer) As UInt128
    ' Negative shift? Shift left instead
    If argRight < 0 Then
      Return argLeft << -argRight
    ElseIf argRight <= 64 Then
      ' Shift bits into the low-order QWord
      Return New UInt128(argLeft.Hi >> argRight, (argLeft.Lo >> argRight) Or (argLeft.Hi << (64 - argRight)))
    Else
      ' High-order QWord is zeroed and remainder of shift moves its bits into the Low-order QWord
      Return New UInt128(0, argLeft.Hi >> (argRight - 64))
    End If
  End Operator

  ' Shift left
  Public Shared Operator <<(ByVal argLeft As UInt128, ByVal argRight As Integer) As UInt128
    ' Negative shift? Shift right instead
    If argRight < 0 Then
      Return argLeft >> -argRight
    ElseIf argRight <= 64 Then
      ' Shift bits into high-or
      Return New UInt128((argLeft.Hi << argRight) Or (argLeft.Lo >> (64 - argRight)), argLeft.Lo << argRight)
    Else
      Return New UInt128(argLeft.Lo << argRight - 64, 0)
    End If
  End Operator

  ' Comparisons
  ' Equality
  Public Shared Operator =(ByVal argLeft As UInt128, ByVal argRight As UInt128) As Boolean
    Return argLeft.Hi = argRight.Hi AndAlso argLeft.Lo = argRight.Lo
  End Operator
  ' Inequality
  Public Shared Operator <>(ByVal argLeft As UInt128, ByVal argRight As UInt128) As Boolean
    Return Not (argLeft = argRight)
  End Operator
  ' Less than
  Public Shared Operator <(ByVal argLeft As UInt128, ByVal argRight As UInt128) As Boolean
    Return argLeft.Hi < argRight.Hi OrElse (argLeft.Hi = argRight.Hi AndAlso argLeft.Lo < argRight.Lo)
  End Operator
  ' Greater than
  Public Shared Operator >(ByVal argLeft As UInt128, ByVal argRight As UInt128) As Boolean
    Return argLeft.Hi > argRight.Hi OrElse (argLeft.Hi = argRight.Hi AndAlso argLeft.Lo > argRight.Lo)
  End Operator
  ' Greater than or Equal
  Public Shared Operator >=(ByVal argLeft As UInt128, ByVal argRight As UInt128) As Boolean
    Return Not (argLeft < argRight)
  End Operator
  ' Less than or Equal
  Public Shared Operator <=(ByVal argLeft As UInt128, ByVal argRight As UInt128) As Boolean
    Return Not (argLeft > argRight)
  End Operator
  ' IsTrue and IsFalse - used for "-Also" shortcutting
  Public Shared Operator IsFalse(ByVal arg128 As UInt128) As Boolean
    Return arg128 = 0
  End Operator
  Public Shared Operator IsTrue(ByVal arg128 As UInt128) As Boolean
    Return (Not arg128) = 0 ' less work to invert arg than to create new FFFF....
  End Operator
End Class

Public Class UDLong
  Inherits UInt128

  ' Combine high-order and low-order parts
  Public Sub New(ByVal argHi As ULong, ByVal argLo As ULong)
    MyBase.New(argHi, argLo)
  End Sub

  ' Copy values from existing UInt128
  Public Sub New(ByVal arg128 As UInt128)
    MyBase.New(arg128)
  End Sub
End Class
