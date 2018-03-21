VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinearSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GaussSeidel(ByVal A As Variant, ByVal b As Variant, ByRef n As Long, ByRef Tolerance As Double)

     Dim LB As Long
     Dim UB As Long
     Dim x As Variant
     Dim xOld As Variant
     Dim lng As Long
     Dim i As Long
     Dim j As Long
     Dim db As Double
     Dim dA As Double
     Dim dx As Variant
     Dim sum As Double
     
     On Error GoTo Catch
     
Try:

     If CheckDimensions(A, b) Then
          LB = LBound(b, 1)
          UB = UBound(b, 1)
          ReDim x(LB To UB, 1 To 1) As Variant
          ReDim dx(LB To UB, 1 To 1) As Variant
          For lng = 1 To n
               xOld = Clone(x)
               sum = 0#
               For i = LB To UB
                    db = b(i, 1)
                    dA = A(i, i)
                    j = LB
                    Do While j < i
                         db = db - A(i, j) * x(j, 1)
                         j = j + 1
                    Loop
                    j = i + 1
                    Do While j <= UB
                         db = db - A(i, j) * xOld(j, 1)
                         j = j + 1
                    Loop
                    x(i, 1) = db / dA
               Next i
               For j = LB To UB
                    dx(j, 1) = x(j, 1) - xOld(j, 1)
                    sum = sum + dx(j, 1) * dx(j, 1)
               Next j
               If Sqr(sum) < Tolerance Then Exit For
          Next lng
     Else
          Call MsgBox("Check dimensions.", vbCritical, "Exception")
     End If

Finally:

     GaussSeidel = x
     Exit Function

Catch:

     Stop: Resume

End Function

Private Function Clone(ByRef x As Variant) As Variant

     Dim i As Long
     Dim LB As Long: LB = LBound(x, 1)
     Dim UB As Long: UB = UBound(x, 1)
     Dim C() As Variant
     ReDim C(LB To UB, 1 To 1) As Variant

     For i = LBound(x, 1) To UBound(x, 1)
          C(i, 1) = x(i, 1)
     Next i
     
     Clone = C
     
End Function

Public Function GaussJordan(ByVal A As Variant, ByVal b As Variant) As Variant

     Dim LB As Long
     Dim UB As Long
     Dim i As Long
     Dim d As Double
     Dim x() As Variant
          
     On Error GoTo Catch
     
Try:

     If CheckDimensions(A, b) Then
          LB = LBound(b, 1)
          UB = UBound(b, 1)
          ReDim x(LB To UB, 1 To 1) As Variant
          Call Triangulate(A, b)
          For i = UB To LB Step -1
               d = A(i, i)
               x(i, 1) = (b(i, 1) - DotProduct(A, x, i)) / d
          Next i
     Else
          Call MsgBox("Check dimensions.", vbCritical, "Exception")
     End If
     
Finally:

     GaussJordan = x
     Exit Function
     
Catch:

     Stop: Resume

End Function

Private Function CheckDimensions(ByRef A As Variant, ByRef b As Variant) As Boolean
     
     Dim bln As Boolean: bln = True
     
     On Error GoTo Catch

Try:

     If LBound(A, 2) <> LBound(b, 1) Then bln = False
     If UBound(A, 2) <> UBound(b, 1) Then bln = False
     
Finally:

     CheckDimensions = bln
     Exit Function
     
Catch:

     bln = False
     Resume Next
     
End Function

Private Sub Triangulate(ByRef A As Variant, ByRef b As Variant)

     Dim d As Double
     Dim dd As Double
     Dim LB As Long: LB = LBound(A, 1)
     Dim UB As Long: UB = UBound(A, 1)
     Dim i As Long
     Dim j As Long
     Dim k As Long

     On Error GoTo Catch
     
Try:

     For i = LB To UB - 1
          d = Pivot(A, b, i)
          For j = i + 1 To UB
               dd = A(j, i) / d
               For k = i + 1 To UB
                    A(j, k) = A(j, k) - dd * A(i, k)
               Next k
               b(j, 1) = b(j, 1) - dd * b(i, 1)
          Next j
     Next i

Finally:

     Exit Sub
     
Catch:

     Stop: Resume

End Sub

Private Function Pivot(ByRef A As Variant, ByRef b As Variant, ByRef q As Long) As Double

     Dim n As Long: n = UBound(b, 1) - LBound(b, 1) + 1
     Dim j As Long
     Dim d As Double: d = 0#
     Dim dd As Double
     Dim i As Long: i = q

     On Error GoTo Catch
     
Try:

     For j = q To n
          dd = Abs(A(j, q))
          If dd > d Then
               d = dd
               i = j
          End If
     Next j
     
     If i > q Then
          Call SwapRow(A, q, i)
          Call SwapRow(b, q, i)
     End If

Finally:

     Pivot = A(q, q)
     Exit Function
     
Catch:

     Stop: Resume

End Function

Private Sub SwapRow(ByRef A As Variant, ByRef q As Long, ByRef i As Long)

     Dim temp As Double: temp = 0#
     Dim j As Long

     On Error GoTo Catch
     
Try:

     For j = LBound(A, 2) To UBound(A, 2)
     
          temp = A(q, j)
          A(q, j) = A(i, j)
          A(i, j) = temp
     
     Next j

Finally:

     Exit Sub

Catch:

     Stop: Resume

End Sub

Private Function DotProduct(ByRef A As Variant, ByRef b As Variant, ByRef i As Long) As Double

     Dim j As Long
     Dim sum As Double: sum = 0#

     For j = LBound(A, 2) To UBound(A, 2)
          
          sum = sum + A(i, j) * b(j, 1)
          
     Next j
     
     DotProduct = sum

End Function

