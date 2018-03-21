VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private varA As Variant
Private varL As Variant
Private varU As Variant
Private intLB As Integer
Private intUB As Integer

Public Property Let A(ByRef RHS As Variant)

     On Error GoTo Catch
     
Try:

     'check that matrix is square
     If LBound(RHS, 1) = LBound(RHS, 2) And UBound(RHS, 1) = UBound(RHS, 2) Then
          
          intLB = LBound(RHS, 1)
          intUB = UBound(RHS, 1)
          varA = RHS
     
     Else
          
          Call MsgBox("Matrix is not square.", vbCritical, "Cannot Factorize")
     
     End If

Finally:

     Exit Property

Catch:
     
     Stop: Resume

End Property

Public Property Get L() As Variant
     
     L = varL

End Property

Public Property Get U() As Variant
     
     U = varU

End Property

Public Property Get C() As Variant
     
     C = varL

End Property

Private Sub InitialiseFactorArrays()
     
     Dim i As Integer
     Dim j As Integer
     
     On Error GoTo Catch

Try:
     
     ReDim varL(intLB To intUB, intLB To intUB) As Double
     ReDim varU(intLB To intUB, intLB To intUB) As Double

     '// initialise arrays
     For i = LBound(varA, 1) To UBound(varA, 1)
          
          For j = LBound(varA, 2) To UBound(varA, 2)
               
               varL(i, j) = 0#
               varU(i, j) = varA(i, j)
          
          Next j
     
     Next i

Finally:

     Exit Sub

Catch:
     
     Stop: Resume

End Sub

Public Sub LU()
     
     Dim i As Integer
     Dim j As Integer
     Dim k As Integer
     
     On Error GoTo Catch
     
Try:

     Call InitialiseFactorArrays
     
     For j = intLB To intUB
          
          For i = j To intUB
               
               If i = j Then
                    varL(i, j) = 1
               Else
                    If varU(j, j) = 0# Then Call Err.Raise(vbObjectError + 1, "FactorLU.Decompose", "Pivot is zero.  Cannot continue.")
                    varL(i, j) = varU(i, j) / varU(j, j)
                    For k = j To intUB
                         varU(i, k) = varU(i, k) - varL(i, j) * varU(j, k)
                    Next k
               End If
          
          Next i
     
     Next j
     
Finally:

     Exit Sub

Catch:
     
     Stop: Resume

End Sub

Public Sub Cholesky()
     
     Dim i As Integer
     Dim j As Integer
     Dim k As Integer
     Dim dblSum As Double
     
     On Error GoTo Catch
     
Try:
     
     Call InitialiseFactorArrays

     If Not CheckSymmetry Then
     
          Call MsgBox("Matrix is not symmetric.", vbCritical, "Can't Determine Cholesky Factors")
     
     Else
               
          For j = intLB To intUB
               
               For i = j To intUB
                    
                    If i < j Then varL(i, j) = 0
                    If i = j Then
                         dblSum = 0
                         k = 1
                         Do While k <= j - 1
                              dblSum = dblSum + varL(i, k) ^ 2
                              k = k + 1
                         Loop
                         varL(i, j) = Sqr(varA(i, j) - dblSum)
                    End If
                    If i > j Then
                         dblSum = 0
                         k = 1
                         Do While k <= j - 1
                              dblSum = dblSum + varL(j, k) * varL(i, k)
                              k = k + 1
                         Loop
                         varL(i, j) = (varA(i, j) - dblSum) / varL(j, j)
                    End If
               
               Next i
          
          Next j
     
     End If
     
Finally:

     Exit Sub

Catch:
     
     If varA(i, j) - dblSum < 0 Then
          Call MsgBox("Matrix is not Positive Definite.", vbCritical, "Can't Determine Cholesky Factors")
          GoTo Finally
     End If

     Stop: Resume

End Sub

Private Function CheckSymmetry() As Boolean

     Dim bln As Boolean: bln = True
     Dim i As Integer, j As Integer
     
     For i = intLB To intUB
          For j = intLB To intUB
               If varA(i, j) <> varA(j, i) Then
                    bln = False
                    GoTo Finally
               End If
          Next j
     Next i
     
Finally:

     CheckSymmetry = bln

End Function
