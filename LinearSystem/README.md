Used to solve the linear algebra equation A.x = b using either:

1. Gauss-Jordan Elimination; or
2. Gauss-Seidel Elimination

ie. A is not inverted when solving for x.

A must be a square matrix.

Usage:

Dim ls as LinearSystem
Dim x_GJ as Variant
Dim x_GS as Variant

set ls = New LinearSystem
x_GJ = ls.GaussJordan(A, b)
x_GS = ls.GaussSeidel(A, b, 10, 0.0001)
