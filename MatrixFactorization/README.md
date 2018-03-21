Used to decompose a square matrix (a range object), "A", into factors using:

1. LU Decomposition; or
2. Cholesky Decomposition.

Returns an array, being the relevant factor.  
For Cholesky, the lower triangle is returned.

Usage:

Dim fac As Factorization  
Dim f_L As Variant  
Dim f_U As Variant  
Dim f_C As Variant  

set fac = New Factorization  
fac.A = A  

fac.LU  
f_L = fac.L  
f_U = fac.U  

fac.C  
f_C = fac.C  
