\ http://lpsolve.sourceforge.net/5.5/CPLEX-format.htm with modifications

Maximize
 obj: x1 + 2 x2 + 3 x3 + x4
Subject To
 c1: - x1 + x2 + x3 + 10 x4 <= 20
 c2: x1 - 3 x2 + x3 <= 30
 c3: x2 - 3.5 x4 = 0
Bounds
-inf <= x1 <= 40
 2 <= x4 <= 3
 x3 <= 4
SOS
set1: S2 :: x4 : 1 x5 : 2
S1 :: x3:6
End