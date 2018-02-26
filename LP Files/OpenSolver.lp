\ Model solved using the solver: CBC
\ Model for sheet Test3
\ Model has 8  Excel constraints giving 8  constraint rows and 4  variables.
MINIMIZE
Obj: +2 C4 +3 D4 +7.12300000000005 _E4 +4 F4
SUBJECT TO
\ H8 <= J8
 +10 C4 +11 D4 +12 F4 <= +11
\ H9 <= J9
 +21 C4 +22 D4 +22.9999999999999 F4 <= +150
\ H10 >= J10
 +22 D4 +100 _E4 -23 F4 >= +110
\ H11 <= J11
 -32 F4 <= +180
\ R3:R4 <= T3:T4
\ (A row with all zero coeffs) <= +44
\ (A row with all zero coeffs) <= +1.4210854715202E-14
\ P5:Q5 >= P7:Q7
\ (A row with all zero coeffs) >= +0
\ (A row with all zero coeffs) >= +0

BOUNDS

GENERAL
_E4
BINARY
D4

END
