Attribute VB_Name = "Easter"
Option Explicit

Public Function EasterMonday(YYYY As Long) As Long

Dim C As Long
Dim N As Long
Dim K As Long
Dim i As Long
Dim j As Long
Dim L As Long
Dim M As Long
Dim D As Long
    
    C = YYYY \ 100
    N = YYYY - 19 * (YYYY \ 19)
    K = (C - 17) \ 25
    i = C - C \ 4 - (C - K) \ 3 + 19 * N + 15
    i = i - 30 * (i \ 30)
    i = i - (i \ 28) * (1 - (i \ 28) * (29 \ (i + 1)) * ((21 - N) \ 11))
    j = YYYY + YYYY \ 4 + i + 2 - C + C \ 4
    j = j - 7 * (j \ 7)
    L = i - j
    M = 3 + (L + 40) \ 44
    D = L + 28 - 31 * (M \ 4)
    D = D + 1
    EasterMonday = DateSerial(YYYY, M, D)
    
End Function
