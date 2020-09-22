Attribute VB_Name = "Functions"
Public Function CircularCollision(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, R1 As Long, Optional R2 As Long = 0) As Boolean
CircularCollision = (Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2))) < R1 + R2
End Function

