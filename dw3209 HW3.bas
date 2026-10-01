Attribute VB_Name = "Module1"
Option Base 0
dw3209


Function getBondPrice(y As Double, face As Double, _
                      couponRate As Double, m As Long, _
                      Optional ppy As Long = 1) As Double

    Dim t As Long
    Dim n As Long
    Dim coupon As Double
    Dim cf As Double
    Dim price As Double

    n = m * ppy
    coupon = face * couponRate / ppy
    price = 0

    For t = 1 To n
        cf = coupon

        If t = n Then
            cf = cf + face
        End If

        price = price + cf / (1 + y / ppy) ^ t
    Next t

    getBondPrice = price

End Function


Function getBondDuration(y As Double, face As Double, _
                         couponRate As Double, m As Long) As Double

    Dim t As Long
    Dim cf As Double
    Dim pvcf As Double
    Dim price As Double
    Dim weightedTime As Double

    price = 0
    weightedTime = 0

    For t = 1 To m
        cf = face * couponRate

        If t = m Then
            cf = cf + face
        End If

        pvcf = cf / (1 + y) ^ t

        price = price + pvcf
        weightedTime = weightedTime + t * pvcf
    Next t

    getBondDuration = weightedTime / price

End Function


Sub TestBondFunctions()

    Dim y As Double
    Dim face As Double
    Dim couponRate As Double
    Dim m As Long

    y = 0.03
    face = 2000000
    couponRate = 0.04
    m = 10

    Debug.Print Round(getBondPrice(y, face, couponRate, m, 1), 0)
    Debug.Print Round(getBondPrice(y, face, couponRate, m, 2), 0)
    Debug.Print Round(getBondDuration(y, face, couponRate, m), 2)

End Sub

