Attribute VB_Name = "Greeks"
Function d1Calc(S As Double, K As Double, T As Double, r As Double, sigma As Double) As Double
    d1Calc = (Log(S / K) + (r + 0.5 * sigma ^ 2) * T) / sigma * Sqr(T)
End Function

Function d2Calc(S As Double, K As Double, T As Double, r As Double, sigma As Double) As Double
    d2Calc = d1Calc(S, K, T, r, sigma) - sigma * Sqr(T)
End Function

Function Nx(x As Double) As Double
    Nx = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function
Function Nprime(x As Double) As Double
    Nprime = Exp(-0.5 * x ^ 2) / Sqr(2 * Application.WorksheetFunction.Pi())
End Function


Function DeltaCalc(S As Double, K As Double, T As Double, r As Double, sigma As Double, option_type As String) As Double
    
    If LCase(option_type) = "Call" Then
        DeltaCalc = Nx(d1Calc(S, K, T, r, sigma))
    ElseIf LCase(option_type) = "put" Then
        DeltaCalc = Nx(d1Calc(S, K, T, r, sigma)) - 1
End Function

Function GammaCalc(S As Double, K As Double, T As Double, r As Double, sigma As Double) As Double
    GammaCalc = (Exp(-0.5 * d1Calc(S, K, T, r, sigma) ^ 2)) / (Sqr(2 * Application.WorksheetFunction.Pi() * T) * sigma * S)
    
End Function

Function VegaCalc(S As Double, K As Double, T As Double, r As Double, sigma As Double) As Double
    VagaCalc = S * Sqr(T) * (Exp(-0.5 * d1Calc(S, K, T, r, sigma) ^ 2)) / (Sqr(2 * Application.WorksheetFunction.Pi())) * 0.01
    
End Function

Function ThetaCalc(S As Double, K As Double, T As Double, r As Double, sigma As Double) As Double
    Dim d1 As Double, d2 As Double
    d1 = d1Calc(S, K, T, r, sigma)
    d2 = d2Calc(S, K, T, r, sigma)
    
  If LCase(option_type) = "call" Then
        Theta = (-S * Nprime(d1) * sigma / (2 * Sqr(T)) - r * K * Exp(-r * T) * Nx(d2)) / 365
    Else
        Theta = (-S * Nprime(d1) * sigma / (2 * Sqr(T)) + r * K * Exp(-r * T) * Nx(-d2)) / 365
    End If

End Function

Function Rho(S As Double, K As Double, T As Double, r As Double, sigma As Double, option_type As String) As Double
    Dim d2 As Double
    d2 = d2Calc(S, K, T, r, sigma)

    If LCase(option_type) = "call" Then
        Rho = K * T * Exp(-r * T) * Nx(d2) * 0.01
    Else
        Rho = -K * T * Exp(-r * T) * Nx(-d2) * 0.01
    End If
End Function
