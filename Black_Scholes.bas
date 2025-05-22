Attribute VB_Name = "BlackScholes"
Function d1Calc(S As Double, K As Double, T As Double, r As Double, sigma As Double) As Double
    d1Calc = (Log(S / K) + (r + 0.5 * sigma ^ 2) * T) / sigma * Sqr(T)
End Function

Function d2Calc(S As Double, K As Double, T As Double, r As Double, sigma As Double) As Double
    d2Calc = d1Calc(S, K, T, r, sigma) - sigma * Sqr(T)
End Function

Function Nx(x As Double) As Double
    Nx = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function

Function BS(S As Double, K As Double, T As Double, r As Double, sigma As Double, option_type As String) As Double
    Dim Nd1 As Double, Nd2 As Double, d1Val As Double, d2Val As Double

    d1Val = d1Calc(S, K, T, r, sigma)
    d2Val = d1Val - sigma * Sqr(T)
    
    Nd1 = Nx(d1Val)
    Nd2 = Nx(d2Val)
    
    If LCase(option_type) = "call" Then
        BS = S * Nd1 - K * Exp(-r * T) * Nd2
    ElseIf option_type = "put" Then
        BS = -S * Nx(-d1Val) + K * Exp(-r * T) * Nx(-d2Val)
    Else
        MsgBox "Veuillez entrer un type d'option valide (call ou put).", vbExclamation, "Erreur de saisie"
    
    End If
    
End Function
