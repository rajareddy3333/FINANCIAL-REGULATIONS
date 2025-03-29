Attribute VB_Name = "Module1"
Function Black76Call(F As Double, K As Double, T As Double, r As Double, sigma As Double) As Double
    ' Black (1976) model to price a call option on a bond future
    ' F = Futures price
    ' K = Strike price (exercise price)
    ' T = Time to maturity (in years)
    ' r = Risk-free interest rate (as a decimal)
    ' sigma = Volatility of futures price (as a decimal)
    
    Dim d1 As Double, d2 As Double, N_d1 As Double, N_d2 As Double
    Dim Pi As Double
    Pi = WorksheetFunction.Pi()
    
    ' Compute d1 and d2
    d1 = (Log(F / K) + (0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
    d2 = d1 - sigma * Sqr(T)
    
    ' Compute cumulative normal distribution values
    N_d1 = WorksheetFunction.Norm_S_Dist(d1, True)
    N_d2 = WorksheetFunction.Norm_S_Dist(d2, True)
    
    ' Black-76 formula for call option price
    Black76Call = Exp(-r * T) * (F * N_d1 - K * N_d2)
End Function

Sub task6()

End Sub
