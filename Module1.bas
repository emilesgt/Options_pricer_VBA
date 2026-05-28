Attribute VB_Name = "Module1"
Option Explicit


' Generate a standard normal random variable
' using the Box-Muller transformation

Private Function RandNorm() As Double
    Dim u1 As Double, u2 As Double
    ' Draw two independent uniform random variables
    u1 = Rnd
    If u1 < 0.000000000001 Then u1 = 0.000000000001
    u2 = Rnd
    ' Box-Muller formula to obtain a N(0,1) variable
    RandNorm = Sqr(-2 * Log(u1)) * Cos(2 * WorksheetFunction.Pi() * u2)
End Function


' Monte Carlo Asian Call (arithmŽtique)
' Inputs:
'   S0      : initial asset price
'   K       : strike
'   r       : risk-free rate
'   q       : dividend yield
'   sigma   : volatility
'   T       : maturity (years)
'   nSteps  : number of time steps
'   nPaths  : number of Monte Carlo simulations
'
' Outputs (ByRef):
'   SE      : standard error of the Monte Carlo estimator
'   CI_L/U  : 95% confidence interval
'
' Function output:
'   AsianCallMC : Monte Carlo price of the option

Private Function AsianCallMC( _
    S0 As Double, K As Double, r As Double, q As Double, _
    sigma As Double, T As Double, nSteps As Long, nPaths As Long, _
    ByRef SE As Double, ByRef CI_L As Double, ByRef CI_U As Double _
) As Double
' Time discretization
    Dim dt As Double, drift As Double, vol As Double
    dt = T / nSteps
    drift = (r - q - 0.5 * sigma ^ 2) * dt
    vol = sigma * Sqr(dt)

    Dim i As Long, j As Long
    ' Variables for asset path simulation
    Dim S As Double, sumS As Double, avgS As Double
    Dim payoff As Double, disc As Double
    disc = Exp(-r * T)

' Variables for online mean and variance (Welford algorithm)
    Dim meanPayoff As Double, M2 As Double
    meanPayoff = 0: M2 = 0

    Dim z As Double, delta As Double
'Monte Carlo loop
    For i = 1 To nPaths
    ' Initialize asset price and sum for averaging
        S = S0
        sumS = 0
' Simulate one price path
        For j = 1 To nSteps
            z = RandNorm() ' standard normal shock
            ' GBM dynamics under risk-neutral measure
            S = S * Exp(drift + vol * z)
            sumS = sumS + S
        Next j
 ' average of prices
        avgS = sumS / nSteps
        payoff = disc * Application.Max(avgS - K, 0) ' Discounted payoff of Asian call
' Update mean and variance incrementally
        delta = payoff - meanPayoff
        meanPayoff = meanPayoff + delta / i
        M2 = M2 + delta * (payoff - meanPayoff)
    Next i

' Estimate variance of the payoff
    Dim variance As Double
    variance = M2 / (nPaths - 1)

' Standard error and confidence interval
    SE = Sqr(variance / nPaths)
    CI_L = meanPayoff - 1.96 * SE
    CI_U = meanPayoff + 1.96 * SE

    AsianCallMC = meanPayoff ' Return Monte Carlo price
End Function


' Closed-form formula for a geometric Asian Call option
' (used as a benchmark for comparison)

Private Function AsianGeoClosedForm( _
    S0 As Double, K As Double, r As Double, q As Double, _
    sigma As Double, T As Double, nSteps As Long _
) As Double

    Dim n As Double: n = nSteps
    
    ' Adjusted volatility for geometric average
    Dim sigmaG2 As Double
    sigmaG2 = sigma ^ 2 * ((n + 1) * (2 * n + 1)) / (6 * n ^ 2)

    Dim sigmaG As Double: sigmaG = Sqr(sigmaG2)
      ' Adjusted drift
    Dim muG As Double
    muG = (r - q - 0.5 * sigma ^ 2) * (n + 1) / (2 * n) + 0.5 * sigmaG2

' Black-Scholes type variables
    Dim d1 As Double, d2 As Double
    d1 = (Log(S0 / K) + (muG + 0.5 * sigmaG2) * T) / (sigmaG * Sqr(T))
    d2 = d1 - sigmaG * Sqr(T)

' Closed-form price
    AsianGeoClosedForm = Exp(-r * T) * _
        (S0 * Exp(muG * T) * WorksheetFunction.Norm_S_Dist(d1, True) _
        - K * WorksheetFunction.Norm_S_Dist(d2, True))
End Function


' runs the Asian option pricing project

Public Sub Run_Asian_Project()

    Dim wsI As Worksheet: Set wsI = Sheets("inputs")
    Dim wsR As Worksheet: Set wsR = Sheets("MC_VBA")

    ' Lecture Inputs
    Dim S0 As Double: S0 = wsI.Range("C3")
    Dim K As Double: K = wsI.Range("D3")
    Dim r As Double: r = wsI.Range("E3")
    Dim q As Double: q = wsI.Range("F3")
    Dim sigma As Double: sigma = wsI.Range("G3")
    Dim T As Double: T = wsI.Range("H3")
    Dim nSteps As Long: nSteps = wsI.Range("I3")
    Dim nPaths As Long: nPaths = wsI.Range("J3")
    Dim seed As Long: seed = wsI.Range("L3")

' Initialize random generator
    Randomize seed

'Run Monte Carlo pricing
    Dim SE As Double, CI_L As Double, CI_U As Double
    Dim priceMC As Double
    priceMC = AsianCallMC(S0, K, r, q, sigma, T, nSteps, nPaths, SE, CI_L, CI_U)
    
' Closed-form geometric Asian price
    Dim priceGeo As Double
    priceGeo = AsianGeoClosedForm(S0, K, r, q, sigma, T, nSteps)

    ' ƒcriture rŽsultats
    wsR.Range("A2") = "Asian Arithmetic Call (MC)"
    wsR.Range("B2") = priceMC

    wsR.Range("A3") = "Std Error"
    wsR.Range("B3") = SE

    wsR.Range("A4") = "CI 95% - Lower"
    wsR.Range("B4") = CI_L

    wsR.Range("A5") = "CI 95% - Upper"
    wsR.Range("B5") = CI_U

    wsR.Range("A7") = "Asian Geometric Call (Closed Form)"
    wsR.Range("B7") = priceGeo

    wsR.Range("A8") = "Difference (MC - Geo)"
    wsR.Range("B8") = priceMC - priceGeo

    

End Sub

' Fill convergence table by increasing the number of simulations
Public Sub Fill_Convergence_Table()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

    Dim wsC As Worksheet: Set wsC = Sheets("Convergence_MC_VBA")
    Dim wsI As Worksheet: Set wsI = Sheets("inputs")
    Dim wsR As Worksheet: Set wsR = Sheets("MC_VBA")
    
' Different numbers of Monte Carlo simulations
    Dim Nlist As Variant
    Nlist = Array(500, 1000, 2000, 5000, 10000, 20000, 50000)

    Dim i As Long
    
    ' Table headers
    wsC.Range("A1") = "Number of simulations"
    wsC.Range("B1") = "Asian Call Price"

' Loop over simulation sizes
    For i = 0 To UBound(Nlist)
        wsI.Range("J3").Value = Nlist(i)   ' change nPaths
        Call Run_Asian_Project              ' run pricing
        wsC.Cells(i + 2, 1).Value = Nlist(i)
        wsC.Cells(i + 2, 2).Value = wsR.Range("B2").Value
    Next i

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub



