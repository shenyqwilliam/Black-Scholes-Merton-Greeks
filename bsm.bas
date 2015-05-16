'Written by William Shen
'Latest revision: 2012/06/21
'Compatible with Excel 2010
'
'
'[Licensing]
'I, the copyright holder of this work, hereby release it into the public domain. This applies worldwide.
'If this is not legally possible:
'I grant any entity the right to use this work for any purpose, without any conditions, unless such conditions are required by law.
'
'
'[Disclaimer]
'This VBA macro may contain mistakes. The author does not assume accountability for any loss caused by this macro.
'In short, use at your own discretion.
'
'
'[Notes]
'This VBA emphasis on efficiency rather than accuracy or code explicitness.
'Basic formulas of option price and most first-order Greeks are based on work by Chen, Lee and Shih (2010).
'Most second- and third-order Greeks are based on article by Haug (2003). Cost of carry "b" in Haug's article equals "r-q" in this Macro, i.e. q=r-b
'CDF of normal distribution uses numerical calculation based on work by Abramowitz and Stegun (1964). If you need more accurate result, use Excel functions instead.
'Since percentage Greeks are subject to convention, they are commented. Make sure to check codes before uncommenting and using them.
'
'
'[References]
'[1] Abramowitz,M. and Stegun,I.A. (1964). Handbook of Mathematical Functions:p932
'    http://people.math.sfu.ca/~cbm/aands/page_932.htm
'[2] Chen,Hong-Yi, Lee,Cheng-Few, and Shih,Weikang (2010). "Derivations and Applications of Greek Letters: Review and Integration", Handbook of Quantitative Finance and Risk Management, III:pp491-503
'    http://centerforpbbefr.rutgers.edu/TaipeiPBFR&D/01-16-09%20papers/5-4%20Greek%20letters.doc
'[3] Haug,Espen Gaarder (2003). "Know Your Weapon Part 1", Wilmott Magazine, May:pp49-57
'    http://www.wilmott.com/pdfs/050527_haug.pdf
'[4] Haug,Espen Gaarder (2003). "Know Your Weapon Part 2", Wilmott Magazine, July
'[5] Haug,Espen Gaarder (2006). The Complete Guide to Option Pricing Formulas 2ed, McGraw-Hill



Option Explicit
Option Compare Text
Option Base 0

Global Const Pi# = 3.14159265358979
Public Const CDFa0# = 0.2316419
Public Const CDFa1# = 0.31938153
Public Const CDFa2# = -0.356563782
Public Const CDFa3# = 1.781477937
Public Const CDFa4# = -1.821255978
Public Const CDFa5# = 1.330274429

'Add to "Insert Formula" dialog box.
'For Excel 2010 only. Not working well so far.
Private Sub BSMAddUDFDesc()
    Dim ArgDesc(7) As String
    ArgDesc(0) = "Specify option type. 0=Call, 1=Put"
    ArgDesc(1) = "Spot price of underlying asset"
    ArgDesc(2) = "Strike price of the option"
    ArgDesc(3) = "Volatility p.a."
    ArgDesc(4) = "Time to expiration"
    ArgDesc(5) = "Risk free interest rate (countinuous). [Default: 0%]"
    ArgDesc(6) = "Dividend rate (continuous). [Default: 0%]"
    ArgDesc(7) = "Output type. 0=Price, 1=Delta, 2=Vega, 3=Theta, 4=Rho, 5=Phi, 6=Dual Delta, 7=Lambda, 8=Vega elasticity, 9=zeta; " & _
                    "11=Gamma, 12=Vanna, 13=Charm, 14=Vomma, 15=DvegaDtime, 16=Vera, 17=Dual Gamma, 18=DzetaDvol, 19=DzetaDtime; " & _
                    "21=Speed, 21=Zomma, 22=Color, 23=Ultima. " & _
                    "[Default: 0 (price)]"
    
    Application.MacroOptions _
        Macro:="PERSONAL.XLSB!BSM", _
        Description:="Black-Scholes-Merton model", _
        Category:="UDF", _
        ArgumentDescriptions:=ArgDesc
End Sub

'Core function
Public Function BSM( _
    ByVal CallOrPut As String, _
    ByVal SpotPrice As Double, _
    ByVal StrikePrice As Double, _
    ByVal VolatilityPerAnnum As Double, _
    ByVal TimeToExpiration As Double, _
    Optional ByVal RiskFreeRate As Double, _
    Optional ByVal DividendRate As Double, _
    Optional ByVal OutputType As String _
) As Double

    Dim OptType As Boolean
    Dim S#, X#, sigma#, tau#, r#, q#
    Dim OutType%
    
    Dim pteVol#, pvX#, ddS#
    Dim d1#, d2#
    Dim Nd1#, Nd2#
    
    'Process call or put
    Select Case CallOrPut
        Case "0", "c", "call"
            OptType = 0
        Case "1", "p", "put"
            OptType = 1
        Case Else
            BSM = "Error: please specify call or put."
            Exit Function
    End Select
    
    'Process output type
    Select Case OutputType
        Case "0", "Price", "Value", "Option Premium", ""
            OutType = 0
        Case "1", "Delta"
            OutType = 1
        Case "2", "Vega", "Kappa"
            OutType = 2
        Case "3", "Theta"
            OutType = 3
        Case "4", "Rho"
            OutType = 4
        Case "5", "Phi", "Rho-2"
            OutType = 5
        Case "6", "Dual Delta"
            OutType = 6
        Case "7", "Lambda", "Omege", "Elasticity", "Delta elasticity"
            OutType = 7
        Case "8", "Vega leverage", "Vega elasticity"
            OutType = 8
        Case "9", "Zeta"
            OutType = 9
        Case "11", "Gamma"
            OutType = 11
        Case "12", "Vanna", "DdeltaDvol", "DvegaDspot"
            OutType = 12
        Case "13", "DdeltaDtime", "Charm", "Delta Decay", "Delta Bleed"
            OutType = 13
        Case "14", "Vomma", "Volga", "DvegaDvol", "Vega Convexity"
            OutType = 14
        Case "15", "DvegaDtime"
            OutType = 15
'        Case "16", "Vera"
'            OutType = 16
        Case "17", "Dual Gamma", "Risk neutral probability density", "RND"
            OutType = 17
        Case "18", "DzetaDvol"
            OutType = 18
        Case "19", "dzetaDtime"
            OutType = 19
        Case "21", "Speed", "DgammaDspot", "Gamma of Gamma"
            OutType = 21
        Case "22", "Zomma", "DgammaDvol"
            OutType = 22
        Case "23", "Color", "Colour", "DgammaDtime", "Gamma Decay", "Gamma Bleed"
            OutType = 23
        Case "24", "Ultima"
            OutType = 24
'        Case "31", "VegaP"
'            OutType = 31
'        Case "32", "GammaP"
'            OutType = 32
'        Case "33", "DvegaPDvol"
'            OutType = 33
'        Case "34", "SpeedP", "DgammaPDspot"
'            OutType = 34
'        Case "35", "ZommaP", "DgammaPDvol"
'            OutType = 35
'        Case "36", "ColorP", "ColourP", "DgammaPDtime"
'            OutType = 36
        Case Else
            BSM = "Error: please specify correct output type."
            Exit Function
    End Select
    
    S = SpotPrice
    X = StrikePrice
    sigma = VolatilityPerAnnum
    tau = TimeToExpiration
    r = RiskFreeRate
    q = DividendRate
    
    
    If tau = 0 Then
        Select Case OutType
            Case 0      'Price
                If OptType = 0 Then
                        If S > X Then BSM = S - X Else BSM = 0
                Else:   If X > S Then BSM = X - S Else BSM = 0
                End If
            Case 1      'Delta
                If OptType = 0 Then
                        If S > X Then BSM = 1 Else If S < X Then BSM = 0 Else BSM = "Error"
                Else:   If X > S Then BSM = -1 Else If X < S Then BSM = 0 Else BSM = "Error"
                End If
            Case 11, 32   'Gamma, GammaP
                If S <> X Then BSM = 0 Else BSM = "Error"
            Case Else
                BSM = "Error: not implemented yet."
        End Select
        
    Else    'tau<>0
        pteVol = sigma * Sqr(tau)   'Volatility from present to expiration.
        pvX = X * Exp(-r * tau)     'Present value of strike.
        ddS = S * Exp(-q * tau)     'Dividend-deducted spot price.
        d1 = (Log(S / X) + (r - q + sigma * sigma / 2) * tau) / pteVol
        d2 = d1 - pteVol
        Nd1 = NormSCDF(d1)
        Nd2 = NormSCDF(d2)
        
        Select Case OutType
            Case 0      'Price
                If OptType = 0 Then
                    BSM = ddS * Nd1 - pvX * Nd2
                Else
                    BSM = pvX * (1 - Nd2) - ddS * (1 - Nd1)
                End If
            Case 1      'Delta
                If OptType = 0 Then
                    BSM = Exp(-q * tau) * Nd1
                Else
                    BSM = Exp(-q * tau) * (Nd1 - 1)
                End If
            Case 2      'Vega
                BSM = ddS * Sqr(tau) * NPrime(d1)
            Case 3      'Theta
                If OptType = 0 Then
                    BSM = q * ddS * Nd1 - (ddS * sigma) / (2 * Sqr(tau)) * NPrime(d1) - r * pvX * Nd2
                Else
                    BSM = -q * ddS * Nd1 - (ddS * sigma) / (2 * Sqr(tau)) * NPrime(d1) + r * pvX * (1 - Nd2)
                End If
            Case 4      'Rho
                If OptType = 0 Then
                    BSM = pvX * tau * Nd2
                Else
                    BSM = pvX * tau * (Nd2 - 1)
                End If
            Case 5      'Phi
                If OptType = 0 Then
                    BSM = -tau * ddS * Nd1
                Else
                    BSM = tau * ddS * (1 - Nd1)
                End If
            Case 6      'Dual Delta
                If OptType = 0 Then
                    BSM = -Exp(-r * tau) * Nd2
                Else
                    Exp (-r * tau) * (1 - Nd2)
                End If
            Case 7      'Lambda
                If OptType = 0 Then
                    BSM = Exp(-q * tau) * Nd1 * S / (ddS * Nd1 - pvX * Nd2)
                Else
                    BSM = Exp(-q * tau) * (Nd1 - 1) * S / (pvX * (1 - Nd2) - ddS * (1 - Nd1))
                End If
            Case 8      'Vega elasticity
                If OptType = 0 Then
                    BSM = ddS * Sqr(tau) * NPrime(d1) * sigma / (ddS * Nd1 - pvX * Nd2)
                Else
                    BSM = ddS * Sqr(tau) * NPrime(d1) * sigma / (pvX * (1 - Nd2) - ddS * (1 - Nd1))
                End If
            Case 9      'Zeta
                If OptType = 0 Then
                    BSM = Nd2
                Else
                    BSM = 1 - Nd2
                End If
            Case 11     'Gamma
                BSM = Exp(-q * tau) / (S * pteVol) * NPrime(d1)
            Case 12     'Vanna
                BSM = -Exp(-q * tau) / sigma * d2 * NPrime(d1)
            Case 13     'Charm
                If OptType = 0 Then
                    BSM = -Exp(-q * tau) * (NPrime(d1) * ((r - q) / pteVol - d2 / (2 * tau)) - q * Nd1)
                Else
                    BSM = -Exp(-q * tau) * (NPrime(d1) * ((r - q) / pteVol - d2 / (2 * tau)) + q * (1 - Nd1))
                End If
            Case 14     'Vomma
                BSM = ddS * Sqr(tau) * NPrime(d1) * d1 * d2 / sigma
            Case 15     'DvegaDtime
                BSM = ddS * Sqr(tau) * NPrime(d1) * (q + (r - q) * d1 / pteVol - (1 + d1 * d2) * 0.5 / tau)
            Case 17     'Dual Gamma
                BSM = Exp(-r * tau) * NPrime(d2) / (X * pteVol)
            Case 18     'DzetaDvol
                If OptType = 0 Then
                    BSM = -NPrime(d2) * d1 / sigma
                Else
                    BSM = NPrime(d2) * d1 / sigma
                End If
            Case 19     'DzetaDtime
                If OptType = 0 Then
                    BSM = NPrime(d2) * ((r - q) / pteVol - d1 / (2 * tau))
                Else
                    BSM = -NPrime(d2) * ((r - q) / pteVol - d1 / (2 * tau))
                End If
            Case 21     'Speed
                BSM = -(Exp(-q * tau) / (S * S * pteVol) * NPrime(d1)) * (1 + d1 / pteVol)
            Case 22     'Zomma
                BSM = Exp(-q * tau) / (S * pteVol) * NPrime(d1) * (d1 * d2 - 1) / sigma
            Case 23     'Color
                BSM = Exp(-q * tau) / (S * pteVol) * NPrime(d1) * (q + (r - q) * d1 / pteVol + (1 - d1 * d2) / (2 * tau))
            Case 24     'Ultima
                BSM = -ddS * Sqr(tau) * NPrime(d1) / (sigma * sigma) * (d1 * d2 * (1 - d1 * d2) + d1 * d1 + d2 * d2)
'            Case 31      'VegaP
'                BSM = sigma * 0.1 * ddS * Sqr(tau) * NPrime(d1)
'            Case 32     'GammaP
'                BSM = Exp(-q * tau) / (100 * pteVol) * NPrime(d1)
'            Case 33     'DvegaPDvol
'                BSM = 0.1 * ddS * Sqr(tau) * NPrime(d1) * d1 * d2
'            Case 34     'SpeedP
'                BSM = -(Exp(-q * tau) / (100 * S * pteVol * pteVol) * NPrime(d1)) * d1
'            Case 35     'ZommaP
'                BSM = Exp(-q * tau) / (100 * pteVol) * NPrime(d1) * (d1 * d2 - 1) / sigma
'            Case 36     'ColorP
'                BSM = Exp(-q * tau) / (100 * pteVol) * NPrime(d1) * (q + (r - q) * d1 / pteVol + (1 - d1 * d2) / (2 * tau))
        End Select
        
    End If
    
End Function

'Standard normal probability density function
Private Function NPrime(ByVal d As Double) As Double
    NPrime = Exp(-d * d * 0.5) / Sqr(2 * Pi)
    
    ''Roughly 3x faster than native Excel functions.
    
    'NPrime = WorksheetFunction.Norm_S_Dist(d, False)            'For Excel 2010
    'NPrime = WorksheetFunction.NormDist(d, 0, 1, False)         'For Excel 2007
End Function

'Standard normal cumulative distribution function
Private Function NormSCDF(ByVal d As Double) As Double
    Dim t#
    
    If d >= 0 Then
        t = 1 / (1 + CDFa0 * d)
        NormSCDF = 1 - NPrime(d) * (CDFa1 * t + CDFa2 * t * t + CDFa3 * t * t * t + CDFa4 * t * t * t * t + CDFa5 * t * t * t * t * t)
    Else: NormSCDF = 1 - NormSCDF(-d)
    End If
    
    ''Numerical approximation, based on work by Abramowitz & Stegun (1964)
    ''Roughly 3x faster than native Excel functions. |epsilon|<7.5E-8
    
    'NormSCDF = WorksheetFunction.Norm_S_Dist(d, True)           'For Excel 2010
    'NormSCDF = WorksheetFunction.NormSDist(d)                   'For Excel 2007
End Function
