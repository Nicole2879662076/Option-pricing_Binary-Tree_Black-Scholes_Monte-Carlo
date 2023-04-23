Attribute VB_Name = "Ä£¿é1"
Function Binomial_eur_Call(Up, Down, Interest, _
        Stock, Exercise, Periods)
    q_up = (Interest - Down) / _
        (Interest * (Up - Down))
    q_down = 1 / Interest - q_up
    Binomial_eur_Call = o
    For Index = o To Periods
        Binomial_eur_Call = Binomial_eur_Call _
            + Application.Combin(Periods, Index) _
            * q_up ^ Index * q_down ^ (Periods - Index) _
            * Application.Max(Stock * Up ^ Index * Down _
            ^ (Periods - Index) - Exercise, 0)
    Next Index
End Function

Function Binomial_eur_put(Up, Down, Interest, _
        Stock, Exercise, Periods)
    Binomial_eur_put = Binomial_eur_Call _
        (Up, Down, Interest, Stock, Exercise, _
        Periods) + Exercise / Interest ^ Periods - Stock
    
    Debug.Print Binomial_eur_put
End Function

Function BSCall(Stock, Exercise, Time, _
            Interest, sigma)
    BSCall = Stock * Application.NormSDist _
            (dOne(Stock, Exercise, Time, Interest, _
            sigma)) - Exercise * Exp(-Time * Interest) * _
            Application.NormSDist(dTwo(Stock, Exercise, _
            Time, Interest, sigma))
End Function
Function BsPut(Stock, Exercise, Time, _
            Interest, sigma)
    BsPut = BSCall(Stock, Exercise, Time, _
            Interest, sigma) + Exercise * _
            Exp(-Interest * Time) - Stock
End Function

Function dOne(Stock, Exercise, Time, _
            Interest, sigma)
    dOne = (Log(Stock / Exercise) + _
            Interest * Time) / (sigma * Sqr(Time)) _
             + 0.5 * sigma * Sqr(Time)
End Function
Function dTwo(Stock, Exercise, Time, _
            Interest, sigma)
    dTwo = dOne(Stock, Exercise, Time, _
            Interest, sigma) - sigma * Sqr(Time)
End Function

