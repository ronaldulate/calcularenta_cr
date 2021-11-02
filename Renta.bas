Attribute VB_Name = "Renta"
'Calcula la Renta 02-11-2021 ronaldulate

'Calcula la renta segun los tramos al 2-Nov-2021
Function CalculaRenta(renta_Bruta)

    Dim saldoRenta As Currency
    Dim rentaBruta As Currency
    Dim rentaNeta As Currency
    Dim exceso As Currency

    Dim tramoRenta1 As Currency
    Dim tramoRenta2 As Currency
    Dim tramoRenta3 As Currency
    Dim tramoRenta4 As Currency

    Dim porcentajeTramo1 As Currency
    Dim porcentajeTramo2 As Currency
    Dim porcentajeTramo3 As Currency
    Dim porcentajeTramo4 As Currency


    'Definiendo los tramos:

    tramoRenta1 = 5157000#
    tramoRenta2 = 7737000#
    tramoRenta3 = 10315000#
    tramoRenta4 = 109337000#

    porcentajeTramo1 = 5 / 100
    porcentajeTramo2 = 10 / 100
    porcentajeTramo3 = 15 / 100
    porcentajeTramo4 = 20 / 100


    
    rentaBruta = renta_Bruta
    rentaNeta = 0#
    
    If rentaBruta > tramoRenta4 Then

    rentaNeta = rentaBruta * 30 / 100

    Else

        saldoRenta = rentaBruta

        If saldoRenta > tramoRenta3 Then
           exceso = saldoRenta - tramoRenta3
           rentaNeta = rentaNeta + (exceso * porcentajeTramo4)
           saldoRenta = tramoRenta3
        End If

        If saldoRenta > tramoRenta2 Then
           exceso = saldoRenta - tramoRenta2
           rentaNeta = rentaNeta + (exceso * porcentajeTramo3)
           saldoRenta = tramoRenta2
        End If

        If saldoRenta > tramoRenta1 Then
           exceso = saldoRenta - tramoRenta1
           rentaNeta = rentaNeta + (exceso * porcentajeTramo2)
           saldoRenta = tramoRenta1
        End If

        If saldoRenta <= tramoRenta1 Then
           rentaNeta = rentaNeta + (saldoRenta * porcentajeTramo1)
           saldoRenta = 0
        End If
  
    End If
    'Return:
    CalculaRenta = rentaNeta
   

End Function
