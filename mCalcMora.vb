Option Explicit On 

Imports System.Math

Module mCalcMora

    Public Function CalcMora(ByVal drUdis As DataRowCollection, ByVal nSaldo As Decimal, ByVal nTasaMoratoria As Decimal, ByVal nDiasMoratorios As Decimal, ByRef nMoratorios As Decimal, ByRef nIvaMoratorios As Decimal, ByVal cFecha As String) As Decimal

        ' Declaración de variables de datos

        Dim cFechaInicial As String
        Dim dFechaInicial As Date
        Dim nUdiFinal As Decimal
        Dim nUdiInicial As Decimal

        dFechaInicial = DateAdd(DateInterval.Day, -nDiasMoratorios, CTOD(cFecha))
        cFechaInicial = DTOC(dFechaInicial)
        nUdiInicial = 0
        nUdiFinal = 0

        nMoratorios = Round(nSaldo * nTasaMoratoria * nDiasMoratorios / 36000, 2)
        nIvaMoratorios = CalcIvaU(drUdis, nSaldo, nTasaMoratoria, cFechaInicial, cFecha, nUdiInicial, nUdiFinal)

    End Function

End Module
