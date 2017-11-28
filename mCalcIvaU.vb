Option Explicit On 

Imports System.Math

Module mCalcIvaU

    ' Esta función recibe el parámetro drUdis como un DataRowCollection,
    ' nSaldo, nTasa (la tasa moratoria), cFechaInicial y cFechaFinal.
    ' Adicionalmente recibe los parámetros nUdiInicial y nUdiFinal los cuales
    ' son recibidos por referencia, lo que significa que su valor es modificado
    ' en esta función

    Public Function CalcIvaU(ByVal drUdis As DataRowCollection, ByVal nSaldo As Decimal, ByVal nTasa As Decimal, ByVal cFechaInicial As String, ByVal cFechaFinal As String, ByRef nUdiInicial As Decimal, ByRef nUdiFinal As Decimal) As Decimal

        ' Declaración de variables de datos

        Dim drUdi As DataRow
        Dim dFechaInicial As Date
        Dim dFechaFinal As Date
        Dim nDias As Integer

        nUdiInicial = 0
        nUdiFinal = 0
        CalcIvaU = 0

        If nSaldo > 0 Then
            dFechaInicial = CTOD(cFechaInicial)
            dFechaFinal = CTOD(cFechaFinal)
            nDias = DateDiff(DateInterval.Day, dFechaInicial, dFechaFinal)
            If nDias > 0 Then
                dFechaInicial = DateAdd(DateInterval.Day, -1, dFechaInicial)
                dFechaFinal = DateAdd(DateInterval.Day, -1, dFechaFinal)
                cFechaInicial = DTOC(dFechaInicial)
                cFechaFinal = DTOC(dFechaFinal)
                For Each drUdi In drUdis
                    If drUdi("Vigencia") = cFechaInicial Then
                        nUdiInicial = drUdi("Udi")
                    End If
                    If drUdi("Vigencia") = cFechaFinal Then
                        nUdiFinal = drUdi("Udi")
                    End If
                Next
                If nUdiFinal <= nUdiInicial Then
                    CalcIvaU = nSaldo * nTasa * nDias / 36000 * 0.15
                Else
                    CalcIvaU = nSaldo * ((nTasa * nDias / 36000) - ((nUdiFinal / nUdiInicial) - 1)) * 0.15
                End If
                CalcIvaU = Round(CalcIvaU, 2)
                If CalcIvaU < 0 Then
                    CalcIvaU = 0
                End If
            End If
        End If

    End Function

End Module
