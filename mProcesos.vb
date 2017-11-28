Option Explicit On

Imports Microsoft.VisualBasic
Imports System.Math

Module mProcesos

    Public Function CTOD(ByVal cFecha As String) As Date

        Dim nDia, nMes, nYear As Integer

        nDia = Val(Right(cFecha, 2))
        nMes = Val(Mid(cFecha, 5, 2))
        nYear = Val(Left(cFecha, 4))

        CTOD = DateSerial(nYear, nMes, nDia)

    End Function

    Public Function DTOC(ByVal dFecha As Date) As String

        Dim cDia, cMes, cYear, sFecha As String

        sFecha = dFecha.ToShortDateString

        cDia = Left(sFecha, 2)
        cMes = Mid(sFecha, 4, 2)
        cYear = Right(sFecha, 4)

        DTOC = cYear & cMes & cDia

    End Function

    Public Function DameUdi(ByVal drUdis As DataRowCollection, ByVal cFecha As String, ByVal cFvenc As String, ByRef nUdiInicial As Decimal, ByRef nUdiFinal As Decimal) As Decimal

        ' Declaración de variables de datos

        Dim drUdi As DataRow
        Dim dFvenc As Date
        Dim dFecha As Date

        dFvenc = DateAdd(DateInterval.Day, -1, CTOD(cFvenc))
        dFecha = DateAdd(DateInterval.Day, -1, CTOD(cFecha))
        cFvenc = DTOC(dFvenc)
        cFecha = DTOC(dFecha)

        nUdiInicial = 0
        nUdiFinal = 0

        For Each drUdi In drUdis
            If drUdi("Vigencia") = cFvenc Then
                nUdiInicial = drUdi("Udi")
            End If
            If drUdi("Vigencia") = cFecha Then
                nUdiFinal = drUdi("Udi")
            End If
        Next

    End Function

    Public Function TraeIVA(ByVal cFecha As String) As Decimal
        Dim nIva As Byte
        Dim cFecha1 As String = "19921111"
        Dim cFecha2 As String = "19950331"
        Dim cFecha3 As String = "19950401"

        If cFecha >= cFecha3 Then
            nIva = 15
        ElseIf cFecha > cFecha1 And cFecha <= cFecha2 Then
            nIva = 10
        ElseIf cFecha <= cFecha1 Then
            nIva = 15
        End If
        TraeIVA = nIva
    End Function

    Public Function Stuff(ByVal Cadena As String, ByVal Lado As String, ByVal Llenarcon As String, ByVal Longitud As Integer) As String

        ' Declaración de variables de datos

        Dim cCadenaAuxiliar As String
        Dim nVeces As Integer
        Dim i As Integer

        nVeces = Longitud - Val(Len(Cadena))

        cCadenaAuxiliar = ""
        For i = 1 To nVeces
            cCadenaAuxiliar = cCadenaAuxiliar & Llenarcon
        Next
        If Lado = "D" Then
            Stuff = Cadena & cCadenaAuxiliar
        Else
            Stuff = cCadenaAuxiliar & Cadena
        End If

    End Function

    Public Function Letras(ByVal numero As String) As String

        'Declaración de variables de datos

        Dim palabras As String
        Dim entero As String
        Dim dec As String
        Dim flag As String
        Dim num As Integer
        Dim x As Integer
        Dim y As Integer

        flag = "N"

        'Dividir parte entera y decimal

        For y = 1 To Len(numero)
            If Mid(numero, y, 1) = "." Then
                flag = "S"
            Else
                If flag = "N" Then
                    entero = entero + Mid(numero, y, 1)
                Else
                    dec = dec + Mid(numero, y, 1)
                End If
            End If
        Next y

        If Len(dec) = 1 Then dec = dec & "0"

        'proceso de conversión

        flag = "N"

        If Val(numero) <= 999999999 Then

            For y = Len(entero) To 1 Step -1

                num = Len(entero) - (y - 1)

                Select Case y

                    Case 3, 6, 9

                        'Asigna las palabras para las centenas

                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If Mid(entero, num + 1, 1) = "0" And Mid(entero, num + 2, 1) = "0" Then
                                    palabras = palabras & "CIEN "
                                Else
                                    palabras = palabras & "CIENTO "
                                End If
                            Case "2"
                                palabras = palabras & "DOSCIENTOS "
                            Case "3"
                                palabras = palabras & "TRESCIENTOS "
                            Case "4"
                                palabras = palabras & "CUATROCIENTOS "
                            Case "5"
                                palabras = palabras & "QUINIENTOS "
                            Case "6"
                                palabras = palabras & "SEISCIENTOS "
                            Case "7"
                                palabras = palabras & "SETECIENTOS "
                            Case "8"
                                palabras = palabras & "OCHOCIENTOS "
                            Case "9"
                                palabras = palabras & "NOVECIENTOS "
                        End Select
                    Case 2, 5, 8

                        'Asigna las palabras para las decenas 

                        Select Case Mid(entero, num, 1)

                            Case "0"
                                flag = "N"
                            Case "1"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    flag = "S"
                                    palabras = palabras & "DIEZ "
                                End If
                                If Mid(entero, num + 1, 1) = "1" Then
                                    flag = "S"
                                    palabras = palabras & "ONCE "
                                End If
                                If Mid(entero, num + 1, 1) = "2" Then
                                    flag = "S"
                                    palabras = palabras & "DOCE "
                                End If
                                If Mid(entero, num + 1, 1) = "3" Then
                                    flag = "S"
                                    palabras = palabras & "TRECE "
                                End If
                                If Mid(entero, num + 1, 1) = "4" Then
                                    flag = "S"
                                    palabras = palabras & "CATORCE "
                                End If
                                If Mid(entero, num + 1, 1) = "5" Then
                                    flag = "S"
                                    palabras = palabras & "QUINCE "
                                End If
                                If Mid(entero, num + 1, 1) > "5" Then
                                    flag = "N"
                                    palabras = palabras & "DIECI"
                                End If
                            Case "2"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "VEINTE "
                                    flag = "S"
                                Else
                                    palabras = palabras & "VEINTI"
                                    flag = "N"
                                End If
                            Case "3"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "TREINTA "
                                    flag = "S"
                                Else
                                    palabras = palabras & "TREINTA Y "
                                    flag = "N"
                                End If
                            Case "4"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "CUARENTA "
                                    flag = "S"
                                Else
                                    palabras = palabras & "CUARENTA Y "
                                    flag = "N"
                                End If
                            Case "5"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "CINCUENTA "
                                    flag = "S"
                                Else
                                    palabras = palabras & "CINCUENTA Y "
                                    flag = "N"
                                End If
                            Case "6"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "SESENTA "
                                    flag = "S"
                                Else
                                    palabras = palabras & "SESENTA Y "
                                    flag = "N"
                                End If
                            Case "7"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "SETENTA "
                                    flag = "S"
                                Else
                                    palabras = palabras & "SETENTA Y "
                                    flag = "N"
                                End If
                            Case "8"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "OCHENTA "
                                    flag = "S"
                                Else
                                    palabras = palabras & "OCHENTA Y "
                                    flag = "N"
                                End If
                            Case "9"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "NOVENTA "
                                    flag = "S"
                                Else
                                    palabras = palabras & "NOVENTA Y "
                                    flag = "N"
                                End If
                        End Select
                    Case 1, 4, 7

                        'Asigna las palabras para las unidades

                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If flag = "N" Then
                                    If y = 1 Then
                                        palabras = palabras & "UN "
                                    Else
                                        palabras = palabras & "UN "
                                    End If
                                End If
                            Case "2"
                                If flag = "N" Then palabras = palabras & "DOS "
                            Case "3"
                                If flag = "N" Then palabras = palabras & "TRES "
                            Case "4"
                                If flag = "N" Then palabras = palabras & "CUATRO "
                            Case "5"
                                If flag = "N" Then palabras = palabras & "CINCO "
                            Case "6"
                                If flag = "N" Then palabras = palabras & "SEIS "
                            Case "7"
                                If flag = "N" Then palabras = palabras & "SIETE "
                            Case "8"
                                If flag = "N" Then palabras = palabras & "OCHO "
                            Case "9"
                                If flag = "N" Then palabras = palabras & "NUEVE "
                        End Select
                End Select

                'Asigna la palabra mil

                If y = 4 Then
                    If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                    (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                    Len(entero) >= 6) Then palabras = palabras & "MIL "
                End If

                'Asigna la palabra millón

                If y = 7 Then
                    If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                        palabras = palabras & "MILLON "
                    Else
                        palabras = palabras & "MILLONES "
                    End If
                End If

            Next y

            'Se une la parte entera y la parte decimal

            If dec <> "" Then
                Letras = "(" & palabras & "PESOS " & dec & "/100 M.N.)"
            Else
                Letras = "(" & palabras & " PESOS 00/100 M.N.)"
            End If

        Else

            Letras = ""

        End If

    End Function

    Public Function Mes(ByVal cFecha As String) As String

        Dim cYear As String
        Dim cMes As String
        Dim cDia As String
        Dim cCadena As String

        cDia = Right(cFecha, 2)
        cMes = Mid(cFecha, 5, 2)
        cYear = Left(cFecha, 4)

        Select Case cMes
            Case "01"
                cCadena = " DE ENERO DE "
            Case "02"
                cCadena = " DE FEBRERO DE "
            Case "03"
                cCadena = " DE MARZO DE "
            Case "04"
                cCadena = " DE ABRIL DE "
            Case "05"
                cCadena = " DE MAYO DE "
            Case "06"
                cCadena = " DE JUNIO DE "
            Case "07"
                cCadena = " DE JULIO DE "
            Case "08"
                cCadena = " DE AGOSTO DE "
            Case "09"
                cCadena = " DE SEPTIEMBRE DE "
            Case "10"
                cCadena = " DE OCTUBRE DE "
            Case "11"
                cCadena = " DE NOVIEMBRE DE "
            Case "12"
                cCadena = " DE DICIEMBRE DE "
        End Select

        Mes = cDia & cCadena & cYear

    End Function

    Public Function Termina(ByVal dInicio As Date, ByVal nPlazo As Integer) As Date

        Dim nDay As Byte
        Dim nYear As Integer
        Dim nMonth As Byte
        Dim nAños As Integer
        Dim nAñosb As Integer
        Dim nLeap As Byte
        Dim i As Integer
        Dim nMes As Integer
        Dim nDia As Integer

        nDay = Day(dInicio)
        nMonth = (Month(dInicio) + (nPlazo Mod 12)) - 1
        nYear = Year(dInicio) + Int(nPlazo / 12)
        If nMonth > 12 Then
            nMonth -= 12
            nYear += 1
        ElseIf nMonth = 0 Then
            nMonth = 12
            nYear -= 1
        End If
        Termina = DateSerial(nYear, nMonth, nDay)

        If nDay <> 6 And nDay <> 16 And nDay <> 20 And nDay <> 25 Then
            Termina = IIf(Month(Termina) > nMonth, DateAdd(DateInterval.Month, -1, Termina), Termina)
            Termina = DayWeek(Termina)
        End If

    End Function

    Public Function DayWeek(ByVal dFecha As Date)

        Dim nDay As Byte
        Dim nYear As Integer
        Dim nMonth As Byte
        Dim nAños As Integer
        Dim nAñosb As Integer
        Dim nLeap As Byte
        Dim i As Integer
        Dim nMes As Integer
        Dim nDia As Integer

        nDay = Day(dFecha)
        nMonth = Month(dFecha)
        nYear = Year(dFecha)

        If nMonth = 12 Then
            nMonth = 1
            nYear += 1
        Else
            nMonth += 1
        End If

        dFecha = DateSerial(nYear, nMonth, 1)
        dFecha = DateAdd(DateInterval.Day, -1, dFecha)

        nDay = Day(dFecha)
        nMonth = Month(dFecha)
        nYear = Year(dFecha)

        nAños = nYear - 1933
        nLeap = 0
        nAñosb = 0

        For i = 1933 To nYear
            nLeap = Leap(i)
            If nLeap = 1 Then
                nAñosb += 1
                nLeap = 0
            End If
        Next

        Select Case nMonth
            Case 1, 10
                nMes = 0
            Case 2, 3, 11
                nMes = 3
            Case 4, 7
                nMes = 6
            Case 5
                nMes = 1
            Case 6
                nMes = 4
            Case 8
                nMes = 2
            Case 9, 12
                nMes = 5
        End Select

        If nMonth = 2 And nDay = 29 Then
            nDay = 28
        End If
        nDia = (nAños + nAñosb + nMes + nDay) Mod 7
        If nDia = 1 Then
            dFecha = DateAdd(DateInterval.Day, -2, dFecha)
        ElseIf nDia = 0 Then
            dFecha = DateAdd(DateInterval.Day, -1, dFecha)
        End If

        DayWeek = dFecha.ToShortDateString

    End Function

    Public Function Leap(ByVal nYear As Integer)

        If nYear Mod 400 = 0 Then
            Leap = 1
        ElseIf nYear Mod 100 = 0 Then
            Leap = 0
        ElseIf nYear Mod 4 = 0 Then
            Leap = 1
        End If

    End Function

    Public Function GeneraLetra(ByVal nLetra As Integer, ByRef cFeven As String, ByVal cFondeo As String)

        Dim cLetra As String
        Dim cNextMonth As String
        Dim nMonth As Integer
        Dim nNextMonth As Integer
        Dim nYear As Integer
        Dim dFeven As Date

        Select Case nLetra
            Case Is < 10
                cLetra = "00" + Trim(Str(nLetra))
            Case Is < 100
                cLetra = "0" + Trim(Str(nLetra))
            Case Else
                cLetra = Trim(Str(nLetra))
        End Select

        If nLetra > 1 Then

            nYear = Val(Left(cFeven, 4))
            nMonth = Val(Mid(cFeven, 5, 2))

            nNextMonth = nMonth

            nNextMonth += 1

            If nNextMonth > 12 Then
                nYear = nYear + Int(nNextMonth / 12)
                nNextMonth = nNextMonth - (Int(nNextMonth / 12) * 12)
            End If

            If nNextMonth < 10 Then
                cNextMonth = "0" + Trim(Str(nNextMonth))
            Else
                cNextMonth = Trim(Str(nNextMonth))
            End If

            cFeven = Trim(Str(nYear)) & cNextMonth & Right(cFeven, 2)

            If cNextMonth = "02" And cFondeo = "03" Then
                If Leap(nYear) = 1 Then
                    cFeven = Trim(Str(nYear)) & cNextMonth & "29"
                Else
                    cFeven = Trim(Str(nYear)) & cNextMonth & "28"
                End If
            End If

        End If

        GeneraLetra = cLetra

    End Function

    Public Function Mpt(ByVal Fecha1 As Date, ByVal Fecha2 As Date) As Integer
        Mpt = ((Year(Fecha1) * 12) + Month(Fecha1)) - ((Year(Fecha2) * 12) + Month(Fecha2))
    End Function

    Function SoloNumeros(ByVal Keyascii As Short, ByRef txtTexto As String) As Short
        If InStr("1234567890.,", Chr(Keyascii)) = 0 Then
            SoloNumeros = 0
        Else
            SoloNumeros = Keyascii
        End If
        If InStr(txtTexto, ".") - 1 > 0 And SoloNumeros = 46 Then
            SoloNumeros = 0
        End If
        Select Case Keyascii
            Case 8
                SoloNumeros = Keyascii
            Case 13
                SoloNumeros = Keyascii
        End Select
    End Function

    Public Function Adanterior(ByRef dtAdeudos As DataTable, ByVal drUdis As DataRowCollection, ByVal drFacturas As DataRowCollection, ByVal cFecha As String) As Decimal

        ' Declaración de Variables

        Dim drFactura As DataRow
        Dim drAnexo As DataRow
        Dim cAnexo As String
        Dim cFeven As String
        Dim cFepag As String
        Dim nSaldoFac As Decimal
        Dim nTasaMoratoria As Decimal
        Dim nMoratorios As Decimal
        Dim nAdeudoAnterior As Decimal
        Dim nIvaMoratorios As Decimal
        Dim nDiasVencido As Integer
        Dim nDiasMoratorios As Integer

        For Each drFactura In drFacturas
            cAnexo = drFactura("Anexo")
            nSaldoFac = drFactura("SaldoFac")
            nDiasVencido = DateDiff(DateInterval.Day, CTOD(drFactura("Feven")), CTOD(cFecha)) + 1
            nDiasMoratorios = 0
            nTasaMoratoria = Round((drFactura("Tasa") + drFactura("Difer")) * 2, 2)
            nMoratorios = 0

            cFeven = drFactura("Feven")
            cFepag = drFactura("fepag")

            If Trim(cFepag) = "" Then
                nDiasMoratorios = DateDiff(DateInterval.Day, CTOD(cFeven), CTOD(cFecha))
            Else
                If cFeven >= cFepag Then
                    nDiasMoratorios = DateDiff(DateInterval.Day, CTOD(cFeven), CTOD(cFecha))
                Else
                    nDiasMoratorios = DateDiff(DateInterval.Day, CTOD(cFepag), CTOD(cFecha))
                End If
            End If
            If nDiasMoratorios < 0 Then
                nDiasMoratorios = 0
            End If

            If nDiasMoratorios > 0 Then
                CalcMora(drUdis, nSaldoFac, nTasaMoratoria, nDiasMoratorios, nMoratorios, nIvaMoratorios, cFecha)
            End If

            If nDiasMoratorios > 0 Then
                nAdeudoAnterior = Round(nSaldoFac + nMoratorios + nIvaMoratorios, 2)
            Else
                nAdeudoAnterior = 0
            End If

            'Buscar si ya existe el anexo en la Tabla de Adeudos

            drAnexo = dtAdeudos.Rows.Find(cAnexo)

            If drAnexo Is Nothing Then
                drAnexo = dtAdeudos.NewRow()
                drAnexo("Anexo") = cAnexo
                drAnexo("AdeudoAnt") = nAdeudoAnterior
                dtAdeudos.Rows.Add(drAnexo)
            Else
                drAnexo("AdeudoAnt") += nAdeudoAnterior
            End If
        Next

    End Function

End Module
