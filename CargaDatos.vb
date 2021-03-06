Imports System.Data.SqlClient
Imports System.Net.Mail
Module CargaDatos

    Sub Main()
        Dim Args As String() = Environment.GetCommandLineArgs()
        If Args.Length = 1 Then
            Console.WriteLine("Sin Argumentos ...")
        Else
            If Args(1).ToUpper = "FACTORAJE" Then
                Console.WriteLine("Cargando pagos de Clientes Factoraje...")
                Carga_FactorajeCliente()
                Console.WriteLine("Cargando pagos de Factoraje a Finagil...")
                Carga_FactorajeANC()
                Carga_FactorajePALM()
            ElseIf Args(1).ToUpper = "MINDS" Then
                Console.WriteLine("Cargando promotores ...")
                Carga_Promotores()
                Console.WriteLine("Cargando clientes ...")
                Carga_Clientes()
                Console.WriteLine("Cargando clientes 2...")
                Carga_ClientesII()
                Console.WriteLine("Cargando cuentas ...")
                Carga_Cuentas()
            ElseIf Args(1).ToUpper = "MINDS2" Then
                Console.WriteLine("Cargando clientes 2...")
                Carga_ClientesII()
            End If
        End If
        Console.WriteLine("Terminado")
    End Sub

    Sub Carga_Promotores()
        Dim ta As New ProductionDataSetTableAdapters.PromotoresTableAdapter
        Dim ta1 As New Minds2DSTableAdapters.layoutsFuncionarioTableAdapter
        Dim PromoOrg As New ProductionDataSet.PromotoresDataTable
        Try
            ta.Fill(PromoOrg)

            For Each r As ProductionDataSet.PromotoresRow In PromoOrg.Rows
                If ta1.Existe(r.Promotor) = 0 Then
                    ta1.Insert(r.Promotor, Trim(r.Nombre), Trim(r.APaterno), Trim(r.AMaterno), Trim(r.Puesto), r.IDPlaza, r.Nacionalidad, CTOD(r.FechaCarga))
                Else
                    ta1.UpdateEmpleado(Trim(r.Nombre), Trim(r.APaterno), Trim(r.AMaterno), Trim(r.Puesto), r.IDPlaza, r.Nacionalidad, CTOD(r.FechaCarga), r.Promotor)
                End If
            Next
        Catch ex As Exception
            EnviaError("ecacerest@lamoderna.com.mx;viapolo@lamoderna.com.mx", ex.Message, "error en PROMOTORES")
        End Try
    End Sub

    Sub Carga_Clientes()
        Dim dsAgil As New DataSet()
        Dim Clientes As New Minds2DSTableAdapters.layoutsKYCTableAdapter
        Dim ClientesORG As New ProductionDataSetTableAdapters.ClientesTableAdapter
        Dim Municipio As New Minds2DSTableAdapters.Cat_MunicipioTableAdapter
        Dim TMunicipio As New Minds2DS.Cat_MunicipioDataTable
        Dim Estado As New Minds2DSTableAdapters.Cat_EstadoTableAdapter
        Dim TEstado As New Minds2DS.Cat_EstadoDataTable
        Dim cMuni As Double = 1
        Dim cEstado As Double = 0
        Dim xEstado As String = ""
        Dim xMuni As String = ""
        Dim cnAgil As New SqlConnection(My.Settings.ConnectionFinagil)
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim drCliente As DataRow
        Dim drDato As DataRow
        Dim drAnexos As DataRow()
        Dim drAnexo As DataRow
        Dim drPlaza As DataRow

        Dim cDia As String
        Dim i As Integer
        Dim cRenglon As String
        Dim cCliente As String
        Dim cDescr As String
        Dim cPromo As String
        'Dim cFecha As String
        Dim cGiro As String
        Dim cIdGiro As String
        Dim cProfGiro As String
        Dim cTipo As String
        Dim nCount As Integer
        Dim nDato As Integer
        Dim nIDEstado As Integer

        Dim aName As New ArrayList()
        Dim cDato As String
        Dim cNombre As String = ""
        Dim cApePaterno As String
        Dim cApeMaterno As String
        Dim cActivo As String = "2"
        Dim cDelegacion As String
        Dim IdSexo As String
        Dim nIdPlazam As Integer
        Dim FechaNac As Date

        Dim dsReporte As New DataSet()
        Dim daCliente As New SqlDataAdapter(cm1)
        Dim daAnexos As New SqlDataAdapter(cm2)
        Dim daPlazas As New SqlDataAdapter(cm3)
        Dim relAnexoCliente As DataRelation

        Try
            cDia = Mid(DTOC(Today), 7, 2) & Mid(DTOC(Today), 5, 2)
            cnAgil.Open()

            With cm1
                .CommandType = CommandType.Text
                .CommandText = "SELECT Descr,Cliente,Promo,Tipo,Calle, Colonia,Delegacion,Copos,Telef1,Giro,Clientes.Plaza, DescPlaza, RFC,Curp,Email1, Fecha1, NombreCliente, ApellidoPaterno, ApellidoMaterno, genero, FechaNac, Abreviado, Nombre_sucursal FROM Clientes " &
                "Inner Join Plazas ON Clientes.Plaza = Plazas.Plaza " &
                "Inner Join sucursales ON Clientes.sucursal = sucursales.id_sucursal " &
                "WHERE        (Cliente BETWEEN N'0' AND N'97337') " &
                "ORDER BY Cliente"
                .Connection = cnAgil
            End With
            daCliente.Fill(dsAgil, "Clientes")

            With cm2
                .CommandType = CommandType.Text
                .CommandText = "select Anexo, Fechacon, Flcan, Cliente FROM Anexos " &
                "UNION Select Anexo, FechaAutorizacion as Fechacon, Flcan, Cliente FROM Avios ORDER BY Cliente"
                .Connection = cnAgil
            End With
            daAnexos.Fill(dsAgil, "Anexos")
            relAnexoCliente = New DataRelation("AnexoCliente", dsAgil.Tables("Clientes").Columns("Cliente"), dsAgil.Tables("Anexos").Columns("Cliente"))
            dsAgil.EnforceConstraints = False
            dsAgil.Relations.Add(relAnexoCliente)
            nCount = 1
            For Each drCliente In dsAgil.Tables("Clientes").Rows
                cCliente = drCliente("Cliente")
                cApePaterno = ""
                cApeMaterno = ""
                cNombre = ""
                cTipo = drCliente("Tipo")
                cGiro = drCliente("Giro")
                cDelegacion = Trim(drCliente("Delegacion"))
                IdSexo = drCliente("Genero")
                nIdPlazam = 0

                Try
                    nIDEstado = Clientes.SacaIdEstado(drCliente("Abreviado").ToString.Trim)
                Catch ex As Exception
                    EnviaError("ecacerest@lamoderna.com.mx;viapolo@lamoderna.com.mx", ex.Message, "error en cLIENTES")
                End Try

                If nIdPlazam = 0 Then
                    Select Case nIDEstado
                        Case 1
                            nIdPlazam = 199009
                        Case 2
                            nIdPlazam = 299009
                        Case 3
                            nIdPlazam = 399009
                        Case 4
                            nIdPlazam = 499009
                        Case 5
                            nIdPlazam = 599009
                        Case 6
                            nIdPlazam = 699009
                        Case 7
                            nIdPlazam = 899009
                        Case 8
                            nIdPlazam = 999009
                        Case 9, 33
                            nIdPlazam = 1001002
                        Case 10
                            nIdPlazam = 1199008
                        Case 11
                            nIdPlazam = 1299003
                        Case 12
                            nIdPlazam = 1399007
                        Case 13
                            nIdPlazam = 1499002
                        Case 14
                            nIdPlazam = 1699001
                        Case 15
                            nIdPlazam = 1899009
                        Case 16
                            nIdPlazam = 2099007
                        Case 17
                            nIdPlazam = 2199006
                        Case 18
                            nIdPlazam = 2299005
                        Case 19
                            nIdPlazam = 2399004
                        Case 20
                            nIdPlazam = 2999007
                        Case 21
                            nIdPlazam = 3299006
                        Case 22
                            nIdPlazam = 3399009
                        Case 23
                            nIdPlazam = 3499003
                        Case 24
                            nIdPlazam = 3599006
                        Case 25
                            nIdPlazam = 3699009
                        Case 26
                            nIdPlazam = 3799003
                        Case 27
                            nIdPlazam = 3899006
                        Case 28
                            nIdPlazam = 3999009
                        Case 29
                            nIdPlazam = 4099001
                        Case 30
                            nIdPlazam = 4399004
                        Case 31
                            nIdPlazam = 4599009
                        Case 32
                            nIdPlazam = 4699007
                    End Select
                End If

                drAnexos = drCliente.GetChildRows("AnexoCliente")
                cPromo = drCliente("Promo")
                If drCliente("Tipo") = "F" Or drCliente("Tipo") = "E" Then
                    cDescr = Trim(drCliente("Descr"))
                    Dim texto() As String = Split(cDescr, " ")

                    nCount = 0
                    aName.Clear()
                    For i = 0 To UBound(texto)
                        aName.Add(texto(i))
                        nCount += 1
                    Next

                    i = 1
                    For Each cDato In aName
                        If aName.Count = 2 Then
                            If cNombre = "" Then
                                cNombre = cDato
                            Else
                                cApePaterno = cDato
                            End If
                        ElseIf i <= nCount - 2 Then
                            If cNombre = "" Then
                                cNombre = cDato
                            Else
                                cNombre = cNombre & " " & cDato
                            End If
                        ElseIf i = nCount - 1 Then
                            cApePaterno = cDato
                        ElseIf i = nCount Then
                            cApeMaterno = cDato
                        End If
                        i += 1
                    Next
                    If IdSexo.ToUpper.Trim = "FEMENINO" Then
                        IdSexo = 2
                    Else
                        IdSexo = 1
                    End If
                    cCliente = drCliente("Cliente")
                    cPromo = drCliente("Promo")
                Else
                    cNombre = Trim(drCliente("Descr"))
                    IdSexo = 0
                    cCliente = drCliente("Cliente")
                    cPromo = drCliente("Promo")

                End If

                nDato = 0
                cActivo = "2"
                For Each drAnexo In drAnexos
                    If nDato = 0 Then
                        'cFecha = drAnexo("Fechacon")
                    End If
                    If drAnexo("Flcan") = "A" Then
                        cActivo = "1"
                    End If
                    nDato += 1
                Next

                If Trim(cGiro) = "" Then
                    cGiro = "18"
                End If

                With cm3
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT IdGiro, ActividadEconomica FROM GirosMinds WHERE Giro = " & cGiro
                    .Connection = cnAgil
                End With
                daPlazas.Fill(dsAgil, "Giros")
                drPlaza = dsAgil.Tables("Giros").Rows(0)

                cIdGiro = drPlaza("IdGiro")
                cProfGiro = drPlaza("ActividadEconomica")
                dsAgil.Tables.Remove("Giros")

                If cTipo = "E" Then
                    cTipo = 3 'FISICA
                ElseIf cTipo = "F" Then
                    cTipo = 1
                Else
                    cTipo = 2
                End If

                Select Case Trim(drCliente("Cliente"))
                    Case "06671"
                        cNombre = "BERTA"
                        cApePaterno = "SANCHEZ"
                        cApeMaterno = "NO PROPORCIONADO"
                    Case "06370"
                        cNombre = cApePaterno
                        cApePaterno = cApeMaterno
                        cApeMaterno = "NO PROPORCIONADO"
                    Case "06790"
                        cNombre = cApePaterno
                        cApePaterno = cApeMaterno
                        cApeMaterno = "NO PROPORCIONADO"
                End Select

                If Trim(cPromo) <> "" Then
                    If Trim(cNombre) <> "" Then
                        If Len(cNombre) > 100 Then
                            cNombre = Mid(cNombre, 1, 100)
                        End If
                        xEstado = "%" & Trim(drCliente("DescPlaza")) & "%"
                        Select Case xEstado
                            Case "%ESTADO DE MEXICO%"
                                xEstado = "%MEXICO%"
                        End Select
                        Estado.FillByEstado(TEstado, xEstado)
                        If TEstado.Rows.Count > 0 Then
                            cEstado = TEstado.Rows(0).Item(0)
                        Else
                            cEstado = 0
                        End If
                        xMuni = "%" & Trim(drCliente("delegacion")) & "%"
                        Municipio.FillByMunicipio(TMunicipio, xMuni, nIDEstado)
                        If TMunicipio.Rows.Count > 0 Then
                            cMuni = TMunicipio.Rows(0).Item(0)
                        Else
                            cMuni = 0
                        End If
                        FechaNac = drCliente("FechaNac")
                        If Clientes.Exsiste(Trim(drCliente("Cliente"))).Value = 0 Then
                            Clientes.Insert(Trim(drCliente("Cliente")), cActivo, 0, 0, 0, 0, 0, 0, cPromo, "Credito", "", cIdGiro, 0, 0, 0, Trim(cNombre), Trim(cApePaterno), Trim(cApeMaterno), cProfGiro, drCliente("RFC"), cTipo, 1, Date.Now.ToShortDateString, Trim(drCliente("Calle")), 0, 0, Trim(drCliente("Colonia")), drCliente("Copos") _
                            , cMuni, nIdPlazam, nIDEstado, 236, 0, 0, 0, 0, 0, 0, 0, 0, Date.Now.ToShortDateString, 1, Trim(drCliente("CURP")), Trim(drCliente("Telef1")), 1, Val(cCliente), Date.Now.ToShortDateString, IdSexo, FechaNac.ToString("dd/MM/yyyy"), cEstado, "", Trim(drCliente("EMail1")), 236, 0, Date.Now.ToShortDateString, 2, Trim(drCliente("Nombre_Sucursal")))
                        Else
                            Clientes.UpdateKYC(cActivo, 0, 0, 0, 0, 0, 0, cPromo, "Credito", "", cIdGiro, 0, 0, 0, Trim(cNombre), Trim(cApePaterno), Trim(cApeMaterno), cProfGiro, drCliente("RFC"), cTipo, 1, Date.Now.ToShortDateString, Trim(drCliente("Calle")), 0, 0, Trim(drCliente("Colonia")), drCliente("Copos") _
                            , cMuni, nIdPlazam, nIDEstado, 236, 0, 0, 0, 0, 0, 0, 0, 0, Date.Now.ToShortDateString, 1, Trim(drCliente("CURP")), Trim(drCliente("Telef1")), 1, Val(cCliente), Date.Now.ToShortDateString, IdSexo, FechaNac.ToString("dd/MM/yyyy"), cEstado, "", Trim(drCliente("EMail1")), 236, 0, Date.Now.ToShortDateString, 2, Trim(drCliente("Nombre_Sucursal")), Trim(drCliente("Cliente")))
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            EnviaError("ecacerest@lamoderna.com.mx;viapolo@lamoderna.com.mx", ex.Message, "error en cLIENTES")
        End Try

        cnAgil.Close()
    End Sub

    Sub Carga_ClientesII()
        Dim cnAgil As SqlConnection = New SqlConnection(My.Settings.ConnectionFinagil)
        Dim cnAgil1 As SqlConnection = New SqlConnection(My.Settings.ConnectionMINDS)
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim daDatos As New SqlDataAdapter(cm1)
        Dim daEstado As New SqlDataAdapter(cm3)
        Dim drDato As DataRow
        Dim drMcpio As DataRow
        Dim drEdo As DataRow
        Dim dsAgil As DataSet = New DataSet()
        Dim nRows As Integer

        Dim strUpdate As String
        Dim cNext As String
        Dim cNint As String
        Dim cNac As String
        Dim cMcipio As String
        Dim nIdMcpio As Decimal
        Dim nIdEdo As Decimal
        Dim nIdEdoNAC As Decimal
        Dim nIdPais As Decimal
        Dim nIdPlaza As Decimal
        Dim xAux As String = ""
        Dim xPlaza As String = ""

        Dim Municipio As New Minds2DSTableAdapters.Cat_MunicipioTableAdapter
        Dim TMunicipio As New Minds2DS.Cat_MunicipioDataTable

        Dim Plazas As New Minds2DSTableAdapters.Cat_PlazaTableAdapter
        Dim TPlazas As New Minds2DS.Cat_PlazaDataTable

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT Datos_PLD.*, Clientes.Nacionalidad, PaisNacimiento, Fecha1, RFC, CURP, Genero, EMail1,SerieFiel, Telef1, descr, Correo " _
            & "FROM Datos_PLD INNER JOIN Clientes ON Clientes.Cliente = Datos_PLD.Cliente INNER JOIN Promotores ON Clientes.Promo = Promotores.Promotor " _
            & "where Clientes.cliente >= '0' " _
            & "ORDER BY Datos_PLD.Cliente"

            .Connection = cnAgil
        End With
        daDatos.Fill(dsAgil, "DatosPLD")

        With cm3
            .CommandType = CommandType.Text
            .CommandText = "SELECT Cat_Estado.* FROM [Cat_Estado]"
            .Connection = cnAgil1
        End With
        daEstado.Fill(dsAgil, "Estados")


        cnAgil1.Open()
        For Each drDato In dsAgil.Tables("DatosPLD").Rows
            If Trim(drDato("PLD_ClaveAE")) = "" Then
                xAux = "El Cliente " & Trim(drDato("descr")) & " no tiene asignada actividad economica."
                EnviaError("mtorres@finagil.com.mx,asangar@finagil.com.mx," & drDato("Correo"), xAux, "PLD_ClaveAE: " & Trim(drDato("descr")))
                EnviaError("ecacerest@lamoderna.com.mx;viapolo@lamoderna.com.mx", xAux, "PLD_ClaveAE: " & Trim(drDato("descr")))
                Continue For
            End If
            If IsDBNull(drDato("PLD_MontoMensual")) Then
                xAux = "El Cliente " & Trim(drDato("descr")) & " no capturado un monto Mensual"
                EnviaError("mtorres@finagil.com.mx,asangar@finagil.com.mx," & drDato("Correo"), xAux, "PLD_MontoMensual: " & Trim(drDato("descr")))
                EnviaError("ecacerest@lamoderna.com.mx;viapolo@lamoderna.com.mx", xAux, "PLD_MontoMensual: " & Trim(drDato("descr")))
                Continue For
            End If
            nIdEdoNAC = 0
            cMcipio = Trim(drDato("PLD_Estado"))
            For Each drEdo In dsAgil.Tables("Estados").Rows
                If SinAcentos(drEdo("Estado")) = SinAcentos(drDato("PLD_Estado")) Then
                    nIdEdo = drEdo("IdEstado")
                    nIdPais = drEdo("IdPais")
                End If
                If SinAcentos(drEdo("Estado")) = SinAcentos(drDato("PLD_EstadoNac")) Then
                    nIdEdoNAC = drEdo("IdEstado")
                End If
            Next

            xAux = "%" & Trim(drDato("pld_delegacion")) & "%"
            Municipio.FillByMunicipio(TMunicipio, xAux, nIdEdo) 'contiene acentos el DS
            If TMunicipio.Rows.Count > 0 Then
                cMcipio = TMunicipio.Rows(0).Item(0)
            Else
                cMcipio = 0
            End If
            If Trim(drDato("pld_ciudad")) = "" Then drDato("pld_ciudad") = drDato("pld_delegacion")
            xPlaza = "%" & Trim(drDato("pld_ciudad")) & "%"
            Plazas.Fill(TPlazas, nIdEdo, SinAcentos((xPlaza)))
            If TPlazas.Rows.Count > 0 Then
                nIdPlaza = TPlazas.Rows(0).Item(0)
            Else
                nIdPlaza = 99999999
            End If

            cNext = IIf(Trim(drDato("PLD_Numext")) = "", "", drDato("PLD_Numext"))
            cNint = IIf(Trim(drDato("PLD_Numint")) = "", "", drDato("PLD_Numint"))
            cNac = IIf(Trim(drDato("Nacionalidad")) = "", 1, IIf(Trim(drDato("Nacionalidad")) = "MEXICANA", 1, 2))
            strUpdate = "UPDATE layoutsKYC SET DirNo = '" & cNext & "'"
            strUpdate = strUpdate & ", Interior = '" & cNint & "'"
            strUpdate = strUpdate & ", IdNacionalidad = '" & cNac & "'"
            strUpdate = strUpdate & ", Calle = '" & drDato("PLD_Calle") & "'"
            strUpdate = strUpdate & ", Colonia = '" & drDato("PLD_Asentamiento") & "'"
            strUpdate = strUpdate & ", CP = '" & drDato("PLD_Copos") & "'"
            strUpdate = strUpdate & ", IdEstado = '" & nIdEdo & "'"
            strUpdate = strUpdate & ", IdPais = '" & nIdPais & "'"
            strUpdate = strUpdate & ", RFC = '" & drDato("RFC") & "'"
            strUpdate = strUpdate & ", idMunicipio = " & cMcipio & ""
            strUpdate = strUpdate & ", idplaza = " & nIdPlaza & ""
            strUpdate = strUpdate & ", CURP = '" & drDato("CURP") & "'"
            strUpdate = strUpdate & ", IdActividadEconomica = '" & drDato("PLD_ClaveAE") & "'"
            strUpdate = strUpdate & ", FIEL = '" & drDato("SerieFiel") & "'"
            strUpdate = strUpdate & ", Correo = '" & drDato("EMail1") & "'"
            strUpdate = strUpdate & ", Telefono = '" & drDato("Telef1") & "'"
            strUpdate = strUpdate & " WHERE nic = '" & drDato("Cliente") & "'"
            Try
                cm2 = New SqlCommand(strUpdate, cnAgil1)
                cm2.ExecuteNonQuery()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                EnviaError("ecacerest@lamoderna.com.mx;viapolo@lamoderna.com.mx", ex.Message, "error en cLIENTES")
            End Try
            If nIdEdoNAC > 0 Then
                strUpdate = "UPDATE layoutsKYC SET id_estadoNacimiento = " & nIdEdoNAC & "  WHERE nic = '" & drDato("Cliente") & "'"
                cm2 = New SqlCommand(strUpdate, cnAgil1)
                cm2.ExecuteNonQuery()
            End If

        Next
        'MsgBox("Datos Actualizados Correctamente", MsgBoxStyle.Information, "Mensaje del Sistema")
    End Sub

    Sub Carga_Cuentas()
        Dim Con1 As New ProductionDataSetTableAdapters.AnexosTableAdapter
        Dim Con2 As New ProductionDataSetTableAdapters.AviosTableAdapter
        Dim dsAgil As New DataSet()
        Dim Pagos As New Minds2DSTableAdapters.layoutsCreditoTableAdapter
        Dim Cuentas As New Minds2DSTableAdapters.layoutsCuentaTableAdapter
        Dim cnAgil As New SqlConnection(My.Settings.ConnectionFinagil)
        Dim cm1 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim cm4 As New SqlCommand()
        Dim drAnexo As DataRow
        Dim drEdoctav As DataRow()
        Dim drDato As DataRow

        Dim cDia As String
        Dim i As Integer
        Dim cMes As String = Date.Now.AddMonths(-1).ToString("yyyyMM") & "%"
        Dim cImporte As String
        Dim cAnexo As String
        Dim cCiclo As String
        Dim cCliente As String
        Dim cFecha As String
        Dim nCount As Integer
        Dim nPago As Decimal
        Dim cProduct As String
        Dim cSucursal As String
        Dim ID_Frecuencia As Integer

        Dim dsReporte As New DataSet()
        Dim daAnexos As New SqlDataAdapter(cm1)
        Dim daAvios As New SqlDataAdapter(cm3)
        Dim daCuentasConcetradoras As New SqlDataAdapter(cm4)
        Dim relAnexoEdoctav As DataRelation

        Try


            cDia = Mid(DTOC(Today), 7, 2) & Mid(DTOC(Today), 5, 2)
            cFecha = DTOC(Today)
            cnAgil.Open()

            With cm1
                .CommandType = CommandType.Text
                .CommandText = "SELECT * FROM Minds_Cuentas "
                .Connection = cnAgil
            End With

            With cm3
                .CommandType = CommandType.Text
                .CommandText = "SELECT * FROM Minds_CuentasAvio"
                .Connection = cnAgil
            End With

            With cm4
                .CommandType = CommandType.Text
                .CommandText = "SELECT * FROM Minds_CuentasConcetradoras where Anexo = 'XX'"
                .Connection = cnAgil
            End With

            ' Este Stored Procedure trae la tabla de amortizaci�n del equipo de todos los contratos activos
            ' con fecha de contrataci�n menor o igual a la de proceso

            daAnexos.Fill(dsAgil, "Anexos")
            daAvios.Fill(dsAgil, "Avios")

            'relAnexoEdoctav = New DataRelation("AnexoEdoctav", dsAgil.Tables("Anexos").Columns("Anexo"), dsAgil.Tables("Edoctav").Columns("Anexo"))
            'dsAgil.EnforceConstraints = False
            'dsAgil.Relations.Add(relAnexoEdoctav)

            For Each drAnexo In dsAgil.Tables("Anexos").Rows
                cAnexo = drAnexo("Anexo")
                'nPago = CDec(drAnexo("PLD_MontoMensul") * 3) ' solicitado por KArla Sanchez 31/05/2019
                nPago = (Con1.PagoMensual(cAnexo, cMes) * 3) ' solicitado por KArla Sanchez 13/04/2020
                cCliente = drAnexo("Cliente")
                cSucursal = drAnexo("Mensu").ToString
                cImporte = drAnexo("MtoFin").ToString
                cFecha = CTOD(drAnexo("Fechacon")).ToShortDateString
                drEdoctav = drAnexo.GetChildRows("AnexoEdoctav")
                Select Case UCase(drAnexo("Vencimiento"))
                    Case "SEMANAL"
                        ID_Frecuencia = 1
                    Case "CATORCENAL"
                        ID_Frecuencia = 2
                    Case "QUINCENAL"
                        ID_Frecuencia = 3
                    Case "MENSUAL"
                        ID_Frecuencia = 4
                    Case "BIMESTRAL"
                        ID_Frecuencia = 5
                    Case "TRIMESTRAL", "TRIMESTRE"
                        ID_Frecuencia = 6
                    Case "SEMESTRAL"
                        ID_Frecuencia = 7
                    Case "ANUAL"
                        ID_Frecuencia = 8
                End Select
                Select Case drAnexo("Tipar")
                    Case "F"
                        cProduct = "1"
                    Case "P"
                        cProduct = "2"
                    Case "R"
                        cProduct = "8"
                    Case "L"
                        cProduct = "11"
                    Case "B"
                        cProduct = "12"
                    Case "S"
                        If drAnexo("LiquidezInmediata") = True Then
                            cProduct = "11"
                        Else
                            cProduct = "3"
                        End If
                End Select
                nCount = 0
                Try
                    If Cuentas.Existe(cAnexo).Value = 0 Then
                        Cuentas.Insert(cAnexo, cCliente, 7, cProduct, cImporte, cFecha, drAnexo("Feven"), 1, nPago.ToString, 1, ID_Frecuencia)
                    Else
                        Cuentas.UpdateCuenta(cCliente, 7, cProduct, cImporte, cFecha, drAnexo("Feven"), 1, nPago.ToString, 1, ID_Frecuencia, cAnexo)
                        Cuentas.UpdateMensualidad(nPago.ToString, cProduct, cAnexo)
                    End If
                Catch ex As Exception
                    EnviaError("ecacerest@lamoderna.com.mx;viapolo@lamoderna.com.mx", ex.Message & "  " & cAnexo, "error en CUENTAS")
                End Try
            Next

            For Each drAnexo In dsAgil.Tables("Avios").Rows
                If "081640011" = drAnexo("Anexo") Then
                    cAnexo = drAnexo("Anexo")
                End If
                cAnexo = drAnexo("Anexo")
                cCliente = drAnexo("Cliente")
                cImporte = drAnexo("LineaActual").ToString
                cFecha = CTOD(drAnexo("FechaAutorizacion")).ToShortDateString
                Select Case drAnexo("Tipar")
                    Case "A"
                        cProduct = "3" ' como simple
                    Case "C"
                        cProduct = "4"
                    Case "H"
                        cProduct = "9"
                End Select

                Select Case UCase(drAnexo("Vencimiento"))
                    Case "SEMANAL"
                        ID_Frecuencia = 1
                    Case "CATORCENAL"
                        ID_Frecuencia = 2
                    Case "QUINCENAL"
                        ID_Frecuencia = 3
                    Case "MENSUAL"
                        ID_Frecuencia = 4
                    Case "BIMESTRAL"
                        ID_Frecuencia = 5
                    Case "TRIMESTRAL", "TRIMESTRE"
                        ID_Frecuencia = 6
                    Case "SEMESTRAL"
                        ID_Frecuencia = 7
                    Case "ANUAL"
                        ID_Frecuencia = 8
                End Select

                nPago = drAnexo("PLD_MontoMensual") * 3 ' solicitado por KArla Sanchez 31/05/2019
                nPago = drAnexo("LineaActual") * 3 ' solicitado por KArla Sanchez 13/04/2020
                If drAnexo("Tipar") <> "AA" Then
                    If Cuentas.Existe(cAnexo).Value = 0 Then
                        Cuentas.Insert(cAnexo, cCliente, 7, cProduct, cImporte, cFecha, drAnexo("Feven"), 1, nPago.ToString, 1, ID_Frecuencia)
                    Else
                        Cuentas.UpdateCuenta(cCliente, 7, cProduct, cImporte, cFecha, drAnexo("Feven"), 1, nPago.ToString, 1, ID_Frecuencia, cAnexo)
                    End If
                    cAnexo = Mid(cAnexo, 1, 9)
                    cCiclo = Mid(cAnexo, 11, 2)
                    Con2.UpdateMinds(cCiclo, cAnexo)
                End If
            Next
        Catch ex As Exception
            EnviaError("ecacerest@lamoderna.com.mx;viapolo@lamoderna.com.mx", ex.Message & "  " & cAnexo, "error en CUENTAS")
        End Try
        cnAgil.Close()
    End Sub

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        If InStr(Mensaje, "No se ha encontrado la ruta de acceso de la red") = 0 Then
            Dim Mensage As New MailMessage("MINDS@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
            Dim Cliente As New SmtpClient(My.Settings.SMTP, My.Settings.SMTP_port)
            Try
                Dim Credenciales As String() = My.Settings.SMTP_creden.Split(",")
                Cliente.Credentials = New System.Net.NetworkCredential(Credenciales(0), Credenciales(1), Credenciales(2))
                Cliente.Send(Mensage)
            Catch ex As Exception
                'ReportError(ex)
            End Try
        Else
            Console.WriteLine("No se ha encontrado la ruta de acceso de la red")
        End If
    End Sub

    Function SinAcentos(ByVal Cad As String) As String
        Cad = Trim(Cad.ToUpper)
        Cad = Cad.Replace("�", "A")
        Cad = Cad.Replace("�", "E")
        Cad = Cad.Replace("�", "I")
        Cad = Cad.Replace("�", "O")
        Cad = Cad.Replace("�", "U")
        Return Cad
    End Function

    Sub Carga_FactorajeCliente()
        Dim tSaldos As New Factor100DSTableAdapters.FacturasConSaldoTableAdapter
        Dim tSal As New Factor100DS.FacturasConSaldoDataTable
        Dim tPag As New Factor100DSTableAdapters.PagosTableAdapter
        Dim tBaan As New BaanDSTableAdapters.PagosBaanTableAdapter
        Dim tBaa As New BaanDS.PagosBaanDataTable
        Dim fac As Factor100DS.FacturasConSaldoRow
        Dim pag As BaanDS.PagosBaanRow
        Dim Serie As String
        Dim Factura As Decimal

        tSaldos.Fill(tSal)
        For Each fac In tSal.Rows
            Serie = Mid(fac.Factura, 1, 3)
            Factura = Mid(fac.Factura, 4, 10)
            tBaan.Fill(tBaa, Serie, Factura)
            For Each pag In tBaa.Rows
                If tPag.ExistePago(fac.Factura, pag.fecha, pag.linea, pag.tipo) <= 0 Then
                    tPag.Insert(fac.Factura, pag.fecha, pag.importe * -1, pag.tipo, pag.linea, False, False)
                End If
            Next
        Next
    End Sub

    Sub Carga_FactorajePALM()
        Dim tPagFactor As New Factor100DSTableAdapters.PagosTableAdapter
        Dim tPAGOS As New BaanDSTableAdapters.PagosPALMFinagilTableAdapter
        Dim tPAG As New BaanDS.PagosPALMFinagilDataTable
        Dim tBaan As New BaanDSTableAdapters.PagosBaanTableAdapter
        Dim fac As BaanDS.PagosPALMFinagilRow

        For Each fac In tPAG.Rows
            'Serie = Mid(fac.factura, 1, 3)
            'Factura = Mid(fac.factura, 4, 10)
            'If "ETP256463" = fac.factura.Trim Then
            'fac.factura = fac.factura.Trim
            'End If

            If fac.t_ccur = "USD" Then
                If tPagFactor.ExistePago(fac.factura.Trim & " USD", fac.fecha, 1, 10) <= 0 Then
                    tPagFactor.Insert(fac.factura.Trim & " USD", fac.fecha, Math.Abs(fac.importe), 10, 1, False, True)
                End If
            Else
                'If "EBA6661" = fac.factura.Trim Then
                '    fac.factura = fac.factura.Trim
                'End If
                If tPagFactor.ExistePago(fac.factura.Trim, fac.fecha, 1, 20) <= 0 Then
                    tPagFactor.Insert(fac.factura.Trim, fac.fecha, Math.Abs(fac.importe), 20, 1, False, True)
                End If
            End If
        Next
        Dim tPAGOS2 As New BaanDSTableAdapters.SugerenciaBaanTableAdapter
        Dim tPAG2 As New BaanDS.SugerenciaBaanDataTable
        Dim tBaan2 As New BaanDSTableAdapters.SugerenciaBaanTableAdapter
        Dim fac2 As BaanDS.SugerenciaBaanRow

        tPAGOS2.Fill(tPAG2, Date.Now.AddDays(-30))
        For Each fac2 In tPAG2.Rows
            'Serie = Mid(fac.factura, 1, 3)
            'Factura = Mid(fac.factura, 4, 10)
            If "ETP278626" = fac2.factura.Trim Then
                fac2.factura = fac2.factura.Trim
            End If


            If fac2.t_ccur = "USD" Then
                If tPagFactor.ExistePago(Mid(fac2.factura.Trim, 1, 16) & " USD", fac2.fecha, fac2.numpago, 100) <= 0 Then
                    tPagFactor.Insert(Mid(fac2.factura.Trim, 1, 16) & " USD", fac2.fecha, Math.Abs(fac2.importe), 100, fac2.numpago, False, True)
                End If
            Else
                'If "EBA6661" = fac2.factura.Trim Then
                'fac2.factura = fac2.factura.Trim
                'End If
                If tPagFactor.ExistePago(Mid(fac2.factura.Trim, 1, 20), fac2.fecha, fac2.numpago, 200) <= 0 Then
                    tPagFactor.Insert(Mid(fac2.factura.Trim, 1, 20), fac2.fecha, Math.Abs(fac2.importe), 200, fac2.numpago, False, True)
                End If
            End If
        Next

        Dim diasMenos As Integer = (-365)
        Dim ta As New Factor100DSTableAdapters.CancelacionesTableAdapter


        Dim tCancel200 As New BaanDSTableAdapters.Cancelaciones200TableAdapter
        Dim tCAN200 As New BaanDS.Cancelaciones200DataTable
        Dim tBaan200 As New BaanDSTableAdapters.Cancelaciones200TableAdapter
        Dim Can200 As BaanDS.Cancelaciones200Row
        tCancel200.CommandTimeout = 60
        tCancel200.Fill(tCAN200, Date.Now.AddDays(diasMenos))
        For Each Can200 In tCAN200.Rows
            If ta.ExisteCancelacion(Can200.t_ttyp, Can200.t_invn) <= 0 Then
                ta.Insert(Can200.t_ttyp, Can200.t_invn, Can200.t_refb.Substring(0, 14), False, Can200.t_odat)
            End If
        Next

        Dim tCancel205 As New BaanDSTableAdapters.Cancelaciones205TableAdapter
        Dim tCAN205 As New BaanDS.Cancelaciones205DataTable
        Dim tBaan205 As New BaanDSTableAdapters.Cancelaciones205TableAdapter
        Dim Can205 As BaanDS.Cancelaciones205Row
        tCancel205.CommandTimeout = 60
        tCancel205.Fill(tCAN205, Date.Now.AddDays(diasMenos))
        For Each Can205 In tCAN205.Rows
            If ta.ExisteCancelacion(Can205.t_ttyp, Can205.t_invn) <= 0 Then
                ta.Insert(Can205.t_ttyp, Can205.t_invn, Can205.t_refb.Substring(0, 14), False, Can205.t_odat)
            End If
        Next

        Dim tCancel208 As New BaanDSTableAdapters.Cancelaciones208TableAdapter
        Dim tCAN208 As New BaanDS.Cancelaciones208DataTable
        Dim tBaan208 As New BaanDSTableAdapters.Cancelaciones208TableAdapter
        Dim Can208 As BaanDS.Cancelaciones208Row
        tCancel208.CommandTimeout = 60
        tCancel208.Fill(tCAN208, Date.Now.AddDays(diasMenos))
        For Each Can208 In tCAN208.Rows
            If ta.ExisteCancelacion(Can208.t_ttyp, Can208.t_invn) <= 0 Then
                ta.Insert(Can208.t_ttyp, Can208.t_invn, Can208.t_refb.Substring(0, 14), False, Can208.t_odat)
            End If
        Next

        Dim tCancel209 As New BaanDSTableAdapters.Cancelaciones209TableAdapter
        Dim tCAN209 As New BaanDS.Cancelaciones209DataTable
        Dim tBaan209 As New BaanDSTableAdapters.Cancelaciones209TableAdapter
        Dim Can209 As BaanDS.Cancelaciones209Row
        tCancel209.CommandTimeout = 60
        tCancel209.Fill(tCAN209, Date.Now.AddDays(diasMenos))
        For Each Can209 In tCAN209.Rows
            If ta.ExisteCancelacion(Can209.t_ttyp, Can209.t_invn) <= 0 Then
                ta.Insert(Can209.t_ttyp, Can209.t_invn, Can209.t_refb.Substring(0, 14), False, Can209.t_odat)
            End If
        Next

        Dim tCancel206 As New BaanDSTableAdapters.Cancelaciones206TableAdapter
        Dim tCAN206 As New BaanDS.Cancelaciones206DataTable
        Dim tBaan206 As New BaanDSTableAdapters.Cancelaciones206TableAdapter
        Dim Can206 As BaanDS.Cancelaciones206Row
        tCancel206.CommandTimeout = 60
        tCancel206.Fill(tCAN206, Date.Now.AddDays(diasMenos))
        For Each Can206 In tCAN206.Rows
            If ta.ExisteCancelacion(Can206.t_ttyp, Can206.t_invn) <= 0 Then
                ta.Insert(Can206.t_ttyp, Can206.t_invn, Can206.t_refb.Substring(0, 14), False, Can206.t_odat)
            End If
        Next

        Dim tCancel207 As New BaanDSTableAdapters.Cancelaciones207TableAdapter
        Dim tCAN207 As New BaanDS.Cancelaciones207DataTable
        Dim tBaan207 As New BaanDSTableAdapters.Cancelaciones207TableAdapter
        Dim Can207 As BaanDS.Cancelaciones207Row
        tCancel207.CommandTimeout = 60
        tCancel207.Fill(tCAN207, Date.Now.AddDays(diasMenos))
        For Each Can207 In tCAN207.Rows
            If ta.ExisteCancelacion(Can207.t_ttyp, Can207.t_invn) <= 0 Then
                ta.Insert(Can207.t_ttyp, Can207.t_invn, Can207.t_refb.Substring(0, 14), False, Can207.t_odat)
            End If
        Next
        ta.QuitaGuiones()
    End Sub

    Sub Carga_FactorajeANC()
        Dim tSaldos As New Factor100DSTableAdapters.FacturasConSaldoTableAdapter
        Dim tSal As New Factor100DS.FacturasConSaldoDataTable
        Dim tCancel As New Factor100DSTableAdapters.CancelacionesTableAdapter
        Dim tBaan As New BaanDSTableAdapters.PagosBaanTableAdapter
        Dim tBaa As New BaanDS.PagosBaanDataTable
        Dim fac As Factor100DS.FacturasConSaldoRow
        Dim pag As BaanDS.PagosBaanRow
        Dim Serie As String
        Dim Factura As Decimal

        tSaldos.FillBy90(tSal)
        For Each fac In tSal.Rows
            Serie = Mid(fac.Factura, 1, 3)
            Factura = Mid(fac.Factura, 4, 10)
            tBaan.FillByANC(tBaa, Serie, Factura)
            For Each pag In tBaa.Rows
                If tCancel.ExisteCancelacion(pag.serie, pag.factura) <= 0 And pag.t_tdoc = "ANC" Then
                    tCancel.Insert(pag.serie, pag.factura, Serie.Trim & Factura, False, pag.fecha)
                End If
            Next
        Next


    End Sub
End Module
