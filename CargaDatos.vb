Imports System.Data.SqlClient
Imports System.Net.Mail
Module CargaDatos
    Dim strConn As String = "Server=SERVER-RAID; DataBase=production; User ID=User_PRO; pwd=User_PRO2015"
    Dim strConn2 As String = "Server=SERVER-MINDS\MINDS; DataBase=PrevencionLavadoDinero; User ID=finagil; pwd=finagil"

    Sub Main()
        Console.WriteLine("Cargando promotores ...")
        Carga_Promotores()
        Console.WriteLine("Cargando clientes ...")
        Carga_Clientes()
        Console.WriteLine("Cargando clientes 2...")
        Carga_ClientesII()
        Console.WriteLine("Cargando cuentas ...")
        Carga_Cuentas()
        Console.WriteLine("Cargando pagos...")
        Carga_Pagos()
        'Console.WriteLine("Cargando pagos de Clientes Factoraje...")
        Carga_FactorajeCliente()
        'Console.WriteLine("Cargando pagos de Factoraje a Finagil...")
        Carga_FactorajeANC()
        Carga_FactorajePALM()



        Console.WriteLine("Terminado")
    End Sub

    Sub Carga_Promotores()
        Dim ta As New CargaMINDS.ProductionDataSetTableAdapters.PromotoresTableAdapter
        Dim ta1 As New CargaMINDS.Minds2DSTableAdapters.layoutsFuncionarioTableAdapter
        Dim PromoOrg As New ProductionDataSet.PromotoresDataTable
        Try
            ta.Fill(PromoOrg)

            For Each r As ProductionDataSet.PromotoresRow In PromoOrg.Rows
                If ta1.Existe(r.Promotor, Trim(r.APaterno)).Value = 0 Then
                    ta1.Insert(r.Promotor, Trim(r.Nombre), Trim(r.APaterno), Trim(r.AMaterno), Trim(r.Puesto), r.IDPlaza, r.Nacionalidad, CTOD(r.FechaCarga))
                Else
                    ta1.UpdateEmpleado(Trim(r.Nombre), Trim(r.APaterno), Trim(r.AMaterno), Trim(r.Puesto), r.IDPlaza, r.Nacionalidad, CTOD(r.FechaCarga), r.Promotor)
                End If
            Next
        Catch ex As Exception
            EnviaError("viapolo@lamoderna.com.mx,viapolo@lamoderna.com.mx", ex.Message, "error en PROMOTORES")
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

        'Clientes.DeleteAll()

        Dim cnAgil As New SqlConnection(strConn)
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

        Dim dsReporte As New DataSet()
        Dim daCliente As New SqlDataAdapter(cm1)
        Dim daAnexos As New SqlDataAdapter(cm2)
        Dim daPlazas As New SqlDataAdapter(cm3)
        Dim relAnexoCliente As DataRelation

        Try



            cDia = Mid(DTOC(Today), 7, 2) & Mid(DTOC(Today), 5, 2)
            'cFecha = dtpProcesar.Value
            cnAgil.Open()

            With cm1
                .CommandType = CommandType.Text
                .CommandText = "SELECT Descr,Cliente,Promo,Tipo,Calle, Colonia,Delegacion,Copos,Telef1,Giro,Clientes.Plaza, DescPlaza, RFC,Curp,Email1, Fecha1, NombreCliente, ApellidoPaterno, ApellidoMaterno, genero FROM Clientes " &
                "Inner Join Plazas ON Clientes.Plaza = Plazas.Plaza " &
                "WHERE        (Cliente BETWEEN N'0' AND N'97337') " &
                "ORDER BY Cliente"


                '"Inner Join Plazas ON Clientes.Plaza = Plazas.Plaza where clientes.cliente = '06671' ORDER BY Cliente"




                '.CommandText = "SELECT Descr,Cliente,Promo,Tipo,Calle, Colonia,Delegacion,Copos,Telef1,Giro,Clientes.Plaza, DescPlaza, RFC,Curp,Email1, Fecha1, NombreCliente, ApellidoPaterno, ApellidoMaterno FROM Clientes " & _
                '"Inner Join Plazas ON Clientes.Plaza = Plazas.Plaza where siebel = 0 or siebel is null ORDER BY Cliente"
                .Connection = cnAgil
            End With
            daCliente.Fill(dsAgil, "Clientes")

            With cm2
                .CommandType = CommandType.Text
                .CommandText = "select Anexo, Fechacon, Flcan, Cliente FROM Anexos " &
                "UNION Select Anexo, FechaAutorizacion as Fechacon, Flcan, Cliente FROM Avios ORDER BY Cliente"
                '"UNION Select Anexo, FechaAutorizacion as Fechacon, Flcan, Cliente FROM Avios where cliente = '05869' ORDER BY Cliente"




                .Connection = cnAgil
            End With
            daAnexos.Fill(dsAgil, "Anexos")

            ' Establecer la relación entre Anexos y Clientes

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
                nIDEstado = drCliente("Plaza")
                IdSexo = drCliente("Genero")
                nIdPlazam = 0
                Select Case drCliente("Plaza")
                    Case Is = "07"
                        nIDEstado = 5
                    Case Is = "08"
                        nIDEstado = 6
                    Case Is = "05"
                        nIDEstado = 7
                    Case Is = "06"
                        nIDEstado = 8
                    Case Is = "12"
                        nIDEstado = 11
                    Case Is = "13"
                        nIDEstado = 12
                    Case Is = "14"
                        nIDEstado = 13
                    Case Is = "15"
                        nIDEstado = 14
                    Case Is = "11"
                        nIDEstado = 15
                    Case Is = "33"
                        nIDEstado = 9
                End Select

                With cm3
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT IdPlazam, NombrePlaza FROM PlazasMinds WHERE IdEstado = " & nIDEstado
                    .Connection = cnAgil
                End With
                daPlazas.Fill(dsAgil, "Plazas")


                For Each drPlaza In dsAgil.Tables("Plazas").Rows
                    If Trim(cDelegacion) = Trim(drPlaza("NombrePlaza")) Then
                        nIdPlazam = drPlaza("IdPlazam")
                    End If
                Next
                dsAgil.Tables.Remove("Plazas")

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
                        'cRenglon = cActivo & "|" & cPromo & "|0|0|0|0|0|0|" & cIdGiro & "|0||0|0|N|N|" & cNombre & "|" & cApePaterno & "|"
                        'cRenglon = cRenglon & cApeMaterno & "|" & cProfGiro & "|" & drCliente("RFC") & "|" & cTipo & "|1|" & CTOD(cFecha).ToShortDateString & "|" & Trim(drCliente("Calle")) & "|0|0|" & Trim(drCliente("Colonia"))
                        'cRenglon = cRenglon & "|" & drCliente("Copos") & "|" & Trim(drCliente("Delegacion")) & "|" & nIdPlazam & "|" & nIDEstado & "|237|0|0|0|0|0|0|0|0|" & (dtpProcesar.Value).ToShortDateString & "|1|" & Trim(drCliente("CURP")) & "|"
                        'cRenglon = cRenglon & Trim(drCliente("Telef1")) & "|1|" & cCliente & "|" & Val(cCliente) & "|" & (dtpProcesar.Value).ToShortDateString & "|0|" & CTOD(drCliente("Fecha1")).ToShortDateString & "|" & Trim(drCliente("DescPlaza")) & "||" & Trim(drCliente("EMail1")) & "|237|"
                        'stmWriter.WriteLine(cRenglon)
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
                        If Clientes.Exsiste(Trim(drCliente("Cliente"))).Value = 0 Then
                            Clientes.Insert(Trim(drCliente("Cliente")), cActivo, 0, 0, 0, 0, 0, 0, cPromo, "Credito", "", cIdGiro, 0, 0, 0, Trim(cNombre), Trim(cApePaterno), Trim(cApeMaterno), cProfGiro, drCliente("RFC"), cTipo, 1, Date.Now.ToShortDateString, Trim(drCliente("Calle")), 0, 0, Trim(drCliente("Colonia")), drCliente("Copos") _
                            , cMuni, nIdPlazam, nIDEstado, 236, 0, 0, 0, 0, 0, 0, 0, 0, Date.Now.ToShortDateString, 1, Trim(drCliente("CURP")), Trim(drCliente("Telef1")), 1, Val(cCliente), Date.Now.ToShortDateString, IdSexo, CTOD(drCliente("Fecha1")).ToShortDateString, cEstado, "", Trim(drCliente("EMail1")), 236, 0, Date.Now.ToShortDateString, 2)
                        Else
                            Clientes.UpdateKYC(cActivo, 0, 0, 0, 0, 0, 0, cPromo, "Credito", "", cIdGiro, 0, 0, 0, Trim(cNombre), Trim(cApePaterno), Trim(cApeMaterno), cProfGiro, drCliente("RFC"), cTipo, 1, Date.Now.ToShortDateString, Trim(drCliente("Calle")), 0, 0, Trim(drCliente("Colonia")), drCliente("Copos") _
                            , cMuni, nIdPlazam, nIDEstado, 236, 0, 0, 0, 0, 0, 0, 0, 0, Date.Now.ToShortDateString, 1, Trim(drCliente("CURP")), Trim(drCliente("Telef1")), 1, Val(cCliente), Date.Now.ToShortDateString, IdSexo, CTOD(drCliente("Fecha1")).ToShortDateString, cEstado, "", Trim(drCliente("EMail1")), 236, 0, Date.Now.ToShortDateString, 2, Trim(drCliente("Cliente")))
                        End If

                        If cTipo = "2" Then
                            cRenglon = cIdGiro & "|" & cActivo & "|1|" & cNombre & "||||" & drCliente("RFC") & "||0|" & Trim(drCliente("Telef1")) & "||||||||||||" & cCliente & "|"
                        Else
                            cRenglon = cIdGiro & "|" & cActivo & "|1||" & cNombre & "|" & cApePaterno & "|" & cApeMaterno & "|" & drCliente("RFC") & "||0|" & Trim(drCliente("Telef1")) & "||||||||||||" & cCliente & "|"
                        End If
                        'DataGridView1.DataSource = dsAgil

                    End If
                End If
            Next
        Catch ex As Exception
            EnviaError("viapolo@lamoderna.com.mx,viapolo@lamoderna.com.mx", ex.Message, "error en cLIENTES")
        End Try

        cnAgil.Close()
    End Sub

    Sub Carga_ClientesII()
        Dim cnAgil As SqlConnection = New SqlConnection(strConn)
        Dim cnAgil1 As SqlConnection = New SqlConnection(strConn2)
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
        Dim xMuni As String = ""
        Dim xPlaza As String = ""

        Dim Municipio As New Minds2DSTableAdapters.Cat_MunicipioTableAdapter
        Dim TMunicipio As New Minds2DS.Cat_MunicipioDataTable

        Dim Plazas As New Minds2DSTableAdapters.Cat_PlazaTableAdapter
        Dim TPlazas As New Minds2DS.Cat_PlazaDataTable

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT Datos_PLD.*, Nacionalidad, PaisNacimiento, Fecha1, RFC, CURP, Genero, EMail1,SerieFiel, Telef1 FROM Datos_PLD INNER JOIN Clientes ON Clientes.Cliente = Datos_PLD.Cliente ORDER BY Cliente"
            '.CommandText = "SELECT Datos_PLD.*, Nacionalidad, PaisNacimiento, Fecha1, RFC, CURP, Genero, EMail1,SerieFiel, Telef1 FROM Datos_PLD INNER JOIN Clientes ON Clientes.Cliente = Datos_PLD.Cliente  where clientes.cliente = '06671' ORDER BY Cliente"

            .Connection = cnAgil
        End With
        daDatos.Fill(dsAgil, "DatosPLD")

        With cm3
            .CommandType = CommandType.Text
            .CommandText = "SELECT Cat_Estado.* FROM [PrevencionLavadoDinero].[dbo].[Cat_Estado]"
            .Connection = cnAgil1
        End With
        daEstado.Fill(dsAgil, "Estados")


        cnAgil1.Open()
        For Each drDato In dsAgil.Tables("DatosPLD").Rows
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

            xMuni = "%" & Trim(drDato("pld_delegacion")) & "%"
            Municipio.FillByMunicipio(TMunicipio, xMuni, nIdEdo) 'contiene acentos el DS
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

            cNext = IIf(Trim(drDato("PLD_Numext")) = "", "0", drDato("PLD_Numext"))
            cNint = IIf(Trim(drDato("PLD_Numint")) = "", "0", drDato("PLD_Numint"))
            cNac = IIf(Trim(drDato("Nacionalidad")) = "", " ", IIf(Trim(drDato("Nacionalidad")) = "MEXICANA", 1, 2))
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
        Dim cnAgil As New SqlConnection(strConn)
        Dim cm1 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim cm4 As New SqlCommand()
        Dim drAnexo As DataRow
        Dim drEdoctav As DataRow()
        Dim drDato As DataRow

        Dim cDia As String
        Dim i As Integer
        Dim cRenglon As String
        Dim cImporte As String
        Dim cAnexo As String
        Dim cCiclo As String
        Dim cCliente As String
        Dim cFecha As String
        Dim cFechafin As String
        Dim cPago As String
        Dim nCount As Integer
        Dim nPago As Decimal
        Dim cProduct As String
        Dim cSucursal As String


        Dim cm2 As New SqlCommand()
        Dim dsReporte As New DataSet()
        Dim daAnexos As New SqlDataAdapter(cm1)
        Dim daEdoctav As New SqlDataAdapter(cm2)
        Dim daAvios As New SqlDataAdapter(cm3)
        Dim daCuentasConcetradoras As New SqlDataAdapter(cm4)
        Dim relAnexoEdoctav As DataRelation

        Try


            cDia = Mid(DTOC(Today), 7, 2) & Mid(DTOC(Today), 5, 2)
            cFecha = DTOC(Today)
            cnAgil.Open()

            With cm1
                .CommandType = CommandType.Text
                .CommandText = "SELECT cliente, Anexo, Fechacon, Mensu, MtoFin, Tipar, Sucursal, LiquidezInmediata FROM Minds_Cuentas "
                .Connection = cnAgil
            End With

            With cm3
                .CommandType = CommandType.Text
                .CommandText = "SELECT * FROM Minds_CuentasAvio"
                '.CommandText = "SELECT Clientes.cliente, Anexo + '-' + ciclo as Anexo, Fechaautorizacion, LineaActual, FechaTerminacion, Tipar FROM Clientes " & _
                '               "INNER JOIN Avios On Avios.Cliente = Clientes.Cliente WHERE Flcan = 'A' and fechaTerminacion >= '20130101' and (minds = 0 or minds is null)"
                .Connection = cnAgil
            End With

            With cm4
                .CommandType = CommandType.Text
                .CommandText = "SELECT * FROM Minds_CuentasConcetradoras where Anexo = 'XX'"
                .Connection = cnAgil
            End With

            ' Este Stored Procedure trae la tabla de amortización del equipo de todos los contratos activos
            ' con fecha de contratación menor o igual a la de proceso

            With cm2
                .CommandType = CommandType.Text
                .CommandText = "SELECT        Edoctav.Anexo, MAX(Edoctav.Letra) AS Letra, MAX(Edoctav.Feven) AS Feven, Edoctav.Abcap, Edoctav.Inter AS Inter, Edoctav.Iva AS Iva, SUM(Edoctav.IvaCapital) AS IvaCapital, " _
                                & "                         SUBSTRING(Edoctav.Feven, 1, 6) AS Mes, Edoctav.Nufac AS NufacT" _
                                & " FROM            Edoctav INNER JOIN" _
                                & " Minds_MovTradicionales ON Edoctav.Anexo = Minds_MovTradicionales.Anexo" _
                                & " WHERE        (Edoctav.Feven > N'20100101')" _
                                & " GROUP BY Edoctav.Anexo, SUBSTRING(Edoctav.Feven, 1, 6), Edoctav.Nufac,Edoctav.Abcap,Edoctav.INTER,Edoctav.IVA" _
                                & " ORDER BY Edoctav.Anexo, Letra "
                .Connection = cnAgil
            End With
            daAnexos.Fill(dsAgil, "Anexos")
            daEdoctav.Fill(dsAgil, "Edoctav")
            daAvios.Fill(dsAgil, "Avios")

            '' ''daCuentasConcetradoras.Fill(dsAgil, "Cuentas") 'SE QUITO A SOLICUTUD DE KARLA SANCCHEZ
            '' ''CONCETRADORAS+++++++++++++++++++++++++++++++++++++
            ' ''For Each drAnexo In dsAgil.Tables("Cuentas").Rows
            ' ''    cAnexo = drAnexo("Anexo")
            ' ''    cCliente = drAnexo("Cliente")
            ' ''    cSucursal = drAnexo("Mensu").ToString
            ' ''    cImporte = drAnexo("MtoFin").ToString
            ' ''    cFecha = CTOD(drAnexo("Fechacon")).ToShortDateString

            ' ''    nCount = 0
            ' ''    ' cProduct = "CREDITO"
            ' ''    ' cSubProduct = "SIMPLE"
            ' ''    cProduct = "3"

            ' ''    cFechafin = "01/01/2030"
            ' ''    cPago = 1

            ' ''    If Cuentas.Existe(cAnexo).Value = 0 Then
            ' ''        Cuentas.Insert(cAnexo, cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago)
            ' ''    End If
            ' ''Next
            '' ''CONCETRADORAS+++++++++++++++++++++++++++++++++++++

            ' Establecer la relación entre Anexos y Edoctav

            relAnexoEdoctav = New DataRelation("AnexoEdoctav", dsAgil.Tables("Anexos").Columns("Anexo"), dsAgil.Tables("Edoctav").Columns("Anexo"))
            dsAgil.EnforceConstraints = False
            dsAgil.Relations.Add(relAnexoEdoctav)

            For Each drAnexo In dsAgil.Tables("Anexos").Rows
                cAnexo = drAnexo("Anexo")
                cCliente = drAnexo("Cliente")
                cSucursal = drAnexo("Mensu").ToString
                cImporte = drAnexo("MtoFin").ToString
                cFecha = CTOD(drAnexo("Fechacon")).ToShortDateString
                drEdoctav = drAnexo.GetChildRows("AnexoEdoctav")
                Select Case drAnexo("Tipar")
                    Case "F"
                        '    cProduct = "ARRENDAMIENTO"
                        '   cSubProduct = "FINANCIERO"
                        cProduct = "1"
                    Case "P"
                        ' cProduct = "ARRENDAMIENTO"
                        ' cSubProduct = "PURO"
                        cProduct = "2"
                    Case "R"
                        'cProduct = "CREDITO"
                        'cSubProduct = "REFACCIONARIO"
                        cProduct = "8"
                    Case "L"
                        'cProduct = "CREDITO"
                        'cSubProduct = "REFACCIONARIO"
                        cProduct = "11"
                    Case "S"
                        ' cProduct = "CREDITO"
                        ' cSubProduct = "SIMPLE"
                        If drAnexo("LiquidezInmediata") = True Then
                            cProduct = "11"
                        Else
                            cProduct = "3"
                        End If
                End Select
                nCount = 0
                nPago = 0
                For Each drDato In drEdoctav
                    'If nCount = 0 Then
                    '    cFechafin = CTOD(drDato("Feven")).ToShortDateString
                    '    nPago = drDato("Abcap") + drDato("Inter") + drDato("iva") + drDato("ivaCapital")
                    '    nCount += 1
                    'End If
                    If Mid(drDato("Feven"), 1, 6) = Date.Now.AddMonths(-1).ToString("yyyyMM") Then
                        cFechafin = CTOD(drDato("Feven")).ToShortDateString
                        nPago += drDato("Abcap") + drDato("Inter") + drDato("iva") + drDato("ivaCapital")
                    End If
                Next
                cPago = nPago.ToString

                cRenglon = cAnexo & "|" & cCliente & "|" & cProduct & "|" & cImporte & "|" & cFecha & "|" & cFechafin & "|1|" & cPago & "|" & cSucursal & "|"
                If Cuentas.Existe(cAnexo).Value = 0 Then
                    Cuentas.Insert(cAnexo, cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago)
                Else
                    'Cuentas.UpdateCuenta(cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago, cAnexo)
                    Cuentas.UpdateMensualidad(cPago, cProduct, cAnexo)
                End If

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
                        ' cProduct = "CREDITO"
                        'cSubProduct = "ANTICIPO DE AVIO"
                        cProduct = "3" ' como simple

                    Case "C"
                        '   cProduct = "CREDITO"
                        '   cSubProduct = "CUENTA CORRIENTE"
                        cProduct = "4"
                    Case "H"
                        ' cProduct = "CREDITO"
                        ' cSubProduct = "AVIO"
                        cProduct = "9"
                End Select
                cFechafin = CTOD(drAnexo("FechaTerminacion")).ToShortDateString
                nPago = drAnexo("LineaActual")
                cPago = nPago.ToString

                If drAnexo("Tipar") <> "AA" Then
                    If Cuentas.Existe(cAnexo).Value = 0 Then
                        Cuentas.Insert(cAnexo, cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago)
                    Else
                        Cuentas.UpdateCuenta(cCliente, 7, cProduct, cImporte, cFecha, cFechafin, 1, cPago, cAnexo)
                    End If

                    cAnexo = Mid(cAnexo, 1, 9)
                    cCiclo = Mid(cAnexo, 11, 2)
                    Con2.UpdateMinds(cCiclo, cAnexo)
                End If
            Next
        Catch ex As Exception
            EnviaError("viapolo@lamoderna.com.mx,viapolo@lamoderna.com.mx", ex.Message & "  " & cAnexo, "error en CUENTAS")
        End Try
        cnAgil.Close()
    End Sub

    Sub Carga_Pagos()

    End Sub

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        If InStr(Mensaje, "No se ha encontrado la ruta de acceso de la red") = 0 Then
            Dim Mensage As New MailMessage("MINDS@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
            Dim Cliente As New SmtpClient("smtp01.cmoderna.com", 26)
            Try
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
        Cad = Cad.Replace("Á", "A")
        Cad = Cad.Replace("É", "E")
        Cad = Cad.Replace("Í", "I")
        Cad = Cad.Replace("Ó", "O")
        Cad = Cad.Replace("Ú", "U")
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

        ''tPAGOS.Fill(tPAG, Date.Now.AddDays(-40)) ' se quito por que ahora se va a la sigenrecia de pago!!!
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
