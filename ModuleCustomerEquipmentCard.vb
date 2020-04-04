Imports System.Data.SqlClient
Module ModuleCustomerEquipmentCard
    Public Function EventsCustomerEquipmentCardGenerarEntrega(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim oNumPoliza As String
        Dim oItemCode As String
        Dim oReciboInt As String
        Dim RutaC As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim Exe As String = Dir(RutaC)
        sPath = Microsoft.VisualBasic.Left(RutaC, Len(RutaC) - Len(Exe)) & "VKLog.log"

        'Obtener Numero de Poliza
        oItemB = oForm.Items.Item("43")
        oEditText = oItemB.Specific
        oNumPoliza = oEditText.Value
        'Obtener Numero Recibo
        oItemB = oForm.Items.Item("44")
        oEditText = oItemB.Specific
        oReciboInt = oEditText.Value
        'Detalle
        oItemB = oForm.Items.Item("45") 'Codigo Articulo
        oEditText = oItemB.Specific
        oItemCode = oEditText.Value
        'Verificar si numero serie existe para el articulo
        Sql = "SELECT COUNT(*) AS Cont FROM OINS WHERE ItemCode = '" & oItemCode & "' AND InternalSN = '" & oReciboInt & "' "
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(Sql)

        If oRs.Fields.Item("Cont").Value = 0 Then
            Try
                Dim BD_Net As String
                oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Sql = "SELECT TOP 1 U_Server AS Server,U_UserSQL AS UserSQL,U_PassSQL AS PassSQL,U_BD AS BD FROM [@CONEXAP]"
                oRs.DoQuery(Sql)
                BD_Net = oRs.Fields.Item("BD").Value
                'Creacion Entrega
                Dim oDeliveryNotes As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes) 'ere
                Dim PorcIva As Decimal
                Dim intComR As Decimal
                Dim DocDate As Date
                Dim Asociado As String
                Dim PrecioNet As Decimal
                Dim dcmCom As Decimal
                Dim intCom As Decimal

                'Obtener datos Cliente
                'Sql = "SELECT T0.strIdAseguradora, T0.strNomAse, T0.intComR, T0.CardCode, LEFT(T0.CardName,30) AS CardName, T0.intIva FROM " & BD_Net & "..appEmisionPolizas T0 "
                'Sql = Sql & " INNER JOIN " & BD_Net & "..appEmisionPolizaRecibosS T1 ON T1.idstrNoPoliza = T0.idstrNoPoliza "
                'Sql = Sql & " WHERE T1.strNoReciboInt = '" & oReciboInt & "' AND T1.BitSap = 0 "
                Sql = " Select [idstrNoPoliza]"
                Sql = Sql & " ,[strIdAseguradora]"
                Sql = Sql & " ,[strNomAse]"
                Sql = Sql & " ,[strNoReciboInt]"
                Sql = Sql & " ,[CardCode]"
                Sql = Sql & " ,[CardName]"
                Sql = Sql & " ,[strRFC]"
                Sql = Sql & " ,[LargoRFC]"
                Sql = Sql & " ,[bintIdCteOperacion]"
                Sql = Sql & " ,[bintIdCliente]"
                Sql = Sql & " ,[intIva]"
                'Sql = Sql & " ,[Asociado]"
                'Sql = Sql & " ,[Ejecutivo]"
                Sql = Sql & " ,[dteAplicacion]"
                Sql = Sql & " ,[ItemCode]"
                Sql = Sql & " ,[ItemName]"
                Sql = Sql & " ,[PrecioNet]"
                Sql = Sql & " ,[intComR]"
                Sql = Sql & " FROM " & BD_Net & "..vw_EntradasPolizas"
                Sql = Sql & " WHERE [strNoReciboInt] = '" & oReciboInt & "'"
                Sql = Sql & " GROUP BY [idstrNoPoliza]"
                Sql = Sql & " ,[strIdAseguradora]"
                Sql = Sql & " ,[strNomAse]"
                Sql = Sql & " ,[strNoReciboInt]"
                Sql = Sql & " ,[CardCode]"
                Sql = Sql & " ,[CardName]"
                Sql = Sql & " ,[strRFC]"
                Sql = Sql & " ,[LargoRFC]"
                Sql = Sql & " ,[bintIdCteOperacion]"
                Sql = Sql & " ,[bintIdCliente]"
                Sql = Sql & " ,[intIva]"
                Sql = Sql & " ,[dteAplicacion]"
                Sql = Sql & " ,[ItemCode]"
                Sql = Sql & " ,[ItemName]"
                Sql = Sql & " ,[PrecioNet]"
                Sql = Sql & " ,[intComR]"
                oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs.DoQuery(Sql)
                oDeliveryNotes.CardCode = oRs.Fields.Item("strIdAseguradora").Value
                oDeliveryNotes.CardName = oRs.Fields.Item("strNomAse").Value
                oDeliveryNotes.UserFields.Fields.Item("U_CardCode").Value = oRs.Fields.Item("CardCode").Value
                oDeliveryNotes.UserFields.Fields.Item("U_CardName").Value = oRs.Fields.Item("CardName").Value
                oDeliveryNotes.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                oDeliveryNotes.Comments = "Documento base Registro Póliza/Fianza " & oNumPoliza & ", Cliente Indirecto " & oRs.Fields.Item("CardCode").Value & ". Numero recibo Apianet " & oReciboInt
                PorcIva = oRs.Fields.Item("intIva").Value
                intComR = oRs.Fields.Item("intComR").Value

                'PARA OBTENER ASOCIADO Y EJECUTIVO
                If conexionSQL2() = 0 Then
                    Sql = "SELECT Asociado, Ejecutivo, dcmCom, intCom "
                    Sql = Sql & " FROM " & BD_Net & "..vw_EntradasPolizas "
                    Sql = Sql & " WHERE strNoReciboInt = '" & oReciboInt & "'"
                    comando.Connection = Sqlconn2
                    comando.CommandText = Sql
                    AsocEje = comando.ExecuteReader

                End If

                Do Until AsocEje.Read.ToString <> True
                    If AsocEje.Item("Asociado").ToString <> "" Then
                        dcmCom = AsocEje.Item("dcmCom").ToString
                        intCom = AsocEje.Item("IntCom").ToString
                    End If
                Loop

                If AsocEje.Read.ToString = True Then
                    AsocEje.Close()
                    Sqlconn2.Close()
                End If

                'Asociado
                oItemB = oForm.Items.Item("ComboAsoc") 'Asociado
                oComboBox = oItemB.Specific
                Try
                    oDeliveryNotes.UserFields.Fields.Item("U_Asociado").Value = oComboBox.Selected.Value
                    Asociado = oComboBox.Selected.Value
                Catch ex As Exception
                    Asociado = ""
                End Try
                'Ejecutivo
                oItemB = oForm.Items.Item("ComboEjec") 'Ejecutivo
                oComboBox = oItemB.Specific
                Try
                    oDeliveryNotes.UserFields.Fields.Item("U_Ejecutivo").Value = oComboBox.Selected.Value
                    oDeliveryNotes.Lines.SalesPersonCode = oComboBox.Selected.Value
                Catch ex As Exception

                End Try

                'Fechas
                'Sql = "SELECT dteAplicacion, dcmPrimaR + dcmDP + dcmRPF AS PrecioNeto FROM " & BD_Net & "..appEmisionPolizaRecibosS "
                'Sql = Sql & "WHERE strNoReciboInt = '" & oReciboInt & "' AND BitSap = 0 "
                'oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRs.DoQuery(Sql)
                oDeliveryNotes.DocDate = oRs.Fields.Item("dteAplicacion").Value
                oDeliveryNotes.DocDueDate = oRs.Fields.Item("dteAplicacion").Value
                oDeliveryNotes.TaxDate = oRs.Fields.Item("dteAplicacion").Value
                DocDate = oRs.Fields.Item("dteAplicacion").Value
                oDeliveryNotes.NumAtCard = oNumPoliza

                oDeliveryNotes.Lines.ItemCode = oItemCode

                oItemB = oForm.Items.Item("46") 'Descripcion Articulo
                oEditText = oItemB.Specific
                oDeliveryNotes.Lines.ItemDescription = oEditText.Value

                oDeliveryNotes.Lines.WarehouseCode = "01"
                oDeliveryNotes.Lines.Quantity = 1
                oDeliveryNotes.Lines.UnitPrice = oRs.Fields.Item("PrecioNet").Value
                PrecioNet = oRs.Fields.Item("PrecioNet").Value
                oDeliveryNotes.Lines.CommisionPercent = intComR
                oDeliveryNotes.Lines.UserFields.Fields.Item("U_NumPoliza").Value = oNumPoliza
                oDeliveryNotes.Lines.UserFields.Fields.Item("U_ReciboInt").Value = oReciboInt
                'Si devuelven o cancelan una poliza, PrecioNet viene negativo
                If oRs.Fields.Item("PrecioNet").Value < 0 Then
                    oDeliveryNotes.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = -1
                Else
                    oDeliveryNotes.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = 1
                End If

                'Impuesto
                Sql = "SELECT Code FROM OSTA WHERE LEFT(Code,2) = 'BE' AND Rate = " & PorcIva
                oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs.DoQuery(Sql)
                oDeliveryNotes.Lines.TaxCode = oRs.Fields.Item("Code").Value

                lRetCode = oDeliveryNotes.Add()

                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    AddLog("Entrega no ha sido creado Fecha " & DocDate & ", poliza " & oNumPoliza & ", Nro registro Apianet " & oReciboInt & ", " & sErrMsg)
                    EventsCustomerEquipmentCardGenerarEntrega = False
                Else
                        'ACTUALIZAR CAMPO DeliveryNo EN OINS, PERO CAMPO ES SOLO DE LECTURA
                    oRs5 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Sql = "SELECT TOP 1 T0.DocEntry FROM ODLN T0 JOIN DLN1 T1 ON T1.DocEntry = T0.DocEntry "
                    Sql = Sql & " WHERE T1.U_NumPoliza = '" & oNumPoliza & "' AND T1.U_ReciboInt = '" & oReciboInt & "' ORDER BY T0.DocEntry Desc"
                    oRs5.DoQuery(Sql)
                    oItemB = oForm.Items.Item("TxtDocNum")
                    oEditText = oItemB.Specific
                    oEditText.Value = oRs5.Fields.Item("DocEntry").Value
                    oRs5 = Nothing

                    Sql = "UPDATE " & BD_Net & "..appEmisionPolizaRecibosS SET BitSap = 1 WHERE strNoReciboInt = '" & oReciboInt & "'"
                        'oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery(Sql)
                    EventsCustomerEquipmentCardGenerarEntrega = True
                    AddOrdenCompra(Asociado, oNumPoliza, oReciboInt, oItemCode, DocDate, dcmCom, intCom1, 0)
                    End If
            Catch ex As Exception
                EventsCustomerEquipmentCardGenerarEntrega = False
            End Try
        Else
                EventsCustomerEquipmentCardGenerarEntrega = True
        End If

    End Function
    Public Sub AddOrdenCompra(ByVal Asociado As String, ByVal idstrNoPoliza As String, ByVal strNoReciboInt As String, ByVal ItemCode As String, ByVal dteAplicacion As Date, ByVal PrecioNet As Decimal, ByVal intCom As Decimal, ByVal intIva As Decimal)
        Dim oPurchesaOrder As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
        oPurchesaOrder.DocDate = dteAplicacion
        oPurchesaOrder.DocDueDate = dteAplicacion
        oPurchesaOrder.TaxDate = dteAplicacion
        oPurchesaOrder.CardCode = Asociado
        oPurchesaOrder.Lines.ItemCode = ItemCode
        oPurchesaOrder.Lines.Quantity = 1
        oPurchesaOrder.Lines.WarehouseCode = "01"
        oPurchesaOrder.Lines.CommisionPercent = intCom
        If intCom > 0 Then
            oPurchesaOrder.Lines.DiscountPercent = 100 - intCom
        Else
            oPurchesaOrder.Lines.DiscountPercent = 0
        End If

        Sql = "SELECT VATLiable FROM OITM WHERE ItemCode = '" & ItemCode & "'"
        oRs6 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs6.DoQuery(Sql)
        If oRs6.Fields.Item("VATLiable").Value = "Y" Then
            oPurchesaOrder.Lines.TaxCode = "W3"
        Else
            oPurchesaOrder.Lines.TaxCode = "W0"
        End If
        oRs6 = Nothing
        oPurchesaOrder.Lines.UnitPrice = PrecioNet
        oPurchesaOrder.Lines.UserFields.Fields.Item("U_NumPoliza").Value = idstrNoPoliza
        oPurchesaOrder.Lines.UserFields.Fields.Item("U_ReciboInt").Value = strNoReciboInt
        'Si devuelven o cancelan una poliza, PrecioNet viene negativo
        If PrecioNet < 0 Then
            oPurchesaOrder.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = -1
        Else
            oPurchesaOrder.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = 1
        End If

        oPurchesaOrder.Comments = "Pago de Comisiones a Asociados"

        lRetCode = oPurchesaOrder.Add()

        If lRetCode <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            AddLog("Orden de Compra " & Asociado & " - " & strNoReciboInt & " no ha sido creada, " & sErrMsg)
        Else
            AddLog("Orden de Compra " & Asociado & " - " & strNoReciboInt & " ha sido creada correctamente")
        End If

    End Sub

    Public Function EventsCustomerEquipmentCard(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form, ByRef oEvent As SAPbouiCOM.ItemEvent) As Boolean
        Dim TmpStatusEvent As Boolean
        TmpStatusEvent = True

        oForm.Title = "Registro de Pólizas/Fianzas"

        'Modifica Numero serie fabricante
        oItemB = oForm.Items.Item("9")
        oStaticText = oItemB.Specific
        'oItemB.Width = oItemB.Width 
        oStaticText.Caption = "Número de Póliza/Fianza"

        'Modifica Numero serie
        oItemB = oForm.Items.Item("4")
        oStaticText = oItemB.Specific
        'oItemB.Width = oItemB.Width 
        oStaticText.Caption = "Número de Póliza/Fianza"

        'Modifica Numero articulo
        oItemB = oForm.Items.Item("16")
        oStaticText = oItemB.Specific
        'oItemB.Width = oItemB.Width 
        oStaticText.Caption = "Código de Póliza/Fianza"

        'Modifica descripcion articulo
        oItemB = oForm.Items.Item("15")
        oStaticText = oItemB.Specific
        oItemB.Width = oItemB.Width + 10
        oStaticText.Caption = "Descripción de Póliza/Fianza"

        'Modifica descripcion para tab Llamada de servicio
        oItemB = oForm.Items.Item("42")
        oFolder = oItemB.Specific
        oFolder.Caption = "Detalle Póliza/Fianza"

        'Oculta campo Tecnico
        oForm.Items.Item("173").Visible = False
        oForm.Items.Item("167").Visible = False

        'Oculta campo Territorio
        oForm.Items.Item("174").Visible = False
        oForm.Items.Item("168").Visible = False

        'Oculta para tab Contrato de servicio
        oForm.Items.Item("41").Visible = False


        '//*************************
        '// Adding a Combo Box para Asociado
        '//*************************
        oItemB = oForm.Items.Item("173")
        oItem = oForm.Items.Add("ComboAsoc", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height
        oItem.FromPane = 1
        oItem.ToPane = 1

        oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        oComboBox.DataBind.SetBound(True, "OINS", "U_Asociado")

        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Sql = "SELECT CAST(SlpCode AS VarChar(20)) [Code], SlpName [Name] FROM OSLP UNION ALL "
        Sql = Sql & "SELECT CardCode [Code], LEFT(CardName,50) [Name] FROM OCRD T0 "
        Sql = Sql & "JOIN OCRG T1 ON T1.GroupCode = T0.GroupCode "
        Sql = Sql & "WHERE T0.CardType = 'S' AND T1.GroupName = 'Asociados'"
        oRs.DoQuery(Sql)
        oItem.DisplayDesc = True
        While Not oRs.EoF
            oComboBox.ValidValues.Add(oRs.Fields.Item("Code").Value, oRs.Fields.Item("Name").Value)
            oRs.MoveNext()
        End While


        '//**********************************
        '// Adding Label para Asociado
        '//**********************************

        oItemB = oForm.Items.Item("167")
        oItem = oForm.Items.Add("StaticTxt4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height
        oItem.FromPane = 1
        oItem.ToPane = 1

        oItem.LinkTo = "ComboAsoc"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Asociado"

        '//*************************
        '// Adding a Combo Box para Ejecutivo
        '//*************************
        oItemB = oForm.Items.Item("174")
        oItem = oForm.Items.Add("ComboEjec", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height
        oItem.FromPane = 1
        oItem.ToPane = 1

        oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        oComboBox.DataBind.SetBound(True, "OINS", "U_Ejecutivo")

        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Sql = "SELECT CAST(SlpCode AS VarChar(20)) [Code], SlpName [Name] FROM OSLP UNION ALL "
        Sql = Sql & "SELECT CardCode [Code], LEFT(CardName,50) [Name] FROM OCRD T0 "
        Sql = Sql & "JOIN OCRG T1 ON T1.GroupCode = T0.GroupCode "
        Sql = Sql & "WHERE T0.CardType = 'S' AND T1.GroupName = 'Asociados'"
        oRs.DoQuery(Sql)
        oItem.DisplayDesc = True
        While Not oRs.EoF
            oComboBox.ValidValues.Add(oRs.Fields.Item("Code").Value, oRs.Fields.Item("Name").Value)
            oRs.MoveNext()
        End While


        '//**********************************
        '// Adding Label para Ejecutivo
        '//**********************************

        oItemB = oForm.Items.Item("168")
        oItem = oForm.Items.Add("StaticTxt5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height
        oItem.FromPane = 1
        oItem.ToPane = 1

        oItem.LinkTo = "ComboEjec"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Ejecutivo"

        '//*************************
        '// Adding a Texto para Numero Entrega
        '//*************************
        oItemB = oForm.Items.Item("172")
        oItem = oForm.Items.Add("TxtDocNum", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top + oItemB.Height + 1
        oItem.Height = oItemB.Height
        oEditText = oItem.Specific
        oEditText.DataBind.SetBound(True, "OINS", "U_DocNum")

        '//**********************************
        '// Adding Label para Numero Entrega
        '//**********************************

        oItemB = oForm.Items.Item("164")
        oItem = oForm.Items.Add("StaticTxt6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width - 15
        oItem.Top = oItemB.Top + oItemB.Height + 1
        oItem.Height = oItemB.Height

        oItem.LinkTo = "TxtDocNum"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Entrega"

        '''//Link
        oItemB = oForm.Items.Item("StaticTxt6")
        oItem = oForm.Items.Add("lblDocNum", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
        oItem.Left = oItemB.Left + oItemB.Width + 5
        oItem.Width = 15
        oItem.Top = oItemB.Top
        oItem.Height = 19


        oItem.LinkTo = "TxtDocNum" '// ID of a EditText present in the Form
        Dim oLinkButton As SAPbouiCOM.LinkedButton = oItem.Specific
        oLinkButton.LinkedObjectType = "15"

        EventsCustomerEquipmentCard = TmpStatusEvent
    End Function

    Public Function EventsCustomerEquipmentCardLoadDatos(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim oReciboInt As String
        Dim Asociado1 As String
        Dim Ejecutivo1 As String
        'Obtener Numero Recibo
        oItemB = oForm.Items.Item("44")
        oEditText = oItemB.Specific
        oReciboInt = oEditText.Value
        'Obtener Datos de Conexion
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Sql = "SELECT TOP 1 U_Server AS Server,U_UserSQL AS UserSQL,U_PassSQL AS PassSQL,U_BD AS BD FROM [@CONEXAP]"
        oRs.DoQuery(Sql)
        Servidor = oRs.Fields.Item("Server").Value
        BD_Net = oRs.Fields.Item("BD").Value
        User = oRs.Fields.Item("UserSQL").Value
        Pass = oRs.Fields.Item("PassSQL").Value
        
        If conexionSQL() = 0 Then
            comando.Connection = Sqlconn
            'Buscar datos de poliza en Apianet
            Sql = " Select [idstrNoPoliza]"
            Sql = Sql & " ,[strIdAseguradora]"
            Sql = Sql & " ,[strNomAse]"
            Sql = Sql & " ,[strNoReciboInt]"
            Sql = Sql & " ,[CardCode]"
            Sql = Sql & " ,[CardName]"
            Sql = Sql & " ,[strRFC]"
            Sql = Sql & " ,[LargoRFC]"
            Sql = Sql & " ,[bintIdCteOperacion]"
            Sql = Sql & " ,[bintIdCliente]"
            Sql = Sql & " ,[intIva]"
            'Sql = Sql & " ,[Asociado]"
            'Sql = Sql & " ,[Ejecutivo]"
            Sql = Sql & " ,[dteAplicacion]"
            Sql = Sql & " ,[ItemCode]"
            Sql = Sql & " ,[ItemName]"
            Sql = Sql & " ,[PrecioNet]"
            Sql = Sql & " ,[intComR]"
            Sql = Sql & " FROM [dbo].[vw_EntradasPolizas]"
            Sql = Sql & " WHERE [strNoReciboInt] = '" & oReciboInt & "'"
            Sql = Sql & " GROUP BY [idstrNoPoliza]"
            Sql = Sql & " ,[strIdAseguradora]"
            Sql = Sql & " ,[strNomAse]"
            Sql = Sql & " ,[strNoReciboInt]"
            Sql = Sql & " ,[CardCode]"
            Sql = Sql & " ,[CardName]"
            Sql = Sql & " ,[strRFC]"
            Sql = Sql & " ,[LargoRFC]"
            Sql = Sql & " ,[bintIdCteOperacion]"
            Sql = Sql & " ,[bintIdCliente]"
            Sql = Sql & " ,[intIva]"
            Sql = Sql & " ,[dteAplicacion]"
            Sql = Sql & " ,[ItemCode]"
            Sql = Sql & " ,[ItemName]"
            Sql = Sql & " ,[PrecioNet]"
            Sql = Sql & " ,[intComR]"
            comando.CommandText = Sql
            LectSN = comando.ExecuteReader
            Do Until LectSN.Read.ToString <> True
                Try
                    'Ingresar numero poliza
                    oItemB = oForm.Items.Item("43")
                    oEditText = oItemB.Specific
                    oEditText.Value = LectSN.Item("idstrNoPoliza").ToString

                    'Ingresar codigo articulo
                    oItemB = oForm.Items.Item("45")
                    oEditText = oItemB.Specific
                    oEditText.Value = LectSN.Item("strIdProd").ToString

                    'Ingresar codigo cliente
                    oItemB = oForm.Items.Item("48")
                    oEditText = oItemB.Specific
                    oEditText.Value = LectSN.Item("CardCode").ToString

                    'PARA OBTENER ASOCIADO Y EJECUTIVO
                    If conexionSQL2() = 0 Then
                        Sql = "SELECT Asociado, Ejecutivo, dcmCom, intCom "
                        Sql = Sql & " FROM " & BD_Net & "..vw_EntradasPolizas "
                        Sql = Sql & " WHERE strNoReciboInt = '" & oReciboInt & "'"
                        comando.Connection = Sqlconn2
                        comando.CommandText = Sql
                        AsocEje = comando.ExecuteReader

                    End If
                    Asociado1 = ""
                    Ejecutivo1 = ""
                    Do Until AsocEje.Read.ToString <> True
                        If AsocEje.Item("Asociado").ToString <> "" Then
                            Asociado1 = AsocEje.Item("Asociado").ToString
                            intCom1 = AsocEje.Item("intCom").ToString
                        Else
                            Ejecutivo1 = AsocEje.Item("Ejecutivo").ToString
                        End If
                    Loop

                    If AsocEje.Read.ToString = True Then
                        AsocEje.Close()
                        Sqlconn2.Close()
                    End If

                    'Ingresar Asociado
                    Try
                        oComboBox = oForm.Items.Item("ComboAsoc").Specific
                        'oEditText = oComboBox.Specific
                        oComboBox.Select(Asociado1, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Catch ex As Exception
                        MsgBox("Asociado no existe en SAP", MsgBoxStyle.Exclamation, "Datos Asociado")
                    End Try

                    'Ingresar Ejecutivo
                    Try
                        oComboBox = oForm.Items.Item("ComboEjec").Specific
                        'oEditText = oComboBox.Specific
                        oComboBox.Select(Ejecutivo1, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Catch ex As Exception
                        MsgBox("Ejecutivo no existe en SAP", MsgBoxStyle.Exclamation, "Datos Ejecutivo")
                    End Try

                Catch ex As Exception
                    'Me.LogVisual.Items.Add("Cliente [" & LectSN.Item(2).ToString & "] [" & LectSN.Item(1).ToString & "] ERROR : " & lErrCode & " - " & sErrMsg)
                    'errores = errores & "cliente [" & LectSN.Item(2).ToString & "] [" & LectSN.Item(1).ToString & "] ERROR : " & lErrCode & " - " & sErrMsg & "<br>"
                End Try
            Loop
            LectSN.Close()
        End If
        EventsCustomerEquipmentCardLoadDatos = True
    End Function


End Module
