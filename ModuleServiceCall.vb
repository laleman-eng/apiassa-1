Option Strict Off
Option Explicit On
Module ModuleServiceCall
    Public Function EventsServiceCall(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form, ByRef oEvent As SAPbouiCOM.ItemEvent) As Boolean
        Dim TmpStatusEvent As Boolean
        TmpStatusEvent = True

        ' oForm.Title = "Detalle de Pólizas/Fianzas"

        'Modifica Numero serie fabricante
        '        oItemB = oForm.Items.Item("30")
        '        oStaticText = oItemB.Specific
        '        oItemB.Width = oItemB.Width + 20
        '        oStaticText.Caption = "Número de Póliza/Fianza"

        'Modifica Numero serie
        '        oItemB = oForm.Items.Item("84")
        '        oStaticText = oItemB.Specific
        '        oItemB.Width = oItemB.Width + 20
        '        oStaticText.Caption = "Número de Póliza/Fianza"

        'Modifica Numero articulo
        '        oItemB = oForm.Items.Item("32")
        '        oStaticText = oItemB.Specific
        '        oItemB.Width = oItemB.Width + 20
        '        oStaticText.Caption = "Código de Póliza/Fianza"

        'Modifica descripcion articulo
        '        oItemB = oForm.Items.Item("66")
        '        oStaticText = oItemB.Specific
        '        oItemB.Width = oItemB.Width + 15
        '        oStaticText.Caption = "Descripción Póliza/Fianza"

        'Modifica descripcion para campo Origen
        '        oItemB = oForm.Items.Item("68")
        '        oStaticText = oItemB.Specific
        '        oStaticText.Caption = "Afianzadora"

        'Modifica descripcion para campo Tipo de Problema
        '        oItemB = oForm.Items.Item("41")
        '        oStaticText = oItemB.Specific
        '        oStaticText.Caption = "Código de CNSF"

        'Modifica descripcion para campo Tipo de Llamada
        '        oItemB = oForm.Items.Item("45")
        '        oStaticText = oItemB.Specific
        '        oItemB.Width = oItemB.Width + 6
        '        oStaticText.Caption = "Código de agente"

        'Modifica descripcion para campo Status de Llamada
        '        oItemB = oForm.Items.Item("39")
        '        oStaticText = oItemB.Specific
        '        oStaticText.Caption = "Status de Póliza"

        'Modifica descripcion para tab Operaciones
        '        oItemB = oForm.Items.Item("53")
        '        oFolder = oItemB.Specific
        '        oFolder.Caption = "Endoso"


        'Oculta campo ID de Llamada
        '        oForm.Items.Item("11").Visible = False
        '        oForm.Items.Item("12").Visible = False
        'Oculta campo Tecnico
        '       oForm.Items.Item("93").Visible = False
        '       oForm.Items.Item("94").Visible = False


        '// add a User Data Source to the form
        'oForm.DataSources.UserDataSources.Add("EditSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
        'oForm.DataSources.UserDataSources.Add("CombSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)


        '//*************************
        '// Adding a EditText Numero Póliza
        '//*************************
        ' oItemB = oForm.Items.Item("12")
        ' oItem = oForm.Items.Add("NumPoliza", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        ' oItem.Left = oItemB.Left - 23
        ' oItem.Width = oItemB.Width
        ' oItem.Top = oItemB.Top
        ' oItem.Height = oItemB.Height


        'oEditText = oItem.Specific
        'oEditText.DataBind.SetBound(True, "OSCL", "U_NumPoliza")

        '//**********************************
        '// Adding Label para Numero Póliza
        '//**********************************

        ' oItemB = oForm.Items.Item("11")
        ' oItem = oForm.Items.Add("StaticTxt3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        ' oItem.Left = oItemB.Left - 23
        ' oItem.Width = oItemB.Width
        ' oItem.Top = oItemB.Top
        ' oItem.Height = oItemB.Height

        'oItem.LinkTo = "NumPoliza"
        'oStaticText = oItem.Specific
        'oStaticText.Caption = "Numero Póliza"

        '//*************************
        '// Adding a Combo Box para Asociado
        '//*************************
        'oItemB = oForm.Items.Item("93")
        'oItem = oForm.Items.Add("ComboAsoc", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        'oItem.Left = oItemB.Left
        'oItem.Width = oItemB.Width
        'oItem.Top = oItemB.Top
        'oItem.Height = oItemB.Height
        'oItem.FromPane = 1
        'oItem.ToPane = 1

        'oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        'oComboBox.DataBind.SetBound(True, "OSCL", "U_Asociado")

        'oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Sql = "SELECT CAST(SlpCode AS VarChar(20)) [Code], SlpName [Name] FROM OSLP UNION ALL "
        'Sql = Sql & "SELECT CardCode [Code], LEFT(CardName,50) [Name] FROM OCRD T0 "
        'Sql = Sql & "JOIN OCRG T1 ON T1.GroupCode = T0.GroupCode "
        'Sql = Sql & "WHERE T0.CardType = 'S' AND T1.GroupName = 'Asociados'"
        'oRs.DoQuery(Sql)
        'oItem.DisplayDesc = True
        'While Not oRs.EoF
        'oComboBox.ValidValues.Add(oRs.Fields.Item("Code").Value, oRs.Fields.Item("Name").Value)
        'oRs.MoveNext()
        'End While


        '//**********************************
        '// Adding Label para Asociado
        '//**********************************

        'oItemB = oForm.Items.Item("94")
        'oItem = oForm.Items.Add("StaticTxt4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        'oItem.Left = oItemB.Left
        'oItem.Width = oItemB.Width
        'oItem.Top = oItemB.Top
        'oItem.Height = oItemB.Height
        'oItem.FromPane = 1
        'oItem.ToPane = 1

        'oItem.LinkTo = "ComboAsoc"
        'oStaticText = oItem.Specific
        'oStaticText.Caption = "Asociado"


        '//*************************
        '// Adding a Combo Box para Ejecutivo
        '//*************************
        'oItemB = oForm.Items.Item("93")
        'oItem = oForm.Items.Add("ComboEjec", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        'oItem.Left = oItemB.Left
        'oItem.Width = oItemB.Width
        'oItem.Top = oItemB.Top + oItemB.Height + 1
        'oItem.Height = oItemB.Height
        'oItem.FromPane = 1
        'oItem.ToPane = 1

        'oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        'oComboBox.DataBind.SetBound(True, "OSCL", "U_Ejecutivo")

        'oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Sql = "SELECT CAST(SlpCode AS VarChar(20)) [Code], SlpName [Name] FROM OSLP UNION ALL "
        'Sql = Sql & "SELECT CardCode [Code], LEFT(CardName,50) [Name] FROM OCRD T0 "
        'Sql = Sql & "JOIN OCRG T1 ON T1.GroupCode = T0.GroupCode "
        'Sql = Sql & "WHERE T0.CardType = 'S' AND T1.GroupName = 'Asociados'"
        'oRs.DoQuery(Sql)
        'oItem.DisplayDesc = True
        'While Not oRs.EoF
        'oComboBox.ValidValues.Add(oRs.Fields.Item("Code").Value, oRs.Fields.Item("Name").Value)
        'oRs.MoveNext()
        'End While


        '//**********************************
        '// Adding Label para Ejecutivo
        '//**********************************

        'oItemB = oForm.Items.Item("94")
        'oItem = oForm.Items.Add("StaticTxt5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        'oItem.Left = oItemB.Left
        'oItem.Width = oItemB.Width
        'oItem.Top = oItemB.Top + oItemB.Height + 1
        'oItem.Height = oItemB.Height
        'oItem.FromPane = 1
        'oItem.ToPane = 1

        'oItem.LinkTo = "ComboEjec"
        'oStaticText = oItem.Specific
        'oStaticText.Caption = "Ejecutivo"


        ''//*************************
        ''// Adding ComboBox para Ramo
        ''//*************************

        'oItemB = oForm.Items.Item("ComboEjec")
        'oItem = oForm.Items.Add("ComboRamo", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        'oItem.Left = oItemB.Left
        'oItem.Width = oItemB.Width
        'oItem.Top = oItemB.Top + oItemB.Height + 1
        'oItem.Height = oItemB.Height
        'oItem.FromPane = 1
        'oItem.ToPane = 1

        'oComboBox = oItem.Specific

        '// bind the text edit item to the defined used data source
        'oComboBox.DataBind.SetBound(True, "OSCL", "U_ramo")
        'oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oRs.DoQuery("SELECT FldValue, Descr FROM UFD1 WHERE TableID = 'OITM' AND FieldID = 6")
        'oItem.DisplayDesc = True
        'While Not oRs.EoF
        '    oComboBox.ValidValues.Add(oRs.Fields.Item("FldValue").Value, oRs.Fields.Item("Descr").Value)
        '    oRs.MoveNext()
        'End While

        '//***************************
        '// Adding a Label para Ramo
        '//***************************

        'oItemB = oForm.Items.Item("StaticTxt5")
        'oItem = oForm.Items.Add("StaticTxt1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        'oItem.Left = oItemB.Left
        'oItem.Width = oItemB.Width
        'oItem.Top = oItemB.Top + oItemB.Height + 1
        'oItem.Height = oItemB.Height
        'oItem.FromPane = 1
        'oItem.ToPane = 1


        'oItem.LinkTo = "ComboRamo"
        'oStaticText = oItem.Specific
        'oStaticText.Caption = "Ramo"


        '//*************************
        '// Adding a Combo Box para SubRamo
        '//*************************
        'oItemB = oForm.Items.Item("ComboRamo")
        'oItem = oForm.Items.Add("ComboSRamo", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        'oItem.Left = oItemB.Left
        'oItem.Width = oItemB.Width
        'oItem.Top = oItemB.Top + oItemB.Height + 1
        'oItem.Height = oItemB.Height
        'oItem.FromPane = 1
        'oItem.ToPane = 1

        'oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        'oComboBox.DataBind.SetBound(True, "OSCL", "U_SubRamo")

        'oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oRs.DoQuery("SELECT FldValue, Descr FROM UFD1 WHERE TableID = 'OITM' AND FieldID = 7")
        'oItem.DisplayDesc = True
        'While Not oRs.EoF
        '    oComboBox.ValidValues.Add(oRs.Fields.Item("FldValue").Value, oRs.Fields.Item("Descr").Value)
        '    oRs.MoveNext()
        'End While


        '//**********************************
        '// Adding Label para SubRamo
        '//**********************************

        'oItemB = oForm.Items.Item("StaticTxt1")
        'oItem = oForm.Items.Add("StaticTxt2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        'oItem.Left = oItemB.Left
        'oItem.Width = oItemB.Width
        'oItem.Top = oItemB.Top + oItemB.Height + 1
        'oItem.Height = oItemB.Height
        'oItem.FromPane = 1
        'oItem.ToPane = 1

        'oItem.LinkTo = "ComboSRamo"
        'oStaticText = oItem.Specific
        'oStaticText.Caption = "Sub Ramo"

        EventsServiceCall = TmpStatusEvent
    End Function
End Module
