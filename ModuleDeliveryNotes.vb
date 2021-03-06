Module ModuleDeliveryNotes

    Public Function EventsDeliveryNotesForm(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form, ByRef oEvent As SAPbouiCOM.ItemEvent) As Boolean
        Dim TmpStatusEvent As Boolean
        TmpStatusEvent = True

        'Oculta campo Empleado de ventas
        oForm.Items.Item("21").Visible = False 'Label
        oForm.Items.Item("20").Visible = False 'SlpCode
        oForm.Items.Item("53").Visible = False 'Boton busqueda

        'Oculta campo Titular
        oForm.Items.Item("230").Visible = False 'Label
        oForm.Items.Item("222").Visible = False 'Titular


        '//*************************
        '// Adding a Combo Box para Asociado
        '//*************************
        oItemB = oForm.Items.Item("20")
        oItem = oForm.Items.Add("ComboAsoc", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height

        oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        oComboBox.DataBind.SetBound(True, "ODLN", "U_Asociado")

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

        oItemB = oForm.Items.Item("21")
        oItem = oForm.Items.Add("StaticTxt1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
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
        oItemB = oForm.Items.Item("222")
        oItem = oForm.Items.Add("ComboEjec", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height

        oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        oComboBox.DataBind.SetBound(True, "ODLN", "U_Ejecutivo")

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

        oItemB = oForm.Items.Item("230")
        oItem = oForm.Items.Add("StaticTxt2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height
        oItem.FromPane = 1
        oItem.ToPane = 1

        oItem.LinkTo = "ComboEjec"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Ejecutivo"

        EventsDeliveryNotesForm = TmpStatusEvent
    End Function

End Module
