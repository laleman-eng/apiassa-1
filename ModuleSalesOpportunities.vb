Module ModuleSalesOpportunities
    Public Function EventsSalesOpportunitiesForm(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form, ByRef oEvent As SAPbouiCOM.ItemEvent) As Boolean
        Dim TmpStatusEvent As Boolean
        TmpStatusEvent = True

        'Cambia Label para Ejecutivos y Asociados
        oItem = oForm.Items.Item("157")
        oStaticText = oItem.Specific
        oStaticText.Caption = "Asociado A"

        oItem = oForm.Items.Item("159")
        oStaticText = oItem.Specific
        oStaticText.Caption = "Nombre Asociado A"

        'Oculta campo Empleado de ventas
        oForm.Items.Item("14").Visible = False 'Label
        oForm.Items.Item("15").Visible = False 'SlpCode

        'Oculta campo Titular
        oForm.Items.Item("184").Visible = False 'Label
        oForm.Items.Item("185").Visible = False 'Titular

        'Oculta campo Territorio
        oForm.Items.Item("150").Visible = False 'Label
        oForm.Items.Item("166").Visible = False 'Titular

        '//*************************
        '// Adding a Combo Box para Ejecutivo A
        '//*************************
        oItemB = oForm.Items.Item("166")
        oItem = oForm.Items.Add("ComboEjec", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height

        oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        oComboBox.DataBind.SetBound(True, "OOPR", "U_SlpCode")

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
        '// Adding Label para Ejecutivo A
        '//**********************************

        oItemB = oForm.Items.Item("150")
        oItem = oForm.Items.Add("StaticTxt1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height
        ''oItem.FromPane = 1
        ''oItem.ToPane = 1

        oItem.LinkTo = "ComboEjec"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Ejecutivo A"


        '//*************************
        '// Adding a Combo Box para Ejecutivo B
        '//*************************
        oItemB = oForm.Items.Item("15")
        oItem = oForm.Items.Add("ComboEjecB", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height

        oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        oComboBox.DataBind.SetBound(True, "OOPR", "U_SlpCode")

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
        '// Adding Label para Ejecutivo B
        '//**********************************

        oItemB = oForm.Items.Item("14")
        oItem = oForm.Items.Add("StaticTxt2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height
        ''oItem.FromPane = 1
        ''oItem.ToPane = 1

        oItem.LinkTo = "ComboEjecB"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Ejecutivo B"


        '//*************************
        '// Adding a Combo Box para Asociado B
        '//*************************
        oItemB = oForm.Items.Item("185")
        oItem = oForm.Items.Add("ComboAsoc", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height

        oComboBox = oItem.Specific

        '// bind the Combo Box item to the defined used data source
        oComboBox.DataBind.SetBound(True, "OOPR", "U_Asociado")

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
        '// Adding Label para Asociado B
        '//**********************************

        oItemB = oForm.Items.Item("184")
        oItem = oForm.Items.Add("StaticTxt3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height
        ''oItem.FromPane = 1
        ''oItem.ToPane = 1

        oItem.LinkTo = "ComboAsoc"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Asociado B"

        EventsSalesOpportunitiesForm = TmpStatusEvent
    End Function
End Module
