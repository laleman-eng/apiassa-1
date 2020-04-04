Module ModuleActivities
    Public Function EventsActivities(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form, ByRef oEvent As SAPbouiCOM.ItemEvent) As Boolean
        Dim TmpStatusEvent As Boolean
        TmpStatusEvent = True

        'Modifica descripcion para campo Hora de Inicio
        oItemB = oForm.Items.Item("20")
        oStaticText = oItemB.Specific
        oStaticText.Caption = "Inicio de Vigencia"

        'Modifica descripcion para campo Hora de Fin
        oItemB = oForm.Items.Item("80")
        oStaticText = oItemB.Specific
        oStaticText.Caption = "Fin de Vigencia"

        'Modifica descripcion para campo Duracion
        oItemB = oForm.Items.Item("77")
        oStaticText = oItemB.Specific
        oStaticText.Caption = "Periodo Vigencia"

        'Oculta campo Numero
        oForm.Items.Item("5").Visible = False
        oForm.Items.Item("8").Visible = False

        '//*************************
        '// Adding a EditText Numero Endoso
        '//*************************
        oItemB = oForm.Items.Item("5")
        oItem = oForm.Items.Add("NumEndoso", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height


        oEditText = oItem.Specific
        oEditText.DataBind.SetBound(True, "OCLG", "U_NumEndoso")

        '//**********************************
        '// Adding Label para Numero Endoso
        '//**********************************

        oItemB = oForm.Items.Item("8")
        oItem = oForm.Items.Add("StaticTxt1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = oItemB.Left
        oItem.Width = oItemB.Width
        oItem.Top = oItemB.Top
        oItem.Height = oItemB.Height


        oItem.LinkTo = "NumEndoso"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Numero Endoso"



        EventsActivities = TmpStatusEvent
    End Function
End Module
