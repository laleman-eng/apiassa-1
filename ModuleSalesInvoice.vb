Module ModuleSalesInvoice
    Public Function EventsSalesSalesInvoiceForm(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form, ByRef oEvent As SAPbouiCOM.ItemEvent) As Boolean
        Dim TmpStatusEvent As Boolean
        Dim oButton As SAPbouiCOM.Button
        TmpStatusEvent = True

        oForm.Items.Add("btnDevol", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem = oForm.Items.Item("btnDevol")
        oItem.Top = oForm.Items.Item("10000330").Top
        oItem.Left = oForm.Items.Item("10000330").Left - oForm.Items.Item("10000330").Width - 5
        oItem.Width = oForm.Items.Item("10000330").Width
        oItem.Height = oForm.Items.Item("10000330").Height
        oItem.FromPane = 0
        oItem.ToPane = 0
        oItem.LinkTo = "10000330"
        oButton = oItem.Specific
        oButton.Caption = "Devoluciones"

        EventsSalesSalesInvoiceForm = TmpStatusEvent
    End Function
End Module
