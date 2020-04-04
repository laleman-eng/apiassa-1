Module RevisarGenerarDevolucion
    Public Function RevisarGrillaGenerarDevolucion(ByRef Code As String) As Boolean
        Dim Card As String
        Dim oText As SAPbouiCOM.EditText
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oColumns As SAPbouiCOM.Columns
        Dim Recibo As String
        Dim oFila As Integer

        If oCodes.ContainsKey(oForm.UniqueID) Then
            oCodes.Remove(oForm.UniqueID)
        End If

        oRs = Nothing

        oText = oForm.Items.Item("4").Specific
        Card = oText.Value

        oMatrix = oForm.Items.Item("38").Specific
        oColumns = oMatrix.Columns


        Sql = " SELECT T0.[Code]"
        Sql = Sql & " ,[U_UReciboInt]"
        Sql = Sql & " FROM [@DETALLEDEVOLUCION] T0 JOIN [@DEVOLUCION] T1 ON T1.Code = T0.Code"
        Sql = Sql & " WHERE T0.[Code] = '" & Code & "' AND T1.[U_DocNume] <> 1"

        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(Sql)

        While Not oRs.EoF
            If Not oCodes.ContainsKey(oForm.UniqueID) Then
                oCodes.Add(oForm.UniqueID, oRs.Fields.Item("Code").Value)
            End If

            oText = oMatrix.Columns.Item("1").Cells.Item(oMatrix.RowCount).Specific
            If oText.Value <> "" Then
                oMatrix.AddRow()
            End If
            oFila = oMatrix.RowCount

            oText = oMatrix.Columns.Item("U_ReciboInt").Cells.Item(oFila).Specific
            Recibo = oRs.Fields.Item("U_UReciboInt").Value
            If oText.Value = Recibo Then
                RevisarGrillaGenerarDevolucion = True
            Else
                RevisarGrillaGenerarDevolucion = False
            End If
            oRs.MoveNext()
        End While

    End Function
End Module
