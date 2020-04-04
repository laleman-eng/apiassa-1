Module ModuleDevolutionSalesInvoice
    Public Function EventsSalesInvoiceGenerarDevolucion(ByRef oSBOApplication As SAPbouiCOM.Application, ByRef oForm As SAPbouiCOM.Form, ByRef bitDevnc As Integer) As Boolean
        Dim Card As String
        Dim oText As SAPbouiCOM.EditText
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oColumns As SAPbouiCOM.Columns
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim Cantidad As String
        Dim Precio As String
        Dim Comisi As String
        Dim oFila As Int32
        Dim i As Int32
        Dim aux As Double
        Dim aux2 As String

        Try
            oForm.Freeze(True)
            If oCodes.ContainsKey(oForm.UniqueID) Then
                oCodes.Remove(oForm.UniqueID)
            End If


            oRs = Nothing

            oText = oForm.Items.Item("4").Specific 'Cliente'

            Card = oText.Value

            oMatrix = oForm.Items.Item("38").Specific 'Detalle 
            oColumns = oMatrix.Columns

            If (bitDevnc = 0) Then  'NC

                Sql = " SELECT TOP 100 T0.[Code]"
                Sql = Sql & " ,T0.[U_DocNume]"
                Sql = Sql & " ,[U_Precio]"
                Sql = Sql & " ,[U_WshCode]"
                Sql = Sql & " ,[U_Quantity]"
                Sql = Sql & " ,[U_ComisionPercent]"
                Sql = Sql & " ,[U_UNumPoliza]"
                Sql = Sql & " ,[U_UReciboInt]"
                Sql = Sql & " ,[U_SignoPoliza]"
                Sql = Sql & " ,[U_SlpCode]"
                Sql = Sql & " ,[U_TaxCode]"
                Sql = Sql & " ,[U_ItemCode]"
                Sql = Sql & " FROM [@DETALLEDEVOLUCION] T0 JOIN [@DEVOLUCION] T1 ON T1.Code = T0.Code"
                Sql = Sql & " WHERE T1.[U_CardCode] = '" & Card & "' AND T1.[U_DocNume] <> 1 AND T1.[U_Devnc] = 0"
                'Sql = Sql & " and T1.[U_DocDate]= '2010-08-24 00:00:00.000' and  '2010-08-24 00:00:00.000'"
            Else 'Devolucion 

                Sql = " SELECT TOP 100 T0.[Code]"
                Sql = Sql & " ,T0.[U_DocNume]"
                Sql = Sql & " ,[U_Precio]"
                Sql = Sql & " ,[U_WshCode]"
                Sql = Sql & " ,[U_Quantity]"
                Sql = Sql & " ,[U_ComisionPercent]"
                Sql = Sql & " ,[U_UNumPoliza]"
                Sql = Sql & " ,[U_UReciboInt]"
                Sql = Sql & " ,[U_SignoPoliza]"
                Sql = Sql & " ,[U_SlpCode]"
                Sql = Sql & " ,[U_TaxCode]"
                Sql = Sql & " ,[U_ItemCode]"
                Sql = Sql & " FROM [@DETALLEDEVOLUCION] T0 JOIN [@DEVOLUCION] T1 ON T1.Code = T0.Code"
                Sql = Sql & " WHERE T1.[U_CardCode] = '" & Card & "' AND T1.[U_DocNume] <> 1 AND T1.[U_Devnc] = 1"

            End If

            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(Sql)
            'If CStr(oRs.Fields.Item("Code").Value) = True Then
            i = 1
            oFila = oMatrix.RowCount
            '''oMatrix.AddRow(oRs.RecordCount)

            While Not oRs.EoF
                If Not oCodes.ContainsKey(oForm.UniqueID) Then
                    oCodes.Add(oForm.UniqueID, oRs.Fields.Item("Code").Value)
                End If
                'Funciones.AddLog(oFila.ToString)
                oText = oMatrix.Columns.Item("1").Cells.Item(oMatrix.RowCount).Specific
                If oText.Value <> "" Then
                    oMatrix.AddRow()
                    oForm.Freeze(False)
                End If

                'Try
                'Funciones.AddLog(oRs.Fields.Item("U_ItemCode").Value)
                oText = oMatrix.Columns.Item("1").Cells.Item(oFila).Specific
                oText.Value = oRs.Fields.Item("U_ItemCode").Value
                'Catch aa As Exception
                '    'Funciones.AddLog("ERROR " + oRs.Fields.Item("U_ItemCode").Value + ", " + aa.Message)
                'End Try

                'Try
                Cantidad = CStr(CInt(oRs.Fields.Item("U_Quantity").Value))
                'Funciones.AddLog("cantidad " + Cantidad)
                oText = oMatrix.Columns.Item("11").Cells.Item(oFila).Specific

                If (oForm.BusinessObject.Type <> "14") Then  'si es Factura
                    oText.Value = "-" & Cantidad
                End If

                'Catch aa As Exception
                '    'Funciones.AddLog("ERROR " + Cantidad + ", " + aa.Message)
                'End Try

                'Try
                Precio = CStr(CDbl(oRs.Fields.Item("U_Precio").Value))
                'Funciones.AddLog("precio " + Precio)
                oText = oMatrix.Columns.Item("14").Cells.Item(oFila).Specific
                oText.Value = Replace(Precio, ",", ".")
                'Catch aa As Exception
                '    'Funciones.AddLog("ERROR " + Precio + ", " + aa.Message)
                'End Try

                'Comentado no se usa 2-17
                'Comisi = CStr(oRs.Fields.Item("U_SlpCode").Value)
                'oCombo = oMatrix.Columns.Item("27").Cells.Item(oFila).Specific
                'oCombo.Select(CStr(oRs.Fields.Item("U_SlpCode").Value), SAPbouiCOM.BoSearchKey.psk_ByValue)

                'Try 
                Comisi = CStr(oRs.Fields.Item("U_ComisionPercent").Value)
                'Funciones.AddLog("comision " + Comisi)

                If (oForm.BusinessObject.Type = "14") Then  'Si es Nota de credito'
                    oText = oMatrix.Columns.Item("15").Cells.Item(oFila).Specific
                    aux = 100 - CDbl(Comisi)
                    aux2 = System.Convert.ToString(aux)
                    aux2 = aux2.Replace(",", ".")
                    oText.Value = aux2
                    'oText.Value = System.Convert.ToString(100 - CDbl(Comisi))
                Else
                    oText = oMatrix.Columns.Item("28").Cells.Item(oFila).Specific
                    oText.Value = Comisi.Replace(",", ".")

                End If

                'Catch aa As Exception
                '    Funciones.AddLog("ERROR " + Comisi + ", " + aa.Message)
                'End Try

                'Try
                '    Funciones.AddLog("u_taxcode " + oRs.Fields.Item("U_TaxCode").Value)
                oText = oMatrix.Columns.Item("160").Cells.Item(oFila).Specific
                oText.Value = oRs.Fields.Item("U_TaxCode").Value
                'Catch aa As Exception
                'Funciones.AddLog("ERROR " + oRs.Fields.Item("U_TaxCode").Value + ", " + aa.Message)
                'End Try

                'Try
                'Funciones.AddLog("U_unimpoliza " + oRs.Fields.Item("U_UNumPoliza").Value)
                oText = oMatrix.Columns.Item("U_NumPoliza").Cells.Item(oFila).Specific
                oText.Value = oRs.Fields.Item("U_UNumPoliza").Value
                'Catch aa As Exception
                'Funciones.AddLog("ERROR " + oRs.Fields.Item("U_UNumPoliza").Value + ", " + aa.Message)
                'End Try

                'Try
                'Funciones.AddLog("u_ureciboint " + oRs.Fields.Item("U_UReciboInt").Value)
                oText = oMatrix.Columns.Item("U_ReciboInt").Cells.Item(oFila).Specific
                oText.Value = oRs.Fields.Item("U_UReciboInt").Value
                'Catch aa As Exception
                '    Funciones.AddLog("ERROR " + oRs.Fields.Item("U_UReciboInt").Value + ", " + aa.Message)
                'End Try

                oCombo = oMatrix.Columns.Item("U_SignoPoliza").Cells.Item(oFila).Specific
                oCombo.Select(CStr(oRs.Fields.Item("U_SignoPoliza").Value), SAPbouiCOM.BoSearchKey.psk_ByValue)

                i = i + 1
                oFila = oFila + 1
                oRs.MoveNext()
            End While

        Catch ex As Exception
            Funciones.AddLog("Error cargar devoluciones, " + ex.Message + ", StackTrace " + ex.StackTrace)
            oSBOApplication.MessageBox("Error cargar devoluciones, " + ex.Message + ", StackTrace " + ex.StackTrace)
        Finally
            oForm.Freeze(False)
        End Try
        EventsSalesInvoiceGenerarDevolucion = True

    End Function
End Module

