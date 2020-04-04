Imports System.Windows.Forms
Imports SAPbouiCOM
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO

Module AtualizarCC
    Public SBOApp As SAPbouiCOM.Application

    Public Sub Cargar_Form(ByRef oForm As SAPbouiCOM.Form)

        oForm.DataSources.UserDataSources.Add("FechIni", BoDataType.dt_DATE, 12)
        oForm.DataSources.UserDataSources.Item("FechIni").ValueEx = Format(Now, "yyyyMMdd")
        oEditText = oForm.Items.Item("FechIni").Specific
        oEditText.DataBind.SetBound(True, "", "FechIni")

        oForm.DataSources.UserDataSources.Add("FechFin", BoDataType.dt_DATE, 12)
        oForm.DataSources.UserDataSources.Item("FechFin").ValueEx = Format(Now, "yyyyMMdd")
        oEditText = oForm.Items.Item("FechFin").Specific
        oEditText.DataBind.SetBound(True, "", "FechFin")

        oForm.DataSources.UserDataSources.Add("Cuenta", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        AddChooseFromList()
        oEditText = oForm.Items.Item("Cuenta").Specific
        oEditText.DataBind.SetBound(True, "", "Cuenta")
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "AcctCode"


    End Sub

    'ChooseFromList para codigo Articulo
    Private Sub AddChooseFromList()
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = oForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL1"

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "Levels"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "3"
            oCon.Relationship = BoConditionRelationship.cr_AND
            oCFL.SetConditions(oCons)

            oCon = oCons.Add()
            oCon.Alias = "GroupMask"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
            oCon.CondVal = "4"
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            AddLog(ex.Message & ",TRACE " & ex.StackTrace)
        End Try
    End Sub

    Public Function Actualizar_CC(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim oJournal As SAPbobsCOM.JournalEntries
        Dim oProgBar As SAPbouiCOM.ProgressBar
        Dim ContErr As Integer = 0
        Dim FechIni As DateTime
        Dim FechFin As DateTime
        Try
            

            FechIni = CDate(oForm.DataSources.UserDataSources.Item("FechIni").Value)


            FechFin = CDate(oForm.DataSources.UserDataSources.Item("FechFin").Value)


            s = "SELECT DISTINCT T0.TransId" & Environment.NewLine
            s += "  FROM OJDT T0" & Environment.NewLine
            s += "  JOIN JDT1 T1 ON T1.TransId = T0.TransId" & Environment.NewLine
            s += "  JOIN OACT T2 ON T2.AcctCode = T1.Account" & Environment.NewLine
            s += " WHERE T0.RefDate BETWEEN '" & FechIni.ToString("yyyyMMdd") & "' AND '" & FechFin.ToString("yyyyMMdd") & "'" & Environment.NewLine
            s += "   AND T2.FatherNum = '" & oForm.DataSources.UserDataSources.Item("Cuenta").Value & "'" & Environment.NewLine
            s += "   AND T0.TransType IN ('13','18')" & Environment.NewLine
            s += "   AND T2.Dim1Relvnt = 'Y'" & Environment.NewLine
            s += "   AND ISNULL(T2.OverCode,'') <> ''" & Environment.NewLine
            s += " ORDER BY T0.TransId ASC" & Environment.NewLine


            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(s)
            If oRs.RecordCount > 0 Then
                oProgBar = SBOApp.StatusBar.CreateProgressBar("Actualizando asientos contables", oRs.RecordCount, True)
                While oRs.EoF = False
                    oProgBar.Value += 1
                    oJournal = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                    If oJournal.GetByKey(Convert.ToInt32(oRs.Fields.Item("TransId").Value)) Then
                        s = "SELECT T1.Line_ID" & Environment.NewLine
                        s += "      ,T1.Account" & Environment.NewLine
                        s += "      ,T2.OverCode 'CC'" & Environment.NewLine
                        s += "  FROM JDT1 T1" & Environment.NewLine
                        s += "  JOIN OACT T2 ON T2.AcctCode = T1.Account" & Environment.NewLine
                        s += " WHERE T1.TransID = " & Convert.ToInt32(oRs.Fields.Item("TransId").Value) & Environment.NewLine
                        s += "   AND T2.FatherNum = '" & oForm.DataSources.UserDataSources.Item("Cuenta").Value & "'" & Environment.NewLine
                        s += "   AND T2.Dim1Relvnt = 'Y'" & Environment.NewLine
                        s += "   AND ISNULL(T2.OverCode,'') <> ''" & Environment.NewLine

                        oRS2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS2.DoQuery(s)
                        While oRS2.EoF = False
                            oJournal.Lines.SetCurrentLine(Convert.ToInt32(oRS2.Fields.Item("Line_ID").Value))
                            oJournal.Lines.CostingCode = oRS2.Fields.Item("CC").Value
                            oRS2.MoveNext()
                        End While

                        If oRS2.RecordCount > 0 Then
                            lRetCode = oJournal.Update()
                            If lRetCode <> 0 Then
                                ContErr += 1
                                oCompany.GetLastError(lErrCode, sErrMsg)
                                AddLog("Error actualizar asiento " & oRs.Fields.Item("TransId").Value & ":" & sErrMsg)
                            End If
                        End If
                        oRS2 = Nothing
                    End If
                    oRs.MoveNext()
                End While
                oProgBar.Stop()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
                GC.WaitForPendingFinalizers()
                GC.Collect()
                If ContErr > 0 Then
                    SBOApp.StatusBar.SetText(ContErr.ToString & " asientos no se han podido actualizar, revisar log", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                End If
            Else
                SBOApp.MessageBox("No se han encontrado registros")
            End If

            Return True
        Catch ex As Exception
            AddLog("Error Actualizar_CC: " & ex.Message & ", TRACE" & ex.StackTrace)
            SBOApp.MessageBox("Error al actualizar centro de costo")
            Return False
        End Try
    End Function
End Module
