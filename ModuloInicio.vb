
Imports System.Windows.Forms
Imports VisualD.MultiFunctions
Imports VisualD.SBOFunctions
Imports VisualD.SBOObjectMg1
Imports VisualD.untLog
Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Imports SAPbobsCOM
Imports System.Globalization



Friend Class Paramconexion


    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Private Sub SetApplication()

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        SboGuiApi = New SAPbouiCOM.SboGuiApi

        '// by following the steps specified above, the following
        '// statment should be suficient for either development or run mode
        'Comentario de prueba
        sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
        If sConnectionString = "" Or sConnectionString = Nothing Then
            sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        End If

        '// connect to a running SBO Application

        SboGuiApi.Connect(sConnectionString)

        '// get an initialized application object

        SBO_Application = SboGuiApi.GetApplication(-1)

    End Sub

    Private Function ConectarCompany2007() As Boolean
        Dim s As String


        Try
            oCompany = SBO_Application.Company.GetDICompany()
            oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
            Return (True)
        Catch ex As Exception

            s = "Error en conexión a SAP SBO 2007 (DI)." + ex.Message + " ** Trace: " + ex.StackTrace
            Funciones.AddLog(s)
            'oCompany.SetStatusBarMessage(s, BoMessageTime.bmt_Short, True)
            MessageBox.Show(s)
            Return (False)
        End Try

    End Function

    Private Sub ConectarCompany()

        Dim sCookie As String
        Dim sConnectionContext As String
        Dim errNum As Integer
        Dim errStr As String
        Dim s As String

        If Not ConectarCompany2007() Then

            Try
                oCompany = New SAPbobsCOM.CompanyClass()

                sCookie = oCompany.GetContextCookie
                sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

                If (oCompany.Connected) Then
                    oCompany.Disconnect()
                End If


                errNum = oCompany.SetSboLoginContext(sConnectionContext)
                If (errNum <> 0) Then
                    Funciones.AddLog("ConectarCompany:  Fallo Login Context " + TMultiFunctions.IntToStr(errNum))
                    SBO_Application.MessageBox("Fallo Login Context  " + TMultiFunctions.IntToStr(errNum), 1, "Ok", "", "")


                End If
                'repeat()
                errNum = oCompany.Connect
                If (errNum <> 0) Then
                    oCompany.GetLastError(errNum, errStr)
                    Funciones.AddLog("ConectarCompany: Fallo conexión a Base de Datos (DIAPI) " + errStr)
                    System.Windows.Forms.Application.Exit()
                End If
                'until(errNum = 0)


                oCompany.XmlExportType = BoXmlExportTypes.xet_ExportImportMode
            Catch ex As Exception
                MessageBox.Show(s)
                s = "Error en conexión a SAP SBO (DI).: " & ex.Message & " ** Trace: " & ex.StackTrace
                MessageBox.Show(s)
                Funciones.AddLog(s)
                System.Windows.Forms.Application.Exit()
            End Try
        End If

    End Sub

    Private Sub Class_Initialize_Renamed()
        Try
            Dim RutaC As String = System.Reflection.Assembly.GetExecutingAssembly.Location
            Dim Exe As String = Dir(RutaC)
            sPath = Microsoft.VisualBasic.Left(RutaC, Len(RutaC) - Len(Exe)) & "VKLog.log"

            SetApplication()

            Try
                ConectarCompany()
            Catch exa As Exception
                Funciones.AddLog("Error ConectarCompany, " + exa.Message)
                MessageBox.Show("El addon ha dejado de funcionar")
                System.Windows.Forms.Application.Exit()
            End Try

            'oCompany = SBO_Application.Company.GetDICompany()
            CrearEstructuraDevol()

            'Cambia descripcion al menu Llamada de servicio
            'SBO_Application.Menus.Item("3587").String = "Detalle de Pólizas/Fianzas"


            'Cambia descripcion al menu Registro de tarjeta
            SBO_Application.Menus.Item("3591").String = "Registro Recibo Pólizas"

            'Cambia Descripcion del Menu Contrato de Servicio
            SBO_Application.Menus.Item("3585").String = "Registro Póliza"


            'Cambia descripcion al menu Servicio
            SBO_Application.Menus.Item("3584").String = "Pólizas/Fianzas"
            'Add field... "Ramo" en Llamada de servicio
            '  oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            ' oUFields.TableName = "OSCL"
            'oUFields.Name = "Ramo"
            'oUFields.Description = "Ramo"
            'oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            'oUFields.EditSize = 10
            'lRetCode = oUFields.Add()
            ''oUFields = Nothing
            'GC.Collect()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            'oUFields = Nothing
            'GC.WaitForPendingFinalizers()
            'GC.Collect()
            'SBO_Application.StatusBar.SetText("Se ha creado campo Ramo, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "SubRamo" en Llamada de servicio
            'oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            'oUFields.TableName = "OSCL"
            'oUFields.Name = "SubRamo"
            'oUFields.Description = "Sub Ramo"
            'oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            'oUFields.EditSize = 10
            'lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            'oUFields = Nothing
            'GC.WaitForPendingFinalizers()
            'GC.Collect()
            'SBO_Application.StatusBar.SetText("Se ha creado campo Sub Ramo, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo o Asociado" en Llamada de servicio
            'oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            'oUFields.TableName = "OSCL"
            'oUFields.Name = "Asociado"
            'oUFields.Description = "Asociado"
            'oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            'oUFields.EditSize = 20
            'lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            'oUFields = Nothing
            'GC.WaitForPendingFinalizers()
            'GC.Collect()
            'SBO_Application.StatusBar.SetText("Se ha creado campo Asociado, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo o Asociado" en Llamada de servicio
            'oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            'oUFields.TableName = "OSCL"
            'oUFields.Name = "Ejecutivo"
            'oUFields.Description = "Ejecutivo"
            'oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            'oUFields.EditSize = 20
            'lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            'oUFields = Nothing
            'GC.WaitForPendingFinalizers()
            'GC.Collect()
            'SBO_Application.StatusBar.SetText("Se ha creado campo Ejecutivo, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "NumPoliza" en Llamda de servicio
            'oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            'oUFields.TableName = "OSCL"
            'oUFields.Name = "NumPoliza"
            'oUFields.Description = "Numero Póliza"
            'oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
            'lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            'oUFields = Nothing
            'GC.WaitForPendingFinalizers()
            'GC.Collect()
            'SBO_Application.StatusBar.SetText("Se ha creado campo Numero Póliza, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "NumEndoso" en Llamada de servicio
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OCLG"
            oUFields.Name = "NumEndoso"
            oUFields.Description = "Numero Endoso"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Número Endoso, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Entrega" en Tarjeta de equipo
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OINS"
            oUFields.Name = "DocNum"
            oUFields.Description = "Entrega"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
            oUFields.EditSize = 11
            lRetCode = oUFields.Add()
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
                'MsgBox(sErrMsg)
            End If
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Asociado, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo o Asociado" en Tarjeta de equipo
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OINS"
            oUFields.Name = "Asociado"
            oUFields.Description = "Asociado"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Asociado, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo o Asociado" en Tarjeta de equipo
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OINS"
            oUFields.Name = "Ejecutivo"
            oUFields.Description = "Ejecutivo"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Ejecutivo, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo o Asociado" en Socios de negocios
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OCRD"
            oUFields.Name = "SlpCode"
            oUFields.Description = "Ejecutivo"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Ejecutivo, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo o Asociado" en Socios de negocios
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OCRD"
            oUFields.Name = "Asociado"
            oUFields.Description = "Asociado"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Asociado, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo A" en Oportunidad de Ventas
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OOPR"
            oUFields.Name = "SlpCode"
            oUFields.Description = "Ejecutivo A"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Ejecutivo A, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo B" en Oportunidad de Ventas
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OOPR"
            oUFields.Name = "SlpCodeB"
            oUFields.Description = "Ejecutivo B"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Ejecutivo B, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo o Asociado" en Oportunidad de Ventas
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "OOPR"
            oUFields.Name = "Asociado"
            oUFields.Description = "Asociado B"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Asociado, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Ejecutivo o Asociado" en Entrega de venta
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "ODLN"
            oUFields.Name = "Asociado"
            oUFields.Description = "Asociado"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Asociado, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Numero Poliza" en Entrega de venta
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "DLN1"
            oUFields.Name = "NumPoliza"
            oUFields.Description = "Numero Poliza"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 40
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Numero Poliza, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Numero Poliza" en Entrega de venta
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "DLN1"
            oUFields.Name = "ReciboInt"
            oUFields.Description = "Numero Recibo"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 40
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Numero Recibo Interno, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            'Add field... "Ejecutivo o Asociado" en Entrega de venta
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "ODLN"
            oUFields.Name = "Ejecutivo"
            oUFields.Description = "Ejecutivo"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Ejecutivo, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Cliente Indirecto" en Entrega de venta
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "ODLN"
            oUFields.Name = "CardCode"
            oUFields.Description = "Cliente Indirecto"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 20
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Cliente Indirecto, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Add field... "Nombre Cliente Indirecto" en Entrega de venta
            oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUFields.TableName = "ODLN"
            oUFields.Name = "CardName"
            oUFields.Description = "Nombre"
            oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUFields.EditSize = 30
            lRetCode = oUFields.Add()
            'oUFields = Nothing
            'GC.Collect()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
            oUFields = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
            SBO_Application.StatusBar.SetText("Se ha creado campo Ejecutivo o Asociado, espere confirmación de creación de campo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'MENSAJE DE ADDON CONECTADO
            'SBO_Application.MessageBox("DI Connected To: " & oCompany.CompanyName & vbNewLine & "This is the new UI,DI connection")
            SBO_Application.StatusBar.SetText("El menú ha sido cargado satisfactoriamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            MessageBox.Show("Error " & ex.Message & " Trace " & ex.StackTrace)
        End Try



            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
            'oUserTablesMD = Nothing
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            'oUserTablesMD = Nothing
            'GC.WaitForPendingFinalizers()
            'GC.Collect()
            oUserTablesMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            oUserTablesMD.TableName = "CONEXAP"
            oUserTablesMD.TableDescription = "Conexion Apianet"
            oUserTablesMD.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterData
            lRetCode = oUserTablesMD.Add
            '// check for errors in the process
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                'oUserTablesMD = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                oUserTablesMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            Else
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                oUserTablesMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
                oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUFields.TableName = "CONEXAP"
                oUFields.Name = "Server"
                oUFields.Description = "Servidor"
                oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUFields.EditSize = 30
                lRetCode = oUFields.Add()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    'MsgBox(sErrMsg)
                End If
                'oUFields = Nothing
                'GC.Collect()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
                oUFields = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()

                oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUFields.TableName = "CONEXAP"
                oUFields.Name = "UserSQL"
                oUFields.Description = "Usuario"
                oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUFields.EditSize = 20
                lRetCode = oUFields.Add()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    'MsgBox(sErrMsg)
                End If
                'oUFields = Nothing
                'GC.Collect()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
                oUFields = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()

                oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUFields.TableName = "CONEXAP"
                oUFields.Name = "PassSQL"
                oUFields.Description = "Password"
                oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUFields.EditSize = 25
                lRetCode = oUFields.Add()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    'MsgBox(sErrMsg)
                End If
                'oUFields = Nothing
                'GC.Collect()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
                oUFields = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()

                oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUFields.TableName = "CONEXAP"
                oUFields.Name = "BD"
                oUFields.Description = "Base Datos"
                oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUFields.EditSize = 30
                lRetCode = oUFields.Add()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    'MsgBox(sErrMsg)
                End If
                'oUFields = Nothing
                'GC.Collect()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
                oUFields = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()

                oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUFields.TableName = "CONEXAP"
                oUFields.Name = "Blanco"
                oUFields.Description = "Blanco"
                oUFields.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUFields.EditSize = 20
                lRetCode = oUFields.Add()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    'MsgBox(sErrMsg)
                End If
                'oUFields = Nothing
                'GC.Collect()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
                oUFields = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()

            End If

            'Crea UDO con tabla de Usuario
            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            oUserObjectMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
            oUserObjectMD.TableName = "CONEXAP"
            oUserObjectMD.Code = "CONEXAP"
            oUserObjectMD.Name = "CONEXAP"
            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
            lRetCode = oUserObjectMD.Add
            '// check for errors in the process
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
                'MsgBox(sErrMsg)
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing

            Try
                oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs.DoQuery("SELECT Code FROM [@CONEXAP]")
            Catch ex As Exception
                SBO_Application.MessageBox("Ha fallado la creacion UDT" & vbNewLine & "Espere unos minutos y vuelva a ingresar a SAP")
            End Try


            'Crea Menus
            AddMenuItems()


    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    Private Sub SBO_Application_AppEvent1(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent

    End Sub

    Private Sub CrearEstructuraDevol()
        Dim SBOFunctions As CSBOFunctions
        Dim SBOMetaData As TSBOObjectMg
        Dim oLog As TObjectWithLog
        Dim sLogFile As String


        sLogFile = TMultiFunctions.GetExePath & "vd.log"
        oLog = New TObjectWithLog(Me.GetType().Name)
        oLog.EnableLog = True
        oLog.LogFile = sLogFile
        oLog.OutLog()
        oLog.OutLog("Iniciando...")
        oLog.OutLog()

        SBOFunctions = New CSBOFunctions
        SBOMetaData = New TSBOObjectMg
        SBOMetaData.oLog = oLog
        SBOMetaData.oCompany = oCompany
        SBOMetaData.oApp = SBO_Application
        SBOMetaData.oSBOf = SBOFunctions


        SBOFunctions.SBOApp = SBO_Application
        SBOFunctions.Cmpny = oCompany
        SBOFunctions.oLog = oLog



        Dim PathXLS As String

        PathXLS = TMultiFunctions.ExtractFilePath(TMultiFunctions.ParamStr(0)) & "\Docs\" & "EDSHA1.xls"

        oLog.OutLog("InitApp: Estructura de datos - SHA1")
        SBO_Application.StatusBar.SetText("Inicializando SHA1", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
        If Not SBOFunctions.ValidEstructSHA1(PathXLS) Then
            If (Not SBOMetaData.SyncTablasUdos("1.1", PathXLS)) Then
                oLog.OutLog("InitApp: sincronización de Estructura de datos fallo")
                SBO_Application.MessageBox("Estructura de datos con problemas, consulte a soporte...", 1, "Ok", "", "")
            End If
        End If

        PathXLS = TMultiFunctions.ExtractFilePath(TMultiFunctions.ParamStr(0)) & "\Docs\" & "EDAPIASSA.xls"

        If Not SBOFunctions.ValidEstructSHA1(PathXLS) Then
            oLog.OutLog("InitApp: Estructura de datos (2)")
            SBO_Application.StatusBar.SetText("Inicializando AddOn APIASSA", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            If (Not SBOMetaData.SyncTablasUdos("1.1", PathXLS)) Then
                SBOFunctions.DeleteSHA1FromTable("EDAPIASSA.xls")
                oLog.OutLog("InitApp: sincronización de Estructura de datos fallo")
                SBO_Application.MessageBox("Estructura de datos con problemas, consulte a soporte...", 1, "Ok", "", "")
            End If
        End If
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            'Tarjeta de equipo
            If (pVal.FormTypeEx = "60150") Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True) Then
                    ' On Form Load
                    'oForm = SBO_Application.Forms.Item(FormUID)
                    BubbleEvent = EventsCustomerEquipmentCard(SBO_Application, SBO_Application.Forms.Item(FormUID), pVal)
                End If
                'ItemUID 3 es cuando se presiona boton crear en la tarjeta de equipo
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                    If (pVal.BeforeAction) And pVal.ItemUID = "3" And (SBO_Application.Forms.Item(FormUID).Mode = BoFormMode.fm_ADD_MODE) Then
                        BubbleEvent = EventsCustomerEquipmentCardGenerarEntrega(SBO_Application, SBO_Application.Forms.Item(FormUID))
                        If Not BubbleEvent Then
                            SBO_Application.SetStatusBarMessage("ha ocurrido un error em la generacion de la entrega, " & sErrMsg, True)
                        End If
                    End If
                End If
                'Traer Datos de poliza desde Apianet
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.BeforeAction = False And pVal.ItemUID = "44" And pVal.ItemChanged = True) Then
                    BubbleEvent = EventsCustomerEquipmentCardLoadDatos(SBO_Application, SBO_Application.Forms.Item(FormUID))
                End If
            End If

            'Campos de usuario Tarjeta de equipo
            If ((pVal.FormTypeEx = "-60150" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And (pVal.BeforeAction = False)) Then
                ' On Form Load
                oForm = SBO_Application.Forms.Item(FormUID)
                'Desactiva campos de usuario
                oItemB = oForm.Items.Item("U_Asociado")
                oItemB.Enabled = False

                oItemB = oForm.Items.Item("U_Ejecutivo")
                oItemB.Enabled = False

                oItemB = oForm.Items.Item("U_DocNum")
                oItemB.Enabled = False
            End If

            'Actividad
            If ((pVal.FormTypeEx = "651" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True)) Then
                ' On Form Load

                'oForm = SBO_Application.Forms.Item(FormUID)
                BubbleEvent = EventsActivities(SBO_Application, SBO_Application.Forms.Item(FormUID), pVal)

            End If


            'Llamada de servicio
            If ((pVal.FormTypeEx = "60110" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True)) Then
                ' On Form Load

                'oForm = SBO_Application.Forms.Item(FormUID)
                BubbleEvent = EventsServiceCall(SBO_Application, SBO_Application.Forms.Item(FormUID), pVal)

            End If


            'Campos de usuario Llamada de servicio
            ' If ((pVal.FormTypeEx = "-60110" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And (pVal.BeforeAction = False)) Then
            ' On Form Load
            'oForm = SBO_Application.Forms.Item(FormUID)
            'Desactiva campos de usuario
            'oItemB = oForm.Items.Item("U_Ramo")
            'oItemB.Enabled = False

            'oItemB = oForm.Items.Item("U_SubRamo")
            'oItemB.Enabled = False

            'oItemB = oForm.Items.Item("U_Asociado")
            'oItemB.Enabled = False

            'oItemB = oForm.Items.Item("U_Ejecutivo")
            'oItemB.Enabled = False

            'oItemB = oForm.Items.Item("U_NumPoliza")
            'oItemB.Enabled = False
            ' End If

            'Entrega de venta
            If (pVal.FormTypeEx = "140") Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True) Then
                    ' On Form Load
                    'oForm = SBO_Application.Forms.Item(FormUID)
                    BubbleEvent = EventsDeliveryNotesForm(SBO_Application, SBO_Application.Forms.Item(FormUID), pVal)
                End If
            End If

            'Campos de usuario Entrega de venta 
            If ((pVal.FormTypeEx = "-140" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And (pVal.BeforeAction = False)) Then
                ' On Form Load
                oForm = SBO_Application.Forms.Item(FormUID)
                'Desactiva campos de usuario
                oItemB = oForm.Items.Item("U_Asociado")
                oItemB.Enabled = False

                oItemB = oForm.Items.Item("U_Ejecutivo")
                oItemB.Enabled = False
            End If

            'Maestro Socio de Negocios
            If (pVal.FormTypeEx = "134") Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True) Then
                    ' On Form Load
                    'oForm = SBO_Application.Forms.Item(FormUID)
                    BubbleEvent = EventsBusinessPartnersForm(SBO_Application, SBO_Application.Forms.Item(FormUID), pVal)
                End If
            End If

            'Campos de usuario Maestro de Socio Negocio 
            If ((pVal.FormTypeEx = "-134" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And (pVal.BeforeAction = False)) Then
                ' On Form Load
                oForm = SBO_Application.Forms.Item(FormUID)
                'Desactiva campos de usuario
                oItemB = oForm.Items.Item("U_SlpCode")
                oItemB.Enabled = False

                oItemB = oForm.Items.Item("U_Asociado")
                oItemB.Enabled = False
            End If

            'Maestro Oportunidades de Ventas
            If (pVal.FormTypeEx = "320") Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True) Then
                    ' On Form Load
                    'oForm = SBO_Application.Forms.Item(FormUID)
                    BubbleEvent = EventsSalesOpportunitiesForm(SBO_Application, SBO_Application.Forms.Item(FormUID), pVal)
                End If
            End If

            'Crear Botón en modulo Nota Credito 
            If (pVal.FormTypeEx = "179") Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True) Then
                    ' On Form Load
                    'oForm = SBO_Application.Forms.Item(FormUID)
                    BubbleEvent = EventsSalesSalesInvoiceForm(SBO_Application, SBO_Application.Forms.Item(FormUID), pVal)
                End If
            End If

            'Crear Botón en modulo Factura 
            If (pVal.FormTypeEx = "133") Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And (pVal.BeforeAction = True) Then
                    ' On Form Load
                    'oForm = SBO_Application.Forms.Item(FormUID)
                    BubbleEvent = EventsSalesSalesInvoiceForm(SBO_Application, SBO_Application.Forms.Item(FormUID), pVal)
                End If
            End If


            'Campos de usuario Oportunidades de Ventas 
            If ((pVal.FormTypeEx = "-320" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) And (pVal.BeforeAction = False)) Then
                ' On Form Load
                oForm = SBO_Application.Forms.Item(FormUID)
                'Desactiva campos de usuario
                oItemB = oForm.Items.Item("U_SlpCode")
                oItemB.Enabled = False

                oItemB = oForm.Items.Item("U_Asociado")
                oItemB.Enabled = False

                oItemB = oForm.Items.Item("U_SlpCodeB")
                oItemB.Enabled = False
            End If
            'Formulario de conexion a Apianet
            If pVal.FormTypeEx = "CA" Then
                If (pVal.EventType = BoEventTypes.et_ITEM_PRESSED) And (pVal.BeforeAction = True) Then
                    oForm = SBO_Application.Forms.Item(FormUID)
                    If (pVal.ItemUID = "1") And oForm.DataSources.DBDataSources.Item("@CONEXAP").GetValue("Code", 0) = "" Then
                        oForm.DataSources.DBDataSources.Item("@CONEXAP").SetValue("Code", 0, "1")
                        oForm.DataSources.DBDataSources.Item("@CONEXAP").SetValue("Name", 0, "1")
                        oForm.DataSources.DBDataSources.Item("@CONEXAP").SetValue("DocEntry", 0, "1")
                    End If
                    SetPass(oForm)
                End If
                If (pVal.EventType = BoEventTypes.et_ITEM_PRESSED) And (pVal.BeforeAction = False) Then
                    oForm = SBO_Application.Forms.Item(FormUID)
                    HidePass(oForm)
                    If (pVal.ItemUID = "1") And (oForm.SupportedModes = 3) And (pVal.ActionSuccess) Then
                        oForm.SupportedModes = 1
                        oForm.Mode = BoFormMode.fm_OK_MODE
                    End If

                End If

                If (pVal.EventType = BoEventTypes.et_KEY_DOWN) And (pVal.BeforeAction = False) And (pVal.ItemUID = "PassTxt") Then
                    If (pVal.CharPressed = 8) Or (pVal.CharPressed = 36) Then
                        oForm = SBO_Application.Forms.Item(FormUID)
                        oForm.DataSources.DBDataSources.Item("@CONEXAP").SetValue("U_PassSQL", 0, "")
                        oPass = ""
                    End If
                ElseIf (pVal.ItemUID = "PassTxt") And ((pVal.ItemUID = "PassTxt") And ((pVal.CharPressed = 9) Or (pVal.CharPressed = 10) Or (pVal.CharPressed = 13))) Then
                    Exit Sub
                ElseIf (pVal.ItemUID = "PassTxt") And (pVal.CharPressed > 0) Then
                    oPass = oPass + Convert.ToChar(pVal.CharPressed)
                    s = ""
                    'oForm.DataSources.DBDataSources.Item("@CONEXAP").SetValue("U_PassSQL", 0, New String("0"c, oPass.Length))
                    oForm.DataSources.DBDataSources.Item("@CONEXAP").SetValue("U_PassSQL", 0, s.PadRight(oPass.Length, "*"))
                End If

            End If

            '179 NC 133 Factura deudores
            If (pVal.FormTypeEx = "179") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                If (pVal.BeforeAction) And pVal.ItemUID = "btnDevol" And (SBO_Application.Forms.Item(FormUID).Mode = BoFormMode.fm_ADD_MODE) Then
                    BubbleEvent = EventsSalesInvoiceGenerarDevolucion(SBO_Application, SBO_Application.Forms.Item(FormUID), 0)
                    If Not BubbleEvent Then
                        SBO_Application.SetStatusBarMessage("ha ocurrido un error en la generacion de la devolución, " & sErrMsg, True)
                    End If
                End If
            End If

            If (pVal.FormTypeEx = "133") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                If (pVal.BeforeAction) And pVal.ItemUID = "btnDevol" And (SBO_Application.Forms.Item(FormUID).Mode = BoFormMode.fm_ADD_MODE) Then
                    BubbleEvent = EventsSalesInvoiceGenerarDevolucion(SBO_Application, SBO_Application.Forms.Item(FormUID), 1)
                    If Not BubbleEvent Then
                        SBO_Application.SetStatusBarMessage("ha ocurrido un error en la generacion de la devolución, " & sErrMsg, True)
                    End If
                End If
            End If
            'NC
            If (pVal.FormTypeEx = "179") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                If (pVal.BeforeAction = False) And pVal.ItemUID = "1" Then
                    If oCodes.ContainsKey(pVal.FormUID) Then
                        ActualizarDevolNC(oCodes.Item(pVal.FormUID))
                        oCodes.Remove(pVal.FormUID)

                    End If
                End If
            End If
            'Devolucion
            If (pVal.FormTypeEx = "133") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                If (pVal.BeforeAction = False) And pVal.ItemUID = "1" Then
                    If oCodes.ContainsKey(pVal.FormUID) Then
                        ActualizarDevolEN(oCodes.Item(pVal.FormUID))
                        oCodes.Remove(pVal.FormUID)

                    End If
                End If
            End If


            If (pVal.FormTypeEx = "ActCC") And (pVal.BeforeAction = False) Then
                If pVal.ItemUID = "Cuenta" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                        oForm = SBO_Application.Forms.Item(FormUID)
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        oEditText = oForm.Items.Item("Cuenta").Specific
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oDataTable = oCFLEvento.SelectedObjects

                        Dim val As String
                        Try
                            val = oDataTable.GetValue("AcctCode", 0).ToString
                        Catch ex As Exception
                            ' SBO_Application.MessageBox(ex.Message)
                            Exit Sub
                        End Try
                        oForm.DataSources.UserDataSources.Item("Cuenta").ValueEx = val

                    End If
                End If

                If pVal.ItemUID = "btn_1" Then
                    oForm = SBO_Application.Forms.Item(FormUID)
                    If oForm.DataSources.UserDataSources.Item("FechIni").Value = "" Then
                        SBO_Application.MessageBox("Debe ingresar fecha inicio")
                    ElseIf oForm.DataSources.UserDataSources.Item("FechFin").Value = "" Then
                        SBO_Application.MessageBox("Debe ingresar fecha final")
                    ElseIf oForm.DataSources.UserDataSources.Item("Cuenta").Value = "" Then
                        SBO_Application.MessageBox("Debe seleccionar Cuenta titulo")
                    Else
                        If SBO_Application.MessageBox("¿ Desea actualizar asientos contables ?", 1, "Aceptar", "Cerrar", "") = 1 Then
                            If Actualizar_CC(oForm) Then
                                SBO_Application.MessageBox("Asientos actualizados satisfactoriamente")
                            End If
                        End If
                    End If
                End If
            End If


            'ACTUALIZAR UN CAMPO CANTIDAD EN EL MODULO DE FACTURA DE CLIENTE
            'If (pVal.FormTypeEx = "133") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
            '    If (pVal.BeforeAction = False) And pVal.ItemUID = "38" And (pVal.ColUID = "U_SignoPoliza") Then
            '        Dim oCombo As SAPbouiCOM.ComboBox
            '        Dim oEdit As SAPbouiCOM.EditText
            '        Dim oMatrix As SAPbouiCOM.Matrix
            '        Dim sValue As String
            '        oForm = SBO_Application.Forms.Item(pVal.FormUID)                
            '        oMatrix = oForm.Items.Item("38").Specific
            '        oCombo = oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific
            '        oEdit = oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific
            '        sValue = oCombo.Selected.Value
            '        If CInt(sValue) < 0 Then
            '            oEdit.Value = CStr(-1)
            '        Else
            '            oEdit.Value = CStr(1)
            '        End If
            '    End If
            'End If

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("Error " & ex.Message & " Trace " & ex.StackTrace, True)
        End Try
    End Sub

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent

        If EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then ' Finalización de la sesión de SBO
            ' Finalizamos la aplicación
            oCompany.Disconnect()
            System.Windows.Forms.Application.Exit()
            'System.Environment.Exit(0)
            End

        ElseIf EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Then  ' Cambio de compañía
            ' Cerramos la conexión, se reiniciará automáticamente luego del cambio de compañía
            oCompany.Disconnect()
            System.Windows.Forms.Application.Exit()
            'System.Environment.Exit(0)
            End

        ElseIf EventType = SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged Then  ' Cambio de idioma
            ' Llamamos a la rutina de adición de opciones de menú
            Class_Initialize_Renamed()

        ElseIf EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Then  ' Pérdida de la comunicación con la UI
            ' Finalizamos la aplicación
            oCompany.Disconnect()
            System.Windows.Forms.Application.Exit()
            'System.Environment.Exit(0)
            End

        End If


    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        If (pVal.MenuUID = "SincApianet") And (pVal.BeforeAction = False) Then
            Try
                Sincronizar()
            Catch ex As Exception
                MessageBox.Show("Error " & ex.Message & " Trace " & ex.StackTrace)
            End Try
        End If

        If (pVal.MenuUID = "SincTarjeta") And (pVal.BeforeAction = False) Then
            Try
                SincronizarTarjeta()
            Catch ex As Exception
                MessageBox.Show("Error " & ex.Message & " Trace " & ex.StackTrace)
            End Try
        End If

        If (pVal.MenuUID = "SincOC") And (pVal.BeforeAction = False) Then
            Try
                SincronizarOC()
            Catch ex As Exception
                MessageBox.Show("Error " & ex.Message & " Trace " & ex.StackTrace)
            End Try
        End If

        If (pVal.MenuUID = "ConfApiassa") And (pVal.BeforeAction = False) Then
            Try
                ConfigConexion()

            Catch ex As Exception

                MessageBox.Show("Error " & ex.Message & " Trace " & ex.StackTrace)
            End Try
        End If

        If (pVal.MenuUID = "ActJournal") And (pVal.BeforeAction = False) Then
            Try
                Try
                    oForm = SBO_Application.Forms.Item("ActCC_")
                    SBO_Application.MessageBox("El Formulario ya existe")
                Catch
                    Try
                        oForm = Nothing
                        LoadFromXML("Actualizar_Asientos.srf")
                        oForm = SBO_Application.Forms.Item("ActCC_")
                        oForm.Freeze(True)
                        SBOApp = SBO_Application
                        Cargar_Form(oForm)
                    Catch ex As Exception
                        SBO_Application.MessageBox(ex.Message)
                        AddLog(ex.Message & ", TRACE " & ex.StackTrace)
                    Finally
                        oForm.Freeze(False)
                    End Try
                End Try

            Catch ex As Exception

                MessageBox.Show("Error " & ex.Message & " Trace " & ex.StackTrace)
            End Try
        End If

    End Sub


    Public Sub AddMenuItems()

        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem

        'Dim i As Integer '// to be used as counter
        'Dim lAddAfter As Integer
        'Dim sXML As String
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

        'Crea menu sincronizador para Oportunidades de Venta

       ' Try
        '    oMenuItem = SBO_Application.Menus.Item("2560")
        '   oMenus = oMenuItem.SubMenus
            '// Create s sub menu
        '  oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        ' oCreationPackage.UniqueID = "SincApianet"
        'oCreationPackage.String = "Sincronizar Apianet"
        '    oMenus.AddEx(oCreationPackage)
        'Catch ex As Exception

        '        End Try


        'Crea menu Actualiar asientos contables con centros de costos
        Try
            oMenuItem = SBO_Application.Menus.Item("1536")
            oMenus = oMenuItem.SubMenus
            '// Create s sub menu
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.Position = 3
            oCreationPackage.UniqueID = "ActJournal"
            oCreationPackage.String = "Actualizar Centro Costo en Asientos"
            oMenus.AddEx(oCreationPackage)
        Catch ex As Exception

        End Try

        'Crea menu sincronizador para Tarjetas de Equipos
        Try
            oMenuItem = SBO_Application.Menus.Item("3584")
            oMenus = oMenuItem.SubMenus
            '// Create s sub menu
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "SincTarjeta"
            oCreationPackage.String = "Sincronizar Registro pago recibos"
            oCreationPackage.Position = 3
            oMenus.AddEx(oCreationPackage)
        Catch ex As Exception

        End Try

        'Crea menu sincronizador para Ordenes de compra
        Try
            oMenuItem = SBO_Application.Menus.Item("3584")
            oMenus = oMenuItem.SubMenus
            '// Create s sub menu
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "SincOC"
            oCreationPackage.String = "Sincronizar Ordenes Compra"
            oCreationPackage.Position = 4
            oMenus.AddEx(oCreationPackage)
        Catch ex As Exception

        End Try

        'menu configuracion conexion sql
        Try
            oMenuItem = SBO_Application.Menus.Item("3328")
            oMenus = oMenuItem.SubMenus
            '// Create s sub menu
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.Position = 10
            oCreationPackage.UniqueID = "ConfApiassa"
            oCreationPackage.String = "Configurar Conexion Apianet"
            oMenus.AddEx(oCreationPackage)
        Catch ex As Exception

        End Try

    End Sub

    ' Sincronizacion de Oportunidades no en produccion.
    Public Sub Sincronizar()
        Dim _nf As CultureInfo = New System.Globalization.CultureInfo("en-US")
        AddLog("Inicio Sincronizacion de Datos - Oportunidades de Venta")

        If (SBO_Application.MessageBox("¿Desea sincronizar Oportunidades de Venta desde Apianet ?", 1, "Si", "No", "") = 1) Then
            'SBO_Application.MessageBox("Se comenzara a sincronizar SAP con Apianet" & vbNewLine & "  Esta Operacion Tardara unos minutos..!!")
            'Datos de Conexion
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql = "SELECT TOP 1 U_Server AS Server,U_UserSQL AS UserSQL,U_PassSQL AS PassSQL,U_BD AS BD FROM [@CONEXAP]"
            oRs.DoQuery(Sql)
            Servidor = oRs.Fields.Item("Server").Value
            BD_Net = oRs.Fields.Item("BD").Value
            User = oRs.Fields.Item("UserSQL").Value
            Pass = oRs.Fields.Item("PassSQL").Value

            barra = SBO_Application.StatusBar.CreateProgressBar("Inicializando Transferencia", 100, True)
            Sql = " Select T0.bintIDOpVen"
            Sql = Sql & " ,T1.CardCode"
            Sql = Sql & " ,RTRIM(LEFT(T1.CardName,100)) AS CardName"
            Sql = Sql & " ,T1.bintIdCteOperacion"
            Sql = Sql & " ,T1.strRFC AS strRFC"
            Sql = Sql & " ,LEN(T1.strRFC) AS LargoRFC"
            Sql = Sql & " ,T1.bintIdCliente"
            Sql = Sql & " ,CAST(T3.strCveSap AS VarChar(20)) AS Asociado"
            Sql = Sql & " ,CAST(T2.strCveSap AS VarChar(20)) AS Ejecutivo"
            Sql = Sql & " ,CAST(A1.strCveSap AS VarChar(20)) AS AsociadoB"
            Sql = Sql & " ,CAST(A0.strCveSap AS VarChar(20)) AS EjecutivoB"
            Sql = Sql & " ,T0.strObservaciones"
            Sql = Sql & " ,T0.strTitular"
            Sql = Sql & " ,T0.intTPla"
            Sql = Sql & " ,T0.dteInicio"
            Sql = Sql & " ,T0.dteCierreEstimado"
            Sql = Sql & " ,T0.dcmMToP"
            Sql = Sql & " ,T0.IntRate"
            Sql = Sql & " ,T0.IntId"
            Sql = Sql & " ,T0.ChnCrdCode"
            Sql = Sql & " ,T0.strNOp "
            Sql = Sql & " ,T0.bigIntEdo"
            Sql = Sql & " ,T0.intPerTPla "
            Sql = Sql & " FROM appOPVenta T0"
            Sql = Sql & " JOIN appMtrClientes T1 ON T1.CardCode = T0.CardCode		LEFT"
            Sql = Sql & " JOIN appMtrContactoAPIASSA T2	ON T2.SlpCode = T0.SlpCode	LEFT"
            Sql = Sql & " JOIN appMtrContactoAPIASSA T3 ON T3.SlpCode = T0.ChnCrdCode LEFT"
            Sql = Sql & " JOIN appMtrContactoAPIASSA A0	ON A0.SlpCode = T0.SlpCodeB		LEFT"
            Sql = Sql & " JOIN appMtrContactoAPIASSA A1	ON A1.SlpCode = T0.ChnCrdCodeB"
            Sql = Sql & " JOIN appOPVentaDet T4 ON T4.strNOp = T0.strNOp"
            Sql = Sql & " WHERE(Not (T0.bintIDOpVen Is NULL))"
            Sql = Sql & " AND (T4.bitSap = 0)"
            Sql = Sql & " AND T0.bitNSAP = False"
            'Sql = Sql & " AND T0.strNOp = 'RV-C03AIN52011-001'" 'DEJAR COMO COMENTARIO ES SOLO PARA PRUEBA
            Sql = Sql & " GROUP BY T0.bintIDOpVen"
            Sql = Sql & " ,T1.CardCode"
            Sql = Sql & " ,RTRIM(LEFT(T1.CardName,100))"
            Sql = Sql & " ,T1.bintIdCteOperacion"
            Sql = Sql & " ,T1.strRFC"
            Sql = Sql & " ,T1.bintIdCliente"
            Sql = Sql & " ,CAST(T3.strCveSap AS VarChar(20))"
            Sql = Sql & " ,CAST(T2.strCveSap AS VarChar(20))"
            Sql = Sql & " ,T0.strObservaciones"
            Sql = Sql & " ,T0.strTitular"
            Sql = Sql & " ,T0.intTPla"
            Sql = Sql & " ,T0.dteInicio"
            Sql = Sql & " ,T0.dteCierreEstimado"
            Sql = Sql & " ,T0.dcmMToP"
            Sql = Sql & " ,T0.IntRate"
            Sql = Sql & " ,T0.IntId"
            Sql = Sql & " ,T0.ChnCrdCode"
            Sql = Sql & " ,T0.strNOp "
            Sql = Sql & " ,T0.bigIntEdo"
            Sql = Sql & " ,T0.intPerTPla"
            Sql = Sql & " ,CAST(A1.strCveSap AS VarChar(20))"
            Sql = Sql & " ,CAST(A0.strCveSap AS VarChar(20))"
            If conexionSQL() = 0 Then
                comando.Connection = Sqlconn
                comando.CommandText = Sql
                LectSN = comando.ExecuteReader

                Do Until LectSN.Read.ToString <> True
                    Try
                        oRS2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Sql = "SELECT TOP 1 COUNT(*) AS Cont,OpprId FROM OOPR WHERE Name = '" & LectSN.Item("strNOp").ToString & "' GROUP BY OpprId"
                        oRS2.DoQuery(Sql)
                        If oRS2.Fields.Item("Cont").Value > 0 Then
                            ActualizarOportunidades(oRS2.Fields.Item("OpprId").Value, LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, LectSN.Item("Ejecutivo").ToString, LectSN.Item("Asociado").ToString, LectSN.Item("strTitular").ToString, LectSN.Item("intTPla").ToString, LectSN.Item("dteCierreEstimado").ToString, LectSN.Item("dcmMToP").ToString, LectSN.Item("IntRate").ToString, LectSN.Item("IntId").ToString, LectSN.Item("ChnCrdCode").ToString, LectSN.Item("strObservaciones").ToString, LectSN.Item("bigIntEdo").ToString, LectSN.Item("intPerTPla").ToString, LectSN.Item("strNOp").ToString, LectSN.Item("dteInicio").ToString, LectSN.Item("bintIDOpVen").ToString)
                        Else
                            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Sql = "SELECT COUNT(*) AS Cont FROM OCRD WHERE CardCode = '" & LectSN.Item("CardCode").ToString & "'"
                            oRs.DoQuery(Sql)
                            If oRs.Fields.Item("Cont").Value = 0 Then
                                CrearSN(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, LectSN.Item("strRFC").ToString, LectSN.Item("LargoRFC").ToString, LectSN.Item("bintIdCteOperacion").ToString, LectSN.Item("bintIdCliente").ToString, LectSN.Item("Ejecutivo").ToString, LectSN.Item("Asociado").ToString)
                                CrearOportunidades(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, LectSN.Item("Ejecutivo").ToString, LectSN.Item("Asociado").ToString, LectSN.Item("EjecutivoB").ToString, LectSN.Item("AsociadoB").ToString, LectSN.Item("strTitular").ToString, LectSN.Item("intTPla").ToString, LectSN.Item("dteCierreEstimado").ToString, Convert.ToString(Convert.ToDouble(LectSN.Item("dcmMToP"), _nf)), LectSN.Item("IntRate").ToString, LectSN.Item("IntId").ToString, LectSN.Item("ChnCrdCode").ToString, LectSN.Item("strObservaciones").ToString, LectSN.Item("bigIntEdo").ToString, LectSN.Item("intPerTPla").ToString, LectSN.Item("strNOp").ToString, LectSN.Item("dteInicio").ToString)
                            Else
                                ActSN(LectSN.Item("CardCode").ToString, LectSN.Item("bintIdCteOperacion").ToString)
                                CrearOportunidades(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, LectSN.Item("Ejecutivo").ToString, LectSN.Item("Asociado").ToString, LectSN.Item("EjecutivoB").ToString, LectSN.Item("AsociadoB").ToString, LectSN.Item("strTitular").ToString, LectSN.Item("intTPla").ToString, LectSN.Item("dteCierreEstimado").ToString, Convert.ToString(Convert.ToDouble(LectSN.Item("dcmMToP"), _nf)), LectSN.Item("IntRate").ToString, LectSN.Item("IntId").ToString, LectSN.Item("ChnCrdCode").ToString, LectSN.Item("strObservaciones").ToString, LectSN.Item("bigIntEdo").ToString, LectSN.Item("intPerTPla").ToString, LectSN.Item("strNOp").ToString, LectSN.Item("dteInicio").ToString)
                            End If
                            oRs = Nothing
                        End If
                        oRS2 = Nothing
                    Catch ex As Exception
                        'Me.LogVisual.Items.Add("Cliente [" & LectSN.Item(2).ToString & "] [" & LectSN.Item(1).ToString & "] ERROR : " & lErrCode & " - " & sErrMsg)
                        'errores = errores & "cliente [" & LectSN.Item(2).ToString & "] [" & LectSN.Item(1).ToString & "] ERROR : " & lErrCode & " - " & sErrMsg & "<br>"
                        'MessageBox.Show("Error " & ex.Message & " Trace " & ex.StackTrace)
                        AddLog(ex.Message)
                        oRs = Nothing

                    End Try
                Loop
                LectSN.Close()
            End If
            Sqlconn.Close()
            SBO_Application.MessageBox("Transferencia terminada")
            barra.Stop()
        End If
    End Sub


    Public Sub SincronizarTarjeta()
        Dim RutaC As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim Exe As String = Dir(RutaC)
        Dim Asociado1 As String
        Dim Ejecutivo1 As String
        Dim dcmCom As Decimal
        Dim intCom As Decimal
        Try
            Bandera = 1
            sPath = Microsoft.VisualBasic.Left(RutaC, Len(RutaC) - Len(Exe)) & "VKLog.log"
            AddLog("Inicio Sincronizacion de Datos - Registro de Pagos Recibidos")

            If (SBO_Application.MessageBox("¿Desea sincronizar Registro de Pagos Recibidos desde Apianet ?", 1, "Si", "No", "") = 1) Then
                'SBO_Application.MessageBox("Se comenzara a sincronizar SAP con Apianet" & vbNewLine & "  Esta Operacion Tardara unos minutos..!!")
                'Datos de Conexion
                oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Sql = "SELECT TOP 1 U_Server AS Server,U_UserSQL AS UserSQL,U_PassSQL AS PassSQL,U_BD AS BD FROM [@CONEXAP]"
                oRs.DoQuery(Sql)
                Servidor = oRs.Fields.Item("Server").Value
                BD_Net = oRs.Fields.Item("BD").Value
                User = oRs.Fields.Item("UserSQL").Value
                Pass = oRs.Fields.Item("PassSQL").Value

                barra = SBO_Application.StatusBar.CreateProgressBar("Inicializando Transferencia", 100, True)
                'Buscar datos de poliza en Apianet
                Sql = " Select [idstrNoPoliza]" + vbNewLine
                Sql += " ,[strIdAseguradora]" + vbNewLine
                Sql += " ,[strNomAse]" + vbNewLine
                Sql += " ,[strNoReciboInt]" + vbNewLine
                Sql += " ,[CardCode]" + vbNewLine
                Sql += " ,[CardName]" + vbNewLine
                Sql += " ,[strRFC]" + vbNewLine
                Sql += " ,[LargoRFC]" + vbNewLine
                Sql += " ,[bintIdCteOperacion]" + vbNewLine
                Sql += " ,[bintIdCliente]" + vbNewLine
                Sql += " ,[intIva]" + vbNewLine
                'Sql +=  " ,[Asociado]" + vbNewLine
                'Sql +=  " ,[Ejecutivo]" + vbNewLine
                Sql += " ,[dteAplicacion]" + vbNewLine
                Sql += " ,[ItemCode]" + vbNewLine
                Sql += " ,[ItemName]" + vbNewLine
                Sql += " ,[PrecioNet]" + vbNewLine 'adicionar 
                Sql += " ,[intComR]" + vbNewLine
                Sql += " ,[StrCvePro]" + vbNewLine
                Sql += " ,[strCancela]" + vbNewLine
                Sql += " ,[strNotC]" + vbNewLine
                Sql += " FROM [dbo].[vw_EntradasPolizasIn]" + vbNewLine
                Sql += "  WHERE BitSap = 0" + vbNewLine
                Sql += " GROUP BY [idstrNoPoliza]" + vbNewLine
                Sql += " ,[strIdAseguradora]" + vbNewLine
                Sql += " ,[strNomAse]" + vbNewLine
                Sql += " ,[strNoReciboInt]" + vbNewLine
                Sql += " ,[CardCode]" + vbNewLine
                Sql += " ,[CardName]" + vbNewLine
                Sql += " ,[strRFC]" + vbNewLine
                Sql += " ,[LargoRFC]" + vbNewLine
                Sql += " ,[bintIdCteOperacion]" + vbNewLine
                Sql += " ,[bintIdCliente]" + vbNewLine
                Sql += " ,[intIva]" + vbNewLine
                Sql += " ,[dteAplicacion]" + vbNewLine
                Sql += " ,[ItemCode]" + vbNewLine
                Sql += " ,[ItemName]" + vbNewLine
                Sql += " ,[PrecioNet]" + vbNewLine
                Sql += " ,[intComR]" + vbNewLine
                Sql += " ,[StrCvePro]" + vbNewLine
                Sql += " ,[strCancela]" + vbNewLine
                Sql += " ,[strNotC]" + vbNewLine
                Sql += " ORDER BY [dteAplicacion], [strNoReciboInt] " + vbNewLine

                If conexionSQL() = 0 Then
                    comando.Connection = Sqlconn
                    comando.CommandText = Sql
                    LectSN = comando.ExecuteReader

                    Do Until LectSN.Read.ToString <> True
                        Try
                            oRs3 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Sql = "SELECT COUNT(*) AS Cont FROM OINS WHERE ItemCode = '" & LectSN.Item("ItemCode").ToString & "' "
                            Sql = Sql & " AND InternalSN = '" & LectSN.Item("strNoReciboInt").ToString & "' "
                            oRs3.DoQuery(Sql)
                            If oRs3.Fields.Item("Cont").Value > 0 Then
                                AddLog("Registro Póliza ya existe, Codigo " & LectSN.Item("ItemCode").ToString & ", Registro Poliza " & LectSN.Item("strNoReciboInt").ToString)
                                If conexionSQL2() = 0 Then
                                    Sql = "UPDATE " & BD_Net & "..appEmisionPolizaRecibosS SET BitSap = 1, dtesap = GETDATE() WHERE strNoReciboInt = '" & LectSN.Item("strNoReciboInt").ToString & "'"
                                    comando.Connection = Sqlconn2
                                    comando.CommandText = Sql
                                    comando.ExecuteNonQuery()
                                    Sqlconn2.Close()
                                End If
                            Else
                                'PARA OBTENER ASOCIADO Y EJECUTIVO
                                If conexionSQL2() = 0 Then
                                    Sql = "SELECT Asociado, Ejecutivo, dcmCom, intCom "
                                    Sql = Sql & " FROM " & BD_Net & "..vw_EntradasPolizas "
                                    Sql = Sql & " WHERE strNoReciboInt = '" & LectSN.Item("strNoReciboInt").ToString & "'"
                                    comando.Connection = Sqlconn2
                                    comando.CommandText = Sql
                                    AsocEje = comando.ExecuteReader

                                End If
                                Asociado1 = ""
                                Ejecutivo1 = ""
                                Do Until AsocEje.Read.ToString <> True
                                    If AsocEje.Item("Asociado").ToString <> "" Then
                                        Asociado1 = AsocEje.Item("Asociado").ToString
                                        dcmCom = AsocEje.Item("dcmCom").ToString
                                        intCom = AsocEje.Item("IntCom").ToString
                                    Else
                                        Ejecutivo1 = AsocEje.Item("Ejecutivo").ToString
                                    End If
                                Loop

                                If AsocEje.Read.ToString = True Then
                                    AsocEje.Close()
                                    Sqlconn2.Close()
                                End If

                                oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Sql = "SELECT COUNT(*) AS Cont FROM OCRD T0 WHERE T0.CardCode = '" & LectSN.Item("CardCode").ToString & "'"
                                oRs.DoQuery(Sql)
                                'ACA VAN LOS CAMBIOS
                                '1 HACER CONSULTA A LA VISTA vw_ClientesOCRD para obtener todos los datos necesarios
                                '
                                If oRs.Fields.Item("Cont").Value = 0 Then
                                    CrearSN(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, LectSN.Item("strRFC").ToString, LectSN.Item("LargoRFC").ToString, LectSN.Item("bintIdCteOperacion").ToString, LectSN.Item("bintIdCliente").ToString, Ejecutivo1, Asociado1)
                                    CrearTarjeta(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, Ejecutivo1, Asociado1, LectSN.Item("idstrNoPoliza").ToString, LectSN.Item("strNoReciboInt").ToString, LectSN.Item("ItemCode").ToString, LectSN.Item("ItemName").ToString, LectSN.Item("strIdAseguradora").ToString, LectSN.Item("strNomAse").ToString, LectSN.Item("dteAplicacion").ToString, LectSN.Item("PrecioNet").ToString, LectSN.Item("intComR").ToString, LectSN.Item("intIva").ToString, intCom, dcmCom, LectSN.Item("strCancela").ToString, LectSN.Item("strNotC").ToString)
                                Else
                                    ActSN(LectSN.Item("CardCode").ToString, LectSN.Item("bintIdCteOperacion").ToString)
                                    CrearTarjeta(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, Ejecutivo1, Asociado1, LectSN.Item("idstrNoPoliza").ToString, LectSN.Item("strNoReciboInt").ToString, LectSN.Item("ItemCode").ToString, LectSN.Item("ItemName").ToString, LectSN.Item("strIdAseguradora").ToString, LectSN.Item("strNomAse").ToString, LectSN.Item("dteAplicacion").ToString, LectSN.Item("PrecioNet").ToString, LectSN.Item("intComR").ToString, LectSN.Item("intIva").ToString, intCom, dcmCom, LectSN.Item("strCancela").ToString, LectSN.Item("strNotC").ToString)
                                End If
                                oRs = Nothing
                            End If
                            oRs3 = Nothing
                        Catch ex As Exception
                            'Me.LogVisual.Items.Add("Cliente [" & LectSN.Item(2).ToString & "] [" & LectSN.Item(1).ToString & "] ERROR : " & lErrCode & " - " & sErrMsg)
                            'errores = errores & "cliente [" & LectSN.Item(2).ToString & "] [" & LectSN.Item(1).ToString & "] ERROR : " & lErrCode & " - " & sErrMsg & "<br>"
                            'MessageBox.Show("Error " & ex.Message & " Trace " & ex.StackTrace)
                            If Bandera = 1 Then
                                AddLog("Registro Poliza no ha sido creado, Numero Poliza " & LectSN.Item("idstrNoPoliza").ToString & ", Registro interno " & LectSN.Item("strNoReciboInt").ToString & ", " & ex.Message)
                            Else
                                AddLog("Entrega no ha sido creado poliza " & LectSN.Item("idstrNoPoliza").ToString & ", Nro registro Apianet " & LectSN.Item("strNoReciboInt").ToString & ", " & ex.Message)
                                Dim oCustomerEquipmentCard As SAPbobsCOM.CustomerEquipmentCards = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
                                oRs4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Sql = "SELECT InsID FROM OINS WHERE ItemCode = '" & LectSN.Item("ItemCode").ToString & "' AND InternalSN = '" & LectSN.Item("strNoReciboInt").ToString & "'"
                                oRs4.DoQuery(Sql)
                                If oCustomerEquipmentCard.GetByKey(oRs4.Fields.Item("InsID").Value) = True Then
                                    lRetCode = oCustomerEquipmentCard.Remove()
                                    If lRetCode <> 0 Then
                                        oCompany.GetLastError(lErrCode, sErrMsg)
                                        AddLog("Registro Poliza quedo ingresado en SAP sin Entrega, poliza " & LectSN.Item("idstrNoPoliza").ToString & ", Nro registro Apianet " & LectSN.Item("strNoReciboInt").ToString & ", " & ex.Message)
                                    Else
                                        AddLog("Registro Poliza ha sido Borrado por no poder realizar la Entrega, poliza " & LectSN.Item("idstrNoPoliza").ToString & ", Nro registro Apianet " & LectSN.Item("strNoReciboInt").ToString & ", " & ex.Message)
                                    End If
                                End If
                                oRs4 = Nothing
                            End If
                            oRs = Nothing

                        End Try
                    Loop
                    LectSN.Close()
                End If
                Sqlconn.Close()
                SBO_Application.MessageBox("Transferencia terminada")
                Try
                    barra.Stop()
                Catch
                End Try
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("Error " & ex.Message)
            AddLog("Error SincronizarTarjeta: " & ex.Message & " Trace " & ex.StackTrace)
        End Try
    End Sub

    Public Sub SincronizarOC()
        Dim RutaC As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim Exe As String = Dir(RutaC)
        Dim Asociado1 As String
        Dim Asociado As String = ""
        Dim Ejecutivo1 As String
        Dim dcmCom As Decimal
        Dim intCom As Decimal
        Dim intIva As String
        Dim intIva2 As String
        Dim CardCode As String
        Dim PrecioNet As Decimal
        Dim oPurchesaOrder As SAPbobsCOM.Documents
        Dim iLineNum As Integer
        Dim strNoReciboInt As String = ""
        Dim tmparr(2500) As String
        Dim arr(2500) As String
        Bandera = 1
        sPath = Microsoft.VisualBasic.Left(RutaC, Len(RutaC) - Len(Exe)) & "VKLog.log"
        AddLog("Inicio Sincronizacion de Datos - Registro de Polizas")

        If (SBO_Application.MessageBox("¿Desea sincronizar Ordenes Compra desde Apianet ?", 1, "Si", "No", "") = 1) Then
            'SBO_Application.MessageBox("Se comenzara a sincronizar SAP con Apianet" & vbNewLine & "  Esta Operacion Tardara unos minutos..!!")
            'Datos de Conexion
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql = "SELECT TOP 1 U_Server AS Server,U_UserSQL AS UserSQL,U_PassSQL AS PassSQL,U_BD AS BD FROM [@CONEXAP]"
            oRs.DoQuery(Sql)
            Servidor = oRs.Fields.Item("Server").Value
            BD_Net = oRs.Fields.Item("BD").Value
            User = oRs.Fields.Item("UserSQL").Value
            Pass = oRs.Fields.Item("PassSQL").Value

            'barra = SBO_Application.StatusBar.CreateProgressBar("Inicializando Transferencia", 100, True)
            SBO_Application.StatusBar.SetText("Inicializando Transferencia", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
            'Buscar datos de poliza en Apianet
            Sql = " Select [idstrNoPoliza]" + vbNewLine
            Sql += " ,[strIdAseguradora]" + vbNewLine
            Sql += " ,[strNomAse]" + vbNewLine
            Sql += " ,[strNoReciboInt]" + vbNewLine
            Sql += " ,[CardCode]" + vbNewLine
            Sql += " ,[CardName]" + vbNewLine
            Sql += " ,[strRFC]" + vbNewLine
            Sql += " ,[LargoRFC]" + vbNewLine
            Sql += " ,[bintIdCteOperacion]" + vbNewLine
            Sql += " ,[bintIdCliente]" + vbNewLine
            Sql += " ,[intIva]" + vbNewLine
            Sql += " ,[Ejecutivo]" + vbNewLine
            Sql += " ,[dteAplicacion]" + vbNewLine
            Sql += " ,[ItemCode]" + vbNewLine
            Sql += " ,[ItemName]" + vbNewLine
            Sql += " ,[PrecioNet]" + vbNewLine 'adicionar 
            Sql += " ,[intComR]" + vbNewLine
            Sql += " ,[StrCvePro]" + vbNewLine
            Sql += " ,[Asociado]" + vbNewLine
            Sql += " ,[dcmCom]" + vbNewLine
            Sql += " ,[intCom]" + vbNewLine
            Sql += " FROM [dbo].[vw_EntradasPolizas]" + vbNewLine
            Sql += " WHERE ISNULL(Asociado,'') <> ''" + vbNewLine
            Sql += "   AND BitOC = 0" + vbNewLine
            Sql += "   AND BitSap = 1" + vbNewLine
            'Sql += "   AND Asociado = 'PROV0192'" + vbNewLine
            'Sql += "   AND idstrNoPoliza = 'D00-3-1-000402086_0000-0-1'" + vbNewLine
            Sql += " GROUP BY [idstrNoPoliza]" + vbNewLine
            Sql += " ,[strIdAseguradora]" + vbNewLine
            Sql += " ,[strNomAse]" + vbNewLine
            Sql += " ,[strNoReciboInt]" + vbNewLine
            Sql += " ,[CardCode]" + vbNewLine
            Sql += " ,[CardName]" + vbNewLine
            Sql += " ,[strRFC]" + vbNewLine
            Sql += " ,[LargoRFC]" + vbNewLine
            Sql += " ,[bintIdCteOperacion]" + vbNewLine
            Sql += " ,[bintIdCliente]" + vbNewLine
            Sql += " ,[intIva]" + vbNewLine
            Sql += " ,[dteAplicacion]" + vbNewLine
            Sql += " ,[ItemCode]" + vbNewLine
            Sql += " ,[ItemName]" + vbNewLine
            Sql += " ,[PrecioNet]" + vbNewLine
            Sql += " ,[intComR]" + vbNewLine
            Sql += " ,[StrCvePro]" + vbNewLine
            Sql += " ,[Asociado]" + vbNewLine
            Sql += " ,[dcmCom]" + vbNewLine
            Sql += " ,[intCom]" + vbNewLine
            Sql += " ,[Ejecutivo]" + vbNewLine
            Sql += " ORDER BY [Asociado], [strNoReciboInt] " + vbNewLine

            If conexionSQL() = 0 Then
                comando.Connection = Sqlconn
                comando.CommandText = Sql
                LectSN = comando.ExecuteReader
                CardCode = ""
                iLineNum = 0
                intIva = "0"
                oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs3 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oPurchesaOrder = oCompany.GetBusinessObject(BoObjectTypes.oPurchaseOrders)
                Do Until LectSN.Read.ToString <> True
                    Try
                        If CardCode <> LectSN.Item("Asociado").ToString.Trim And CardCode <> "" Then
                            CardCode = LectSN.Item("Asociado").ToString.Trim
                            iLineNum = 0
                            'ReDim arr()

                            lRetCode = oPurchesaOrder.Add()

                            If lRetCode <> 0 Then
                                oCompany.GetLastError(lErrCode, sErrMsg)
                                AddLog("Orden de Compra " + Asociado + " no ha sido creada, " + sErrMsg)
                                oPurchesaOrder.SaveXML("C:\pruebaOC.xml")
                            Else
                                AddLog("Sincronizacion OC Asociado: " & Asociado)
                                If conexionSQL2() = 0 Then
                                    For i As Integer = 0 To arr.Length - 1
                                        If arr(i) = "" Then
                                            Exit For
                                        End If
                                        Sql = "UPDATE " & BD_Net & "..appEmisionPolizaRecibosS SET BitOC = 1, dteOC = GETDATE() WHERE strNoReciboInt = '" & arr(i) & "'"
                                        comando.Connection = Sqlconn2
                                        'AddLog(Sql)
                                        comando.CommandText = Sql
                                        comando.ExecuteNonQuery()
                                    Next
                                    Sqlconn2.Close()
                                End If
                            End If
                            oPurchesaOrder = Nothing
                            'oPurchesaOrder = oCompany.GetBusinessObject(BoObjectTypes.oProductionOrders)
                            arr = tmparr
                        End If

                        CardCode = LectSN.Item("Asociado").ToString.Trim
                        Asociado = CardCode
                        If iLineNum = 0 Then
                            oPurchesaOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            oPurchesaOrder.DocDate = Date.Now
                            oPurchesaOrder.DocDueDate = Date.Now
                            oPurchesaOrder.TaxDate = Date.Now
                            oPurchesaOrder.CardCode = CardCode
                            oPurchesaOrder.Comments = "Pago de Comisiones a Asociados"
                        Else
                            oPurchesaOrder.Lines.Add()
                        End If

                        arr(iLineNum) = LectSN.Item("strNoReciboInt").ToString.Trim

                        ''PARA OBTENER ASOCIADO Y EJECUTIVO
                        'If conexionSQL2() = 0 Then
                        '    Sql = "SELECT Asociado, Ejecutivo, dcmCom, intCom "
                        '    Sql = Sql & " FROM " & BD_Net & "..vw_EntradasPolizas "
                        '    Sql = Sql & " WHERE strNoReciboInt = '" & LectSN.Item("strNoReciboInt").ToString & "'"
                        '    Sql = Sql & "   AND Asociado = '" & LectSN.Item("Asociado").ToString & "'"
                        '    comando.Connection = Sqlconn2
                        '    comando.CommandText = Sql
                        '    AsocEje = comando.ExecuteReader

                        'End If
                        Asociado1 = ""
                        Ejecutivo1 = ""
                        dcmCom = 0
                        intCom = 0
                        'Do Until AsocEje.Read.ToString <> True
                        '    If AsocEje.Item("Asociado").ToString <> "" Then
                        '        Asociado1 = AsocEje.Item("Asociado").ToString
                        '        dcmCom = AsocEje.Item("dcmCom").ToString
                        '        intCom = AsocEje.Item("IntCom").ToString
                        '    Else
                        '        Ejecutivo1 = AsocEje.Item("Ejecutivo").ToString
                        '    End If
                        'Loop

                        'If AsocEje.Read.ToString = True Then
                        '    AsocEje.Close()
                        '    Sqlconn2.Close()
                        'End If
                        If Asociado <> "" Then
                            Asociado1 = Asociado.Trim
                            dcmCom = LectSN.Item("dcmCom").ToString.Trim
                            intCom = LectSN.Item("IntCom").ToString.Trim
                        Else
                            Ejecutivo1 = LectSN.Item("Ejecutivo").ToString.Trim
                        End If

                        oPurchesaOrder.Lines.ItemCode = LectSN.Item("ItemCode").ToString.Trim
                        oPurchesaOrder.Lines.WarehouseCode = "01"
                        If LectSN.Item("StrCvePro").ToString.Trim <> "" Then
                            oPurchesaOrder.Lines.ProjectCode = LectSN.Item("StrCvePro").ToString.Trim
                        End If
                        oPurchesaOrder.Lines.CommisionPercent = intCom
                        If intCom > 0 Then
                            oPurchesaOrder.Lines.DiscountPercent = 100 - intCom
                        Else
                            oPurchesaOrder.Lines.DiscountPercent = 100
                        End If
                        'Impuesto
                        intIva = LectSN.Item("intIva").ToString
                        Sql = "SELECT Code,Name, Rate FROM OSTC WHERE [Lock] <> 'Y' AND [ValidForAP] = 'Y' AND RATE ='" & intIva & "'"
                        oRs.DoQuery(Sql)
                        intIva2 = oRs.Fields.Item("Rate").Value

                        If oRs.Fields.Item("Rate").Value = "16" Then
                            oPurchesaOrder.Lines.TaxCode = "W3"
                        Else
                            If oRs.Fields.Item("Rate").Value = "0" Then
                                oPurchesaOrder.Lines.TaxCode = "W3"
                            Else
                                oPurchesaOrder.Lines.TaxCode = oRs.Fields.Item("Code").Value
                            End If
                        End If
                        oPurchesaOrder.Lines.UserFields.Fields.Item("U_NumPoliza").Value = LectSN.Item("idstrNoPoliza").ToString.Trim
                        oPurchesaOrder.Lines.UserFields.Fields.Item("U_ReciboInt").Value = LectSN.Item("strNoReciboInt").ToString.Trim
                        oPurchesaOrder.Lines.UserFields.Fields.Item("U_dteAplicacion").Value = LectSN.Item("dteAplicacion").ToString

                        'Si devuelven o cancelan una poliza, PrecioNet viene negativo
                        PrecioNet = dcmCom

                        If PrecioNet < 0 Then
                            oPurchesaOrder.Lines.UnitPrice = PrecioNet * -1 ' lo deja en Negativo segu  conversado con Juan Manuel el 201401081514
                            oPurchesaOrder.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = -1
                            oPurchesaOrder.Lines.Quantity = -1
                        Else
                            oPurchesaOrder.Lines.UnitPrice = PrecioNet
                            oPurchesaOrder.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = 1
                            oPurchesaOrder.Lines.Quantity = 1
                        End If

                        iLineNum = iLineNum + 1

                    Catch ex As Exception
                        AddLog("Error SincronizarOC, " + ex.Message)
                    End Try
                Loop

                'Ultimo registro
                If oPurchesaOrder IsNot Nothing Then
                    lRetCode = oPurchesaOrder.Add()
                    If lRetCode <> 0 Then
                        oCompany.GetLastError(lErrCode, sErrMsg)
                        AddLog("Orden de Compra " + Asociado + " no ha sido creada, " + sErrMsg)
                        oPurchesaOrder.SaveXML("C:\pruebaOC.xml")
                    Else
                        AddLog("Ultima Sincronizacion OC Asociado: " & Asociado)
                        If conexionSQL2() = 0 Then
                            For i As Integer = 0 To arr.Length - 1
                                If arr(i) = "" Then
                                    Exit For
                                End If
                                Sql = "UPDATE " & BD_Net & "..appEmisionPolizaRecibosS SET BitOC = 1, dteOC = GETDATE() WHERE strNoReciboInt = '" & arr(i) & "'"
                                comando.Connection = Sqlconn2
                                'AddLog(Sql)
                                comando.CommandText = Sql
                                comando.ExecuteNonQuery()
                            Next
                            Sqlconn2.Close()
                        End If
                    End If
                End If

                oPurchesaOrder = Nothing
                LectSN.Close()
                oRs = Nothing
                oRs3 = Nothing
            End If

            Sqlconn.Close()
            SBO_Application.StatusBar.SetText("Transferencia terminada", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
            SBO_Application.MessageBox("Transferencia terminada")
            'barra.Stop()
        End If
    End Sub


    Public Sub CrearSN(ByVal CardCode As String, ByVal CardName As String, ByVal LicTradNum As String, ByVal LargoRFC As Integer, ByVal TpoCliente As String, ByVal bintIdCliente As String, ByVal Ejecutivo As String, ByVal Asociado As String)
        Dim oBusinessParteners As SAPbobsCOM.BusinessPartners = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        Dim CodeIndirecto As Integer
        Dim CodeProspecto As Integer
        oBusinessParteners.CardCode = CardCode
        oBusinessParteners.CardName = CardName
        oBusinessParteners.FederalTaxID = LicTradNum
        If LargoRFC = 12 Then
            oBusinessParteners.CompanyPrivate = SAPbobsCOM.BoCardCompanyTypes.cCompany
        Else
            oBusinessParteners.CompanyPrivate = SAPbobsCOM.BoCardCompanyTypes.cPrivate
        End If

        Sql = "SELECT GroupCode, GroupName FROM OCRG WHERE GroupName IN ('Indirectos','Prospectos')"
        oRs6 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs6.DoQuery(Sql)
        If oRs6.Fields.Item("GroupName").Value = "Indirectos" Then
            CodeIndirecto = oRs6.Fields.Item("GroupCode").Value
        Else
            CodeProspecto = oRs6.Fields.Item("GroupCode").Value
        End If
        oRs6 = Nothing

        If TpoCliente = "3" Then
            oBusinessParteners.CardType = SAPbobsCOM.BoCardTypes.cCustomer
            oBusinessParteners.GroupCode = CodeIndirecto
        Else
            oBusinessParteners.CardType = SAPbobsCOM.BoCardTypes.cLid
            oBusinessParteners.GroupCode = CodeProspecto
        End If
        If Ejecutivo = "" Then
            oBusinessParteners.UserFields.Fields.Item("U_SlpCode").Value = "-1"
        Else
            oBusinessParteners.UserFields.Fields.Item("U_SlpCode").Value = Ejecutivo
        End If

        If Asociado = "" Then
            oBusinessParteners.UserFields.Fields.Item("U_Asociado").Value = "-1"
        Else
            oBusinessParteners.UserFields.Fields.Item("U_Asociado").Value = Asociado
        End If

        'oBusinessParteners.FreeText = Observaciones
        'BUSCAR TELEFONO Y EMAIL - bintIdCliente
        If conexionSQL2() = 0 Then
            Sql = "SELECT bintIdCliente,bintAppCatTiposGen, strTelefono FROM " & BD_Net & ".dbo.appMtrClientesTELEFONOS WHERE bintIdCliente = '" & bintIdCliente & "'"
            comando2.Connection = Sqlconn2
            comando2.CommandText = Sql
            Detalle = comando2.ExecuteReader
            Dim I As Integer = 1
            Do Until Detalle.Read.ToString <> True
                If I = 1 Then
                    oBusinessParteners.Phone1 = Detalle.Item("strTelefono").ToString
                End If
                If I = 2 Then
                    oBusinessParteners.Phone2 = Detalle.Item("strTelefono").ToString
                End If
                If I = 3 Then
                    oBusinessParteners.Cellular = Detalle.Item("strTelefono").ToString
                End If
                If I = 4 Then
                    oBusinessParteners.Fax = Detalle.Item("strTelefono").ToString
                End If
                I = I + 1
            Loop
            Detalle.Close()
            Sqlconn2.Close()
        End If

        If conexionSQL2() = 0 Then
            Sql = "SELECT bintIdCliente, bintAppCatTiposGen,strMedioEl FROM " & BD_Net & ".dbo.appMtrClientesMElectr  WHERE bintIdCliente = '" & bintIdCliente & "'"
            comando2.Connection = Sqlconn2
            comando2.CommandText = Sql
            Detalle = comando2.ExecuteReader
            Dim I As Integer = 1
            Do Until Detalle.Read.ToString <> True
                If I = 1 Then
                    oBusinessParteners.EmailAddress = Detalle.Item("strMedioEl").ToString
                End If
            Loop
            Detalle.Close()
            Sqlconn2.Close()
        End If


        lRetCode = oBusinessParteners.Add()
        If lRetCode <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            AddLog("SN " & CardCode & " - " & CardName & " - " & LicTradNum & " no ha sido creado, " & sErrMsg)

        Else
            ''
        End If
    End Sub


    Public Sub ActSN(ByVal CardCode As String, ByVal TpoCliente As String)
        Dim oBusinessPartener As SAPbobsCOM.BusinessPartners = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        Dim CodeIndirecto As Integer
        Dim CodeProspecto As Integer
        Sql = "SELECT GroupCode, GroupName FROM OCRG WHERE GroupName IN ('Indirectos','Prospectos')"
        oRs6 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs6.DoQuery(Sql)
        If oRs6.Fields.Item("GroupName").Value = "Indirectos" Then
            CodeIndirecto = oRs6.Fields.Item("GroupCode").Value
        Else
            CodeProspecto = oRs6.Fields.Item("GroupCode").Value
        End If
        oRs6 = Nothing

        If oBusinessPartener.GetByKey(CardCode) = True Then
            If TpoCliente = "3" Then
                oBusinessPartener.CardType = SAPbobsCOM.BoCardTypes.cCustomer
                oBusinessPartener.GroupCode = CodeIndirecto
            Else
                oBusinessPartener.CardType = SAPbobsCOM.BoCardTypes.cLid
                oBusinessPartener.GroupCode = CodeProspecto
            End If
            lRetCode = oBusinessPartener.Update
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
                AddLog("Cliente " & CardCode & " no ha sido actualizado, " & sErrMsg)
            Else
            End If
        End If

    End Sub


    Public Sub CrearOportunidades(ByVal CardCode As String, ByVal CardName As String, ByVal Ejecutivo As String, ByVal Asociado As String, ByVal EjecutivoB As String, ByVal AsociadoB As String, ByVal strTitular As String, ByVal intTPla As String, ByVal dteCierreEstimado As Date, ByVal dcmMToP As String, ByVal IntRate As String, ByVal IntId As String, ByVal ChnCrdCode As String, ByVal strObservaciones As String, ByVal bigIntEdo As String, ByVal intPerTPla As String, ByVal strNOp As String, ByVal dteInicio As Date)
        Dim oSalesOpportunities As SAPbobsCOM.SalesOpportunities = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesOpportunities)
        oSalesOpportunities.CardCode = CardCode
        oSalesOpportunities.OpportunityName = strNOp
        If Ejecutivo = "" Then
            oSalesOpportunities.UserFields.Fields.Item("U_SlpCode").Value = "-1"
        Else
            oSalesOpportunities.UserFields.Fields.Item("U_SlpCode").Value = Ejecutivo
        End If

        If Asociado <> "" Then
            oSalesOpportunities.BPChanelCode = Asociado
        End If

        If EjecutivoB = "" Then
            oSalesOpportunities.UserFields.Fields.Item("U_SlpCodeB").Value = "-1"
        Else
            oSalesOpportunities.UserFields.Fields.Item("U_SlpCodeB").Value = EjecutivoB
        End If

        If AsociadoB = "" Then
            oSalesOpportunities.UserFields.Fields.Item("U_Asociado").Value = "-1"
        Else
            oSalesOpportunities.UserFields.Fields.Item("U_Asociado").Value = AsociadoB
        End If

        'oSalesOpportunities.DataOwnershipfield = strTitular
        'oSalesOpportunities.ClosingDate = intTPla 'LO CALCULA SAP
        If intPerTPla = 1 Then
            oSalesOpportunities.ClosingType = SAPbobsCOM.BoSoClosedInTypes.sos_Months
        End If
        If intPerTPla = 2 Then
            oSalesOpportunities.ClosingType = SAPbobsCOM.BoSoClosedInTypes.sos_Weeks
        Else
            oSalesOpportunities.ClosingType = SAPbobsCOM.BoSoClosedInTypes.sos_Days
        End If
        oSalesOpportunities.StartDate = dteInicio
        oSalesOpportunities.PredictedClosingDate = dteCierreEstimado

        oSalesOpportunities.InterestLevel = IntRate
        'oSalesOpportunities.Interests.InterestId = IntId
        'oSalesOpportunities.BPChanelCode = ChnCrdCode
        oSalesOpportunities.Remarks = strObservaciones
        If bigIntEdo = 1 Then
            oSalesOpportunities.Status = SAPbobsCOM.BoSoOsStatus.sos_Open
        Else
            If bigIntEdo = 2 Then
                oSalesOpportunities.Status = SAPbobsCOM.BoSoOsStatus.sos_Sold
            Else
                oSalesOpportunities.Status = SAPbobsCOM.BoSoOsStatus.sos_Missed
            End If
        End If

        'ETAPA strNOp
        Dim I As Integer = 1
        Dim X As Integer = 0
        If conexionSQL2() = 0 Then
            Sql = "SELECT dteIniEta, dteFinEt, Step_Id, StrMovimiento, biginTPartida AS LineNum FROM " & BD_Net & ".dbo.appOPVentaDet WHERE strNOp = '" & strNOp & "' AND BitSAP = 0"
            Sql = Sql & " ORDER BY biginTPartida"
            comando2.Connection = Sqlconn2
            comando2.CommandText = Sql
            Detalle = comando2.ExecuteReader
            I = 1
            Do Until Detalle.Read.ToString <> True
                If I > 1 Then
                    oSalesOpportunities.Lines.Add()
                End If
                oSalesOpportunities.Lines.StartDate = Detalle.Item("dteIniEta").ToString
                oSalesOpportunities.Lines.ClosingDate = Detalle.Item("dteFinEt").ToString
                oSalesOpportunities.Lines.StageKey = Detalle.Item("Step_Id").ToString
                oSalesOpportunities.Lines.Remarks = Detalle.Item("StrMovimiento").ToString
                oSalesOpportunities.Lines.MaxLocalTotal = dcmMToP
                I = I + 1
                X = Detalle.Item("LineNum").ToString
            Loop
            Detalle.Close()
            Sqlconn2.Close()
        End If

        lRetCode = oSalesOpportunities.Add()
        If lRetCode <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            AddLog("Oportunidad " & CardCode & " - " & CardName & " no ha sido creado, " & sErrMsg)
            oSalesOpportunities.SaveXML("C:\OportunidadVenta.xml")
        Else
            If conexionSQL2() = 0 Then
                Sql = "UPDATE " & BD_Net & ".dbo.appOPVentaDet SET bitSap = 1 WHERE strNOp = '" & strNOp & "' AND biginTPartida <= " & X
                comando.Connection = Sqlconn2
                comando.CommandText = Sql
                comando.ExecuteNonQuery()
                Sqlconn2.Close()
            End If
        End If

    End Sub


    Public Sub ActualizarOportunidades(ByVal OpprId As Decimal, ByVal CardCode As String, ByVal CardName As String, ByVal Ejecutivo As String, ByVal Asociado As String, ByVal strTitular As String, ByVal intTPla As String, ByVal dteCierreEstimado As Date, ByVal dcmMToP As String, ByVal IntRate As String, ByVal IntId As String, ByVal ChnCrdCode As String, ByVal strObservaciones As String, ByVal bigIntEdo As String, ByVal intPerTPla As String, ByVal strNOp As String, ByVal dteInicio As Date, ByVal bintIDOpVen As String)
        Dim sPaso As String
        Dim oSalesOpportunities As SAPbobsCOM.SalesOpportunities = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesOpportunities)
        'oSalesOpportunities.CardCode = CardCode
        'oSalesOpportunities.OpportunityName = strNOp
        If oSalesOpportunities.GetByKey(OpprId) = True Then
            If Ejecutivo = "" Then
                oSalesOpportunities.UserFields.Fields.Item("U_SlpCode").Value = "-1"
            Else
                oSalesOpportunities.UserFields.Fields.Item("U_SlpCode").Value = Ejecutivo
            End If

            If Asociado = "" Then
                oSalesOpportunities.UserFields.Fields.Item("U_Asociado").Value = "-1"
            Else
                oSalesOpportunities.UserFields.Fields.Item("U_Asociado").Value = Asociado
            End If

            'oSalesOpportunities.DataOwnershipfield = strTitular
            'oSalesOpportunities.ClosingDate = intTPla 'LO CALCULA SAP
            If intPerTPla = 1 Then
                oSalesOpportunities.ClosingType = SAPbobsCOM.BoSoClosedInTypes.sos_Months
            End If
            If intPerTPla = 2 Then
                oSalesOpportunities.ClosingType = SAPbobsCOM.BoSoClosedInTypes.sos_Weeks
            Else
                oSalesOpportunities.ClosingType = SAPbobsCOM.BoSoClosedInTypes.sos_Days
            End If
            oSalesOpportunities.StartDate = dteInicio
            oSalesOpportunities.PredictedClosingDate = dteCierreEstimado
            oSalesOpportunities.Lines.SetCurrentLine(0)
            oSalesOpportunities.Lines.MaxLocalTotal = dcmMToP

            oSalesOpportunities.InterestLevel = IntRate
            'oSalesOpportunities.Interests.InterestId = IntId
            oSalesOpportunities.BPChanelCode = ChnCrdCode
            oSalesOpportunities.Remarks = strObservaciones
            If bigIntEdo = 1 Then
                oSalesOpportunities.Status = SAPbobsCOM.BoSoOsStatus.sos_Open
            Else
                If bigIntEdo = 2 Then
                    oSalesOpportunities.Status = SAPbobsCOM.BoSoOsStatus.sos_Sold
                Else
                    oSalesOpportunities.Status = SAPbobsCOM.BoSoOsStatus.sos_Missed
                End If
            End If

            lRetCode = oSalesOpportunities.Update
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
                AddLog("Oportunidad " & CardCode & " - " & CardName & " no ha sido actualizada, " & sErrMsg)
                oSalesOpportunities.SaveXML("C:\OportunidadVenta.xml")
            Else



                'ETAPA strNOp
                Dim I As Integer = 1
                Dim X As Integer = 0
                If conexionSQL2() = 0 Then
                    Sql = "SELECT dteIniEta, dteFinEt, Step_Id, StrMovimiento, bitSap, biginTPartida AS LineNum "
                    Sql = Sql & " FROM " & BD_Net & ".dbo.appOPVentaDet WHERE strNOp = '" & strNOp & "' AND bitSap = 0 "
                    Sql = Sql & " ORDER BY  biginTPartida"
                    comando2.Connection = Sqlconn2
                    comando2.CommandText = Sql
                    Detalle = comando2.ExecuteReader
                    I = 1
                    Do Until Detalle.Read.ToString <> True
                        oSalesOpportunities = Nothing
                        oSalesOpportunities = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesOpportunities)

                        If oSalesOpportunities.GetByKey(OpprId) Then
                            If oSalesOpportunities.Lines.Count > 0 Then
                                oSalesOpportunities.Lines.Add()
                            End If
                            oSalesOpportunities.Lines.SetCurrentLine(oSalesOpportunities.Lines.Count - 1)
                            oSalesOpportunities.Lines.StartDate = Detalle.Item("dteIniEta").ToString
                            oSalesOpportunities.Lines.ClosingDate = Detalle.Item("dteFinEt").ToString
                            sPaso = Detalle.Item("Step_Id").ToString
                            oSalesOpportunities.Lines.StageKey = sPaso
                            oSalesOpportunities.Lines.Remarks = Detalle.Item("StrMovimiento").ToString
                            oSalesOpportunities.Lines.MaxLocalTotal = dcmMToP
                            'Sincronizar Step_Id es 6 o 7 dejar tabla appOPVenta campo bitNSAP como true
                            I = I + 1
                            X = Detalle.Item("LineNum").ToString

                            lRetCode = oSalesOpportunities.Update
                            If lRetCode <> 0 Then
                                oCompany.GetLastError(lErrCode, sErrMsg)
                                AddLog("Oportunidad " & CardCode & " - " & CardName & " no ha sido actualizada, " & sErrMsg)
                                oSalesOpportunities.SaveXML("C:\OportunidadVenta2.xml")
                            Else
                                If conexionSQL2() = 0 Then
                                    Sql = "UPDATE " & BD_Net & ".dbo.appOPVentaDet SET bitSap = 1 WHERE strNOp = '" & strNOp & "' AND biginTPartida <= " & X
                                    comando.Connection = Sqlconn2
                                    comando.CommandText = Sql
                                    comando.ExecuteNonQuery()

                                    If sPaso = "6" Or sPaso = "7" Then
                                        Sql = "UPDATE " & BD_Net & ".dbo.appOPVenta SET bitNSAP = true WHERE bintIDOpVen = " + bintIDOpVen
                                        comando.CommandText = Sql
                                        comando.ExecuteNonQuery()
                                    End If

                                    Sqlconn2.Close()
                                End If
                            End If
                        End If
                    Loop
                    Detalle.Close()
                    Sqlconn2.Close()
                End If
            End If

        Else
            AddLog("Oportunidad " & CardCode & " - " & CardName & " no ha sido actualizada, no se encontro registro")
        End If
        oSalesOpportunities = Nothing
    End Sub


    Public Sub CrearContrato(ByVal CardCode As String, ByVal CardName As String, ByVal idstrNoPoliza As String, ByVal strNoReciboInt As String, ByVal ItemCode As String, ByVal Apoderado As String, ByVal APODECOMI As Double, ByVal Asociado1 As String, ByVal ASOCOMI As Double, ByVal Empleado As String, ByVal EMPCOMI As Double)
        Dim oServiceContract As SAPbobsCOM.ServiceContracts = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts)
        Dim idContract As Integer

        'Obtener Información necesaria de la Poliza
        Sql = "select [dteIniVigencia]" + vbNewLine
        Sql += ",[dteFinVigencia]" + vbNewLine
        Sql += ",[txtObservaciones]" + vbNewLine
        Sql += ",[dteCancela]" + vbNewLine
        Sql += ",[dcmPR]" + vbNewLine
        Sql += ",[CveMoneda]" + vbNewLine
        Sql += ",[dcmComR]" + vbNewLine
        Sql += ",[idstrNoPoliza]" + vbNewLine
        Sql += ",[strIdProd]" + vbNewLine
        Sql += ",[strNomProd]" + vbNewLine
        Sql += ",[idFormaPago]" + vbNewLine
        Sql += ",[CountRec]" + vbNewLine
        Sql += "FROM [dbo].appEmisionPolizas" + vbNewLine
        Sql += "WHERE idstrNoPoliza = '" & idstrNoPoliza & "'"

        If conexionSQL() = 0 Then
            comando.Connection = Sqlconn
            comando.CommandText = Sql
            LectPoliza = comando.ExecuteReader

            Do Until LectPoliza.Read.ToString <> True
                Try
                    oRs8 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Sql = "SELECT COUNT(*) as Contador, ContractID FROM OCTR  WHERE U_NUMPOLIZA = '" + idstrNoPoliza + "' Group by ContractID"
                    oRs8.DoQuery(Sql)
                    If oRs8.Fields.Item("Contador").Value > 0 Then  ' Si ya existe Contrato
                        AddLog("El Contrato " + idstrNoPoliza + " ya existe, se cargará el registro de pago")
                        idContract = CInt(oRs8.Fields.Item("ContractID").Value)

                        If oServiceContract.GetByKey(idContract) = True Then

                            oServiceContract.Lines.Add()
                            oServiceContract.Lines.ManufacturerSerialNum = idstrNoPoliza
                            oServiceContract.Lines.InternalSerialNum = strNoReciboInt
                            oServiceContract.Lines.ItemCode = LectPoliza.Item("strIdProd").ToString
                            oServiceContract.Lines.ItemName = Left(LectPoliza.Item("strNomProd").ToString, 100)
                            oServiceContract.Lines.StartDate = LectPoliza.Item("dteIniVigencia").ToString
                            oServiceContract.Lines.EndDate = LectPoliza.Item("dteFinVigencia").ToString '16/08/2017 0:00:00" 

                            If LectPoliza.Item("dteCancela").ToString <> "" Then
                                oServiceContract.TerminationDate = LectPoliza.Item("dteCancela").ToString '"16/08/2017 0:00:00"
                            End If

                            lRetCode = oServiceContract.Update()
                            If lRetCode <> 0 Then
                                oCompany.GetLastError(lErrCode, sErrMsg)
                                AddLog("Error en la modificacion de la Poliza: " + idstrNoPoliza + " Mensaje Error: " + sErrMsg)
                            Else
                                AddLog("Poliza: " + idstrNoPoliza + " Modificada Satisfactoriamente Recibo: " + strNoReciboInt)
                            End If
                            'Revisar Cual caso tiene diversos registros y borrar tarjeta y crearlo
                            ' Revisar que ocurre cuando la tarjeta existe pero no hay Contrato 
                        Else
                            AddLog("Contrato " + idContract + "no ha sido actualizado, no se encontro el Registro")
                        End If
                    Else
                        AddLog("Cargando Registro Poliza: " + idstrNoPoliza)
                        ' Encabezado  
                        If Apoderado <> "" Then
                            oServiceContract.UserFields.Fields.Item("U_SlpCode").Value = Apoderado
                            oServiceContract.UserFields.Fields.Item("U_APODECOMI").Value = APODECOMI
                            oServiceContract.UserFields.Fields.Item("U_ASOCIADO").Value = Apoderado
                        End If
                        If Asociado1 <> "" Then
                            oServiceContract.UserFields.Fields.Item("U_ASOCIADO").Value = Asociado1
                            oServiceContract.UserFields.Fields.Item("U_ASOCOMI").Value = ASOCOMI
                        End If
                        If Empleado <> "" Then
                            oServiceContract.UserFields.Fields.Item("U_EMPLEADO").Value = Empleado
                            oServiceContract.UserFields.Fields.Item("U_EMPCOMI").Value = EMPCOMI
                        End If
                        oServiceContract.UserFields.Fields.Item("U_NUMPOLIZA").Value = idstrNoPoliza
                        oServiceContract.UserFields.Fields.Item("U_MONEPOLI").Value = Left(LectPoliza.Item("CveMoneda").ToString, 3)
                        oServiceContract.UserFields.Fields.Item("U_VLRCOMI").Value = CDbl(LectPoliza.Item("dcmComR").ToString)
                        oServiceContract.UserFields.Fields.Item("U_VLRPOLIZ").Value = CDbl(LectPoliza.Item("dcmPR").ToString)
                        oServiceContract.CustomerCode = CardCode
                        oServiceContract.CustomerName = CardName
                        oServiceContract.Description = idstrNoPoliza
                        oServiceContract.Status = 0
                        oServiceContract.StartDate = LectPoliza.Item("dteIniVigencia").ToString
                        oServiceContract.EndDate = LectPoliza.Item("dteFinVigencia").ToString  '"16/08/2017 0:00:00"
                        oServiceContract.UserFields.Fields.Item("U_METOPAGO").Value = LectPoliza.Item("idFormaPago").ToString
                        oServiceContract.UserFields.Fields.Item("U_CANTRECIBO").Value = CDbl(LectPoliza.Item("CountRec").ToString)
                        oServiceContract.UserFields.Fields.Item("U_ITEMCODE").Value = ItemCode

                        If LectPoliza.Item("dteCancela").ToString <> "" Then
                            oServiceContract.TerminationDate = LectPoliza.Item("dteCancela").ToString '"16/08/2017 0:00:00"
                        End If

                        'CTR1 Lineal
                        oServiceContract.Lines.ManufacturerSerialNum = idstrNoPoliza
                        oServiceContract.Lines.InternalSerialNum = strNoReciboInt
                        oServiceContract.Lines.ItemCode = LectPoliza.Item("strIdProd").ToString
                        oServiceContract.Lines.ItemName = Left(LectPoliza.Item("strNomProd").ToString, 100)
                        oServiceContract.Lines.StartDate = LectPoliza.Item("dteIniVigencia").ToString
                        oServiceContract.Lines.EndDate = LectPoliza.Item("dteFinVigencia").ToString '"16/08/2017 0:00:00" 

                        If LectPoliza.Item("dteCancela").ToString <> "" Then
                            oServiceContract.Lines.TerminationDate = LectPoliza.Item("dteCancela").ToString '"16/08/2017 0:00:00" '
                        End If

                        oServiceContract.Add()

                        If lRetCode <> 0 Then
                            oCompany.GetLastError(lErrCode, sErrMsg)
                            AddLog("Error en la creacion de la Poliza: " + idstrNoPoliza + " Mensaje Error: " + sErrMsg)

                            Dim oCustomerEquipmentCard As SAPbobsCOM.CustomerEquipmentCards = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
                            oRs4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Sql = "SELECT InsID FROM OINS WHERE ItemCode = '" & ItemCode & "' AND InternalSN = '" & strNoReciboInt & "'"
                            oRs4.DoQuery(Sql)
                            If oCustomerEquipmentCard.GetByKey(oRs4.Fields.Item("InsID").Value) = True Then
                                lRetCode = oCustomerEquipmentCard.Remove()
                                If lRetCode <> 0 Then
                                    oCompany.GetLastError(lErrCode, sErrMsg)
                                    AddLog("Registro Poliza quedo ingresado en SAP  el contrato Poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
                                Else
                                    AddLog("Registro Poliza ha sido Borrado por no poder Crear el contrato, Fecha , poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
                                End If
                            End If
                            oRs4 = Nothing
                        Else
                            AddLog("Poliza: " + idstrNoPoliza + " Creada Satisfactoriamente")
                        End If
                        End If
                Catch ex As Exception
                    AddLog("Error lectura Vista appEmisionPolizas para la creacion del Contrato : " & ex.Message & " Trace " & ex.StackTrace)
                End Try
            Loop
            Sqlconn.Close()
        End If
    End Sub


    Public Sub CrearTarjeta(ByVal CardCode As String, ByVal CardName As String, ByVal Ejecutivo As String, ByVal Asociado As String, ByVal idstrNoPoliza As String, ByVal strNoReciboInt As String, ByVal ItemCode As String, ByVal ItemName As String, ByVal strIdAseguradora As String, ByVal strNomAse As String, ByVal dteAplicacion As Date, ByVal PrecioNet As Decimal, ByVal intComR As Decimal, ByVal intIva As Decimal, ByVal intCom As Decimal, ByVal dcmCom As Decimal, ByVal strCancela As String, ByVal strNotC As String)
        Dim oCustomerEquipmentCard As SAPbobsCOM.CustomerEquipmentCards = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
        Dim tipoComi As Integer
        Dim EMPCOMI As Double
        Dim APODECOMI As Double
        Dim ASOCOMI As Double
        Dim Asociado1 As String
        Dim Apoderado As String
        Dim Empleado As String

        Asociado1 = ""
        Apoderado = ""
        Empleado = ""
        'NUEVO OBTENER LAS COMISIONES 
        Sql = " Select [bintIdPoliza]" + vbNewLine
        Sql += " ,[intPoCom]" + vbNewLine
        Sql += " ,[strCveSap]" + vbNewLine
        Sql += " ,[strRSocial]" + vbNewLine
        Sql += " ,[bintTpoS]" + vbNewLine
        Sql += " FROM [dbo].[vw_ComisionPoliza]" + vbNewLine
        Sql += "  WHERE idstrNoPoliza = '" + idstrNoPoliza + "'"

        If conexionSQL() = 0 Then
            comando.Connection = Sqlconn
            comando.CommandText = Sql
            LectComision = comando.ExecuteReader

            Do Until LectComision.Read.ToString <> True
                Try
                    tipoComi = CInt(LectComision.Item("bintTpoS").ToString)
                    Select Case tipoComi
                        Case 1
                            Apoderado = LectComision.Item("strCveSap").ToString
                            APODECOMI = CDbl(LectComision.Item("intPoCom").ToString)
                        Case 2
                            Asociado1 = LectComision.Item("strCveSap").ToString
                            ASOCOMI = CDbl(LectComision.Item("intPoCom").ToString)
                        Case 3
                            Empleado = LectComision.Item("strCveSap").ToString
                            EMPCOMI = CDbl(LectComision.Item("intPoCom").ToString)
                    End Select
                Catch ex As Exception
                    AddLog("Error en la obtención de comisiones del Contrato: " & idstrNoPoliza & " Mensaje: " & ex.Message & " Trace " & ex.StackTrace)
                End Try
            Loop

            oCustomerEquipmentCard.CustomerCode = CardCode
            oCustomerEquipmentCard.CustomerName = CardName
            oCustomerEquipmentCard.ManufacturerSerialNum = idstrNoPoliza
            oCustomerEquipmentCard.InternalSerialNum = strNoReciboInt

            If Apoderado <> "" Then
                oCustomerEquipmentCard.UserFields.Fields.Item("U_Ejecutivo").Value = Apoderado
                oCustomerEquipmentCard.UserFields.Fields.Item("U_APODECOMI").Value = APODECOMI
                oCustomerEquipmentCard.UserFields.Fields.Item("U_Asociado").Value = Apoderado
            End If

            If Asociado1 <> "" Then
                oCustomerEquipmentCard.UserFields.Fields.Item("U_Asociado").Value = Asociado1
                oCustomerEquipmentCard.UserFields.Fields.Item("U_ASOCOMI").Value = ASOCOMI
            End If

            If Empleado <> "" Then
                oCustomerEquipmentCard.UserFields.Fields.Item("U_EMPLEADO").Value = Empleado
                oCustomerEquipmentCard.UserFields.Fields.Item("U_EMPCOMI").Value = EMPCOMI
            End If

            oCustomerEquipmentCard.ItemCode = ItemCode
            oCustomerEquipmentCard.ItemDescription = ItemName

            lRetCode = oCustomerEquipmentCard.Add()

            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
                AddLog("Registro Poliza " & CardCode & " - " & CardName & " - " & strNoReciboInt & " no ha sido creado, " & sErrMsg)
            Else
                AddLog("Registro Poliza " & CardCode & " - " & CardName & " - " & strNoReciboInt & "  creada, " & sErrMsg) 
                CrearContrato(CardCode, CardName, idstrNoPoliza, strNoReciboInt, ItemCode, Apoderado, APODECOMI, Asociado1, ASOCOMI, Empleado, EMPCOMI)
                If PrecioNet >= 0 Then
                    If strCancela <> "C" Then
                        CrearEntrega(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, LectSN.Item("strIdAseguradora").ToString, LectSN.Item("strNomAse").ToString, Ejecutivo, Asociado, LectSN.Item("idstrNoPoliza").ToString, LectSN.Item("strNoReciboInt").ToString, LectSN.Item("ItemCode").ToString, LectSN.Item("ItemName").ToString, LectSN.Item("dteAplicacion").ToString, LectSN.Item("PrecioNet").ToString, LectSN.Item("intComR").ToString, LectSN.Item("intIva").ToString, LectSN.Item("StrCvePro").ToString.Trim)
                    End If
                Else
                    If strCancela <> "C" And strNotC = "N" Then  'NC
                        CrearDevolEN(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, LectSN.Item("strIdAseguradora").ToString, LectSN.Item("strNomAse").ToString, Ejecutivo, Asociado, LectSN.Item("idstrNoPoliza").ToString, LectSN.Item("strNoReciboInt").ToString, LectSN.Item("ItemCode").ToString, LectSN.Item("ItemName").ToString, LectSN.Item("dteAplicacion").ToString, LectSN.Item("PrecioNet").ToString, LectSN.Item("intComR").ToString, LectSN.Item("intIva").ToString, 0)
                    Else ' Devolucion
                        CrearDevolEN(LectSN.Item("CardCode").ToString, LectSN.Item("CardName").ToString, LectSN.Item("strIdAseguradora").ToString, LectSN.Item("strNomAse").ToString, Ejecutivo, Asociado, LectSN.Item("idstrNoPoliza").ToString, LectSN.Item("strNoReciboInt").ToString, LectSN.Item("ItemCode").ToString, LectSN.Item("ItemName").ToString, LectSN.Item("dteAplicacion").ToString, LectSN.Item("PrecioNet").ToString, LectSN.Item("intComR").ToString, LectSN.Item("intIva").ToString, 1)
                    End If
                End If
                If Asociado <> "" And AddOC = 1 Then
                    'If dcmCom >= 0 Then
                    '*****************CrearOrdenCompra(Asociado, LectSN.Item("idstrNoPoliza").ToString, LectSN.Item("strNoReciboInt").ToString, LectSN.Item("ItemCode").ToString, LectSN.Item("ItemName").ToString, LectSN.Item("dteAplicacion").ToString, dcmCom, intCom, LectSN.Item("intIva").ToString)
                    'Else
                    '    'CrearDevolOC(Asociado, LectSN.Item("idstrNoPoliza").ToString, LectSN.Item("strNoReciboInt").ToString, LectSN.Item("ItemCode").ToString, LectSN.Item("ItemName").ToString, LectSN.Item("dteAplicacion").ToString, dcmCom, intCom, LectSN.Item("intIva").ToString)
                    'End If
                End If
            End If
        End If

    End Sub


    Public Sub CrearOrdenCompra(ByVal Asociado As String, ByVal idstrNoPoliza As String, ByVal strNoReciboInt As String, ByVal ItemCode As String, ByVal ItemName As String, ByVal dteAplicacion As Date, ByVal PrecioNet As Decimal, ByVal intCom As Decimal, ByVal intIva As Decimal)
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
        oPurchesaOrder.Lines.UserFields.Fields.Item("U_NumPoliza").Value = idstrNoPoliza
        oPurchesaOrder.Lines.UserFields.Fields.Item("U_ReciboInt").Value = strNoReciboInt
        'Si devuelven o cancelan una poliza, PrecioNet viene negativo
        If PrecioNet < 0 Then
            oPurchesaOrder.Lines.UnitPrice = PrecioNet * -1
            oPurchesaOrder.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = -1
        Else
            oPurchesaOrder.Lines.UnitPrice = PrecioNet
            oPurchesaOrder.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = 1
        End If
        oPurchesaOrder.Comments = "Pago de Comisiones a Asociados"

        lRetCode = oPurchesaOrder.Add()

        If lRetCode <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            AddLog("Orden de Compra " & Asociado & " - " & strNoReciboInt & " no ha sido creada, " & sErrMsg)
        End If

    End Sub


    Public Sub CrearEntrega(ByVal CardCode As String, ByVal CardName As String, ByVal strIdAseguradora As String, ByVal strNomAse As String, ByVal Ejecutivo As String, ByVal Asociado As String, ByVal idstrNoPoliza As String, ByVal strNoReciboInt As String, ByVal ItemCode As String, ByVal ItemName As String, ByVal dteAplicacion As Date, ByVal PrecioNet As Decimal, ByVal intComR As Decimal, ByVal intIva As Decimal, ByVal StrCvePro As String)
        Bandera = 2
        Dim oDeliveryNote As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
        oDeliveryNote.CardCode = strIdAseguradora
        'oDeliveryNote.CardNameName = strNomAse
        oDeliveryNote.UserFields.Fields.Item("U_CardCode").Value = CardCode
        oDeliveryNote.UserFields.Fields.Item("U_CardName").Value = Left(CardName, 30)
        oDeliveryNote.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
        If Ejecutivo = "" Then
            oDeliveryNote.UserFields.Fields.Item("U_Ejecutivo").Value = "-1"
        Else
            oDeliveryNote.UserFields.Fields.Item("U_Ejecutivo").Value = Ejecutivo
        End If

        If Asociado = "" Then
            oDeliveryNote.UserFields.Fields.Item("U_Asociado").Value = "-1"
        Else
            oDeliveryNote.UserFields.Fields.Item("U_Asociado").Value = Asociado
        End If

        oDeliveryNote.DocDate = dteAplicacion
        oDeliveryNote.DocDueDate = dteAplicacion
        oDeliveryNote.TaxDate = dteAplicacion
        oDeliveryNote.Comments = "Documento base Registro Póliza/Fianza " & idstrNoPoliza & ", Cliente Indirecto " & CardCode & ". Numero recibo Apianet " & strNoReciboInt
        oDeliveryNote.NumAtCard = idstrNoPoliza

        oDeliveryNote.Lines.ItemCode = ItemCode
        'oDeliveryNote.Lines.ItemDescription = ItemName
        oDeliveryNote.Lines.UnitPrice = PrecioNet
        oDeliveryNote.Lines.WarehouseCode = "01"
        oDeliveryNote.Lines.Quantity = 1
        oDeliveryNote.Lines.CommisionPercent = intComR
        If StrCvePro <> "" Then
            oDeliveryNote.Lines.ProjectCode = StrCvePro
        End If
        oDeliveryNote.Lines.UserFields.Fields.Item("U_NumPoliza").Value = idstrNoPoliza
        oDeliveryNote.Lines.UserFields.Fields.Item("U_ReciboInt").Value = strNoReciboInt
        'Si devuelven o cancelan una poliza, PrecioNet viene negativo
        oDeliveryNote.Lines.UserFields.Fields.Item("U_SignoPoliza").Value = 1
        If Ejecutivo = "" Then
            oDeliveryNote.Lines.SalesPersonCode = "-1"
        Else
            oDeliveryNote.Lines.SalesPersonCode = Ejecutivo
        End If

        'Impuesto
        Sql = "SELECT Code FROM OSTA WHERE LEFT(Code,2) = 'BE' AND Rate = " & intIva
        oRS2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS2.DoQuery(Sql)
        oDeliveryNote.Lines.TaxCode = oRS2.Fields.Item("Code").Value
        oRS2 = Nothing

        lRetCode = oDeliveryNote.Add()

        If lRetCode <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            AddLog("Entrega no ha sido creado Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
            Dim oCustomerEquipmentCard As SAPbobsCOM.CustomerEquipmentCards = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
            oRs4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql = "SELECT InsID FROM OINS WHERE ItemCode = '" & ItemCode & "' AND InternalSN = '" & strNoReciboInt & "'"
            oRs4.DoQuery(Sql)
            If oCustomerEquipmentCard.GetByKey(oRs4.Fields.Item("InsID").Value) = True Then
                lRetCode = oCustomerEquipmentCard.Remove()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    AddLog("Registro Poliza quedo ingresado en SAP sin Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
                Else
                    AddLog("Registro Poliza ha sido Borrado por no poder realizar la Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
                End If
            End If
            oRs4 = Nothing
            AddOC = 0
        Else
            'ACTUALIZAR CAMPO DeliveryNo EN OINS, PERO CAMPO ES SOLO DE LECTURA
            Dim oCustomerEquipmentCard As SAPbobsCOM.CustomerEquipmentCards = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
            oRs4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql = "SELECT InsID FROM OINS WHERE ItemCode = '" & ItemCode & "' AND InternalSN = '" & strNoReciboInt & "'"
            oRs4.DoQuery(Sql)
            If oCustomerEquipmentCard.GetByKey(oRs4.Fields.Item("InsID").Value) = True Then
                oRs5 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Sql = "SELECT TOP 1 T0.DocEntry FROM ODLN T0 JOIN DLN1 T1 ON T1.DocEntry = T0.DocEntry "
                Sql = Sql & " WHERE T1.U_NumPoliza = '" & idstrNoPoliza & "' AND T1.U_ReciboInt = '" & strNoReciboInt & "' ORDER BY T0.DocEntry Desc"
                oRs5.DoQuery(Sql)
                oCustomerEquipmentCard.UserFields.Fields.Item("U_DocNum").Value = oRs5.Fields.Item("DocEntry").Value
                lRetCode = oCustomerEquipmentCard.Update()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    AddLog("Registro Poliza quedo ingresado en SAP sin actualizar el numero Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
                Else
                    AddLog("Registro Poliza se ha actualizado numero de Entrega: " & oRs5.Fields.Item("DocEntry").Value & ", Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt)
                End If
                oRs5 = Nothing
            End If
            oRs4 = Nothing
            If conexionSQL2() = 0 Then
                Sql = "UPDATE " & BD_Net & "..appEmisionPolizaRecibosS SET BitSap = 1, dtesap = GETDATE() WHERE strNoReciboInt = '" & strNoReciboInt & "'"
                comando.Connection = Sqlconn2
                comando.CommandText = Sql
                comando.ExecuteNonQuery()
                Sqlconn2.Close()
            End If
            AddOC = 1
        End If

    End Sub


    Public Sub ActualizarDevolEN(ByVal Code As String)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        'Dim ExistGrilla As Boolean
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql = "SELECT U_CardCode FROM [@DEVOLUCION] WHERE Code= '" & Code & "'"
            oRs.DoQuery(Sql)

            oRs7 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql1 = "SELECT Code FROM [@DEVOLUCION] WHERE U_CardCode= '" & CStr(oRs.Fields.Item("U_CardCode").Value) & "' AND [U_DocNume] <> 1 AND [U_Devnc] = 1"
            oRs7.DoQuery(Sql1)

            While Not oRs7.EoF
                oGeneralService = oCompany.GetCompanyService.GetGeneralService("VM_POLIZADEVOL")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("Code", CStr(oRs7.Fields.Item("Code").Value))
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                'ExistGrilla = RevisarGrillaGenerarDevolucion(CStr(oRs7.Fields.Item("Code").Value))
                'If Not ExistGrilla Then
                oGeneralData.SetProperty("U_DocNume", 1)
                'End If
                oGeneralService.Update(oGeneralData)
                oRs7.MoveNext()
            End While
            oRs7.DoQuery("UPDATE D SET D.U_DocNume=0 from [@DEVOLUCION] D INNER JOIN [@DETALLEDEVOLUCION] V ON  D.Code=V.Code" & _
                          " WHERE [U_Devnc] = 1 AND (SElECt COUNT(*) FrOM INV1 I Where I.U_ReciboInt=V.U_UReciboInt)=0 ")


        Catch ex As Exception

            AddLog("Error " & ex.Message & " Trace " & ex.StackTrace)
            SBO_Application.StatusBar.SetText("Devolución no ha sido actualizado en tabla " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Public Sub ActualizarDevolNC(ByVal Code As String)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        'Dim ExistGrilla As Boolean
        Try
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql = "SELECT U_CardCode FROM [@DEVOLUCION] WHERE Code= '" & Code & "'"
            oRs.DoQuery(Sql)

            oRs7 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql1 = "SELECT Code FROM [@DEVOLUCION] WHERE U_CardCode= '" & CStr(oRs.Fields.Item("U_CardCode").Value) & "' AND [U_DocNume] <> 1 AND [U_Devnc] = 0"
            oRs7.DoQuery(Sql1)

            While Not oRs7.EoF
                oGeneralService = oCompany.GetCompanyService.GetGeneralService("VM_POLIZADEVOL")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("Code", CStr(oRs7.Fields.Item("Code").Value))
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                'ExistGrilla = RevisarGrillaGenerarDevolucion(CStr(oRs7.Fields.Item("Code").Value))
                'If Not ExistGrilla Then
                oGeneralData.SetProperty("U_DocNume", 1)
                'End If
                oGeneralService.Update(oGeneralData)
                oRs7.MoveNext()
            End While
            oRs7.DoQuery("UPDATE D SET D.U_DocNume=0 from [@DEVOLUCION] D INNER JOIN [@DETALLEDEVOLUCION] V ON  D.Code=V.Code" & _
                          " WHERE [U_Devnc] = 0 AND (SElECt COUNT(*) FrOM RIN1 I Where I.U_ReciboInt=V.U_UReciboInt)=0 ")


        Catch ex As Exception

            AddLog("Error " & ex.Message & " Trace " & ex.StackTrace)
            SBO_Application.StatusBar.SetText("Devolución no ha sido actualizado en tabla " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub



    Public Sub CrearDevolEN(ByVal CardCode As String, ByVal CardName As String, ByVal strIdAseguradora As String, ByVal strNomAse As String, ByVal Ejecutivo As String, ByVal Asociado As String, ByVal idstrNoPoliza As String, ByVal strNoReciboInt As String, ByVal ItemCode As String, ByVal ItemName As String, ByVal dteAplicacion As Date, ByVal PrecioNet As Decimal, ByVal intComR As Decimal, ByVal intIva As Decimal, ByVal devnc As Integer)
        Bandera = 2
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCustomerEquipmentCard As SAPbobsCOM.CustomerEquipmentCards
        Dim Cod As String
        Try
            'oCompanyService = oCompany.GetCompanyService
            oGeneralService = oCompany.GetCompanyService.GetGeneralService("VM_POLIZADEVOL")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

            oRs7 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql = "SELECT max(convert(integer,T0.Code))+1 as Code FROM [@DEVOLUCION] T0 "
            oRs7.DoQuery(Sql)
            If CStr(oRs7.Fields.Item("Code").Value) = False Then
                Cod = "1"
            Else
                Cod = CStr(CInt(oRs7.Fields.Item("Code").Value))
            End If
            oGeneralData.SetProperty("Code", Cod)
            oGeneralData.SetProperty("U_CardCode", strIdAseguradora)
            oGeneralData.SetProperty("U_UCardCode", CardCode)
            oGeneralData.SetProperty("U_UCardName", Left(CardName, 30))
            oGeneralData.SetProperty("U_Devnc", devnc)
            'UserTableDev.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
            If Ejecutivo = "" Then
                oGeneralData.SetProperty("U_UEjecutivo", "-1")
            Else
                oGeneralData.SetProperty("U_UEjecutivo", Ejecutivo)
            End If

            If Asociado = "" Then
                oGeneralData.SetProperty("U_UAsociado", "-1")
            Else
                oGeneralData.SetProperty("U_UAsociado", Asociado)
            End If
            oGeneralData.SetProperty("U_DocDate", dteAplicacion)
            oGeneralData.SetProperty("U_DocDueDate", dteAplicacion)
            oGeneralData.SetProperty("U_TaxDate", dteAplicacion)
            oGeneralData.SetProperty("U_DocNume", 0)
            oGeneralData.SetProperty("U_Coments", "Documento base Registro Póliza/Fianza " & idstrNoPoliza & ", Cliente Indirecto " & CardCode & ". Numero recibo Apianet " & strNoReciboInt)

            oChildren = oGeneralData.Child("DETALLEDEVOLUCION")
            oChild = oChildren.Add

            oChild.SetProperty("U_ItemCode", ItemCode)
            oChild.SetProperty("U_Precio", CDbl(PrecioNet * -1))
            oChild.SetProperty("U_WshCode", "01")
            oChild.SetProperty("U_Quantity", 1)
            oChild.SetProperty("U_ComisionPercent", CDbl(intComR))
            oChild.SetProperty("U_UNumPoliza", idstrNoPoliza)
            oChild.SetProperty("U_UReciboInt", strNoReciboInt)
            'Si devuelven o cancelan una poliza, PrecioNet viene negativo
            oChild.SetProperty("U_SignoPoliza", -1)
            If Ejecutivo = "" Then
                oChild.SetProperty("U_SlpCode", -1) '"-1")
            Else
                oChild.SetProperty("U_SlpCode", CInt(Ejecutivo))
            End If

            'Impuesto
            Sql = "SELECT Code FROM OSTA WHERE LEFT(Code,2) = 'BE' AND Rate = " & intIva
            oRS2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS2.DoQuery(Sql)
            oChild.SetProperty("U_TaxCode", oRS2.Fields.Item("Code").Value)
            oRS2 = Nothing

            oGeneralParams = oGeneralService.Add(oGeneralData)

            AddLog("Registro Devolucion Poliza: " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", Fecha " & dteAplicacion)

            'oCustomerEquipmentCard = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
            'oRs4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Sql = "SELECT InsID FROM OINS WHERE ItemCode = '" & ItemCode & "' AND InternalSN = '" & strNoReciboInt & "'"
            'oRs4.DoQuery(Sql)
            'If oCustomerEquipmentCard.GetByKey(oRs4.Fields.Item("InsID").Value) = True Then
            '    oRs5 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    Sql = "SELECT TOP 1 T0.Code FROM [@DEVOLUCION] T0 JOIN [@DETALLEDEVOLUCION] T1 ON T1.Code = T0.Code "
            '    Sql = Sql & " WHERE T1.U_UNumPoliza = '" & idstrNoPoliza & "' AND T1.U_UReciboInt = '" & strNoReciboInt & "' ORDER BY T0.Code Desc"
            '    oRs5.DoQuery(Sql)
            '    'oCustomerEquipmentCard.UserFields.Fields.Item("U_DocNum").Value = oRs5.Fields.Item("Code").Value
            '    'lRetCode = oCustomerEquipmentCard.Update()
            '    If lRetCode <> 0 Then
            '        oCompany.GetLastError(lErrCode, sErrMsg)
            '        AddLog("Registro Poliza quedo ingresado en SAP sin actualizar el numero Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
            '    Else
            '        AddLog("Registro Poliza se ha actualizado numero de Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt)
            '    End If
            '    oRs5 = Nothing
            'End If
            'oRs4 = Nothing

            If conexionSQL2() = 0 Then
                Sql = "UPDATE " & BD_Net & "..appEmisionPolizaRecibosS SET BitSap = 1 WHERE strNoReciboInt = '" & strNoReciboInt & "'"
                comando.Connection = Sqlconn2
                comando.CommandText = Sql
                comando.ExecuteNonQuery()
                Sqlconn2.Close()
            End If
            AddOC = 1

        Catch ex As Exception

            AddLog("Error " & ex.Message & " Trace " & ex.StackTrace)
            SBO_Application.StatusBar.SetText("Devolución no ha sido creado en tabla " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'oCompany.GetLastError(lErrCode, sErrMsg)
            AddLog("Entrega no ha sido creado Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & ex.Message)
            oCustomerEquipmentCard = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
            oRs4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Sql = "SELECT InsID FROM OINS WHERE ItemCode = '" & ItemCode & "' AND InternalSN = '" & strNoReciboInt & "'"
            oRs4.DoQuery(Sql)
            If oCustomerEquipmentCard.GetByKey(oRs4.Fields.Item("InsID").Value) = True Then
                lRetCode = oCustomerEquipmentCard.Remove()
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    AddLog("Registro Poliza quedo ingresado en SAP sin Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
                Else
                    AddLog("Registro Poliza ha sido Borrado por no poder realizar la Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
                End If
            End If
            oRs4 = Nothing
            AddOC = 0
        End Try

        'If lRetCode <> 0 Then
        'oCompany.GetLastError(lErrCode, sErrMsg)
        'AddLog("Entrega no ha sido creado Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
        'Dim oCustomerEquipmentCard As SAPbobsCOM.CustomerEquipmentCards = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
        'oRs4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Sql = "SELECT InsID FROM OINS WHERE ItemCode = '" & ItemCode & "' AND InternalSN = '" & strNoReciboInt & "'"
        'oRs4.DoQuery(Sql)
        'If oCustomerEquipmentCard.GetByKey(oRs4.Fields.Item("InsID").Value) = True Then
        '    lRetCode = oCustomerEquipmentCard.Remove()
        '    If lRetCode <> 0 Then
        '        oCompany.GetLastError(lErrCode, sErrMsg)
        '        AddLog("Registro Poliza quedo ingresado en SAP sin Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
        '    Else
        '        AddLog("Registro Poliza ha sido Borrado por no poder realizar la Entrega, Fecha " & dteAplicacion & ", poliza " & idstrNoPoliza & ", Nro registro Apianet " & strNoReciboInt & ", " & sErrMsg)
        '    End If
        'End If
        'oRs4 = Nothing
        'AddOC = 0
        'Else
        'ACTUALIZAR CAMPO DeliveryNo EN OINS, PERO CAMPO ES SOLO DE LECTURA

        'End If

    End Sub


    Sub ConfigConexion()
        Dim oItem As SAPbouiCOM.Item
        Dim oButton As SAPbouiCOM.Button
        Dim oStaticText As SAPbouiCOM.StaticText
        Dim oEditText As SAPbouiCOM.EditText
        'Dim oComboBox As SAPbouiCOM.ComboBox
        Dim oDBDSDetalle As SAPbouiCOM.DBDataSource
        '// add a new form
        Dim oCreationParams As SAPbouiCOM.FormCreationParams

        oCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)

        oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed
        oCreationParams.UniqueID = "ConecApianet"
        oCreationParams.FormType = "CA"
        oCreationParams.ObjectType = "CONEXAP"

        oForm = SBO_Application.Forms.AddEx(oCreationParams)

        '// set the form properties
        oForm.Title = "Conexion Apianet"
        oForm.Left = 400
        oForm.Top = 100
        oForm.ClientHeight = 110
        oForm.ClientWidth = 240

        '// Adding an Ok button
        oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem.Left = 6
        oItem.Width = 65
        oItem.Top = 90
        oItem.Height = 19

        oButton = oItem.Specific

        oButton.Caption = "Ok"

        oForm.DefButton = "1"



        '// Adding a Cancel button
        oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem.Left = 75
        oItem.Width = 65
        oItem.Top = 90
        oItem.Height = 19

        oButton = oItem.Specific

        oButton.Caption = "Cancel"

        '//************************
        '// Adding a Rectangle
        '//***********************

        oItem = oForm.Items.Add("Rect1", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE)
        oItem.Left = 0
        oItem.Width = 235
        oItem.Top = 1
        oItem.Height = 85

        '// Adding a Text Server
        oItem = oForm.Items.Add("ServerTxt", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 117
        oItem.Width = 110
        oItem.Top = 8
        oItem.Height = 14

        oEditText = oItem.Specific

        '// bind the text edit item to the defined used data source
        oEditText.DataBind.SetBound(True, "@CONEXAP", "U_Server")


        '// Adding a Static Server
        oItem = oForm.Items.Add("SeverLbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 7
        oItem.Width = 108
        oItem.Top = 8
        oItem.Height = 14

        oItem.LinkTo = "ServerText"
        oStaticText = oItem.Specific
        oStaticText.Caption = "Servidor"

        '// Adding a Text User
        oItem = oForm.Items.Add("UserTxt", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 117
        oItem.Width = 110
        oItem.Top = 25
        oItem.Height = 14

        oEditText = oItem.Specific

        '// bind the text edit item to the defined used data source
        oEditText.DataBind.SetBound(True, "@CONEXAP", "U_UserSQL")

        '// Adding a Static User
        oItem = oForm.Items.Add("UserLbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 7
        oItem.Width = 108
        oItem.Top = 25
        oItem.Height = 14

        oItem.LinkTo = "UserText"

        oStaticText = oItem.Specific

        oStaticText.Caption = "Usuario"

        '// Adding a Text Password
        oItem = oForm.Items.Add("PassTxt", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 117
        oItem.Width = 110
        oItem.Top = 42
        oItem.Height = 14
        oItem.Visible = True
        oEditText = oItem.Specific

        '// bind the text edit item to the defined used data source
        oEditText.DataBind.SetBound(True, "@CONEXAP", "U_PassSQL")

        ''// Adding a Text 2 Password para usar **
        'oItem = oForm.Items.Add("PassTxt1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        'oItem.Left = 117
        'oItem.Width = 110
        'oItem.Top = 42
        'oItem.Height = 14
        ''oItem.Visible = False
        'oEditText = oItem.Specific

        ''// bind the text edit item to the defined used data source
        'oEditText.DataBind.SetBound(True, "@CONEXAP", "U_Blanco")



        '// Adding a Static Password
        oItem = oForm.Items.Add("PassLbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 7
        oItem.Width = 108
        oItem.Top = 42
        oItem.Height = 14

        oItem.LinkTo = "PassText"

        oStaticText = oItem.Specific
        oStaticText.Caption = "Password"

        '// Adding a Text Base de Datos
        oItem = oForm.Items.Add("BDTxt", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 117
        oItem.Width = 110
        oItem.Top = 59
        oItem.Height = 14

        oEditText = oItem.Specific

        '// bind the text edit item to the defined used data source
        oEditText.DataBind.SetBound(True, "@CONEXAP", "U_BD")

        '// Adding a Static BD
        oItem = oForm.Items.Add("BDLbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 7
        oItem.Width = 108
        oItem.Top = 59
        oItem.Height = 14

        oItem.LinkTo = "BDText"

        oStaticText = oItem.Specific
        oStaticText.Caption = "Base Datos"

        ''preguntar si existen registros en la tabla de parametros
        ''if existen registros colorcar el formualrio en modo update
        '' si no existen colocar en modo add
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Sql = "SELECT COUNT(*) AS Cont FROM [@CONEXAP]"
        oRs.DoQuery(Sql)
        oForm.AutoManaged = True
        oForm.Visible = True

        If oRs.Fields.Item("Cont").Value > 0 Then
            oForm.SupportedModes = 1
            oForm.Mode = BoFormMode.fm_OK_MODE
        Else
            oForm.SupportedModes = 3
            oForm.Mode = BoFormMode.fm_ADD_MODE
        End If

        If oForm.SupportedModes = 1 Then
            oDBDSDetalle = oForm.DataSources.DBDataSources.Item("@CONEXAP")
            oDBDSDetalle.Query(Nothing)
            oPass = GetPass(oForm)
            HidePass(oForm)
        End If

    End Sub


    Public Sub LoadFromXML(ByVal FileName As String)
        Dim oXMLDoc1 As XmlDocument
        Dim RutaC As String = System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim Exe As String = Dir(RutaC)
        Dim sPath As String
        sPath = Microsoft.VisualBasic.Left(RutaC, Len(RutaC) - Len(Exe))
        Try
            oXMLDoc1 = New XmlDocument
            oXMLDoc1.Load(sPath & FileName)
            SBO_Application.LoadBatchActions(oXMLDoc1.InnerXml)
        Catch ex As Exception
            SBO_Application.MessageBox("Error " & ex.Message & " Trace " & ex.StackTrace)
        End Try

    End Sub


    Public Sub SaveAsXML()
        '//**********************************************************************
        '//
        '// always use XML to work with user forms.
        '// after creating your form save it as an XML file
        '//
        '//**********************************************************************

        Dim oXmlDoc As Xml.XmlDocument

        oXmlDoc = New Xml.XmlDocument

        Dim sXmlString As String

        '// get the form as an XML string
        sXmlString = oForm.GetAsXML

        '// load the form's XML string to the
        '// XML document object
        oXmlDoc.LoadXml(sXmlString)

        '// save the XML Document
        Dim sPath As String

        sPath = IO.Directory.GetParent(SBO_Application.StartupPath).ToString

        oXmlDoc.Save((sPath & "\MySimpleForm.xml"))

    End Sub

End Class