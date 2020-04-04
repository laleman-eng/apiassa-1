Imports System.Windows.Forms
Imports SAPbouiCOM
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO

Module Funciones
    Public Function conexionSQL() As Integer

        Sqlconn = New SqlConnection
        If Sqlconn.State = 1 Then
            Exit Function
        Else
            Sqlconn.ConnectionString = "server=" & Servidor & ";user id=" & User & "; password=" & Pass & "; database=" & BD_Net & "; pooling=false" 'TU CONEXION
            Try
                Sqlconn.Open() 'ABRE TU CONEXION
                conexionSQL = 0
            Catch excepcion As SqlException ' DECLARAS EXCEPCION 
                conexionSQL = 1
                MsgBox("Error de conexión base Origen", MsgBoxStyle.Exclamation, "Datos de Acceso")
                'Sqlconn.Close() 'CIERRA TU CONEXION
            End Try
        End If
    End Function
    Public Function GetPass(ByVal oForm As SAPbouiCOM.Form) As String
        GetPass = oForm.DataSources.DBDataSources.Item("@CONEXAP").GetValue("U_PassSQL", 0)

    End Function

    Public Sub HidePass(ByVal oForm As SAPbouiCOM.Form)
        oForm.DataSources.DBDataSources.Item("@CONEXAP").SetValue("U_PassSQL", 0, "*****")

    End Sub

    Public Sub SetPass(ByVal oForm As SAPbouiCOM.Form)
        oForm.DataSources.DBDataSources.Item("@CONEXAP").SetValue("U_PassSQL", 0, oPass)

    End Sub


    Public Function conexionSQL2() As Integer

        Sqlconn2 = New SqlConnection
        If Sqlconn2.State = 1 Then
            Exit Function
        Else
            Sqlconn2.ConnectionString = "server=" & Servidor & ";user id=" & User & "; password=" & Pass & "; database=" & BD_Net & "; pooling=false" 'TU CONEXION
            Try
                Sqlconn2.Open() 'ABRE TU CONEXION
                conexionSQL2 = 0
            Catch excepcion As SqlException ' DECLARAS EXCEPCION 
                conexionSQL2 = 1
                MsgBox("Error de conexión base Origen", MsgBoxStyle.Exclamation, "Datos de Acceso")
                'Sqlconn.Close() 'CIERRA TU CONEXION
            End Try
        End If
    End Function

    Sub AddLog(ByVal Mensaje As String)
        Dim Arch As StreamWriter
        Arch = New StreamWriter(sPath, True)
        Try
            Arch.WriteLine(String.Format("{0:dd-MM-yyyy hh:mm:ss}", DateTime.Now) & " " & Mensaje)
        Finally
            Arch.Flush()
            Arch.Close()
        End Try
    End Sub


End Module
