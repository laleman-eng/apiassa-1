Imports System.Data.SqlClient
Imports System.Collections.Generic


Module MainSub
    'VARIABLES
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public oCompany As SAPbobsCOM.Company
    Public oCodes As New Dictionary(Of String, String)
    Public Card As String
    Public oForm As SAPbouiCOM.Form
    Public oFormB As SAPbouiCOM.Form
    Public oItem As SAPbouiCOM.Item
    Public oItemB As SAPbouiCOM.Item
    Public ObjItem As SAPbouiCOM.Item
    Public oStaticText As SAPbouiCOM.StaticText
    Public s As String
    Public oEditText As SAPbouiCOM.EditText
    Public oEditText1 As SAPbouiCOM.EditText
    Public oComboBox As SAPbouiCOM.ComboBox
    Public oFolder As SAPbouiCOM.Folder
    Public oMenu As SAPbouiCOM.Menus
    Public oUFields As SAPbobsCOM.UserFieldsMD
    Public oRs As SAPbobsCOM.Recordset
    Public oRS2 As SAPbobsCOM.Recordset
    Public oRs3 As SAPbobsCOM.Recordset
    Public oRs4 As SAPbobsCOM.Recordset
    Public oRs5 As SAPbobsCOM.Recordset
    Public oRs6 As SAPbobsCOM.Recordset
    Public oRs7 As SAPbobsCOM.Recordset
    Public oRs8 As SAPbobsCOM.Recordset
    Public lRetCode As Integer
    Public Sql As String
    Public Sql1 As String
    Public sErrMsg As String
    Public lErrCode As Integer
    Public sPath As String
    Public oPass As String
    Public barra As SAPbouiCOM.ProgressBar
    Public Servidor As String
    Public BD_Net As String
    Public User As String
    Public Pass As String
    Public Sqlconn As New SqlConnection
    Public Sqlconn2 As New SqlConnection
    Public comando As New SqlCommand
    Public comando2 As New SqlCommand
    Public LectSN As SqlDataReader
    Public Detalle As SqlDataReader
    Public AsocEje As SqlDataReader
    Public LectPoliza As SqlDataReader
    Public LectComision As SqlDataReader
    Public texto As String
    Public Bandera As Integer
    Public intCom1 As Integer
    Public AddOC As Integer


    Public Sub Main()
        Dim oParamconexion As Paramconexion

        oParamconexion = New Paramconexion

        ' Starting the Application
        System.Windows.Forms.Application.Run()
    End Sub
End Module
