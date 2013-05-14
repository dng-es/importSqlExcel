Imports Microsoft.Office.Interop
Imports System.Data.OleDb

Public Class Conexion

    'BBDD target CONNECTION
    Dim BBDD_SERVER As String = "192.168.0.8"
    Dim BBDD_NAME As String = "bbdd_target"
    Dim BBDD_USER As String = "root"
    Dim BBDD_PASS As String = "*******"
    Dim BBDD_PORT As Integer = 3306
    Dim ConexString As String = "Persist Security Info=False;" & _
                    "database=" & BBDD_NAME & _
                    ";server= " & BBDD_SERVER & _
                    ";user id=" & BBDD_USER & _
                    ";Password=" & BBDD_PASS & _
                    ";port=" & BBDD_PORT

    'BBDD origen CONNECTION
    Dim BBDD_SERVER_ORIGEN As String = "192.168.0.8"
    Dim BBDD_NAME_ORIGEN As String = "bbdd_origen"
    Dim BBDD_USER_ORIGEN As String = "root"
    Dim BBDD_PASS_ORIGEN As String = "******"
    Dim BBDD_PORT_ORIGEN As Integer = 3306
    Dim OrigenConexString As String = "Persist Security Info=False;" & _
                "database=" & BBDD_NAME_ORIGEN & _
                ";server= " & BBDD_SERVER_ORIGEN & _
                ";user id=" & BBDD_USER_ORIGEN & _
                ";Password=" & BBDD_PASS & _
                ";port=" & BBDD_PORT_ORIGEN

    Dim ds As DataSet
    Dim MySQL_cmd As MySqlClient.MySqlCommand
    Dim MySQL_cn As New MySqlClient.MySqlConnection(ConexString)
    Public MySQL_da As MySqlClient.MySqlDataAdapter
    Dim MySQL_cn_origen As New MySqlClient.MySqlConnection(OrigenConexString)

    Public Function RunSql(ByVal SQL_query As String) As Boolean
        Try
            If MySQL_cn.State <> 1 Then MySQL_cn.Open()

            MySQL_cmd = New MySqlClient.MySqlCommand
            MySQL_cmd.Connection = MySQL_cn
            MySQL_cmd.CommandText = SQL_query

            Dim SQL_result As Boolean = Boolean.Parse(MySQL_cmd.ExecuteNonQuery > 0)
            MySQL_cn.Close()
            Return SQL_result

        Catch ex As Exception
            If Not ex.Message.Contains("Duplicate entry") Then MsgBox(ex.Message, MsgBoxStyle.Critical, ex.Source)
            MySQL_cn.Close()
            Return False
        End Try
    End Function

    Public Function GetData(ByVal SQL_query As String, ByVal TableReturn As String) As DataSet
        Try
            If MySQL_cn.State <> 1 Then MySQL_cn.Open()

            'Dim MySQL_cmd2 As MySqlClient.MySqlCommand
            'MySQL_cmd2.CommandTimeout = "30000"

            MySQL_da = New MySqlClient.MySqlDataAdapter(SQL_query, MySQL_cn)
            MySQL_da.SelectCommand.CommandTimeout = 10000

            ds = New DataSet
            MySQL_da.Fill(ds, TableReturn)
            MySQL_cn.Close()
            Return ds

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ex.Source)
            MySQL_cn.Close()
            Return ds
        End Try
    End Function

    Public Function GetDataOrigen(ByVal SQL_query As String, ByVal TableReturn As String) As DataSet
        Try
            If MySQL_cn_origen.State <> 1 Then MySQL_cn_origen.Open()

            MySQL_da = New MySqlClient.MySqlDataAdapter(SQL_query, MySQL_cn_origen)
            MySQL_da.SelectCommand.CommandTimeout = 10000

            ds = New DataSet
            MySQL_da.Fill(ds, TableReturn)
            MySQL_cn_origen.Close()
            Return ds
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ex.Source)
            MySQL_cn_origen.Close()
            Return ds
        End Try
    End Function

    Public Function GetMaxValue(ByVal FieldName As String, ByVal TableName As String, ByVal SQLCondition As String) As Integer
        Try
            Dim ds As New DataSet
            ds = GetData("SELECT ifnull(max(" & FieldName & "),0) as max_value from " & TableName & SQLCondition, "max_value")
            Return Integer.Parse(ds.Tables("max_value").Rows(0).Item(0))
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Function GetLastInsertId() As Integer
        Dim ds As New DataSet
        ds = GetData("SELECT LAST_INSERT_ID()", "LastId")
        Return Integer.Parse(ds.Tables("LastId").Rows(0).Item(0))
    End Function

    Public Function InsertDate(ByVal DateValue As Date) As String
        Return Format(DateValue, "yyyy-MM-dd")
    End Function

    Public Function InsertDateTime(ByVal DateValue As Date) As String
        Return Format(DateValue, "yyyy-MM-dd HH:mm:ss")
    End Function

    Public Function InsertarDouble(ByVal DoubleValue As Double) As String
        Return Replace(DoubleValue.ToString, ",", ".")
    End Function

    Protected Friend Function ObtenerCadena(ByVal TableName As String, ByVal FieldName As String, _
                                  Optional ByVal SqlCondition As String = "") As String
        Dim consulta As String = ""
        Dim Cadena As String = ""

        Try
            If MySQL_cn.State <> 1 Then MySQL_cn.Open()
            If SqlCondition <> "" Then SqlCondition = " WHERE " & SqlCondition

            consulta = "SELECT " & FieldName & _
                       " FROM " & TableName & SqlCondition

            MySQL_da = New MySqlClient.MySqlDataAdapter(consulta, MySQL_cn)

            ds = New DataSet
            MySQL_da.Fill(ds, "Cadena")

            Cadena = ds.Tables("Cadena").Rows(0).Item(0).ToString
            MySQL_cn.Close()
            Return Cadena

        Catch ex As Exception
            MySQL_cn.Close()
            Return ""
        End Try
    End Function

    Protected Friend Function CheckData(ByVal SQL_query As String) As Boolean
        Try
            If MySQL_cn.State <> 1 Then MySQL_cn.Open()

            MySQL_cmd = New MySqlClient.MySqlCommand
            MySQL_cmd.Connection = MySQL_cn
            MySQL_cmd.CommandText = SQL_query

            Dim Solucion As Boolean = MySQL_cmd.ExecuteReader.HasRows

            MySQL_cn.Close()
            Return Solucion
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ex.Source)
            MySQL_cn.Close()
            Return False
        End Try
    End Function

    Public Function ReadExcel(ByVal path As String) As DataSet
        Dim ExcelSheet As String = ""
        Dim err As Boolean = False
        Dim ConexString As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                    "Data Source=" + path & _
                                    ";Extended Properties=Excel 8.0"
        Dim con As New System.Data.OleDb.OleDbConnection(ConexString)
        Dim adaptador As New OleDbDataAdapter
        Dim ds As New DataSet

        Dim Exc As Excel.Application = New Excel.Application
        Exc.Workbooks.OpenText(path, , , , _
          Excel.XlTextQualifier.xlTextQualifierNone, , True)

        Dim Wb As Excel.Workbook = Exc.ActiveWorkbook
        Dim Ws As Excel.Worksheet = Wb.ActiveSheet
        ExcelSheet = Ws.Name

        Try
            con.Open()
            adaptador = New OleDbDataAdapter("SELECT * FROM [" & ExcelSheet & "$]", con)

            adaptador.Fill(ds, "XLData")
            con.Close()
        Catch ex As Exception
            Console.Write(ex.Message.ToString)
            con.Close()
            err = True
            ds = Nothing
        End Try

        If err Then
            Try
                Ws.Rows.Delete(Ws.Rows(0))

                con.Open()
                adaptador = New OleDbDataAdapter("SELECT * FROM [" & ExcelSheet & "$]", con)

                adaptador.Fill(ds, "XLData")
                con.Close()
            Catch ex As Exception
                Console.Write(ex.Message.ToString)
                con.Close()
                ds = Nothing
                Exc.Quit()
            End Try
        End If

        Exc.Quit()
        Return ds
    End Function
End Class
