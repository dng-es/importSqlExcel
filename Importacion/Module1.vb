Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

Module Module1
    Dim conexion As New Conexion
    Dim Sql As String

    Sub Main()
        Try
            ImportTable()
            ImportExcel()

            Console.Write("IMPORT FINISHED: ALL OK")
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Sub

    Public Sub ImportTable()
        Dim ds As DataSet = conexion.GetDataOrigen("SELECT * FROM ocio_contenidos WHERE activo=1", "cabecera")
        Dim user_name As String = ""
        Dim user_date As Date
        For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
            user_name = ds.Tables(0).Rows(i).Item(9).ToString
            user_date = ds.Tables(0).Rows(i).Item(7).ToString

            Sql = "INSERT INTO table_name (user,date_user) VALUES('" & user_name & "','" & conexion.InsertDate(user_date) & "');"
            conexion.RunSql(Sql)
            Console.Write("import: " & user_name & vbCrLf)
        Next
        Console.Write("---------SQL IMPORT FINISHED-----------" & vbCrLf)
    End Sub

    Private Sub ImportExcel()
        Try
            Dim ds As New DataSet
            ds = conexion.ReadExcel("C:/excel_file.xls")
            Dim user_name As String = ""
            Dim user_date As Date = Now
            For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                'USER_NAME
                Try
                    user_name = Double.Parse(ds.Tables(0).Rows(i).Item(3).ToString)
                Catch ex As Exception
                    Console.Write(ex.Message)
                End Try

                'USER_DATE
                Try
                    user_date = ds.Tables(0).Rows(i).Item(4).ToString
                Catch ex As Exception
                    Console.Write(ex.Message)
                End Try

                Sql = "INSERT INTO table_name (user,date_user) VALUES('" & user_name & "','" & conexion.InsertDate(user_date) & "');"
                conexion.RunSql(Sql)
                Console.Write(vbTab & "import: " & user_name & vbCrLf)
            Next
            Console.Write("---------EXCEL IMPORT FINISHED-----------" & vbCrLf)
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Sub
End Module

