Imports System
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.WindowsAPICodePack.Dialogs
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Globalization
Imports System.Globalization.CultureInfo

Public Class Form1
    Public cn As New SqlConnection
    Private counter As Double = 1
    Private Function adoconecta()
        Try
            Dim strcn As String
            strcn = "Server=LAPTOP-JOSUE\MSSQLSERVER14;database=Migracion;user id=migracionUser;password=ODS2020*;multipleactiveresultsets=true;"

            cn = New SqlConnection(strcn)
            cn.Open()

            Return cn
        Catch ex As Exception

        End Try

    End Function
    Private Function adoconectaExcel()
        '   Try
        Dim cn2 As System.Data.OleDb.OleDbConnection
        Dim strcn As String
        strcn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & (Label1.Text) & ";" & "Extended Properties=" & Convert.ToChar(34).ToString() & "Excel 12.0" + Convert.ToChar(34).ToString() & ";"
        cn2 = New System.Data.OleDb.OleDbConnection(strcn)
        cn2.Open()
        Return cn2
        '  Catch ex As Exception

        '    End Try

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Dim years(5) As String
        Dim years(4) As String
        '        years(0) = "2015"
        years(0) = "2016"
        years(1) = "2017"
        years(2) = "2018"
        years(3) = "2019"
        'years(5) = "2020"

        'years(0) = "2010"
        'years(1) = "2011"
        'years(2) = "2013"
        'years(3) = "2015"
        'years(4) = "2016"
        'years(5) = "2017"
        'years(6) = "2018"
        'years(7) = "2019"

        Label2.Text = "Trabajando en ello..."
        'Create a new excel file

        Dim xlApp As New Excel.Application
        Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
        Dim excelFileName As String = "c:\archivos\Book-" & System.DateTime.Now.ToString.Replace(".", "").Replace(" ", "").Replace("/", "").Replace("-", "").Replace(":", "") & ".xlsx"

        excelBook.SaveAs(excelFileName)
        excelBook.Close()
        Dim cuenta As Integer = 1

        ProgressBar1.Maximum = years.Length
        Dim xls As New Microsoft.Office.Interop.Excel.Application
        Dim xlsWorkBook As Microsoft.Office.Interop.Excel.Workbook
        xlsWorkBook = xls.Workbooks.Open(excelFileName)
        For Each year As String In years
            ProgressBar1.Value = cuenta
            If Not IsNothing(year) Then
                ProcessData(year, xls, xlsWorkBook)
            End If

            ' Application.DoEvents()

            cuenta = cuenta + 1
        Next
        xlsWorkBook.Close()
        xls.Quit()
        Label2.Text = "Tarea finalizada"
        Button1.Enabled = False
        counter = 1
        MessageBox.Show("Datos migrados con exito en el archivo " & excelFileName)

    End Sub

    Private Sub ProcessData(year, xls, xlsWorkBook)
        Dim ssql As String
        Dim conexion = adoconectaExcel()

        'ssql = "SELECT * FROM [GH_CRND_" & year & "$]"
        ssql = "SELECT * FROM [GH_CHE_" & year & "$]"

        Dim cmd As New OleDbCommand(ssql, conexion)
        Dim lect As OleDbDataReader = cmd.ExecuteReader


        'Modifies excel file

        Dim xlsWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        While lect.Read

            xlsWorkSheet = xlsWorkBook.Sheets("hoja1")

            If counter = 1 Then
                xlsWorkSheet.Cells(1, 1) = "CodigoPlanta"
                xlsWorkSheet.Cells(1, 2) = "FechaHora"
                xlsWorkSheet.Cells(1, 3) = "Potencia"
                counter = counter + 1
            End If

            Dim columCounter = 1

            While columCounter <= 24

                Dim hora As Integer = columCounter - 1
                        If columCounter = 1 Then
                            hora = 0
                        End If
                Dim fecha1 As Date
                ' Si es el primer dia se asigna manual
                If Not IsDBNull(lect(0)) Then
                    fecha1 = CDate(lect(0))
                    Dim fecha As DateTime = New DateTime(fecha1.Year, fecha1.Month, fecha1.Day, hora, 0, 0, 0)
                    xlsWorkSheet.Cells(counter, 1) = txtCodPlanta.Text
                    xlsWorkSheet.Cells(counter, 2) = fecha
                    xlsWorkSheet.Cells(counter, 3) = lect(columCounter)

                    Label2.Text = counter & " filas escritas"
                    counter = counter + 1
                Else
                    fecha1 = CDate("01/01/" & year)
                End If

                columCounter = columCounter + 1

            End While

            '            counter = counter + 1


        End While

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Dim openDialog As OpenFileDialog = New OpenFileDialog()
        openDialog.Title = "Select A File"
        openDialog.Filter = "All Files (*.*)|*.*"
        If openDialog.ShowDialog() = DialogResult.OK Then

            Dim file As String = openDialog.FileName
            Label1.Text = file

        End If
        Button1.Enabled = True

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End
    End Sub
End Class
