Imports System.Data.OleDb
Imports System.IO
Imports ADODB
Imports System.Windows.Forms.Application
Imports ADOX

Public Class CSVCompare
    Dim tablePOS As String = "PosData"
    Dim tableDC As String = "CollectorData"
    Dim ExcelPath As String
    Dim connTable As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StartupPath & "\Database\inOut.mdb")
    Dim conn As New Connection


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            connTable.Open()
        Catch ex As Exception
            ' Part 1: Create Access Database file using ADOX
            If (Not Directory.Exists(StartupPath & "/Database")) Then
                Directory.CreateDirectory(StartupPath & "/Database")
            End If
            Try
                Call New Catalog().Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StartupPath & "\Database\inOut.mdb;Jet OLEDB:Engine Type=5")
            Catch exe As Exception

            End Try

            ' Part 2: Create one Table using OLEDB Provider 
            'Get database schema
            connTable.Open()
            Dim dbSchema As DataTable = connTable.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, (New Object() {Nothing, Nothing, tablePOS, "TABLE"}))
            connTable.Close()

            ' If the table exists, the count = 1
            If dbSchema.Rows.Count > 0 Then
                ' do whatever you want to do if the table exists
            Else
                'do whatever you want to do if the table does not exist
                ' e.g. create a table
                connTable.Open()
                Call New OleDbCommand("CREATE TABLE [" + tablePOS + "] ([Id] COUNTER, [DataCode] TEXT(50), [DataDescription] TEXT(100))", connTable).ExecuteNonQuery()
                Call New OleDbCommand("CREATE TABLE [" + tableDC + "] ([Id] COUNTER, [DataCode] TEXT(50))", connTable).ExecuteNonQuery()
            End If
        End Try
        connTable.Close()
    End Sub

    Private Sub ImportFromPOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportFromPOSToolStripMenuItem.Click
        connTable.Open()

        Dim OpenFileDialog1 As New OpenFileDialog With {
            .Filter = "Excel Files (*)|*.csv;"
        }

        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            ExcelPath = OpenFileDialog1.FileName
            Try
                ClearTablePos()
                Call New OleDbCommand($"INSERT INTO [PosData] SELECT F1 AS DataCode, F2 AS DataDescription FROM [Text;FMT=Delimited;HDR=No;CharacterSet=850;DATABASE={Path.GetDirectoryName(ExcelPath)}].[{ExcelPath.Split("\")(ExcelPath.Split("\").Length - 1)}]", connTable).ExecuteNonQuery()
                MessageBox.Show("Import Successful.")
            Catch ex As Exception
                MessageBox.Show("Import Failure.")
                'MessageBox.Show(ex.ToString)
            End Try
        End If
        connTable.Close()
        Call DataGridShow()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call DataGridShow()
    End Sub

    Private Sub ExportResultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportResultToolStripMenuItem.Click
        subExportDGVToCSV($"CompareResult_{Date.Now.ToString("yyyy-MM-dd HH-mm-ss")}.csv", DataGridView3)
        Process.Start(StartupPath)
    End Sub

    Private Sub ImportFromDCToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportFromDCToolStripMenuItem.Click
        connTable.Open()
        Dim OpenFileDialog1 As New OpenFileDialog With {
            .Filter = "Excel Files (*)|*.csv;"
        }

        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            ExcelPath = OpenFileDialog1.FileName
            Try
                ClearTableCollector()
                Call New OleDbCommand("INSERT INTO [CollectorData] SELECT F1 AS DataCode FROM [Text;FMT=Delimited;HDR=No;CharacterSet=850;DATABASE=" & Path.GetDirectoryName(ExcelPath) & "].[" & ExcelPath.Split("\")(ExcelPath.Split("\").Length - 1) & "]", connTable).ExecuteNonQuery()
                MessageBox.Show("Import Successful.")
            Catch ex As Exception
                MessageBox.Show("Import Failure.")
                'MessageBox.Show(ex.ToString)
            End Try
        End If
        connTable.Close()
        Call DataGridShow()
    End Sub
    Sub ClearTablePos()
        Try
            Call New OleDbCommand("DELETE * FROM PosData", connTable).ExecuteNonQuery()
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Sub ClearTableCollector()
        Try
            Call New OleDbCommand("DELETE * FROM CollectorData", connTable).ExecuteNonQuery()
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub DataGridShow()
        connTable.Open()
        Dim dt1 As New DataTable
        'Dim dt2 As New DataTable
        Dim dt3 As New DataTable

        Call New DataSet().Tables.Add(dt1)
        'Call New DataSet().Tables.Add(dt2)
        Call New DataSet().Tables.Add(dt3)

        'da = New OleDbDataAdapter("SELECT [PosData.DataDescription],[PosData.DataCode],[CollectorData.DataCode] FROM [PosData] INNER JOIN [CollectorData] ON [PosData.DataCode] = [CollectorData.DataCode]", connTable)

        Call New OleDbDataAdapter("SELECT [PosData.DataDescription] AS Name,[PosData.DataCode] AS Pos,[CollectorData.DataCode] AS Collector FROM [PosData] INNER JOIN [CollectorData] ON [PosData].DataCode LIKE [CollectorData].DataCode", connTable).Fill(dt1)
        DataGridView1.DataSource = dt1.DefaultView

        ''Call New OleDbDataAdapter("SELECT [PosData.DataDescription] AS Name,[PosData.DataCode] AS Pos,[CollectorData.DataCode] AS Collector FROM [PosData] INNER JOIN [CollectorData] ON [PosData].DataCode NOT LIKE [CollectorData].DataCode AND [CollectorData].DataCode NOT LIKE [PosData].DataCode", connTable).Fill(dt2)
        'Call New OleDbDataAdapter("SELECT [PosData.DataDescription] AS Name,[PosData.DataCode] AS Pos,[CollectorData.DataCode] AS Collector FROM [PosData] LEFT JOIN [CollectorData] ON [PosData].DataCode LIKE [CollectorData].DataCode", connTable).Fill(dt2)
        'DataGridView2.DataSource = dt2.DefaultView

        Call New OleDbDataAdapter("SELECT [PosData.DataDescription] AS Name,[PosData.DataCode] AS Pos,[CollectorData.DataCode] AS Collector FROM [PosData] LEFT JOIN [CollectorData] ON [PosData].DataCode LIKE [CollectorData].DataCode UNION SELECT [PosData.DataDescription] AS Name,[PosData.DataCode] AS Pos,[CollectorData.DataCode] AS Collector FROM [CollectorData] LEFT JOIN [PosData] ON [PosData].DataCode LIKE [CollectorData].DataCode", connTable).Fill(dt3)
        DataGridView3.DataSource = dt3.DefaultView

        connTable.Close()
    End Sub
    Private Sub subExportDGVToCSV(ByVal strExportFileName As String, ByVal DataGridView As DataGridView, Optional ByVal blnWriteColumnHeaderNames As Boolean = False, Optional ByVal strDelimiterType As String = ",")

        Dim sr As StreamWriter = File.CreateText(strExportFileName)
        Dim strDelimiter As String = strDelimiterType
        Dim intColumnCount As Integer = DataGridView.Columns.Count - 1
        Dim strRowData As String = ""


        For intX As Integer = 0 To intColumnCount
                strRowData += Replace(DataGridView.Columns(intX).Name, strDelimiter, "") & IIf(intX < intColumnCount, strDelimiter, "")
            Next intX
            sr.WriteLine(strRowData)

        For intX As Integer = 0 To DataGridView.Rows.Count - 1
            strRowData = ""
            For intRowData As Integer = 0 To intColumnCount
                Try
                    strRowData += Replace(DataGridView.Rows(intX).Cells(intRowData).Value, strDelimiter, "") & IIf(intRowData < intColumnCount, strDelimiter, "")
                Catch ex As Exception
                    strRowData += ","
                End Try
            Next intRowData
            sr.WriteLine(strRowData)
        Next intX
        sr.Close()
    End Sub
End Class