Imports System.Data.SqlClient
Imports System.IO
Imports Scripting
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class Form2

    Dim fso As New FileSystemObject
    Dim ts, tsLog As TextStream
    Dim ConfigFile1 As String = "QueryConfig.ini"
    Dim Query1, outputfolder, connection1, Query2, rowfolder, count As String
    Public Sub ReadConfigFile1()
        If IO.File.Exists(ConfigFile1) Then
            Me.ts = fso.OpenTextFile(ConfigFile1)
            Dim ConfigParam As String
            Do While Not Me.ts.AtEndOfStream
                ConfigParam = ts.ReadLine.Trim
                If (ConfigParam).StartsWith("Query1=") Then
                    Query1 = ConfigParam.Replace("Query1=", "")
                ElseIf (ConfigParam).StartsWith("outputfolder=") Then
                    outputfolder = ConfigParam.Replace("outputfolder=", "")
                ElseIf (ConfigParam).StartsWith("connection=") Then
                    connection1 = ConfigParam.Replace("connection=", "")
                ElseIf (ConfigParam).StartsWith("Query2=") Then
                    Query2 = ConfigParam.Replace("Query2=", "")
                ElseIf (ConfigParam).StartsWith("rowfolder=") Then
                    rowfolder = ConfigParam.Replace("rowfolder=", "")
                ElseIf (ConfigParam).StartsWith("count=") Then
                    count = ConfigParam.Replace("count=", "")

                End If
            Loop
            Me.ts.Close()
            ts = Nothing
            'Me.fso = Nothing
        Else
            ts = fso.CreateTextFile(ConfigFile1, False)
            ts.WriteLine("Query1=SELECT '''' + CONVERT(nvarchar, [Invoice No:]) AS [Invoice No:],[LPO NO:],[Invoice Date]AS [Invoice Date] FROM [Indoguna].[dbo].[IndogunaDB]")
            ts.WriteLine("outputfolder=C:\Users\Akshay\Desktop\csvfilepath")
            ts.WriteLine("rowfolder=C:\Users\Akshay\Desktop\last")
            ts.WriteLine("Data Source=DESKTOP-LPHK7HD;Initial Catalog=EROS;User ID=sa;Password=p@ssw0rd")
            ts.WriteLine("Query2=SELECT [Invoice No],''''+CONVERT(VARCHAR(10), [InvoiceDate], 105) AS [InvoiceDate],[Entity],[Customer Code],[Total Amount Inclusive Vat],[Total Amount Exclusive Vat],[Customer Code 1],[Site Code],[Site Name],[DN Number],''''+[Tax Cr Note Number] as [Tax Cr Note Number],''''+CONVERT(VARCHAR(10), [TaxCredit Note Date], 105) AS [TaxCredit Note Date],[Agreement Number],[Internal Doc Ref] FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY [Invoice No]) AS row_num FROM [EROS].[dbo].[dat_q_to_ex1]) AS numbered_rows WHERE row_num >")
            ts.Close()
            ts = Nothing
            'fso = Nothing
        End If

    End Sub

    Private Sub frmMain_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ReadConfigFile1()


        ' 


        'connection =Data Source=10.8.71.25; Initial Catalog=DMS;User Id=sa;Password=WeB@$20#$21#####
        'connection=Data Source=DESKTOP-LPHK7HD;Initial Catalog=Indoguna;User ID=sa;Password=p@ssw0rd ###LOCAL
        Dim connectionString As String = connection1
        Dim connection As New SqlConnection(connectionString)
        connection.Open()
        'Dim curdate As String = "SELECT CAST( GETDATE() AS Date );"
        'Dim command As New SqlCommand(curdate, connection)
        'Dim reader As SqlDataReader = command.ExecuteReader()
        'Dim date1 As String =

        Dim excelFilePath As String = rowfolder + "\" + "LastRowNumber.xlsx"
        Dim Value As String = ReadExcelFile(excelFilePath)

        'SQL SELECT query to retrieve data from the Invoice table
        'Dim query1 As String = "SELECT '''' + CONVERT(nvarchar, [Invoice No:]) AS [Invoice No:],[LPO NO:],
        '[Invoice Date]
        'AS [Invoice Date] FROM [Indoguna].[dbo].[IndogunaDB] where [Invoice Date] = FORMAT(GETDATE(), 'dd/MM/yyyy') ;"

        If Value = 0 Then

            Dim command As New SqlCommand(Query1, connection)
            Dim reader As SqlDataReader = command.ExecuteReader()

            'Calculate column widths based on data
            Dim invoiceWidth As Integer = 40
            Dim lpoWidth As Integer = 40
            Dim dateWidth As Integer = 40

            While reader.Read()
                Dim invoice_no As String = reader("Invoice No").ToString()
                Dim InvoiceDate As String = reader("InvoiceDate").ToString()
                Dim Entity As String = reader("Entity").ToString()
                Dim Customer_Code As String = reader("Customer Code").ToString()
                Dim Total_Inc_Vat As String = reader("Total Amount Inclusive Vat").ToString()
                Dim Total_Exc_Vat As String = reader("Total Amount Exclusive Vat").ToString()
                Dim Customer_Code_1 As String = reader("Customer Code 1").ToString()
                Dim Site_Code As String = reader("Site Code").ToString()
                Dim Site_Name As String = reader("Site Name").ToString()
                Dim DN_Number As String = reader("DN Number").ToString()
                Dim Tax_Number As String = reader("Tax Cr Note Number").ToString()
                Dim TaxCredit_Date As String = reader("TaxCredit Note Date").ToString()
                Dim Agreement_Number As String = reader("Agreement Number").ToString()
                Dim Internal_Doc As String = reader("Internal Doc Ref").ToString()
                'worksheet.Column(3)
                ' Update column widths based on data length if necessary
                invoiceWidth = Math.Max(invoiceWidth, invoice_no.Length)
                lpoWidth = Math.Max(lpoWidth, InvoiceDate.Length)
                dateWidth = Math.Max(dateWidth, Entity.Length)
                dateWidth = Math.Max(dateWidth, Customer_Code.Length)
                dateWidth = Math.Max(dateWidth, Total_Inc_Vat.Length)
                dateWidth = Math.Max(dateWidth, Total_Exc_Vat.Length)
                dateWidth = Math.Max(dateWidth, Customer_Code_1.Length)
                dateWidth = Math.Max(dateWidth, Site_Code.Length)
                dateWidth = Math.Max(dateWidth, Site_Name.Length)
                dateWidth = Math.Max(dateWidth, DN_Number.Length)
                dateWidth = Math.Max(dateWidth, Tax_Number.Length)
                dateWidth = Math.Max(dateWidth, TaxCredit_Date.Length)
                dateWidth = Math.Max(dateWidth, Agreement_Number.Length)
                dateWidth = Math.Max(dateWidth, Internal_Doc.Length)

            End While

            reader.Close()

            ' Create a new CSV file to save the data
            Dim csvFilePath As String = outputfolder + "\" + "Daily_Report.csv"
            Using writer As New StreamWriter(csvFilePath, False)
                ' Write the header row with the calculated column widths
                writer.WriteLine($"{"invoice_no".PadRight(invoiceWidth)},{"InvoiceDate".PadRight(lpoWidth)},{"Entity".PadRight(dateWidth)},{"Customer Code".PadRight(dateWidth)}, {"Total Amount Inclusive Vat".PadRight(dateWidth)},{"Total Amount Exclusive Vat".PadRight(dateWidth)},{"Customer Code 1".PadRight(dateWidth)},{"Site Code".PadRight(dateWidth)},{"Site Name".PadRight(dateWidth)},{"DN Number".PadRight(dateWidth)},{"Tax Cr Note Number".PadRight(dateWidth)},{"TaxCredit Note Date".PadRight(dateWidth)},{"Agreement Number".PadRight(dateWidth)},{"Internal Doc Ref".PadRight(dateWidth)}")

                ' Reopen the reader to read the data again for writing
                reader = command.ExecuteReader()

                ' Write the data rows
                While reader.Read()

                    'Dim Invoice_date1 As DateTime = DirectCast(reader("Invoice Date"), DateTime)
                    'MsgBox("Invoice_date1")

                    Dim invoice_no As String = reader("Invoice No").ToString().PadRight(invoiceWidth)
                    Dim InvoiceDate As String = reader("InvoiceDate").ToString().PadRight(lpoWidth)
                    Dim Entity As String = reader("Entity").ToString().PadRight(dateWidth)
                    Dim Customer_Code As String = reader("Customer Code").ToString().PadRight(dateWidth)
                    Dim Total_Inc_Vat As String = reader("Total Amount Inclusive Vat").ToString().PadRight(dateWidth)
                    Dim Total_Exc_Vat As String = reader("Total Amount Exclusive Vat").ToString().PadRight(dateWidth)
                    Dim Customer_Code_1 As String = reader("Customer Code 1").ToString().PadRight(dateWidth)
                    Dim Site_Code As String = reader("Site Code").ToString().PadRight(dateWidth)
                    Dim Site_Name As String = reader("Site Name").ToString().PadRight(dateWidth)
                    Dim DN_Number As String = reader("DN Number").ToString().PadRight(dateWidth)
                    Dim Tax_Cr_Number As String = reader("Tax Cr Note Number").ToString().PadRight(dateWidth)
                    Dim TaxCredit_Date As String = reader("TaxCredit Note Date").ToString().PadRight(dateWidth)
                    Dim Agreement_Number As String = reader("Agreement Number").ToString().PadRight(dateWidth)
                    Dim Internal_Doc_Ref As String = reader("Internal Doc Ref").ToString().PadRight(dateWidth)


                    ' Write the data to the CSV file
                    writer.WriteLine($"{invoice_no},{InvoiceDate},{Entity},{Customer_Code},{Total_Inc_Vat},{Total_Exc_Vat},{Customer_Code_1},{Site_Code},{Site_Name},{DN_Number},{Tax_Cr_Number},{TaxCredit_Date},{Agreement_Number},{Internal_Doc_Ref}")
                End While

                reader.Close()
            End Using
        Else ' query number two----------------------------------------------------------------------------------------------------------------------------------------------------------
            'MessageBox.Show(Value)
            Dim query3 As String = Query2 + Value
            'Dim query3 As String = "SELECT [Invoice No],[InvoiceDate],[Entity],[Customer Code],[Total Amount Inclusive Vat],[Total Amount Exclusive Vat],[Customer Code 1],[Site Code],[Site Name],[DN Number],''''+[Tax Cr Note Number] as [Tax Cr Note Number],[TaxCredit Note Date],[Agreement Number],[Internal Doc Ref] FROM (SELECT *, ROW_NUMBER() OVER (ORDER BY [Invoice No]) AS row_num FROM [EROS].[dbo].[dat_q_to_ex1]) AS numbered_rows WHERE row_num >" + Value

            Dim command As New SqlCommand(query3, connection)
            Dim reader As SqlDataReader = command.ExecuteReader()

            'Calculate column widths based on data
            Dim invoiceWidth As Integer = 40
            Dim lpoWidth As Integer = 40
            Dim dateWidth As Integer = 40

            While reader.Read()
                Dim invoice_no As String = reader("Invoice No").ToString()
                Dim InvoiceDate As String = reader("InvoiceDate").ToString()
                Dim Entity As String = reader("Entity").ToString()
                Dim Customer_Code As String = reader("Customer Code").ToString()
                Dim Total_Inc_Vat As String = reader("Total Amount Inclusive Vat").ToString()
                Dim Total_Exc_Vat As String = reader("Total Amount Exclusive Vat").ToString()
                Dim Customer_Code_1 As String = reader("Customer Code 1").ToString()
                Dim Site_Code As String = reader("Site Code").ToString()
                Dim Site_Name As String = reader("Site Name").ToString()
                Dim DN_Number As String = reader("DN Number").ToString()
                Dim Tax_Number As String = reader("Tax Cr Note Number").ToString()
                Dim TaxCredit_Date As String = reader("TaxCredit Note Date").ToString()
                Dim Agreement_Number As String = reader("Agreement Number").ToString()
                Dim Internal_Doc As String = reader("Internal Doc Ref").ToString()

                ' Update column widths based on data length if necessary
                invoiceWidth = Math.Max(invoiceWidth, invoice_no.Length)
                lpoWidth = Math.Max(lpoWidth, InvoiceDate.Length)
                dateWidth = Math.Max(dateWidth, Entity.Length)
                dateWidth = Math.Max(dateWidth, Customer_Code.Length)
                dateWidth = Math.Max(dateWidth, Total_Inc_Vat.Length)
                dateWidth = Math.Max(dateWidth, Total_Exc_Vat.Length)
                dateWidth = Math.Max(dateWidth, Customer_Code_1.Length)
                dateWidth = Math.Max(dateWidth, Site_Code.Length)
                dateWidth = Math.Max(dateWidth, Site_Name.Length)
                dateWidth = Math.Max(dateWidth, DN_Number.Length)
                dateWidth = Math.Max(dateWidth, Tax_Number.Length)
                dateWidth = Math.Max(dateWidth, TaxCredit_Date.Length)
                dateWidth = Math.Max(dateWidth, Agreement_Number.Length)
                dateWidth = Math.Max(dateWidth, Internal_Doc.Length)

            End While

            reader.Close()

            ' Create a new CSV file to save the data
            Dim csvFilePath As String = outputfolder + "\" + "Daily_Report.csv"
            Using writer As New StreamWriter(csvFilePath, False)
                ' Write the header row with the calculated column widths
                writer.WriteLine($"{"invoice_no".PadRight(invoiceWidth)},{"InvoiceDate".PadRight(lpoWidth)},{"Entity".PadRight(dateWidth)},{"Customer Code".PadRight(dateWidth)}, {"Total Amount Inclusive Vat".PadRight(dateWidth)},{"Total Amount Exclusive Vat".PadRight(dateWidth)},{"Customer Code 1".PadRight(dateWidth)},{"Site Code".PadRight(dateWidth)},{"Site Name".PadRight(dateWidth)},{"DN Number".PadRight(dateWidth)},{"Tax Cr Note Number".PadRight(dateWidth)},{"TaxCredit Note Date".PadRight(dateWidth)},{"Agreement Number".PadRight(dateWidth)},{"Internal Doc Ref".PadRight(dateWidth)}")

                ' Reopen the reader to read the data again for writing
                reader = command.ExecuteReader()

                ' Write the data rows
                While reader.Read()

                    'Dim Invoice_date1 As DateTime = DirectCast(reader("Invoice Date"), DateTime)
                    'MsgBox("Invoice_date1")

                    Dim invoice_no As String = reader("Invoice No").ToString().PadRight(invoiceWidth)
                    Dim InvoiceDate As String = reader("InvoiceDate").ToString().PadRight(lpoWidth)
                    Dim Entity As String = reader("Entity").ToString().PadRight(dateWidth)
                    Dim Customer_Code As String = reader("Customer Code").ToString().PadRight(dateWidth)
                    Dim Total_Inc_Vat As String = reader("Total Amount Inclusive Vat").ToString().PadRight(dateWidth)
                    Dim Total_Exc_Vat As String = reader("Total Amount Exclusive Vat").ToString().PadRight(dateWidth)
                    Dim Customer_Code_1 As String = reader("Customer Code 1").ToString().PadRight(dateWidth)
                    Dim Site_Code As String = reader("Site Code").ToString().PadRight(dateWidth)
                    Dim Site_Name As String = reader("Site Name").ToString().PadRight(dateWidth)
                    Dim DN_Number As String = reader("DN Number").ToString().PadRight(dateWidth)
                    Dim Tax_Cr_Number As String = reader("Tax Cr Note Number").ToString().PadRight(dateWidth)
                    Dim TaxCredit_Date As String = reader("TaxCredit Note Date").ToString().PadRight(dateWidth)
                    Dim Agreement_Number As String = reader("Agreement Number").ToString().PadRight(dateWidth)
                    Dim Internal_Doc_Ref As String = reader("Internal Doc Ref").ToString().PadRight(dateWidth)


                    ' Write the data to the CSV file
                    writer.WriteLine($"{invoice_no},{InvoiceDate},{Entity},{Customer_Code},{Total_Inc_Vat},{Total_Exc_Vat},{Customer_Code_1},{Site_Code},{Site_Name},{DN_Number},{Tax_Cr_Number},{TaxCredit_Date},{Agreement_Number},{Internal_Doc_Ref}")
                End While

                reader.Close()
            End Using






        End If

        Dim lastRowNumber1 As Integer = GetLastRowNumberAndSaveToExcel()
        ' MessageBox.Show($"Last Row Number: {lastRowNumber1} saved to Excel.")

        connection.Close()
        Environment.Exit(0)
    End Sub

    Public Function GetLastRowNumber() As Integer
        Dim lastRowNumber As Integer = -1 ' Initialize with a default value

        Dim connectionString As String = connection1
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            ' Dim query As String = " SELECT COUNT(*) FROM [EROS].[dbo].[dat_q_to_ex1]"
            Using command As New SqlCommand(count, connection)
                Dim result As Object = command.ExecuteScalar()
                If result IsNot DBNull.Value Then
                    lastRowNumber = Convert.ToInt32(result)
                End If
            End Using
        End Using

        Return lastRowNumber
    End Function

    Public Function GetLastRowNumberAndSaveToExcel() As Integer
        Dim lastRowNumber As Integer = GetLastRowNumber() ' Get the last row number

        ' Create a new Excel package
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Using excelPackage As New ExcelPackage()

            Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets.Add("LastRowNumber")

            ' Add the last row number to the Excel worksheet
            worksheet.Cells("A1").Value = "Last Row Number"
            worksheet.Cells("B1").Value = lastRowNumber

            ' Save the Excel package to a file
            Dim excelFilePath1 As String = rowfolder + "\" + "LastRowNumber.xlsx" ' Change this to the desired file path
            excelPackage.SaveAs(New FileInfo(excelFilePath1))
        End Using

        Return lastRowNumber
    End Function

    Public Function ReadExcelFile(filePath As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Using excelPackage As New ExcelPackage(New FileInfo(filePath))
            Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets(0) ' Assuming you want to read the first worksheet

            ' Read data from specific cells
            Dim lastRowNumber2 As Integer = CInt(worksheet.Cells("B1").Value) ' Read the value from cell B1

            ' Process the read data
            Console.WriteLine("Last Row Number: " & lastRowNumber2)
            'MessageBox.Show($"Last Row Number: {lastRowNumber2} value")
            Return lastRowNumber2
        End Using
    End Function



End Class
