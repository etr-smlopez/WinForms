Imports System.IO
Imports System.Net
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class Form1

    '        Dim apiUrl As String = "http://192.168.0.20/api/Dashboard/"
    Dim apiUrl As String = "https://localhost:7251/api/Dashboard/"

    Private salesOrderCache As Dictionary(Of Integer, JObject) = New Dictionary(Of Integer, JObject)()
    Private returnCache As Dictionary(Of Integer, JObject) = New Dictionary(Of Integer, JObject)()
    Private invoiceCache As Dictionary(Of Integer, JObject) = New Dictionary(Of Integer, JObject)()
    Private stockTransferCache As Dictionary(Of Integer, JObject) = New Dictionary(Of Integer, JObject)()

    Dim baseUrl As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dgTableHeaders()

        Try

            Dim jsonResponse As String = SendGetRequest(apiUrl)
            Dim data As JObject = JObject.Parse(jsonResponse)

            Dim salesOrders As JObject = DirectCast(data("SalesOrder"), JObject)
            Dim _return As JObject = DirectCast(data("Returns"), JObject)
            Dim invoices As JObject = DirectCast(data("Invoice"), JObject)
            Dim stockTransfers As JObject = DirectCast(data("StockTransfer"), JObject)

            For Each salesOrderItem As KeyValuePair(Of String, JToken) In salesOrders
                Dim salesOrderJson As JObject = DirectCast(salesOrderItem.Value, JObject)
                Dim salesOrderKey As Integer = Integer.Parse(salesOrderItem.Key)
                Dim row As DataGridViewRow = New DataGridViewRow()
                row.CreateCells(dgSalesOrder)
                row.Cells(0).Value = salesOrderKey
                row.Cells(1).Value = salesOrderJson("orderStatus")
                row.Cells(2).Value = salesOrderJson("sO_SONumber")
                Dim soDate As DateTime = salesOrderJson("sO_SODate").ToObject(Of DateTime)()
                row.Cells(3).Value = soDate.ToString("yyyy-MM-dd")
                row.Cells(4).Value = salesOrderJson("sO_CreatedBy").ToString().ToUpper()
                dgSalesOrder.Rows.Add(row)
                salesOrderCache.Add(salesOrderKey, salesOrderJson)
            Next

            For Each returnsItem As KeyValuePair(Of String, JToken) In _return
                Dim returnsJson As JObject = DirectCast(returnsItem.Value, JObject)
                Dim returnsKey As Integer = Integer.Parse(returnsItem.Key)
                Dim row As DataGridViewRow = New DataGridViewRow()
                row.CreateCells(dgSalesOrder)
                row.Cells(0).Value = returnsKey
                row.Cells(1).Value = returnsJson("status")
                row.Cells(2).Value = returnsJson("rsNumber")
                Dim returnsDate As DateTime = returnsJson("returnDate").ToObject(Of DateTime)()
                row.Cells(3).Value = returnsDate.ToString("yyyy-MM-dd")
                row.Cells(4).Value = returnsJson("employeeName").ToString().ToUpper()
                dgReturns.Rows.Add(row)
                returnCache.Add(returnsKey, returnsJson)
            Next

            For Each invoicesItem As KeyValuePair(Of String, JToken) In invoices
                Dim invoicesJson As JObject = DirectCast(invoicesItem.Value, JObject)
                Dim invoicesKey As Integer = Integer.Parse(invoicesItem.Key)
                Dim row As DataGridViewRow = New DataGridViewRow()
                row.CreateCells(dgSalesOrder)
                row.Cells(0).Value = invoicesKey
                row.Cells(1).Value = invoicesJson("description")
                row.Cells(2).Value = invoicesJson("invoiceNumber")
                Dim invDate As DateTime = invoicesJson("invoiceDate").ToObject(Of DateTime)()
                row.Cells(3).Value = invDate.ToString("yyyy-MM-dd")
                row.Cells(4).Value = invoicesJson("cust_storename").ToString().ToUpper()
                dgInvoice.Rows.Add(row)
                invoiceCache.Add(invoicesKey, invoicesJson)
            Next

            For Each stockTransfersItem As KeyValuePair(Of String, JToken) In stockTransfers
                Dim stockTransfersJson As JObject = DirectCast(stockTransfersItem.Value, JObject)
                Dim stockTransfersKey As Integer = Integer.Parse(stockTransfersItem.Key)
                Dim row As DataGridViewRow = New DataGridViewRow()
                row.CreateCells(dgSalesOrder)
                row.Cells(0).Value = stockTransfersKey
                row.Cells(1).Value = stockTransfersJson("status")
                row.Cells(2).Value = stockTransfersJson("transferSlipNo")
                Dim invDate As DateTime = stockTransfersJson("transferDate").ToObject(Of DateTime)()
                row.Cells(3).Value = invDate.ToString("yyyy-MM-dd")
                row.Cells(4).Value = stockTransfersJson("createdBy").ToString().ToUpper()
                dgStockTransfer.Rows.Add(row)
                stockTransferCache.Add(stockTransfersKey, stockTransfersJson)
            Next

        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message & "", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgTableHeaders()
        dgSalesOrder.Columns.Add("salesOrderKey", "SalesOrder Key")
        dgSalesOrder.Columns.Add("orderStatus", "Status")
        dgSalesOrder.Columns.Add("sO_SONumber", "Transaction Number")
        dgSalesOrder.Columns.Add("sO_SODate", "Transaction Date")
        dgSalesOrder.Columns.Add("sO_CreatedBy", "Created By")
        '    DataGridView1.Columns("salesOrderKeyColumn").Visible = False
        For Each column As DataGridViewColumn In dgSalesOrder.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        dgReturns.Columns.Add("returnsKey", "Returns Key")
        dgReturns.Columns.Add("status", "Status")
        dgReturns.Columns.Add("rsNumber", "Transaction Number")
        dgReturns.Columns.Add("returnDate", "Transaction Date")
        dgReturns.Columns.Add("employeeName", "Created By")
        '    DataGridView2.Columns("returnsKey").Visible = False
        For Each column As DataGridViewColumn In dgReturns.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        dgInvoice.Columns.Add("invoiceKey", "Invoice Key")
        dgInvoice.Columns.Add("description", "Status")
        dgInvoice.Columns.Add("invoiceNumber", "Invoice Number")
        dgInvoice.Columns.Add("invoiceDate", "Invoice Date")
        dgInvoice.Columns.Add("cust_storename", "Created By")
        '     dgInvoice.Columns("salesOrderKeyColumn").Visible = False
        For Each column As DataGridViewColumn In dgInvoice.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        dgStockTransfer.Columns.Add("stockTransferKey", "Stock Transfer Key")
        dgStockTransfer.Columns.Add("status", "Status")
        dgStockTransfer.Columns.Add("transferSlipNo", "Stock Transfer Number")
        dgStockTransfer.Columns.Add("transferDate", "Stock Transfer Date")
        dgStockTransfer.Columns.Add("createdBy", "Created By")
        '     dgStockTransfer.Columns("returnsKey").Visible = False
        For Each column As DataGridViewColumn In dgStockTransfer.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next
    End Sub

    Private Function SendGetRequest(url As String) As String
        Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        request.Method = "GET"

        Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Using reader As New StreamReader(response.GetResponseStream())
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Try
            Select Case ComboBox1.SelectedItem.ToString()
                Case "salesorder"
                    '   baseUrl = "http://192.168.0.20/api/Dashboard/salerorder/"
                    baseUrl = "https://localhost:7251/api/Dashboard/salerorder/"
                Case "returns"
                    '        baseUrl = "http://192.168.0.20/api/Dashboard/returns/"
                    baseUrl = "https://localhost:7251/api/Dashboard/returns/"
                Case "invoice"
                    '            baseUrl = "http://192.168.0.20/api/Dashboard/invoice/"
                    baseUrl = "https://localhost:7251/api/Dashboard/invoice/"
                Case "stocktransfer"
                    '             baseUrl = "http://192.168.0.20/api/Dashboard/stocktransfer/"
                    baseUrl = "https://localhost:7251/api/Dashboard/stocktransfer/"
            End Select

            Dim tableID As String = TextBox1.Text
            Dim postapiUrl As String = baseUrl & tableID.ToString()
            Dim postDataJson As String = GetJsonData()
            Dim jsonResponse As String = SendPostRequest(postapiUrl, postDataJson)

            CacheUpdate(tableID, postDataJson)


            MessageBox.Show("POST request successful:" & jsonResponse & " ", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("An error occurred:" & ex.Message & "", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CacheUpdate(tableID As String, postDataJson As String)
        Select Case ComboBox1.SelectedItem.ToString()
            Case "salesorder"
                Dim salesOrderJson As JObject = JObject.Parse(postDataJson)
                Dim salesOrderKey As Integer = Integer.Parse(tableID)
                salesOrderCache(salesOrderKey) = salesOrderJson
                dgSalesOrderUpdate(tableID)
            Case "returns"
                Dim returnsJson As JObject = JObject.Parse(postDataJson)
                Dim returnsKey As Integer = Integer.Parse(tableID)
                returnCache(returnsKey) = returnsJson
                dgReturnsUpdate(tableID)
            Case "invoice"
                Dim invoicesJson As JObject = JObject.Parse(postDataJson)
                Dim invoicesKey As Integer = Integer.Parse(tableID)
                invoiceCache(invoicesKey) = invoicesJson
                dgInvoiceUpdate(tableID)
            Case "stocktransfer"
                Dim stockTransfersJson As JObject = JObject.Parse(postDataJson)
                Dim stockTransfersKey As Integer = Integer.Parse(tableID)
                stockTransferCache(stockTransfersKey) = stockTransfersJson
                dgStockTransferUpdate(tableID)
        End Select
    End Sub

    Private Sub dgSalesOrderUpdate(tableID As String)
        Dim rowIndex As Integer = -1
        For Each row As DataGridViewRow In dgSalesOrder.Rows
            If row.Cells("salesOrderKey").Value IsNot Nothing AndAlso row.Cells("salesOrderKey").Value.ToString() = tableID Then
                rowIndex = row.Index
                Exit For
            End If
        Next
        If rowIndex <> -1 Then
            Dim salesOrderJson As JObject = salesOrderCache(Integer.Parse(tableID))
            dgSalesOrder.Rows(rowIndex).Cells("orderStatus").Value = salesOrderJson("orderStatus")
            dgSalesOrder.Rows(rowIndex).Cells("sO_SONumber").Value = salesOrderJson("sO_SONumber")
            Dim soDate As DateTime = salesOrderJson("sO_SODate").ToObject(Of DateTime)()
            dgSalesOrder.Rows(rowIndex).Cells("sO_SODate").Value = soDate.ToString("yyyy-MM-dd")
            dgSalesOrder.Rows(rowIndex).Cells("sO_CreatedBy").Value = salesOrderJson("sO_CreatedBy").ToString().ToUpper()
        Else
            Dim salesOrderJson As JObject = salesOrderCache(Integer.Parse(tableID))
            Dim row As DataGridViewRow = New DataGridViewRow()
            row.CreateCells(dgSalesOrder)
            row.Cells(0).Value = tableID
            row.Cells(1).Value = salesOrderJson("orderStatus")
            row.Cells(2).Value = salesOrderJson("sO_SONumber")
            Dim soDate As DateTime = DateTime.Now
            row.Cells(3).Value = soDate.ToString("yyyy-MM-dd")
            row.Cells(4).Value = salesOrderJson("sO_CreatedBy").ToString().ToUpper()
            dgSalesOrder.Rows.Add(row)
            dgSalesOrder.Refresh()
        End If
    End Sub
    Private Sub dgReturnsUpdate(tableID As String)
        Dim rowIndex As Integer = -1
        For Each row As DataGridViewRow In dgReturns.Rows
            If row.Cells("returnsKey").Value IsNot Nothing AndAlso row.Cells("returnsKey").Value.ToString() = tableID Then
                rowIndex = row.Index
                Exit For
            End If
        Next
        If rowIndex <> -1 Then
            Dim returnsJson As JObject = returnCache(Integer.Parse(tableID))
            dgReturns.Rows(rowIndex).Cells("status").Value = returnsJson("status")
            dgReturns.Rows(rowIndex).Cells("rsNumber").Value = returnsJson("rsNumber")
            Dim soDate As DateTime = returnsJson("returnDate").ToObject(Of DateTime)()
            dgReturns.Rows(rowIndex).Cells("returnDate").Value = soDate.ToString("yyyy-MM-dd")
            dgReturns.Rows(rowIndex).Cells("employeeName").Value = returnsJson("employeeName").ToString().ToUpper()
        Else
            Dim returnsJson As JObject = returnCache(Integer.Parse(tableID))
            Dim row As DataGridViewRow = New DataGridViewRow()
            row.CreateCells(dgReturns)
            row.Cells(0).Value = tableID
            row.Cells(1).Value = returnsJson("status")
            row.Cells(2).Value = returnsJson("rsNumber")
            Dim soDate As DateTime = DateTime.Now
            row.Cells(3).Value = soDate.ToString("yyyy-MM-dd")
            row.Cells(4).Value = returnsJson("employeeName").ToString().ToUpper()
            dgReturns.Rows.Add(row)
            dgReturns.Refresh()
        End If
    End Sub
    Private Sub dgInvoiceUpdate(tableID As String)
        Dim rowIndex As Integer = -1
        For Each row As DataGridViewRow In dgInvoice.Rows
            If row.Cells("invoiceKey").Value IsNot Nothing AndAlso row.Cells("invoiceKey").Value.ToString() = tableID Then
                rowIndex = row.Index
                Exit For
            End If
        Next
        If rowIndex <> -1 Then
            Dim invoiceJson As JObject = invoiceCache(Integer.Parse(tableID))
            dgInvoice.Rows(rowIndex).Cells("description").Value = invoiceJson("description")
            dgInvoice.Rows(rowIndex).Cells("invoiceNumber").Value = invoiceJson("invoiceNumber")
            Dim soDate As DateTime = invoiceJson("invoiceDate").ToObject(Of DateTime)()
            dgInvoice.Rows(rowIndex).Cells("invoiceDate").Value = soDate.ToString("yyyy-MM-dd")
            dgInvoice.Rows(rowIndex).Cells("cust_storename").Value = invoiceJson("cust_storename").ToString().ToUpper()
        Else
            Dim invoiceJson As JObject = invoiceCache(Integer.Parse(tableID))
            Dim row As DataGridViewRow = New DataGridViewRow()
            row.CreateCells(dgInvoice)
            row.Cells(0).Value = tableID
            row.Cells(1).Value = invoiceJson("description")
            row.Cells(2).Value = invoiceJson("invoiceNumber")
            Dim soDate As DateTime = DateTime.Now
            row.Cells(3).Value = soDate.ToString("yyyy-MM-dd")
            row.Cells(4).Value = invoiceJson("cust_storename").ToString().ToUpper()
            dgInvoice.Rows.Add(row)
            dgInvoice.Refresh()
        End If
    End Sub
    Private Sub dgStockTransferUpdate(tableID As String)
        Dim rowIndex As Integer = -1
        For Each row As DataGridViewRow In dgStockTransfer.Rows
            If row.Cells("stockTransferKey").Value IsNot Nothing AndAlso row.Cells("stockTransferKey").Value.ToString() = tableID Then
                rowIndex = row.Index
                Exit For
            End If
        Next
        If rowIndex <> -1 Then
            Dim stockTransferJson As JObject = stockTransferCache(Integer.Parse(tableID))
            dgStockTransfer.Rows(rowIndex).Cells("status").Value = stockTransferJson("status")
            dgStockTransfer.Rows(rowIndex).Cells("transferSlipNo").Value = stockTransferJson("transferSlipNo")
            Dim soDate As DateTime = stockTransferJson("transferDate").ToObject(Of DateTime)()
            dgStockTransfer.Rows(rowIndex).Cells("transferDate").Value = soDate.ToString("yyyy-MM-dd")
            dgStockTransfer.Rows(rowIndex).Cells("createdBy").Value = stockTransferJson("createdBy").ToString().ToUpper()
        Else
            Dim stockTransferJson As JObject = stockTransferCache(Integer.Parse(tableID))
            Dim row As DataGridViewRow = New DataGridViewRow()
            row.CreateCells(dgStockTransfer)
            row.Cells(0).Value = tableID
            row.Cells(1).Value = stockTransferJson("status")
            row.Cells(2).Value = stockTransferJson("transferSlipNo")
            Dim soDate As DateTime = DateTime.Now
            row.Cells(3).Value = soDate.ToString("yyyy-MM-dd")
            row.Cells(4).Value = stockTransferJson("createdBy").ToString().ToUpper()
            dgStockTransfer.Rows.Add(row)
            dgStockTransfer.Refresh()
        End If
    End Sub

    Private Function GetJsonData() As String
        Dim jsonData As String = ""

        Select Case ComboBox1.SelectedItem.ToString()
            Case "salesorder"
                Dim salesOrder As New SalesOrder With {
                .salesOrderKey = TextBox1.Text,
                .orderStatus = TextBox2.Text,
                .sO_SONumber = TextBox3.Text,
                .sO_SODate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                .sO_CreatedBy = TextBox5.Text
            }
                jsonData = JsonConvert.SerializeObject(salesOrder)

            Case "returns"
                Dim returns As New Returns With {
                .returnsKey = TextBox1.Text,
                .status = TextBox2.Text,
                .rsNumber = TextBox3.Text,
                .returnDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                .employeeName = TextBox5.Text
            }
                jsonData = JsonConvert.SerializeObject(returns)

            Case "invoice"
                Dim invoice As New Invoice With {
                .invoiceKey = TextBox1.Text,
                .description = TextBox2.Text,
                .invoiceNumber = TextBox3.Text,
                .invoiceDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                .cust_storename = TextBox5.Text
            }
                jsonData = JsonConvert.SerializeObject(invoice)

            Case "stocktransfer"
                Dim stocktransfer As New StockTransfer With {
                .stockTransferKey = TextBox1.Text,
                .status = TextBox2.Text,
                .transferSlipNo = TextBox3.Text,
                .transferDate = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                .createdBy = TextBox5.Text
            }
                jsonData = JsonConvert.SerializeObject(stocktransfer)

            Case Else
                MessageBox.Show("Invalid selection", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select

        Return jsonData
    End Function

    Private Function SendPostRequest(url As String, postDataJson As String) As String
        Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        request.Method = "POST"
        request.ContentType = "application/json"

        Dim dataBytes As Byte() = Encoding.UTF8.GetBytes(postDataJson)
        request.ContentLength = dataBytes.Length

        Using requestStream As Stream = request.GetRequestStream()
            requestStream.Write(dataBytes, 0, dataBytes.Length)
        End Using

        Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Using reader As New StreamReader(response.GetResponseStream())
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function


End Class


Public Class SalesOrder
    Public Property salesOrderKey As Integer
    Public Property orderStatus As String
    Public Property sO_SONumber As String
    Public Property sO_SODate As Date
    Public Property sO_CreatedBy As String
End Class

Public Class Returns
    Public Property returnsKey As Integer
    Public Property status As String
    Public Property rsNumber As String
    Public Property returnDate As Date
    Public Property employeeName As String
End Class

Public Class Invoice
    Public Property invoiceKey As Integer
    Public Property description As String
    Public Property invoiceNumber As String
    Public Property invoiceDate As Date
    Public Property cust_storename As String
End Class

Public Class StockTransfer
    Public Property stockTransferKey As Integer
    Public Property status As String
    Public Property transferSlipNo As String
    Public Property transferDate As Date
    Public Property createdBy As String
End Class
