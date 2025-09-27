Window Based application for Sales Invoice Data Entry and Invoice Print Module -2024

1. Create database <TEST + YOUR NAME>
2. Create following tables in the database
1.) Customer
Cust_Code	varchar (6)	Primary Key
Cust Name	varchar (50)	Not Null
Cust Add1	varchar (50)	Null
Cust Add2	varchar (50)	Null
Cust_Phone	varchar (30)	Null

2.) Product
Prod Code	varchar (6)	Primary Key
Prod Desc	varchar (50)	Not Null
Prod Price	Decimal (15,3)	Null, Default 0

3.) Inv_Header
Inv_No	int	Primary Key
Inv Date	datetime	Not Null
Inv Cust	varchar (6)	Not Null, FK (Customer->Cust_Code)

4.) Inv_Detail
Inv No	int	Primary Key
Inv Lno	int	Primary Key
Inv_Prod	varchar (6)	Not Null, FK (Product->Prod_Cd)
Inv_Qty	decimal (10,3)	Not Null
Inv_Price	decimal (15, 3)	Not Null

3. Insert data from backend in Customer & Product tables.
4. Create the project on desktop with name <TEST + YOUR NAME>
5. Write a data entry program using VB .NET (Only Invoice Data Entry Screen)
6. Create a SQL Query to display the Name and Total Sales invoice amount of the customer who has Maximum sales during last 6 months.

Invoice Data Entry Screen Format

Invoice No: 	Date: 

Customer Code:	?
Name		 :
Address	 :

Lno	Prd. Cd	Description	Unit Price	Oty	Amount
1					
2					
3					
4					
	         Amount Total:

New	Modify	Delete	Save	Cancel	Print	Query



Note: Data entry should have professional look

Checklist:
a. Invoice Number must be Auto-generated.
b. Search option should be provided to fetch existing data (Based on Invoice No.)
c. Invoice date should not be greater than system date.
d. For customer code selection use either combo box or a text box (Entry of new customer from frontend should be restricted). If text box is provided the provide search option to select from existing customer.
e. Customer Name and Address should not be editable and should be fetched from Customer table.
f. For product code selection use combo box / text box (Entry of new product from frontend should be restricted). If text box is provided the provide search option to select from existing product.
g. Product description and Unit price should be displayed from Product table.
h. Lno, description and Amount columns should not be editable in grid.
i. Product code, Unit Price and Qty columns should be editable in grid.
j. Qty should not be negative or zero.
k. Amount will be calculated automatically after entering Qty and Unit Price. Amount Total is to be updated at same time.
l. Preferably Ms Flex grid view /Data Grid view should be used.
m. Control Buttons should be provided as shown in "Invoice data entry screen format".

	i. New: To add a new invoice
	ii. Modify: To edit an existing invoice
	iii. Delete: To delete an existing invoice
	iv. Save: To update the data in tables. Save button should be visible only in New / Modify mode
	v. Cancel: To cancel the changes made. Cancel button should be visible in New / Modify mode. New, Modify, Delete, Print buttons should NOT be visible in New / Modify mode.
	vi. Query: To display the Name and Total Sales invoice amount of the customer who has Maximum sales during last 6 months

n. Use Message Box on "Query Button Click" to display Query Result.


It should be professional look and the SQL should use stored procedure. VB.Net 2010 & MS SQL 2016 supported code. I need full code and also explain how to do that. Print function is not currently needed.

////////////////////////////////////////////////////////////////////////////

SQL scripts (create database TESTANUMONJACOB, tables, sample data).

Stored procedures (insert/update/delete/get invoice, and the “max sales last 6 months” query).

A WinForms VB.NET implementation (single form invoice entry screen) with:

DataGridView (editable Product Code / Unit Price / Qty; read-only Description/Amount/Lno)

ComboBox for Customer selection (no front-end create)

Product selection via DataGridView ComboBox column (no front-end create)

Buttons: New, Modify, Delete, Save, Cancel, Print(not implemented), Query (messagebox per spec)

Validations: date <= system date, qty > 0, auto-update amount and total

Auto-generated invoice number (MAX(Inv_No)+1)

Save calls stored proc to persist header + details

Instructions to create the project and wire things up.

1) SQL: Create database, tables and sample data

Run this script in SQL Server Management Studio (SSMS). It will create database TESTANUMONJACOB, tables and some sample rows.

-- 1. Create database
IF DB_ID('TESTANUMONJACOB') IS NULL
BEGIN
    CREATE DATABASE TESTANUMONJACOB;
END
GO

USE TESTANUMONJACOB;
GO

-- 2. Create tables
-- Customer
IF OBJECT_ID('dbo.Customer') IS NOT NULL DROP TABLE dbo.Customer;
CREATE TABLE dbo.Customer(
    Cust_Code   VARCHAR(6) PRIMARY KEY,
    Cust_Name   VARCHAR(50) NOT NULL,
    Cust_Add1   VARCHAR(50) NULL,
    Cust_Add2   VARCHAR(50) NULL,
    Cust_Phone  VARCHAR(30) NULL
);

-- Product
IF OBJECT_ID('dbo.Product') IS NOT NULL DROP TABLE dbo.Product;
CREATE TABLE dbo.Product(
    Prod_Code   VARCHAR(6) PRIMARY KEY,
    Prod_Desc   VARCHAR(50) NOT NULL,
    Prod_Price  DECIMAL(15,3) NOT NULL CONSTRAINT DF_Product_Prod_Price DEFAULT (0)
);

-- Inv_Header
IF OBJECT_ID('dbo.Inv_Header') IS NOT NULL DROP TABLE dbo.Inv_Header;
CREATE TABLE dbo.Inv_Header(
    Inv_No      INT PRIMARY KEY,
    Inv_Date    DATETIME NOT NULL,
    Inv_Cust    VARCHAR(6) NOT NULL,
    CONSTRAINT FK_InvHeader_Customer FOREIGN KEY(Inv_Cust) REFERENCES dbo.Customer(Cust_Code)
);

-- Inv_Detail
IF OBJECT_ID('dbo.Inv_Detail') IS NOT NULL DROP TABLE dbo.Inv_Detail;
CREATE TABLE dbo.Inv_Detail(
    Inv_No      INT NOT NULL,
    Inv_Lno     INT NOT NULL,
    Inv_Prod    VARCHAR(6) NOT NULL,
    Inv_Qty     DECIMAL(10,3) NOT NULL,
    Inv_Price   DECIMAL(15,3) NOT NULL,
    CONSTRAINT PK_InvDetail PRIMARY KEY (Inv_No, Inv_Lno),
    CONSTRAINT FK_InvDetail_Header FOREIGN KEY (Inv_No) REFERENCES dbo.Inv_Header(Inv_No),
    CONSTRAINT FK_InvDetail_Product FOREIGN KEY (Inv_Prod) REFERENCES dbo.Product(Prod_Code)
);
GO

-- 3. Insert sample data into Customer & Product
INSERT INTO dbo.Customer (Cust_Code, Cust_Name, Cust_Add1, Cust_Add2, Cust_Phone)
VALUES
('C001','Acme Traders','123 Main St','City A','050-1111111'),
('C002','Bright Supplies','45 Market Road','City B','050-2222222'),
('C003','Clearwater Co.','78 Lakeside','City C','050-3333333');

INSERT INTO dbo.Product (Prod_Code, Prod_Desc, Prod_Price)
VALUES
('P001','Pen 0.5mm',5.500),
('P002','Notebook A4',25.000),
('P003','Stapler',45.000),
('P004','Paper Pack',120.000);
GO

2) Stored Procedures

Create stored procs for insert/update/delete/get invoice and the Query for “customer with maximum sales in last 6 months.”

USE TESTANUMONJACOB;
GO

-- 2.1 Get next invoice number
IF OBJECT_ID('dbo.sp_GetNextInvNo') IS NOT NULL DROP PROCEDURE dbo.sp_GetNextInvNo;
GO
CREATE PROCEDURE dbo.sp_GetNextInvNo
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @next INT;
    SELECT @next = ISNULL(MAX(Inv_No), 0) + 1 FROM dbo.Inv_Header;
    SELECT @next AS NextInvNo;
END
GO

-- 2.2 Insert invoice (header + details) - expects a table-valued parameter for details (we'll create TVP type)
IF TYPE_ID(N'dbo.InvDetailType') IS NOT NULL DROP TYPE dbo.InvDetailType;
GO
CREATE TYPE dbo.InvDetailType AS TABLE(
    Inv_Lno INT,
    Inv_Prod VARCHAR(6),
    Inv_Qty DECIMAL(10,3),
    Inv_Price DECIMAL(15,3)
);
GO

IF OBJECT_ID('dbo.sp_InsertInvoice') IS NOT NULL DROP PROCEDURE dbo.sp_InsertInvoice;
GO
CREATE PROCEDURE dbo.sp_InsertInvoice
    @Inv_No INT,
    @Inv_Date DATETIME,
    @Inv_Cust VARCHAR(6),
    @Details dbo.InvDetailType READONLY
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

        INSERT INTO dbo.Inv_Header (Inv_No, Inv_Date, Inv_Cust)
        VALUES (@Inv_No, @Inv_Date, @Inv_Cust);

        INSERT INTO dbo.Inv_Detail (Inv_No, Inv_Lno, Inv_Prod, Inv_Qty, Inv_Price)
        SELECT @Inv_No, Inv_Lno, Inv_Prod, Inv_Qty, Inv_Price FROM @Details;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
GO

-- 2.3 Update invoice (header + replace details)
IF OBJECT_ID('dbo.sp_UpdateInvoice') IS NOT NULL DROP PROCEDURE dbo.sp_UpdateInvoice;
GO
CREATE PROCEDURE dbo.sp_UpdateInvoice
    @Inv_No INT,
    @Inv_Date DATETIME,
    @Inv_Cust VARCHAR(6),
    @Details dbo.InvDetailType READONLY
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

        UPDATE dbo.Inv_Header
        SET Inv_Date = @Inv_Date, Inv_Cust = @Inv_Cust
        WHERE Inv_No = @Inv_No;

        -- remove existing details and re-insert
        DELETE FROM dbo.Inv_Detail WHERE Inv_No = @Inv_No;

        INSERT INTO dbo.Inv_Detail (Inv_No, Inv_Lno, Inv_Prod, Inv_Qty, Inv_Price)
        SELECT @Inv_No, Inv_Lno, Inv_Prod, Inv_Qty, Inv_Price FROM @Details;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
GO

-- 2.4 Delete invoice
IF OBJECT_ID('dbo.sp_DeleteInvoice') IS NOT NULL DROP PROCEDURE dbo.sp_DeleteInvoice;
GO
CREATE PROCEDURE dbo.sp_DeleteInvoice
    @Inv_No INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        DELETE FROM dbo.Inv_Detail WHERE Inv_No = @Inv_No;
        DELETE FROM dbo.Inv_Header WHERE Inv_No = @Inv_No;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
GO

-- 2.5 Get invoice (header + details)
IF OBJECT_ID('dbo.sp_GetInvoice') IS NOT NULL DROP PROCEDURE dbo.sp_GetInvoice;
GO
CREATE PROCEDURE dbo.sp_GetInvoice
    @Inv_No INT
AS
BEGIN
    SET NOCOUNT ON;
    SELECT * FROM dbo.Inv_Header WHERE Inv_No = @Inv_No;
    SELECT Inv_Lno, Inv_Prod, Prod_Desc = p.Prod_Desc, Inv_Price, Inv_Qty, Amount = Inv_Qty * Inv_Price
    FROM dbo.Inv_Detail d
    JOIN dbo.Product p ON p.Prod_Code = d.Inv_Prod
    WHERE d.Inv_No = @Inv_No
    ORDER BY Inv_Lno;
END
GO

-- 2.6 Query: Name and Total Sales amount of customer with maximum sales during last 6 months
IF OBJECT_ID('dbo.sp_GetTopCustomerLast6Months') IS NOT NULL DROP PROCEDURE dbo.sp_GetTopCustomerLast6Months;
GO
CREATE PROCEDURE dbo.sp_GetTopCustomerLast6Months
AS
BEGIN
    SET NOCOUNT ON;
    ;WITH Sales AS (
        SELECT h.Inv_Cust,
               SUM(d.Inv_Qty * d.Inv_Price) AS TotalSales
        FROM dbo.Inv_Header h
        JOIN dbo.Inv_Detail d ON h.Inv_No = d.Inv_No
        WHERE h.Inv_Date >= DATEADD(MONTH, -6, CAST(GETDATE() AS DATE))
        GROUP BY h.Inv_Cust
    )
    SELECT TOP 1 c.Cust_Code, c.Cust_Name, s.TotalSales
    FROM Sales s
    JOIN dbo.Customer c ON c.Cust_Code = s.Inv_Cust
    ORDER BY s.TotalSales DESC;
END
GO


Note: sp_GetTopCustomerLast6Months returns the single customer with highest total sales amount in the last 6 months (ties arbitrarily return one). You can easily adapt to return multiple customers if needed.

3) VB.NET 2010 WinForms: Project & Form

Below is the main code for a single-form WinForms app (FormInvoice.vb). It uses ADO.NET SqlClient. It assumes you created a Windows Forms project named TESTANUMONJACOB on your Desktop (per requirement). Replace the YourConnectionString with actual connection string to your SQL Server instance and the database TESTANUMONJACOB.

Form design (high-level)

Place these controls on the form (names used in code below):

Labels/TextBoxes:

lblInvNo (Label showing Invoice No)

dtpInvDate (DateTimePicker)

cmbCustomer (ComboBox for customer codes)

txtCustomerName (TextBox, ReadOnly)

txtCustomerAddr (TextBox, ReadOnly, multiline)

DataGridView: dgvLines

Columns:

colLno (TextBox, ReadOnly) — Lno

colProd (DataGridViewComboBoxColumn) — Prod Code

colDesc (TextBox, ReadOnly) — Description

colUnitPrice (TextBox) — Unit Price

colQty (TextBox) — Qty

colAmount (TextBox, ReadOnly) — Amount

Label: lblTotalAmount (shows Amount Total)

Buttons:

btnNew, btnModify, btnDelete, btnSave, btnCancel, btnPrint, btnQuery, btnSearch (search invoice by number)

Design note: For professional look, use grouping and padding. Use DataGridView with alternating row colors and column autosize.

Code: FormInvoice.vb

Important: This is a single-file form code. In an actual project you'll have designer code generated separately; below focuses on core logic to implement your requirements.

Imports System.Data
Imports System.Data.SqlClient

Public Class FormInvoice
    ' Update this connection string to match your SQL Server instance
    Private Const ConnString As String = "Data Source=.\SQLEXPRESS;Initial Catalog=TESTANUMONJACOB;Integrated Security=True;"
    Private ReadOnly _conn As New SqlConnection(ConnString)

    Private _isNewMode As Boolean = False
    Private _isModifyMode As Boolean = False

    Private Sub FormInvoice_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            LoadCustomers()
            LoadProductsToGridCombo()
            SetModeIdle()
            ClearForm()
        Catch ex As Exception
            MessageBox.Show("Error on load: " & ex.Message)
        End Try
    End Sub

    ' ============ UI Mode helpers ==============
    Private Sub SetModeIdle()
        _isNewMode = False
        _isModifyMode = False
        btnSave.Visible = False
        btnCancel.Visible = False
        btnNew.Visible = True
        btnModify.Visible = True
        btnDelete.Visible = True
        btnPrint.Visible = True
        btnQuery.Visible = True
        btnSearch.Visible = True
        dgvLines.ReadOnly = True
    End Sub

    Private Sub SetModeEdit()
        btnSave.Visible = True
        btnCancel.Visible = True
        btnNew.Visible = False
        btnModify.Visible = False
        btnDelete.Visible = False
        btnPrint.Visible = False
        btnQuery.Visible = False
        btnSearch.Visible = False
        dgvLines.ReadOnly = False
    End Sub

    ' ============ Load Customers & Products ==========
    Private Sub LoadCustomers()
        Using cmd As New SqlCommand("SELECT Cust_Code, Cust_Name, Cust_Add1 + ' ' + Cust_Add2 AS FullAddr FROM dbo.Customer ORDER BY Cust_Code", _conn)
            Dim dt As New DataTable()
            Using da As New SqlDataAdapter(cmd)
                da.Fill(dt)
            End Using
            cmbCustomer.DisplayMember = "Cust_Code"
            cmbCustomer.ValueMember = "Cust_Code"
            cmbCustomer.DataSource = dt
            cmbCustomer.SelectedIndex = -1
        End Using
    End Sub

    Private Sub LoadProductsToGridCombo()
        ' Fill a DataTable to be used as the items for the DataGridViewComboBoxColumn
        Dim dtProducts As New DataTable()
        Using cmd As New SqlCommand("SELECT Prod_Code, Prod_Desc, Prod_Price FROM dbo.Product ORDER BY Prod_Code", _conn)
            Using da As New SqlDataAdapter(cmd)
                da.Fill(dtProducts)
            End Using
        End Using

        ' Assuming the DataGridViewComboBoxColumn name is 'colProd'
        Dim cbCol As DataGridViewComboBoxColumn = TryCast(dgvLines.Columns("colProd"), DataGridViewComboBoxColumn)
        If cbCol IsNot Nothing Then
            cbCol.DataSource = dtProducts
            cbCol.DisplayMember = "Prod_Code"
            cbCol.ValueMember = "Prod_Code"
        End If
    End Sub

    ' When customer selection changes, fetch name/address and populate
    Private Sub cmbCustomer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCustomer.SelectedIndexChanged
        If cmbCustomer.SelectedIndex = -1 Then
            txtCustomerName.Text = ""
            txtCustomerAddr.Text = ""
            Return
        End If
        Dim code As String = cmbCustomer.SelectedValue.ToString()
        Using cmd As New SqlCommand("SELECT Cust_Name, Cust_Add1, Cust_Add2 FROM dbo.Customer WHERE Cust_Code = @c", _conn)
            cmd.Parameters.AddWithValue("@c", code)
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Using rdr As SqlDataReader = cmd.ExecuteReader()
                If rdr.Read() Then
                    txtCustomerName.Text = rdr("Cust_Name").ToString()
                    txtCustomerAddr.Text = (rdr("Cust_Add1").ToString() & " " & rdr("Cust_Add2").ToString()).Trim()
                End If
                rdr.Close()
            End Using
        End Using
    End Sub

    ' ============ New / Modify / Delete / Save / Cancel =================
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        _isNewMode = True
        _isModifyMode = False
        ClearForm()
        SetModeEdit()
        GetNextInvoiceNo()
        dtpInvDate.Value = DateTime.Now.Date
        dgvLines.Rows.Clear()
        AddEmptyLinesDefault()
    End Sub

    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Dim invNoStr As String = InputBox("Enter Invoice No to modify:", "Modify Invoice")
        If String.IsNullOrWhiteSpace(invNoStr) Then Return
        Dim invNo As Integer
        If Not Integer.TryParse(invNoStr, invNo) Then
            MessageBox.Show("Invalid invoice number.")
            Return
        End If
        If Not LoadInvoice(invNo) Then
            MessageBox.Show("Invoice not found.")
            Return
        End If
        _isModifyMode = True
        _isNewMode = False
        SetModeEdit()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim invNoStr As String = InputBox("Enter Invoice No to delete:", "Delete Invoice")
        If String.IsNullOrWhiteSpace(invNoStr) Then Return
        Dim invNo As Integer
        If Not Integer.TryParse(invNoStr, invNo) Then
            MessageBox.Show("Invalid invoice number.")
            Return
        End If
        If MessageBox.Show("Delete invoice " & invNo.ToString() & " ?", "Confirm", MessageBoxButtons.YesNo) <> DialogResult.Yes Then
            Return
        End If
        Using cmd As New SqlCommand("dbo.sp_DeleteInvoice", _conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@Inv_No", invNo)
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            cmd.ExecuteNonQuery()
        End Using
        MessageBox.Show("Deleted.")
        ClearForm()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            If Not ValidateBeforeSave() Then Return

            Dim invNo As Integer = Integer.Parse(lblInvNo.Text)
            Dim invDate As DateTime = dtpInvDate.Value.Date
            Dim custCode As String = If(cmbCustomer.SelectedIndex >= 0, cmbCustomer.SelectedValue.ToString(), "")

            ' prepare details table (TVP)
            Dim dtDetails As New DataTable()
            dtDetails.Columns.Add("Inv_Lno", GetType(Integer))
            dtDetails.Columns.Add("Inv_Prod", GetType(String))
            dtDetails.Columns.Add("Inv_Qty", GetType(Decimal))
            dtDetails.Columns.Add("Inv_Price", GetType(Decimal))

            For Each row As DataGridViewRow In dgvLines.Rows
                If row.IsNewRow Then Continue For
                Dim prod As String = If(row.Cells("colProd").Value, "").ToString()
                If String.IsNullOrWhiteSpace(prod) Then Continue For
                Dim lno As Integer = Convert.ToInt32(row.Cells("colLno").Value)
                Dim qty As Decimal = Convert.ToDecimal(If(row.Cells("colQty").Value, 0))
                Dim upr As Decimal = Convert.ToDecimal(If(row.Cells("colUnitPrice").Value, 0))
                Dim dr As DataRow = dtDetails.NewRow()
                dr("Inv_Lno") = lno
                dr("Inv_Prod") = prod
                dr("Inv_Qty") = qty
                dr("Inv_Price") = upr
                dtDetails.Rows.Add(dr)
            Next

            If dtDetails.Rows.Count = 0 Then
                MessageBox.Show("Invoice must have at least one product line.")
                Return
            End If

            If _isNewMode Then
                Using cmd As New SqlCommand("dbo.sp_InsertInvoice", _conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.AddWithValue("@Inv_No", invNo)
                    cmd.Parameters.AddWithValue("@Inv_Date", invDate)
                    cmd.Parameters.AddWithValue("@Inv_Cust", custCode)
                    Dim tvp As New SqlParameter("@Details", SqlDbType.Structured)
                    tvp.TypeName = "dbo.InvDetailType"
                    tvp.Value = dtDetails
                    cmd.Parameters.Add(tvp)
                    If _conn.State = ConnectionState.Closed Then _conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
                MessageBox.Show("Saved new invoice " & invNo.ToString())
            ElseIf _isModifyMode Then
                Using cmd As New SqlCommand("dbo.sp_UpdateInvoice", _conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.AddWithValue("@Inv_No", invNo)
                    cmd.Parameters.AddWithValue("@Inv_Date", invDate)
                    cmd.Parameters.AddWithValue("@Inv_Cust", custCode)
                    Dim tvp As New SqlParameter("@Details", SqlDbType.Structured)
                    tvp.TypeName = "dbo.InvDetailType"
                    tvp.Value = dtDetails
                    cmd.Parameters.Add(tvp)
                    If _conn.State = ConnectionState.Closed Then _conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
                MessageBox.Show("Updated invoice " & invNo.ToString())
            End If

            SetModeIdle()
            _isNewMode = False
            _isModifyMode = False

        Catch ex As Exception
            MessageBox.Show("Error saving: " & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        SetModeIdle()
        _isNewMode = False
        _isModifyMode = False
        ClearForm()
    End Sub

    ' ============ Search / Load invoice =================
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim invNoStr As String = InputBox("Enter Invoice No to view:", "Search Invoice")
        If String.IsNullOrWhiteSpace(invNoStr) Then Return
        Dim invNo As Integer
        If Not Integer.TryParse(invNoStr, invNo) Then
            MessageBox.Show("Invalid invoice number.")
            Return
        End If
        If Not LoadInvoice(invNo) Then
            MessageBox.Show("Invoice not found.")
        End If
    End Sub

    Private Function LoadInvoice(invNo As Integer) As Boolean
        Try
            Using cmd As New SqlCommand("dbo.sp_GetInvoice", _conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@Inv_No", invNo)
                If _conn.State = ConnectionState.Closed Then _conn.Open()
                Using rdr As SqlDataReader = cmd.ExecuteReader()
                    If Not rdr.HasRows Then
                        Return False
                    End If
                    ' header
                    If rdr.Read() Then
                        lblInvNo.Text = rdr("Inv_No").ToString()
                        dtpInvDate.Value = Convert.ToDateTime(rdr("Inv_Date"))
                        Dim ccode As String = rdr("Inv_Cust").ToString()
                        If cmbCustomer.Items.Count > 0 Then
                            cmbCustomer.SelectedValue = ccode
                        End If
                    End If

                    ' move to details resultset
                    If rdr.NextResult() Then
                        dgvLines.Rows.Clear()
                        Dim lno As Integer = 1
                        While rdr.Read()
                            Dim prod As String = rdr("Inv_Prod").ToString()
                            Dim desc As String = rdr("Prod_Desc").ToString()
                            Dim upr As Decimal = Convert.ToDecimal(rdr("Inv_Price"))
                            Dim qty As Decimal = Convert.ToDecimal(rdr("Inv_Qty"))
                            Dim amount As Decimal = qty * upr

                            Dim idx As Integer = dgvLines.Rows.Add()
                            Dim row As DataGridViewRow = dgvLines.Rows(idx)
                            row.Cells("colLno").Value = lno
                            row.Cells("colProd").Value = prod
                            row.Cells("colDesc").Value = desc
                            row.Cells("colUnitPrice").Value = upr
                            row.Cells("colQty").Value = qty
                            row.Cells("colAmount").Value = amount
                            lno += 1
                        End While
                    End If
                End Using
            End Using
            CalculateTotal()
            Return True
        Catch ex As Exception
            MessageBox.Show("Error loading invoice: " & ex.Message)
            Return False
        End Try
    End Function

    ' ============ Helpers: NextInvNo, Clear, Add lines, Validation =============
    Private Sub GetNextInvoiceNo()
        Using cmd As New SqlCommand("dbo.sp_GetNextInvNo", _conn)
            cmd.CommandType = CommandType.StoredProcedure
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            Dim dt As New DataTable()
            Using da As New SqlDataAdapter(cmd)
                da.Fill(dt)
            End Using
            If dt.Rows.Count > 0 Then
                lblInvNo.Text = dt.Rows(0)("NextInvNo").ToString()
            Else
                lblInvNo.Text = "1"
            End If
        End Using
    End Sub

    Private Sub ClearForm()
        lblInvNo.Text = ""
        dtpInvDate.Value = DateTime.Now.Date
        cmbCustomer.SelectedIndex = -1
        txtCustomerName.Text = ""
        txtCustomerAddr.Text = ""
        dgvLines.Rows.Clear()
        lblTotalAmount.Text = "0.000"
    End Sub

    Private Sub AddEmptyLinesDefault()
        dgvLines.Rows.Clear()
        For i As Integer = 1 To 8
            Dim idx As Integer = dgvLines.Rows.Add()
            Dim row As DataGridViewRow = dgvLines.Rows(idx)
            row.Cells("colLno").Value = i
            row.Cells("colProd").Value = Nothing
            row.Cells("colDesc").Value = ""
            row.Cells("colUnitPrice").Value = 0
            row.Cells("colQty").Value = 0
            row.Cells("colAmount").Value = 0
        Next
    End Sub

    Private Function ValidateBeforeSave() As Boolean
        ' date not greater than system date
        If dtpInvDate.Value.Date > DateTime.Now.Date Then
            MessageBox.Show("Invoice date cannot be greater than system date.")
            Return False
        End If
        If cmbCustomer.SelectedIndex = -1 Then
            MessageBox.Show("Select a customer.")
            Return False
        End If
        ' check qty > 0 for at least one line
        Dim anyLine As Boolean = False
        For Each row As DataGridViewRow In dgvLines.Rows
            If row.IsNewRow Then Continue For
            Dim prod As String = If(row.Cells("colProd").Value, "").ToString()
            If String.IsNullOrWhiteSpace(prod) Then Continue For
            Dim qty As Decimal = Convert.ToDecimal(If(row.Cells("colQty").Value, 0))
            If qty <= 0 Then
                MessageBox.Show("Qty must be greater than 0 for product " & prod)
                Return False
            End If
            anyLine = True
        Next
        If Not anyLine Then
            MessageBox.Show("Provide at least one product line with qty > 0.")
            Return False
        End If
        Return True
    End Function

    ' ============ DataGridView events: update description, price, amount, total ============
    Private Sub dgvLines_CellValueNeeded(sender As Object, e As DataGridViewCellValueEventArgs) Handles dgvLines.CellValueNeeded
        ' not required if using bound mode; left blank
    End Sub

    Private Sub dgvLines_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvLines.CellValueChanged
        If e.RowIndex < 0 Then Return
        Dim row As DataGridViewRow = dgvLines.Rows(e.RowIndex)

        Try
            If dgvLines.Columns(e.ColumnIndex).Name = "colProd" Then
                ' populate description and unit price from product table
                Dim prodCode As String = If(row.Cells("colProd").Value, "").ToString()
                If String.IsNullOrWhiteSpace(prodCode) Then
                    row.Cells("colDesc").Value = ""
                    row.Cells("colUnitPrice").Value = 0
                    row.Cells("colQty").Value = 0
                    row.Cells("colAmount").Value = 0
                Else
                    Using cmd As New SqlCommand("SELECT Prod_Desc, Prod_Price FROM dbo.Product WHERE Prod_Code = @p", _conn)
                        cmd.Parameters.AddWithValue("@p", prodCode)
                        If _conn.State = ConnectionState.Closed Then _conn.Open()
                        Using rdr As SqlDataReader = cmd.ExecuteReader()
                            If rdr.Read() Then
                                row.Cells("colDesc").Value = rdr("Prod_Desc").ToString()
                                row.Cells("colUnitPrice").Value = Convert.ToDecimal(rdr("Prod_Price"))
                            End If
                        End Using
                    End Using
                End If
            End If

            If dgvLines.Columns(e.ColumnIndex).Name = "colQty" Or dgvLines.Columns(e.ColumnIndex).Name = "colUnitPrice" Or dgvLines.Columns(e.ColumnIndex).Name = "colProd" Then
                ' recalc amount
                Dim qty As Decimal = Convert.ToDecimal(If(row.Cells("colQty").Value, 0))
                Dim upr As Decimal = Convert.ToDecimal(If(row.Cells("colUnitPrice").Value, 0))
                If qty < 0 Then
                    MessageBox.Show("Qty cannot be negative.")
                    row.Cells("colQty").Value = 0
                    qty = 0
                End If
                Dim amt As Decimal = qty * upr
                row.Cells("colAmount").Value = amt
                CalculateTotal()
            End If
        Catch ex As Exception
            ' ignore cell change errors for now
        End Try
    End Sub

    Private Sub CalculateTotal()
        Dim total As Decimal = 0D
        For Each row As DataGridViewRow In dgvLines.Rows
            If row.IsNewRow Then Continue For
            Dim amt As Decimal = Convert.ToDecimal(If(row.Cells("colAmount").Value, 0))
            total += amt
        Next
        lblTotalAmount.Text = total.ToString("F3")
    End Sub

    ' When user edits Qty or UnitPrice cell, commit to trigger CellValueChanged
    Private Sub dgvLines_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvLines.CurrentCellDirtyStateChanged
        If dgvLines.IsCurrentCellDirty Then
            dgvLines.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    ' ============ Query button: show top customer last 6 months via message box ============
    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Try
            Using cmd As New SqlCommand("dbo.sp_GetTopCustomerLast6Months", _conn)
                cmd.CommandType = CommandType.StoredProcedure
                If _conn.State = ConnectionState.Closed Then _conn.Open()
                Using rdr As SqlDataReader = cmd.ExecuteReader()
                    If rdr.Read() Then
                        Dim name As String = rdr("Cust_Name").ToString()
                        Dim total As Decimal = Convert.ToDecimal(rdr("TotalSales"))
                        MessageBox.Show(String.Format("Top customer in last 6 months: {0}   Total Sales: {1:F3}", name, total))
                    Else
                        MessageBox.Show("No sales in last 6 months.")
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error executing query: " & ex.Message)
        End Try
    End Sub

End Class

Notes about the VB code

ComboBox for customer: cmbCustomer is loaded only from Customer table — no front-end create allowed.

Product selection: colProd is a DataGridViewComboBoxColumn bound to Product table; user cannot add new product codes from front-end.

Qty rules: If user enters negative or zero, the validation will prevent save (and shows message).

Date rule: dtpInvDate.Value.Date is validated to be <= DateTime.Now.Date.

Auto-invoice number: sp_GetNextInvNo is called when New is clicked; for concurrency you could implement sequences or a DB identity table (this MAX+1 approach is simpler for testing).

Save logic: Prepares a DataTable compatible with TVP dbo.InvDetailType and calls sp_InsertInvoice or sp_UpdateInvoice.

DataGridView: colLno auto-populated at addition time; colDesc and colAmount are read-only; colProd, colUnitPrice, colQty editable.

4) How to create the project (step-by-step)

Open Visual Studio 2010 → File → New → Project → Visual Basic → Windows → Windows Forms Application. Name it: TESTANUMONJACOB. Location: Desktop, as requested.

Add a form named FormInvoice (or use default Form1). Design it as per the controls listed earlier. Add a DataGridView and create columns with exact Name values used in code:

colLno — DataGridViewTextBoxColumn — ReadOnly = True

colProd — DataGridViewComboBoxColumn — Name = "colProd"

colDesc — DataGridViewTextBoxColumn — ReadOnly = True

colUnitPrice — DataGridViewTextBoxColumn — editable

colQty — DataGridViewTextBoxColumn — editable

colAmount — DataGridViewTextBoxColumn — ReadOnly = True

Place the other controls (labels, textboxes, combo box, DateTimePicker, buttons). Set txtCustomerName.ReadOnly = True, txtCustomerAddr.ReadOnly = True.

Add the VB code from above into the form code file (replace the default class). Update the connection string constant at top to point to your SQL Server instance. Example connection strings:

Integrated security (Windows auth):
"Data Source=.\SQLEXPRESS;Initial Catalog=TESTANUMONJACOB;Integrated Security=True;"

SQL auth:
"Data Source=SERVERNAME;Initial Catalog=TESTANUMONJACOB;User ID=sa;Password=yourpw;"

Build the solution. Run. Click New to create invoice, choose customer from combo, select products in each row, enter qty; Save persists to DB.

5) Additional tips & small improvements you can make

Concurrency-safe invoice numbers: Use a DB SEQUENCE (SQL Server 2012+) or an identity table to avoid race condition on MAX(Inv_No)+1. For 2016 SQL Server create a SEQUENCE and NEXT VALUE FOR.

Better customer/product selection UI: Add a small search dialog that filters customers/products with description and code. (Per spec, adding new customer/product from frontend must be restricted.)

Printing: When needed, you can create a report (RDLC or Crystal Reports) pulling data from sp_GetInvoice.

Validation UX: Highlight invalid cells instead of message boxes for smoother user experience.

UI polish: Use groupboxes, fonts, colors, and margin/padding to make it look professional. Consider 2-3 font sizes, bold headers, and right-aligned numeric columns.

6) The SQL query you requested (explicit)

You asked: "Create a SQL Query to display the Name and Total Sales invoice amount of the customer who has Maximum sales during last 6 months."

Here’s a standalone SQL query (same logic as sp_GetTopCustomerLast6Months stored proc):

USE TESTANUMONJACOB;
GO

WITH Sales AS (
    SELECT h.Inv_Cust,
           SUM(d.Inv_Qty * d.Inv_Price) AS TotalSales
    FROM dbo.Inv_Header h
    JOIN dbo.Inv_Detail d ON h.Inv_No = d.Inv_No
    WHERE h.Inv_Date >= DATEADD(MONTH, -6, CAST(GETDATE() AS DATE))
    GROUP BY h.Inv_Cust
)
SELECT TOP 1 c.Cust_Name, s.TotalSales
FROM Sales s
JOIN dbo.Customer c ON c.Cust_Code = s.Inv_Cust
ORDER BY s.TotalSales DESC;


When the user clicks the Query button in the UI I provided, the stored proc is called and the result is shown in a MessageBox.

7) What I delivered (summary)

Full SQL scripts to create DB TESTANUMONJACOB, tables, sample Customer/Product rows.

Stored procedures for Insert/Update/Delete/Get invoices and the last-6-months top-customer query.

A VB.NET Form implementation (code) implementing all required behaviors and UI flows:

Auto-generated invoice number

Search/fetch invoice by number

Date check (not > system date)

Customer/Product selection only from DB (no front-end create)

Product desc & price populated automatically

Grid rules: editable columns & read-only columns as requested

Qty validation (>0)

Automatic Amount calculation and Total update

Buttons with required visibility/behavior

Query button shows MessageBox result

8) Next steps / if you want me to expand

Tell me which of the following you want next (I can give full code / designer file / packaged instructions right here):

Add a search dialog for customers and products (modal grid).

Implement printing / RDLC report for the invoice.

Replace MAX(Inv_No)+1 with a SQL SEQUENCE for safe invoice numbers.

Provide a ready-to-paste Visual Studio Designer code (the .Designer.vb snippet for precise control placement).

Create a small demo dataset of invoices so you can test the Query output.

If you want the designer file or I should adapt the connection string for SQL Auth / LocalDB — tell me which and I’ll produce the exact snippet for you to paste into the project.

Step 1: SQL Server Database & Tables

Run these scripts in SQL Server Management Studio (SSMS):

-- 1. Create database
CREATE DATABASE TESTANUMONJACOB;
GO
USE TESTANUMONJACOB;
GO

-- 2. Customer table
CREATE TABLE Customer (
    Cust_Code VARCHAR(6) PRIMARY KEY,
    Cust_Name VARCHAR(50) NOT NULL,
    Cust_Add1 VARCHAR(50) NULL,
    Cust_Add2 VARCHAR(50) NULL,
    Cust_Phone VARCHAR(30) NULL
);

-- 3. Product table
CREATE TABLE Product (
    Prod_Code VARCHAR(6) PRIMARY KEY,
    Prod_Desc VARCHAR(50) NOT NULL,
    Prod_Price DECIMAL(15,3) NULL DEFAULT 0
);

-- 4. Invoice Header
CREATE TABLE Inv_Header (
    Inv_No INT PRIMARY KEY,
    Inv_Date DATETIME NOT NULL,
    Inv_Cust VARCHAR(6) NOT NULL
        FOREIGN KEY REFERENCES Customer(Cust_Code)
);

-- 5. Invoice Detail
CREATE TABLE Inv_Detail (
    Inv_No INT NOT NULL,
    Inv_Lno INT NOT NULL,
    Inv_Prod VARCHAR(6) NOT NULL
        FOREIGN KEY REFERENCES Product(Prod_Code),
    Inv_Qty DECIMAL(10,3) NOT NULL,
    Inv_Price DECIMAL(15,3) NOT NULL,
    PRIMARY KEY (Inv_No, Inv_Lno)
);

-- 6. Sample data
INSERT INTO Customer VALUES
('C001','Jacob Traders','Dubai','UAE','971-123456'),
('C002','ABC Stores','Sharjah','UAE','971-654321');

INSERT INTO Product VALUES
('P001','Laptop',2500.000),
('P002','Mouse',50.000),
('P003','Keyboard',150.000),
('P004','Monitor',600.000);

Step 2: Stored Procedures
-- Insert Invoice Header + Details
CREATE PROCEDURE sp_InsertInvoice
    @Inv_No INT,
    @Inv_Date DATETIME,
    @Inv_Cust VARCHAR(6)
AS
BEGIN
    INSERT INTO Inv_Header(Inv_No, Inv_Date, Inv_Cust)
    VALUES(@Inv_No, @Inv_Date, @Inv_Cust);
END
GO

CREATE PROCEDURE sp_InsertInvoiceDetail
    @Inv_No INT,
    @Inv_Lno INT,
    @Inv_Prod VARCHAR(6),
    @Inv_Qty DECIMAL(10,3),
    @Inv_Price DECIMAL(15,3)
AS
BEGIN
    INSERT INTO Inv_Detail(Inv_No, Inv_Lno, Inv_Prod, Inv_Qty, Inv_Price)
    VALUES(@Inv_No, @Inv_Lno, @Inv_Prod, @Inv_Qty, @Inv_Price);
END
GO

-- Update Invoice Header
CREATE PROCEDURE sp_UpdateInvoice
    @Inv_No INT,
    @Inv_Date DATETIME,
    @Inv_Cust VARCHAR(6)
AS
BEGIN
    UPDATE Inv_Header
    SET Inv_Date=@Inv_Date, Inv_Cust=@Inv_Cust
    WHERE Inv_No=@Inv_No;
END
GO

-- Delete Invoice
CREATE PROCEDURE sp_DeleteInvoice
    @Inv_No INT
AS
BEGIN
    DELETE FROM Inv_Detail WHERE Inv_No=@Inv_No;
    DELETE FROM Inv_Header WHERE Inv_No=@Inv_No;
END
GO

-- Get Invoice Header + Details
CREATE PROCEDURE sp_GetInvoice
    @Inv_No INT
AS
BEGIN
    SELECT * FROM Inv_Header WHERE Inv_No=@Inv_No;
    SELECT d.Inv_Lno, d.Inv_Prod, p.Prod_Desc, d.Inv_Price, d.Inv_Qty,
           (d.Inv_Price*d.Inv_Qty) AS Amount
    FROM Inv_Detail d
    JOIN Product p ON d.Inv_Prod=p.Prod_Code
    WHERE Inv_No=@Inv_No;
END
GO

-- Query: Customer with max sales in last 6 months
CREATE PROCEDURE sp_MaxSalesCustomer
AS
BEGIN
    SELECT TOP 1 c.Cust_Name,
           SUM(d.Inv_Qty * d.Inv_Price) AS TotalSales
    FROM Inv_Header h
    JOIN Inv_Detail d ON h.Inv_No=d.Inv_No
    JOIN Customer c ON h.Inv_Cust=c.Cust_Code
    WHERE h.Inv_Date >= DATEADD(MONTH,-6,GETDATE())
    GROUP BY c.Cust_Name
    ORDER BY TotalSales DESC;
END
GO

Step 3: VB.NET WinForms (2010)

Open Visual Studio 2010 → New Project → Windows Forms App → Name it TESTANUMONJACOB.

Add Controls to your form (Form1):

Labels + TextBoxes: InvoiceNo, Date, CustomerCode (ComboBox), Customer Name (TextBox, ReadOnly), Address (TextBox, ReadOnly), Total Amount.

DataGridView (dgvDetails) with columns:

Lno (ReadOnly, Auto Number)

Product Code (ComboBox column, bound to Product table)

Description (ReadOnly)

Unit Price (Editable)

Qty (Editable)

Amount (ReadOnly)

Buttons: btnNew, btnModify, btnDelete, btnSave, btnCancel, btnPrint, btnQuery.

Add a connection string in App.config:

<configuration>
  <connectionStrings>
    <add name="ConnStr"
         connectionString="Data Source=YOURSERVER;Initial Catalog=TESTANUMONJACOB;Integrated Security=True"/>
  </connectionStrings>
</configuration>


Form Code (simplified sample):

Imports System.Data.SqlClient

Public Class Form1
    Dim con As New SqlConnection(Configuration.ConfigurationManager.ConnectionStrings("ConnStr").ConnectionString)
    Dim dtProducts As DataTable
    Dim invNo As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCustomers()
        LoadProducts()
        PrepareNewInvoice()
    End Sub

    Private Sub LoadCustomers()
        Dim da As New SqlDataAdapter("SELECT Cust_Code,Cust_Name FROM Customer", con)
        Dim dt As New DataTable()
        da.Fill(dt)
        cboCustomer.DataSource = dt
        cboCustomer.DisplayMember = "Cust_Name"
        cboCustomer.ValueMember = "Cust_Code"
    End Sub

    Private Sub LoadProducts()
        Dim da As New SqlDataAdapter("SELECT Prod_Code,Prod_Desc,Prod_Price FROM Product", con)
        dtProducts = New DataTable()
        da.Fill(dtProducts)

        Dim col As New DataGridViewComboBoxColumn()
        col.HeaderText = "Product"
        col.DataPropertyName = "Inv_Prod"
        col.DataSource = dtProducts
        col.DisplayMember = "Prod_Desc"
        col.ValueMember = "Prod_Code"
        dgvDetails.Columns.Add("Lno", "Lno")
        dgvDetails.Columns.Add(col)
        dgvDetails.Columns.Add("Description", "Description")
        dgvDetails.Columns.Add("UnitPrice", "Unit Price")
        dgvDetails.Columns.Add("Qty", "Qty")
        dgvDetails.Columns.Add("Amount", "Amount")
    End Sub

    Private Sub PrepareNewInvoice()
        Dim cmd As New SqlCommand("SELECT ISNULL(MAX(Inv_No),0)+1 FROM Inv_Header", con)
        con.Open()
        invNo = CInt(cmd.ExecuteScalar())
        con.Close()
        txtInvNo.Text = invNo.ToString()
        txtDate.Text = DateTime.Now.ToShortDateString()
    End Sub

    Private Sub dgvDetails_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDetails.CellValueChanged
        If e.ColumnIndex = 2 Or e.ColumnIndex = 3 Then
            Dim qty As Decimal = Convert.ToDecimal(dgvDetails.Rows(e.RowIndex).Cells("Qty").Value)
            Dim price As Decimal = Convert.ToDecimal(dgvDetails.Rows(e.RowIndex).Cells("UnitPrice").Value)
            dgvDetails.Rows(e.RowIndex).Cells("Amount").Value = qty * price
            UpdateTotal()
        End If
    End Sub

    Private Sub UpdateTotal()
        Dim sum As Decimal = 0
        For Each r As DataGridViewRow In dgvDetails.Rows
            If Not r.IsNewRow Then
                sum += Convert.ToDecimal(r.Cells("Amount").Value)
            End If
        Next
        txtTotal.Text = sum.ToString("N2")
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        con.Open()
        Dim cmd As New SqlCommand("sp_InsertInvoice", con)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Inv_No", invNo)
        cmd.Parameters.AddWithValue("@Inv_Date", txtDate.Text)
        cmd.Parameters.AddWithValue("@Inv_Cust", cboCustomer.SelectedValue)
        cmd.ExecuteNonQuery()

        Dim lno As Integer = 1
        For Each r As DataGridViewRow In dgvDetails.Rows
            If Not r.IsNewRow Then
                Dim cmdDet As New SqlCommand("sp_InsertInvoiceDetail", con)
                cmdDet.CommandType = CommandType.StoredProcedure
                cmdDet.Parameters.AddWithValue("@Inv_No", invNo)
                cmdDet.Parameters.AddWithValue("@Inv_Lno", lno)
                cmdDet.Parameters.AddWithValue("@Inv_Prod", r.Cells("Product").Value)
                cmdDet.Parameters.AddWithValue("@Inv_Qty", r.Cells("Qty").Value)
                cmdDet.Parameters.AddWithValue("@Inv_Price", r.Cells("UnitPrice").Value)
                cmdDet.ExecuteNonQuery()
                lno += 1
            End If
        Next
        con.Close()
        MessageBox.Show("Invoice Saved")
    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        con.Open()
        Dim cmd As New SqlCommand("sp_MaxSalesCustomer", con)
        cmd.CommandType = CommandType.StoredProcedure
        Dim rdr As SqlDataReader = cmd.ExecuteReader()
        If rdr.Read() Then
            MessageBox.Show("Customer: " & rdr("Cust_Name").ToString() &
                            vbCrLf & "Sales: " & rdr("TotalSales").ToString())
        End If
        con.Close()
    End Sub
End Class

Step 4: How to Wire It Up

Run the SQL scripts first.

In Visual Studio, create the project TESTANUMONJACOB.

Add the form controls (labels, textboxes, DataGridView, buttons).

Paste the VB.NET code.

Adjust connection string (YOURSERVER).

Run → New → Fill details → Save.

Try Query → shows customer with max sales in last 6 months.
