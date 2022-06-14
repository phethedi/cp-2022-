Public Class 9712245588083

    Public IDKey As Long

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Dim dbconnect As New OleDb.OleDbConnection
            Dim command As New OleDb.OleDbCommand
            Dim Child As New Data.DataTable

            dbconnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Student\Desktop\Database61.mdb ;user id = ; password = ;"
            dbconnect.Open()
            command.Connection = dbconnect
            command.CommandText = "SELECT * FROM ContactDetails"
            Child.Load(command.ExecuteReader())

            dbconnect.Close()

            DataGridView1.DataSource = Child

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub

    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Try
            Dim dbconnect As New OleDb.OleDbConnection
            Dim command As New OleDb.OleDbCommand
            Dim Child As New Data.DataTable

            dbconnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Student\Desktop\Database61.mdb ;user id = ; password = ;"
            dbconnect.Open()
            command.Connection = dbconnect
            command.CommandText = "INSERT INTO ContactDetails (Name, Surname, CellNumber ) VALUES (@Name, @Surname, @CellNumber)"

            Dim paramName = New OleDb.OleDbParameter("@Name", txtName.Text)
            Dim paramSurname = New OleDb.OleDbParameter("@Surname", txtSurname.Text)
            Dim paramCellNumber = New OleDb.OleDbParameter("@CellNumber", txtCellnumber.Text)

            command.Parameters.Add(paramName)
            command.Parameters.Add(paramSurname)
            command.Parameters.Add(paramCellNumber)

            command.ExecuteNonQuery()

            dbconnect.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString())

        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try
            Dim dbconnect As New OleDb.OleDbConnection
            Dim command As New OleDb.OleDbCommand
            Dim Child As New Data.DataTable

            dbconnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Student\Desktop\Database61.mdb ;user id = ; password = ;"
            dbconnect.Open()
            command.Connection = dbconnect
            command.CommandText = "UPDATE ContactDetails SET Name=@Name, Surname=@Surname, CellNumber=@CellNumber WHERE ID = @id"

            Dim paramName = New OleDb.OleDbParameter("@Name", txtName.Text)
            Dim paramSurname = New OleDb.OleDbParameter("@Surname", txtSurname.Text)
            Dim paramCellNumber = New OleDb.OleDbParameter("@CellNumber", txtCellnumber.Text)

            command.Parameters.Add(paramName)
            command.Parameters.Add(paramSurname)
            command.Parameters.Add(paramCellNumber)

            command.ExecuteNonQuery()

            dbconnect.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Try
            Dim dbconnect As New OleDb.OleDbConnection
            Dim command As New OleDb.OleDbCommand
            Dim Child As New Data.DataTable

            dbconnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Student\Desktop\Database61.mdb ;user id = ; password = ;"
            dbconnect.Open()
            command.Connection = dbconnect
            command.CommandText = "SELECT * FROM ContactDetails WHERE name=@name"

            Dim paramName = New OleDb.OleDbParameter("@Name", txtName.Text)

            command.Parameters.Add(paramName)
            Child.Load(command.ExecuteReader())

            dbconnect.Close()

            If Child.Rows.Count = 0 Then
                MessageBox.Show("No Contact found, Re-enter")
            ElseIf Child.Rows.Count = 1 Then
                txtName.Text = Child.Rows(0)("Name")
                txtSurname.Text = Child.Rows(0)("Surname")
                txtCellnumber.Text = Child.Rows(0)("cellNumber")
            End If

        Catch ex As Exception
            MsgBox(ex.ToString())

        End Try
    End Sub

    Private Sub btnViewAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewAll.Click
        Try
            Dim dbconnect As New OleDb.OleDbConnection
            Dim command As New OleDb.OleDbCommand
            Dim Child As New Data.DataTable

            dbconnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Student\Desktop\Database61.mdb ;user id = ; password = ;"
            dbconnect.Open()
            command.Connection = dbconnect
            command.CommandText = "SELECT * FROM ContactDetails"
            Child.Load(command.ExecuteReader())

            dbconnect.Close()

            DataGridView1.DataSource = Child

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            Dim dbconnect As New OleDb.OleDbConnection
            Dim command As New OleDb.OleDbCommand
            Dim Child As New Data.DataTable

            dbconnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Student\Desktop\Database61.mdb ;user id = ; password = ;"
            dbconnect.Open()
            command.Connection = dbconnect
            command.CommandText = "DELETE FROM ContactDetails WHERE Name=@Name, Surname=@Surname, CellNumber=@CellNumber"

            Dim paramName = New OleDb.OleDbParameter("@Name", DataGridView1.CurrentRow.Cells("Name").Value)

            command.Parameters.Add(paramName)
            command.ExecuteNonQuery()

            dbconnect.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub
    Private Sub updGrid()
        Try
            'Create the conection object
            Dim dbconnect As New OleDb.OleDbConnection

            'Create the command object
            Dim command As New OleDb.OleDbCommand

            'Create a table object to hold the selected data for the grid
            Dim Child As New Data.DataTable

            'set the connection string
            dbconnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= C:\Users\Student\Desktop\Database61.mdb ;user id = ; password = ;"

            'Call the open method
            dbconnect.Open()

            'initialise the connection property to the command object
            command.Connection = dbconnect

            'Create the SQL query to retrieve all the records from contactdetails table
            command.CommandText = "SELECT * FROM ContactDetails"

            'Put the retrieved data into the table
            Child.Load(command.ExecuteReader())

            'Close the db connection
            dbconnect.Close()

            'Link the retrieved data table to the datagridview
            DataGridView1.DataSource = Child


        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
    End Sub

End Class

