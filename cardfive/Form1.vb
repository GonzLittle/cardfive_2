Imports System.Data.OleDb
Imports MetroFramework

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'DateTimePicker1.Format = DateTimePickerFormat.Custom
        'DateTimePicker1.CustomFormat = "MM/dd/yyyy"

        '--------for id---------
        MetroTextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
        MetroTextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
        Dim DataCollection As New AutoCompleteStringCollection()
        getData(DataCollection)
        MetroTextBox1.AutoCompleteCustomSource = DataCollection
        '--------end for id---------


        '--------for position---------
        MetroTextBox13.AutoCompleteMode = AutoCompleteMode.Suggest
        MetroTextBox13.AutoCompleteSource = AutoCompleteSource.CustomSource
        Dim _position As New AutoCompleteStringCollection()
        getposition(_position)
        MetroTextBox13.AutoCompleteCustomSource = _position
        '--------end for posotion---------


        '--------for office---------
        MetroTextBox14.AutoCompleteMode = AutoCompleteMode.Suggest
        MetroTextBox14.AutoCompleteSource = AutoCompleteSource.CustomSource
        Dim _department As New AutoCompleteStringCollection()
        getdepartment(_department)
        MetroTextBox14.AutoCompleteCustomSource = _department
        '--------end for office---------


        MetroTextBox1.Focus()

    End Sub

    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

    End Sub

    Public Sub Loadpicture()
        Try
            Dim idno As String = MetroTextBox1.Text.ToString
            Dim folder As String = "\\pmis_server\pictures"
            Dim filename As String = System.IO.Path.Combine(folder, idno & ".jpg")
            'Dim folder1 As String = "C:\Users\User\Desktop\HRIS\Displaydata\Resources"
            'Dim filename1 As String = System.IO.Path.Combine(folder, "error image.jpg")
            If Not System.IO.File.Exists(filename) Then
                ' file does not exist
                PictureBox1.Image = Nothing
            Else
                PictureBox1.Image = Image.FromFile(filename)
                PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                PictureBox1.Image = Image.FromFile(filename)
                PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Sub loadsig()
        Try
            Dim idno As String = MetroTextBox1.Text.ToString
            Dim folder As String = "\\pmis_server\signs"
            Dim filename As String = System.IO.Path.Combine(folder, idno & ".jpg")
            'Dim folder1 As String = "C:\Users\User\Desktop\HRIS\Displaydata\Resources"
            'Dim filename1 As String = System.IO.Path.Combine(folder, "error image.jpg")
            If Not System.IO.File.Exists(filename) Then
                ' file does not exist
                PictureBox2.Image = Nothing


            Else
                PictureBox2.Image = Image.FromFile(filename)
                PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
                PictureBox2.Image = Image.FromFile(filename)
                PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage

            End If

        Catch ex As Exception

        End Try
    End Sub



    Private Sub MetroTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MetroTextBox1.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And MetroTextBox1.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If

    End Sub

    Private Sub MetroTextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MetroTextBox7.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And MetroTextBox7.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub MetroTextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MetroTextBox8.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And MetroTextBox8.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub MetroTextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MetroTextBox9.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And MetroTextBox9.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub MetroTextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MetroTextBox10.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And MetroTextBox10.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub MetroTextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MetroTextBox11.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And MetroTextBox11.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If
    End Sub
    Private Sub MetroTextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MetroTextBox12.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And MetroTextBox12.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If

    End Sub
    Private Sub MetroTextBox20_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MetroTextBox20.KeyPress
        If (Not e.KeyChar = ChrW(Keys.Back) And ("0123456789.").IndexOf(e.KeyChar) = -1) Or (e.KeyChar = "." And MetroTextBox20.Text.ToCharArray().Count(Function(c) c = ".") > 0) Then
            e.Handled = True
        End If

    End Sub

    Public Sub clearfields()
        MetroTextBox1.Clear()
        MetroTextBox2.Clear()
        MetroTextBox3.Clear()
        MetroTextBox4.Clear()
        MetroTextBox5.Clear()
        DateTimePicker1.Value = Date.Now
        MetroTextBox7.Clear()
        MetroTextBox8.Clear()
        MetroTextBox9.Clear()
        MetroTextBox10.Clear()
        MetroTextBox11.Clear()
        MetroTextBox12.Clear()
        MetroTextBox13.Clear()
        MetroTextBox14.Clear()
        MetroTextBox15.Clear()
        MetroTextBox16.Clear()
        MetroTextBox17.Clear()
        MetroTextBox18.Clear()
        MetroTextBox19.Clear()
        MetroTextBox20.Clear()
        Label21.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If conmed.State = ConnectionState.Open Then

        ElseIf conmed.State = ConnectionState.Closed Then
            conmed.Open()
        End If
        Dim _id = MetroTextBox1.Text.Trim
        Try


            If _id = "" Then
                MetroMessageBox.Show(Me, vbCrLf & "Provide a valid Employee ID. Don't leave empty fields." & vbCrLf & vbCrLf & vbCrLf & "©(PGLU-HRMD Series of 2017)", "PGLU ID Manager.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                deletedata()
                Dim InsertQuery As String
                'Dim conmed As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\pmis_server\IDPGLUdb2013\id_pglu.mdb")
                'InsertQuery = ("insert into dbo_idpers ( pers_id, last_name, first_name, mi, nick_name) values (@t1, @t2, @t3, @t4, @t5)")
                InsertQuery = ("insert into dbo_idpers ( pers_id, last_name, first_name, mi, nick_name, birthdate, pagibig_rdn, gsis_id, gsis_umid_crn, philhealth, tin, gsis_bpno, pos_name, dept_name, blood, allergies, medical_condition, person, address, c_telno ) values (@t1, @t2, @t3, @t4, @t5, @t6, @t7, @t8, @t9, @t10, @t11, @t12, @t13, @t14, @t15, @t16, @t17, @t18, @t19, @t20)")


                Dim cmd As OleDbCommand = New OleDbCommand(InsertQuery, conmed)

                cmd.Parameters.Add(New OleDbParameter("@t1", OleDbType.VarChar, 5, "pers_id"))
                cmd.Parameters.Add(New OleDbParameter("@t2", OleDbType.VarChar, 255, "last_name"))
                cmd.Parameters.Add(New OleDbParameter("@t3", OleDbType.VarChar, 255, "first_name"))
                cmd.Parameters.Add(New OleDbParameter("@t4", OleDbType.VarChar, 5, "mi"))
                cmd.Parameters.Add(New OleDbParameter("@t5", OleDbType.VarChar, 50, "nick_name"))
                cmd.Parameters.Add(New OleDbParameter("@t6", OleDbType.Date, 50, "birthdate"))
                cmd.Parameters.Add(New OleDbParameter("@t7", OleDbType.VarChar, 25, "pagibig_rdn"))
                cmd.Parameters.Add(New OleDbParameter("@t8", OleDbType.VarChar, 25, "gsis_id"))
                cmd.Parameters.Add(New OleDbParameter("@t9", OleDbType.VarChar, 25, "gsis_umid_crn"))
                cmd.Parameters.Add(New OleDbParameter("@t10", OleDbType.VarChar, 25, "philhealth"))
                cmd.Parameters.Add(New OleDbParameter("@t11", OleDbType.VarChar, 25, "tin"))
                cmd.Parameters.Add(New OleDbParameter("@t12", OleDbType.VarChar, 25, "gsis_bpno"))
                cmd.Parameters.Add(New OleDbParameter("@t13", OleDbType.VarChar, 255, "pos_name"))
                cmd.Parameters.Add(New OleDbParameter("@t14", OleDbType.VarChar, 255, "dept_name"))
                cmd.Parameters.Add(New OleDbParameter("@t15", OleDbType.VarChar, 5, "blood"))
                cmd.Parameters.Add(New OleDbParameter("@t16", OleDbType.VarChar, 100, "allergies"))
                cmd.Parameters.Add(New OleDbParameter("@t17", OleDbType.VarChar, 100, "medical_condition"))
                cmd.Parameters.Add(New OleDbParameter("@t18", OleDbType.VarChar, 255, "person"))
                cmd.Parameters.Add(New OleDbParameter("@t19", OleDbType.VarChar, 255, "address"))
                cmd.Parameters.Add(New OleDbParameter("@t20", OleDbType.VarChar, 40, "c_telno"))


                cmd.Parameters("@t1").Value = _id
                cmd.Parameters("@t2").Value = MetroTextBox2.Text.Trim
                cmd.Parameters("@t3").Value = MetroTextBox3.Text.Trim
                cmd.Parameters("@t4").Value = MetroTextBox4.Text.Trim
                cmd.Parameters("@t5").Value = MetroTextBox5.Text.Trim
                cmd.Parameters("@t6").Value = DateTimePicker1.Text.Trim
                cmd.Parameters("@t7").Value = MetroTextBox7.Text.Trim
                cmd.Parameters("@t8").Value = MetroTextBox8.Text.Trim
                cmd.Parameters("@t9").Value = MetroTextBox9.Text.Trim
                cmd.Parameters("@t10").Value = MetroTextBox10.Text.Trim
                cmd.Parameters("@t11").Value = MetroTextBox11.Text.Trim
                cmd.Parameters("@t12").Value = MetroTextBox12.Text.Trim
                cmd.Parameters("@t13").Value = MetroTextBox13.Text.Trim
                cmd.Parameters("@t14").Value = MetroTextBox14.Text.Trim
                cmd.Parameters("@t15").Value = MetroTextBox15.Text.Trim
                cmd.Parameters("@t16").Value = MetroTextBox16.Text.Trim
                cmd.Parameters("@t17").Value = MetroTextBox17.Text.Trim
                cmd.Parameters("@t18").Value = MetroTextBox18.Text.Trim
                cmd.Parameters("@t19").Value = MetroTextBox19.Text.Trim
                cmd.Parameters("@t20").Value = MetroTextBox20.Text.Trim

                cmd.ExecuteReader()
                MetroMessageBox.Show(Me, vbCrLf & "Record Saved." & vbCrLf & vbCrLf & vbCrLf & "©(PGLU-HRMD Series of 2017)", "HR Document Tracking", MessageBoxButtons.OK, MessageBoxIcon.Question)
                clearfields()
                MetroTextBox1.Focus()

            End If

            MetroGrid1.Columns.Clear()
        Catch ex As Exception
            MetroMessageBox.Show(Me, vbCrLf & "Error inserting record!" & ex.Message & vbCrLf & vbCrLf & vbCrLf & "©(PGLU-HRMD Series of 2017)", "HR Document Tracking", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    'SHOW RECORD FOR LIVE SEARCH and show it to datagrid tables
    Public Sub Showrecord()

        Dim searchname As String = MetroTextBox1.Text.Trim
        If searchname = "" Then
            clearfields()
        Else
            Dim dba As New OleDbDataAdapter("SELECT * FROM dbo_idpers Where pers_id = '" & searchname & "' order by last_name asc", conmed)
            Dim dbset As New DataSet
            dba.Fill(dbset)
            Me.MetroGrid1.DataSource = dbset.Tables(0).DefaultView

            'Label20.Text = DataGridView1.Rows.Count
            Label21.Text = CType(MetroGrid1.Rows.Count, String)
        End If


    End Sub

    Private Sub MetroTextBox1_TextChanged(sender As Object, e As EventArgs) Handles MetroTextBox1.TextChanged
        Loadpicture()
        loadsig()
        Showrecord()

    End Sub


    Private Sub getData(ByVal dataCollection As AutoCompleteStringCollection)

        Dim adapter As New OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim sql As String = "SELECT DISTINCT [pers_id] FROM [dbo_idpers]"
        Try

            idcommand = New OleDbCommand(sql, conmed)
            adapter.SelectCommand = idcommand
            adapter.Fill(ds)
            adapter.Dispose()
            idcommand.Dispose()

            For Each row As DataRow In ds.Tables(0).Rows
                dataCollection.Add(row(0).ToString.Trim())
            Next
        Catch ex As Exception
            MessageBox.Show("Can not open connection ! ")
        End Try
    End Sub

    Private Sub getposition(ByVal _position As AutoCompleteStringCollection)
        Dim _pos As String = MetroTextBox13.Text.Trim

        Dim adapter As New OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim sql As String = "SELECT DISTINCT [pos_name] FROM [dbo_idpers] where pos_name Like '%" & _pos & "%'  order by [pos_name] asc "
        'Dim sql As String = "SELECT DISTINCT [pos_name] FROM [dbo_idpers]"
        Try

            idcommand = New OleDbCommand(sql, conmed)
            adapter.SelectCommand = idcommand
            adapter.Fill(ds)
            adapter.Dispose()
            idcommand.Dispose()

            For Each row As DataRow In ds.Tables(0).Rows
                _position.Add(row(0).ToString.Trim())
            Next
        Catch ex As Exception
            MessageBox.Show("Can not open connection ! ")
        End Try
    End Sub
    Private Sub getdepartment(ByVal _department As AutoCompleteStringCollection)
        Dim _dept As String = MetroTextBox14.Text.Trim

        Dim adapter As New OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim sql As String = "SELECT DISTINCT [dept_name] FROM [dbo_idpers] where dept_name Like '%" & _dept & "%' order by [dept_name] asc"
        'Dim sql As String = "SELECT DISTINCT [dept_name] FROM [dbo_idpers]"
        Try

            idcommand = New OleDbCommand(sql, conmed)
            adapter.SelectCommand = idcommand
            adapter.Fill(ds)
            adapter.Dispose()
            idcommand.Dispose()

            For Each row As DataRow In ds.Tables(0).Rows
                _department.Add(row(0).ToString.Trim())
            Next
        Catch ex As Exception
            MessageBox.Show("Can not open connection ! ")
        End Try
    End Sub

    Private Sub MetroGrid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles MetroGrid1.CellContentClick
        MetroTextBox2.Text = MetroGrid1.CurrentRow.Cells(1).Value.ToString.Trim
        MetroTextBox3.Text = MetroGrid1.CurrentRow.Cells(2).Value.ToString.Trim
        MetroTextBox4.Text = MetroGrid1.CurrentRow.Cells(3).Value.ToString.Trim
        MetroTextBox5.Text = MetroGrid1.CurrentRow.Cells(4).Value.ToString.Trim
        DateTimePicker1.Text = MetroGrid1.CurrentRow.Cells(5).Value.ToString.Trim
        MetroTextBox7.Text = MetroGrid1.CurrentRow.Cells(6).Value.ToString.Trim
        MetroTextBox8.Text = MetroGrid1.CurrentRow.Cells(7).Value.ToString.Trim
        MetroTextBox9.Text = MetroGrid1.CurrentRow.Cells(8).Value.ToString.Trim
        MetroTextBox10.Text = MetroGrid1.CurrentRow.Cells(9).Value.ToString.Trim
        MetroTextBox11.Text = MetroGrid1.CurrentRow.Cells(10).Value.ToString.Trim
        MetroTextBox12.Text = MetroGrid1.CurrentRow.Cells(11).Value.ToString.Trim
        MetroTextBox13.Text = MetroGrid1.CurrentRow.Cells(12).Value.ToString.Trim
        MetroTextBox14.Text = MetroGrid1.CurrentRow.Cells(13).Value.ToString.Trim
        MetroTextBox15.Text = MetroGrid1.CurrentRow.Cells(14).Value.ToString.Trim
        MetroTextBox16.Text = MetroGrid1.CurrentRow.Cells(15).Value.ToString.Trim
        MetroTextBox17.Text = MetroGrid1.CurrentRow.Cells(16).Value.ToString.Trim
        MetroTextBox18.Text = MetroGrid1.CurrentRow.Cells(17).Value.ToString.Trim
        MetroTextBox19.Text = MetroGrid1.CurrentRow.Cells(18).Value.ToString.Trim
        MetroTextBox20.Text = MetroGrid1.CurrentRow.Cells(19).Value.ToString.Trim

        'FOR ID FRONT

        Button2.Text = MetroGrid1.CurrentRow.Cells(1).Value.ToString.Trim + ", " + MetroGrid1.CurrentRow.Cells(2).Value.ToString.Trim + " " + MetroGrid1.CurrentRow.Cells(3).Value.ToString.Trim + "."
        Button3.Text = MetroGrid1.CurrentRow.Cells(12).Value.ToString.Trim
        Button4.Text = MetroGrid1.CurrentRow.Cells(0).Value.ToString.Trim
        Button5.Text = MetroGrid1.CurrentRow.Cells(4).Value.ToString.Trim

        'FOR ID BACK
        Label24.Text = MetroGrid1.CurrentRow.Cells(13).Value.ToString.Trim
        Label25.Text = DateTimePicker1.Text.Trim
        Label26.Text = MetroGrid1.CurrentRow.Cells(10).Value.ToString.Trim
        Label27.Text = MetroGrid1.CurrentRow.Cells(8).Value.ToString.Trim
        Label28.Text = MetroGrid1.CurrentRow.Cells(7).Value.ToString.Trim
        Label29.Text = MetroGrid1.CurrentRow.Cells(6).Value.ToString.Trim
        Label30.Text = MetroGrid1.CurrentRow.Cells(9).Value.ToString.Trim

        Label31.Text = MetroGrid1.CurrentRow.Cells(14).Value.ToString.Trim
        Label32.Text = MetroGrid1.CurrentRow.Cells(15).Value.ToString.Trim
        Label33.Text = MetroGrid1.CurrentRow.Cells(16).Value.ToString.Trim

        Button7.Text = MetroGrid1.CurrentRow.Cells(17).Value.ToString.Trim
        Button8.Text = MetroGrid1.CurrentRow.Cells(18).Value.ToString.Trim
        Button9.Text = MetroGrid1.CurrentRow.Cells(19).Value.ToString.Trim

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim _FRM As New ValidityDate
        _FRM.TextBox1.Text = Button6.Text.Trim
        _FRM.ShowDialog()

    End Sub

    Private Sub deletedata()
        Dim _idnum As String = MetroTextBox1.Text.Trim
        idcommand = New OleDbCommand("delete from dbo_idpers where pers_id = '" & _idnum & "'", conmed)
        idcommand.ExecuteNonQuery()
    End Sub

End Class
