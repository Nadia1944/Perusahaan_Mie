Public Class WarehouseRM
    Sub KosongkanForm1()
        txtId_RM.Text = ""
        txtNama_RM.Text = ""
        txtJumlah_Stok_RM.Text = ""
        txtSatuan_RM.Text = ""
    End Sub
    Sub MatikanForm1()
        txtId_RM.Enabled = False
        txtNama_RM.Enabled = False
        txtJumlah_Stok_RM.Enabled = False
        txtSatuan_RM.Enabled = False

    End Sub
    Sub HidupkanForm1()
        txtId_RM.Enabled = True
        txtNama_RM.Enabled = True
        txtJumlah_Stok_RM.Enabled = True
        txtSatuan_RM.Enabled = True

    End Sub
    Sub TampilkanData1()
        Call koneksiDB()
        DA = New OleDb.OleDbDataAdapter("select * from Warehouse_RM ", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DataGridView1.DataSource = DS.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub Customer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call MatikanForm()
        Call TampilkanData()
    End Sub
    Private Sub btn_input_Click(sender As Object, e As EventArgs) Handles btn_input.Click
        Call HidupkanForm()
        Call KosongkanForm()
    End Sub

    Private Sub btn_save_Click(sender As Object, e As EventArgs) Handles btn_save.Click
        If txtId_RM.Text = "" Or txtNama_RM.Text = "" Or txtJumlah_Stok_RM.Text = "" Or txtSatuan_RM.Text = "" Then
            MsgBox("Data Warehouse_RM Belum Lengkap")
            Exit Sub
        Else
            Call koneksiDB()
            CMD = New OleDb.OleDbCommand(" select * from Warehouse_RM where Id_RM='" & txtId_RM.Text & "'", Conn)
            DM = CMD.ExecuteReader
            DM.Read()
            If Not DM.HasRows Then
                Call koneksiDB()
                Dim simpan As String
                simpan = "insert into Warehouse_RM values ('" & txtId_RM.Text &
               "', '" & txtNama_RM.Text & "',  '" & txtJumlah_Stok_RM.Text & "', '" & txtSatuan_RM.Text & "')"
                CMD = New OleDb.OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()
                MsgBox("Input Data Sukses")
            Else
                MsgBox("Data Sudah Ada")
            End If
            Call MatikanForm()
            Call KosongkanForm()
            Call TampilkanData()
        End If
    End Sub

    Private Sub btn_edit_Click(sender As Object, e As EventArgs) Handles btn_edit.Click
        If txtId_RM.Text = "" Or txtNama_RM.Text = "" Or txtJumlah_Stok_RM.Text = "" Or txtSatuan_RM.Text = Then
            MsgBox("Data Warehouse_RM Belum Lengkap")
            Exit Sub
        Else
            Call koneksiDB()
            CMD = New OleDb.OleDbCommand("update Warehouse_RM set Nama_RM = '" &
           txtNama_RM.Text & "', Jumlah_Stok_RM = '" & txtJumlah_Stok_RM.Text & "', No_Hp = '" & txtSatuan_RM.Text & "' 
           where Id_RM ='" & txtId_RM.Text & "'", Conn)
            DM = CMD.ExecuteReader
            MsgBox("Update Data Berhasil")
        End If
        Call KosongkanForm()
        Call MatikanForm()
        Call TampilkanData()
    End Sub

    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        If txtId_RM.Text = "" Then
            MsgBox("Tidak ada data yang dipilih")
            Exit Sub
        Else
            If MessageBox.Show(" Are you sure to delete this data?", "Konfirmasi", MessageBoxButtons.YesNoCancel) Then
                Call koneksiDB()
                CMD = New OleDb.OleDbCommand(" delete from Customer where Id_RM = '" & txtId_RM.Text & "'", Conn)
                DM = CMD.ExecuteReader
                MsgBox("Data Berhasil Dihapus")
                Call MatikanForm()
                Call KosongkanForm()
                Call TampilkanData()
            Else
                Call KosongkanForm()
                Call TampilkanData()
            End If
        End If
    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click
        Call MatikanForm()
        Call KosongkanForm()
    End Sub

    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        On Error Resume Next
        txtId_RM.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        txtNama_RM.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        txtJumlah_Stok_RM.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        txtSatuan_RM.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        Call HidupkanForm()
        txtId_RM.Enabled = False
    End Sub

    Sub KosongkanForm()
        txtId_RM2.Text = ""
        txtNama_RM2.Text = ""
        txtJumlah_Stok_RM2.Text = ""
        txtSatuan_RM2.Text = ""
    End Sub
    Sub MatikanForm()
        txtId_RM2.Enabled = False
        txtNama_RM2.Enabled = False
        txtJumlah_Stok_RM2.Enabled = False
        txtSatuan_RM2.Enabled = False

    End Sub
    Sub HidupkanForm()
        txtId_RM2.Enabled = True
        txtNama_RM2.Enabled = True
        txtJumlah_Stok_RM2.Enabled = True
        txtSatuan_RM2.Enabled = True

    End Sub
    Sub TampilkanData()
        Call koneksiDB()
        DA = New OleDb.OleDbDataAdapter("select * from Warehouse_RM ", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DataGridView2.DataSource = DS.Tables(0)
        DataGridView2.ReadOnly = True
    End Sub
    Private Sub Customer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call MatikanForm()
        Call TampilkanData()
    End Sub
    Private Sub btn_input_Click(sender As Object, e As EventArgs) Handles btn_input2.Click
        Call HidupkanForm()
        Call KosongkanForm()
    End Sub

    Private Sub btn_save2_Click(sender As Object, e As EventArgs) Handles btn_save2.Click
        If txtId_RM2.Text = "" Or txtNama_RM2.Text = "" Or txtJumlah_Stok_RM2.Text = "" Or txtSatuan_RM2.Text = "" Then
            MsgBox("Data Warehouse_RM Belum Lengkap")
            Exit Sub
        Else
            Call koneksiDB()
            CMD = New OleDb.OleDbCommand(" select * from Warehouse_RM where Id_RM ='" & txtId_RM2.Text & "'", Conn)
            DM = CMD.ExecuteReader
            DM.Read()
            If Not DM.HasRows Then
                Call koneksiDB()
                Dim simpan As String
                simpan = "insert into Warehouse_RM values ('" & txtId_RM2.Text &
               "', '" & txtNama_RM2.Text & "',  '" & txtJumlah_Stok_RM2.Text & "', '" & txtSatuan_RM2.Text & "')"
                CMD = New OleDb.OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()
                MsgBox("Input Data Sukses")
            Else
                MsgBox("Data Sudah Ada")
            End If
            Call MatikanForm()
            Call KosongkanForm()
            Call TampilkanData()
        End If
    End Sub

    Private Sub btn_edit2_Click(sender As Object, e As EventArgs) Handles btn_edit2.Click
        If txtId_RM2.Text = "" Or txtNama_RM2.Text = "" Or txtJumlah_Stok_RM2.Text = "" Or txtSatuan_RM2.Text = Then
            MsgBox("Data Warehouse_RM Belum Lengkap")
            Exit Sub
        Else
            Call koneksiDB()
            CMD = New OleDb.OleDbCommand("update Warehouse_RM set Nama_RM = '" &
           txtNama_RM2.Text & "', Jumlah_Stok_RM = '" & txtJumlah_Stok_RM2.Text & "', No_Hp = '" & txtSatuan_RM2.Text & "' 
           where Id_RM ='" & txtId_RM2.Text & "'", Conn)
            DM = CMD.ExecuteReader
            MsgBox("Update Data Berhasil")
        End If
        Call KosongkanForm()
        Call MatikanForm()
        Call TampilkanData()
    End Sub

    Private Sub btn_delete2_Click(sender As Object, e As EventArgs) Handles btn_delete2.Click
        If txtId_RM2.Text = "" Then
            MsgBox("Tidak ada data yang dipilih")
            Exit Sub
        Else
            If MessageBox.Show(" Are you sure to delete this data?", "Konfirmasi", MessageBoxButtons.YesNoCancel) Then
                Call koneksiDB()
                CMD = New OleDb.OleDbCommand(" delete from Customer where Id_RM = '" & txtId_RM2.Text & "'", Conn)
                DM = CMD.ExecuteReader
                MsgBox("Data Berhasil Dihapus")
                Call MatikanForm()
                Call KosongkanForm()
                Call TampilkanData()
            Else
                Call KosongkanForm()
                Call TampilkanData()
            End If
        End If
    End Sub

    Private Sub btn_cancel2_Click(sender As Object, e As EventArgs) Handles btn_cancel2.Click
        Call MatikanForm()
        Call KosongkanForm()
    End Sub

    Private Sub btn_exit2_Click(sender As Object, e As EventArgs) Handles btn_exit2.Click
        Me.Close()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        On Error Resume Next
        txtId_RM2.Text = DataGridView2.Rows(e.RowIndex).Cells(0).Value
        txtNama_RM2.Text = DataGridView2.Rows(e.RowIndex).Cells(1).Value
        txtJumlah_Stok_RM2.Text = DataGridView2.Rows(e.RowIndex).Cells(2).Value
        txtSatuan_RM.Text = DataGridView2.Rows(e.RowIndex).Cells(3).Value
        Call HidupkanForm()
        txtId_RM.Enabled = False
    End Sub

End Class
