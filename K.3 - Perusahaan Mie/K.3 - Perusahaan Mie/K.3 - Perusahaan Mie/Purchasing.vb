Public Class Purchasing
    Sub NomorPO()
        Call koneksiDB()
        CMD = New OleDb.OleDbCommand("select * from Order_RM where No_PO in (select max(No_PO) from Order_RM)", Conn)
        DM = CMD.ExecuteReader
        DM.Read()
        Dim urutankode As String
        Dim hitung As Long
        If Not DM.HasRows Then
            urutankode = "PO" + Format(Now, "yyMMdd") + "001"
        Else
            hitung = Microsoft.VisualBasic.Right(DM.GetString(0), 9) + 1
            urutankode = "PO" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & hitung, 3)
        End If
        txtnoPO.Text = urutankode
    End Sub
    Sub kondisiawal()
        tglPO.Value = Today
    End Sub
    Sub KosongkanForm()
        Namapemesan.Text = ""
        Supplier.Text = ""
        namaRM.Text = ""
        kodeorderRM.Text = ""
        namaRM.Text = ""
        jumlahorder.Text = ""
        satuanRM.Text = ""
        hargaRM.Text = ""
        txtnoPO.Focus()
    End Sub
    Sub MatikanForm()
        txtnoPO.Enabled = False
        Namapemesan.Enabled = False
        Supplier.Enabled = False
        namaRM.Enabled = False
        kodeorderRM.Enabled = False
        satuanRM.Enabled = False
        hargaRM.Enabled = False
        jumlahorder.Enabled = False
    End Sub
    Sub HidupkanForm()
        txtnoPO.Enabled = True
        Namapemesan.Enabled = True
        Supplier.Enabled = True
        namaRM.Enabled = True
        kodeorderRM.Enabled = True
        satuanRM.Enabled = True
        hargaRM.Enabled = True
        jumlahorder.Enabled = True
        tglPO.Enabled = True
        DateTimePicker2.Enabled = True
        cmbTOP.Enabled = True
    End Sub
    Sub Hidupkanbtn()
        Save.Enabled = True
        Edit.Enabled = True
        btnexit.Enabled = True
        Input.Enabled = True
        CetakPO.Enabled = True
        Delete.Enabled = True
        Cancel.Enabled = True
    End Sub
    Sub Matikanbtn()
        Edit.Enabled = False
        Delete.Enabled = False
        Cancel.Enabled = False
    End Sub
    Sub TampilkanData()
        Call koneksiDB()
        DA = New OleDb.OleDbDataAdapter("Select * from Order_RM", Conn)
        DS = New DataSet
        DA.Fill(DS)
        dgvPO.DataSource = DS.Tables(0)
        dgvPO.ReadOnly = True
    End Sub
    Private Sub Purchasing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call MatikanForm()
        Call TampilkanData()
        Call Matikanbtn()
        tglPO.Enabled = False
        DateTimePicker2.Enabled = False
        cmbTOP.Enabled = False
        txtnoPO.Text = NoPO.Text
        Call kondisiawal()
        Call NomorPO()
    End Sub
    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles Cancel.Click
        Call KosongkanForm()
        Call MatikanForm()
    End Sub
    Private Sub Input_Click(sender As Object, e As EventArgs) Handles Input.Click
        Call NomorPO()
        Call HidupkanForm()
        Call KosongkanForm()
        Call Hidupkanbtn()
    End Sub
    Private Sub btnexit_Click(sender As Object, e As EventArgs) Handles btnexit.Click
        Dim menu As New Main_Menu
        Me.Close()
        menu.Show()
    End Sub
End Class