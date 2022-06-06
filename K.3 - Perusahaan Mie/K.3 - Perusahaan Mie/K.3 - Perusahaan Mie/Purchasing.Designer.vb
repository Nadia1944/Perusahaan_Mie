<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Purchasing
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.btnexit = New System.Windows.Forms.Button()
        Me.CetakPO = New System.Windows.Forms.Button()
        Me.Input = New System.Windows.Forms.Button()
        Me.dgvPO = New System.Windows.Forms.DataGridView()
        Me.Save = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.satuanRM = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.hargaRM = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.jumlahorder = New System.Windows.Forms.TextBox()
        Me.namaRM = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.kodeorderRM = New System.Windows.Forms.TextBox()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.Delete = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Edit = New System.Windows.Forms.Button()
        Me.Namapemesan = New System.Windows.Forms.TextBox()
        Me.Supplier = New System.Windows.Forms.TextBox()
        Me.tglPO = New System.Windows.Forms.DateTimePicker()
        Me.txtnoPO = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbTOP = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.txtnamapotokasir = New System.Windows.Forms.TextBox()
        Me.btntelusuri = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.tglterima = New System.Windows.Forms.DateTimePicker()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.NoPO = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.tglkirim = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.totalbiaya = New System.Windows.Forms.TextBox()
        Me.namapengirim = New System.Windows.Forms.TextBox()
        Me.NoInvoice = New System.Windows.Forms.TextBox()
        Me.poto = New System.Windows.Forms.TextBox()
        Me.btnsavedetailbayar = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnuploadbukti = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvPO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(7, 10)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(682, 496)
        Me.TabControl1.TabIndex = 59
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(674, 470)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Buat Pesanan"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Cancel)
        Me.GroupBox1.Controls.Add(Me.btnexit)
        Me.GroupBox1.Controls.Add(Me.CetakPO)
        Me.GroupBox1.Controls.Add(Me.Input)
        Me.GroupBox1.Controls.Add(Me.dgvPO)
        Me.GroupBox1.Controls.Add(Me.Save)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.DateTimePicker2)
        Me.GroupBox1.Controls.Add(Me.Delete)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Edit)
        Me.GroupBox1.Controls.Add(Me.Namapemesan)
        Me.GroupBox1.Controls.Add(Me.Supplier)
        Me.GroupBox1.Controls.Add(Me.tglPO)
        Me.GroupBox1.Controls.Add(Me.txtnoPO)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cmbTOP)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(11, 19)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(650, 435)
        Me.GroupBox1.TabIndex = 32
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Input PO"
        '
        'Cancel
        '
        Me.Cancel.Location = New System.Drawing.Point(24, 378)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(75, 34)
        Me.Cancel.TabIndex = 52
        Me.Cancel.Text = "Cancel"
        Me.Cancel.UseVisualStyleBackColor = True
        '
        'btnexit
        '
        Me.btnexit.Location = New System.Drawing.Point(553, 378)
        Me.btnexit.Name = "btnexit"
        Me.btnexit.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.btnexit.Size = New System.Drawing.Size(75, 34)
        Me.btnexit.TabIndex = 56
        Me.btnexit.Text = "EXIT"
        Me.btnexit.UseVisualStyleBackColor = True
        '
        'CetakPO
        '
        Me.CetakPO.Location = New System.Drawing.Point(426, 378)
        Me.CetakPO.Name = "CetakPO"
        Me.CetakPO.Size = New System.Drawing.Size(75, 34)
        Me.CetakPO.TabIndex = 51
        Me.CetakPO.Text = "CETAK PO"
        Me.CetakPO.UseVisualStyleBackColor = True
        '
        'Input
        '
        Me.Input.Location = New System.Drawing.Point(107, 378)
        Me.Input.Name = "Input"
        Me.Input.Size = New System.Drawing.Size(75, 34)
        Me.Input.TabIndex = 46
        Me.Input.Text = "Input"
        Me.Input.UseVisualStyleBackColor = True
        '
        'dgvPO
        '
        Me.dgvPO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPO.Location = New System.Drawing.Point(291, 91)
        Me.dgvPO.Name = "dgvPO"
        Me.dgvPO.RowTemplate.Height = 25
        Me.dgvPO.Size = New System.Drawing.Size(351, 265)
        Me.dgvPO.TabIndex = 30
        '
        'Save
        '
        Me.Save.Location = New System.Drawing.Point(346, 378)
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(75, 34)
        Me.Save.TabIndex = 50
        Me.Save.Text = "SAVE"
        Me.Save.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.satuanRM)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.hargaRM)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.jumlahorder)
        Me.GroupBox2.Controls.Add(Me.namaRM)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.kodeorderRM)
        Me.GroupBox2.Location = New System.Drawing.Point(58, 128)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(204, 206)
        Me.GroupBox2.TabIndex = 29
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Detail Bahan Baku"
        '
        'satuanRM
        '
        Me.satuanRM.Location = New System.Drawing.Point(99, 131)
        Me.satuanRM.Name = "satuanRM"
        Me.satuanRM.Size = New System.Drawing.Size(81, 20)
        Me.satuanRM.TabIndex = 35
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(7, 138)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(61, 13)
        Me.Label13.TabIndex = 34
        Me.Label13.Text = "Satuan RM"
        '
        'hargaRM
        '
        Me.hargaRM.Location = New System.Drawing.Point(99, 168)
        Me.hargaRM.Name = "hargaRM"
        Me.hargaRM.Size = New System.Drawing.Size(81, 20)
        Me.hargaRM.TabIndex = 33
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(9, 175)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(56, 13)
        Me.Label12.TabIndex = 32
        Me.Label12.Text = "Harga RM"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(7, 105)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(69, 13)
        Me.Label11.TabIndex = 31
        Me.Label11.Text = "Jumlah Order"
        '
        'jumlahorder
        '
        Me.jumlahorder.Location = New System.Drawing.Point(99, 98)
        Me.jumlahorder.Name = "jumlahorder"
        Me.jumlahorder.Size = New System.Drawing.Size(81, 20)
        Me.jumlahorder.TabIndex = 30
        '
        'namaRM
        '
        Me.namaRM.Location = New System.Drawing.Point(99, 65)
        Me.namaRM.Name = "namaRM"
        Me.namaRM.Size = New System.Drawing.Size(81, 20)
        Me.namaRM.TabIndex = 25
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(7, 72)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(55, 13)
        Me.Label10.TabIndex = 24
        Me.Label10.Text = "Nama RM"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(7, 36)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(81, 13)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "Kode Order RM"
        '
        'kodeorderRM
        '
        Me.kodeorderRM.Location = New System.Drawing.Point(99, 29)
        Me.kodeorderRM.Name = "kodeorderRM"
        Me.kodeorderRM.Size = New System.Drawing.Size(81, 20)
        Me.kodeorderRM.TabIndex = 23
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Location = New System.Drawing.Point(442, 14)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(167, 20)
        Me.DateTimePicker2.TabIndex = 27
        '
        'Delete
        '
        Me.Delete.Location = New System.Drawing.Point(267, 378)
        Me.Delete.Name = "Delete"
        Me.Delete.Size = New System.Drawing.Size(75, 34)
        Me.Delete.TabIndex = 48
        Me.Delete.Text = "Delete"
        Me.Delete.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(297, 21)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(71, 13)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Tanggal Kirim"
        '
        'Edit
        '
        Me.Edit.Location = New System.Drawing.Point(187, 378)
        Me.Edit.Name = "Edit"
        Me.Edit.Size = New System.Drawing.Size(75, 34)
        Me.Edit.TabIndex = 47
        Me.Edit.Text = "Edit"
        Me.Edit.UseVisualStyleBackColor = True
        '
        'Namapemesan
        '
        Me.Namapemesan.Location = New System.Drawing.Point(152, 91)
        Me.Namapemesan.Name = "Namapemesan"
        Me.Namapemesan.Size = New System.Drawing.Size(132, 20)
        Me.Namapemesan.TabIndex = 22
        '
        'Supplier
        '
        Me.Supplier.Location = New System.Drawing.Point(152, 64)
        Me.Supplier.Name = "Supplier"
        Me.Supplier.Size = New System.Drawing.Size(132, 20)
        Me.Supplier.TabIndex = 20
        '
        'tglPO
        '
        Me.tglPO.Location = New System.Drawing.Point(152, 39)
        Me.tglPO.Name = "tglPO"
        Me.tglPO.Size = New System.Drawing.Size(132, 20)
        Me.tglPO.TabIndex = 19
        '
        'txtnoPO
        '
        Me.txtnoPO.Location = New System.Drawing.Point(152, 14)
        Me.txtnoPO.Name = "txtnoPO"
        Me.txtnoPO.Size = New System.Drawing.Size(132, 20)
        Me.txtnoPO.TabIndex = 18
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(297, 45)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(139, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Jangka Waktu Pembayaran"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(60, 96)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Nama Pemesan"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(63, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Supplier"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(63, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Tanggal PO"
        '
        'cmbTOP
        '
        Me.cmbTOP.FormattingEnabled = True
        Me.cmbTOP.Items.AddRange(New Object() {"Net d Days", "Net EOM", "In Advance"})
        Me.cmbTOP.Location = New System.Drawing.Point(442, 42)
        Me.cmbTOP.Name = "cmbTOP"
        Me.cmbTOP.Size = New System.Drawing.Size(167, 21)
        Me.cmbTOP.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(65, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "No.PO"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.txtnamapotokasir)
        Me.TabPage2.Controls.Add(Me.btntelusuri)
        Me.TabPage2.Controls.Add(Me.GroupBox4)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(674, 470)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Transaksi"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'txtnamapotokasir
        '
        Me.txtnamapotokasir.Location = New System.Drawing.Point(952, 271)
        Me.txtnamapotokasir.Name = "txtnamapotokasir"
        Me.txtnamapotokasir.Size = New System.Drawing.Size(97, 20)
        Me.txtnamapotokasir.TabIndex = 20
        '
        'btntelusuri
        '
        Me.btntelusuri.Location = New System.Drawing.Point(1071, 271)
        Me.btntelusuri.Name = "btntelusuri"
        Me.btntelusuri.Size = New System.Drawing.Size(75, 23)
        Me.btntelusuri.TabIndex = 18
        Me.btntelusuri.Text = "Telusuri File"
        Me.btntelusuri.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.tglterima)
        Me.GroupBox4.Controls.Add(Me.Button1)
        Me.GroupBox4.Controls.Add(Me.Button2)
        Me.GroupBox4.Controls.Add(Me.Button3)
        Me.GroupBox4.Controls.Add(Me.NoPO)
        Me.GroupBox4.Controls.Add(Me.Label19)
        Me.GroupBox4.Controls.Add(Me.Label18)
        Me.GroupBox4.Controls.Add(Me.tglkirim)
        Me.GroupBox4.Controls.Add(Me.Label17)
        Me.GroupBox4.Controls.Add(Me.totalbiaya)
        Me.GroupBox4.Controls.Add(Me.namapengirim)
        Me.GroupBox4.Controls.Add(Me.NoInvoice)
        Me.GroupBox4.Controls.Add(Me.poto)
        Me.GroupBox4.Controls.Add(Me.btnsavedetailbayar)
        Me.GroupBox4.Controls.Add(Me.PictureBox1)
        Me.GroupBox4.Controls.Add(Me.btnuploadbukti)
        Me.GroupBox4.Controls.Add(Me.Label16)
        Me.GroupBox4.Controls.Add(Me.Label15)
        Me.GroupBox4.Controls.Add(Me.Label14)
        Me.GroupBox4.Location = New System.Drawing.Point(15, 14)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(634, 483)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Detail Pembayaran"
        '
        'tglterima
        '
        Me.tglterima.Location = New System.Drawing.Point(131, 172)
        Me.tglterima.Name = "tglterima"
        Me.tglterima.Size = New System.Drawing.Size(151, 20)
        Me.tglterima.TabIndex = 69
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(131, 254)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 34)
        Me.Button1.TabIndex = 68
        Me.Button1.Text = "Cancel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(220, 254)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 34)
        Me.Button2.TabIndex = 67
        Me.Button2.Text = "CETAK"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(130, 215)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 34)
        Me.Button3.TabIndex = 66
        Me.Button3.Text = "Input"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'NoPO
        '
        Me.NoPO.Location = New System.Drawing.Point(131, 16)
        Me.NoPO.Name = "NoPO"
        Me.NoPO.Size = New System.Drawing.Size(151, 20)
        Me.NoPO.TabIndex = 65
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(29, 23)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(42, 13)
        Me.Label19.TabIndex = 64
        Me.Label19.Text = "No. PO"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(29, 177)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(87, 13)
        Me.Label18.TabIndex = 62
        Me.Label18.Text = "Tanggal Diterima"
        '
        'tglkirim
        '
        Me.tglkirim.Location = New System.Drawing.Point(131, 140)
        Me.tglkirim.Name = "tglkirim"
        Me.tglkirim.Size = New System.Drawing.Size(151, 20)
        Me.tglkirim.TabIndex = 61
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(29, 142)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(71, 13)
        Me.Label17.TabIndex = 60
        Me.Label17.Text = "Tanggal Kirim"
        '
        'totalbiaya
        '
        Me.totalbiaya.Location = New System.Drawing.Point(131, 108)
        Me.totalbiaya.Name = "totalbiaya"
        Me.totalbiaya.Size = New System.Drawing.Size(151, 20)
        Me.totalbiaya.TabIndex = 59
        '
        'namapengirim
        '
        Me.namapengirim.Location = New System.Drawing.Point(131, 80)
        Me.namapengirim.Name = "namapengirim"
        Me.namapengirim.Size = New System.Drawing.Size(151, 20)
        Me.namapengirim.TabIndex = 58
        '
        'NoInvoice
        '
        Me.NoInvoice.Location = New System.Drawing.Point(131, 50)
        Me.NoInvoice.Name = "NoInvoice"
        Me.NoInvoice.Size = New System.Drawing.Size(151, 20)
        Me.NoInvoice.TabIndex = 57
        '
        'poto
        '
        Me.poto.Location = New System.Drawing.Point(474, 25)
        Me.poto.Name = "poto"
        Me.poto.Size = New System.Drawing.Size(126, 20)
        Me.poto.TabIndex = 56
        '
        'btnsavedetailbayar
        '
        Me.btnsavedetailbayar.Location = New System.Drawing.Point(220, 215)
        Me.btnsavedetailbayar.Name = "btnsavedetailbayar"
        Me.btnsavedetailbayar.Size = New System.Drawing.Size(75, 34)
        Me.btnsavedetailbayar.TabIndex = 55
        Me.btnsavedetailbayar.Text = "SAVE"
        Me.btnsavedetailbayar.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(325, 50)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(274, 186)
        Me.PictureBox1.TabIndex = 54
        Me.PictureBox1.TabStop = False
        '
        'btnuploadbukti
        '
        Me.btnuploadbukti.Location = New System.Drawing.Point(325, 25)
        Me.btnuploadbukti.Name = "btnuploadbukti"
        Me.btnuploadbukti.Size = New System.Drawing.Size(75, 20)
        Me.btnuploadbukti.TabIndex = 49
        Me.btnuploadbukti.Text = "Telusuri"
        Me.btnuploadbukti.UseVisualStyleBackColor = True
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(29, 50)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(62, 13)
        Me.Label16.TabIndex = 42
        Me.Label16.Text = "No. Invoice"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(29, 80)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(78, 13)
        Me.Label15.TabIndex = 41
        Me.Label15.Text = "Nama Pengirim"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(29, 111)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(60, 13)
        Me.Label14.TabIndex = 40
        Me.Label14.Text = "Total Biaya"
        '
        'Purchasing
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(703, 514)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "Purchasing"
        Me.Text = "Purchasing"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvPO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Cancel As Button
    Friend WithEvents btnexit As Button
    Friend WithEvents CetakPO As Button
    Friend WithEvents Input As Button
    Friend WithEvents dgvPO As DataGridView
    Friend WithEvents Save As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents satuanRM As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents hargaRM As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents jumlahorder As TextBox
    Friend WithEvents namaRM As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents kodeorderRM As TextBox
    Friend WithEvents DateTimePicker2 As DateTimePicker
    Friend WithEvents Delete As Button
    Friend WithEvents Label9 As Label
    Friend WithEvents Edit As Button
    Friend WithEvents Namapemesan As TextBox
    Friend WithEvents Supplier As TextBox
    Friend WithEvents tglPO As DateTimePicker
    Friend WithEvents txtnoPO As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbTOP As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents txtnamapotokasir As TextBox
    Friend WithEvents btntelusuri As Button
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents tglterima As DateTimePicker
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents NoPO As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents tglkirim As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents totalbiaya As TextBox
    Friend WithEvents namapengirim As TextBox
    Friend WithEvents NoInvoice As TextBox
    Friend WithEvents poto As TextBox
    Friend WithEvents btnsavedetailbayar As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btnuploadbukti As Button
    Friend WithEvents Label16 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label14 As Label
End Class
