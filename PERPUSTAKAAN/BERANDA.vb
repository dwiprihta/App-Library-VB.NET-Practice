Imports MySql.Data.MySqlClient
Imports CrystalDecisions.CrystalReports.Engine
Public Class BERANDA

    '(=========================================BAGIAN DATA BUKU========================================================)

    'TAMPIL DATA
    Sub tampildt()
        Try
            Call bukadb()
            da = New MySqlDataAdapter("SELECT * FROM buku", conn)
            ds = New DataSet
            da.Fill(ds, "buku")
            DataGridView1.DataSource = ds.Tables("buku")
            DataGridView1.ReadOnly = True
            With DataGridView1
                .Columns(0).HeaderText = "KODE BUKU"
                .Columns(1).HeaderText = "JUDUL BUKU"
                .Columns(2).HeaderText = "TAHUN TERBIT"
                .Columns(3).HeaderText = "EXEMLAR"
                .Columns(4).HeaderText = "PENGARANG"
                .Columns(0).Width = 100
                .Columns(1).Width = 200
                .Columns(2).Width = 50
                .Columns(3).Width = 70
                .Columns(4).Width = 150
            End With
            DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Yellow
        Catch ex As Exception
            MsgBox("MENAMPILKAN DATA GAGAL, CEK KEMBALI KONEKSI ANDA !!!", MsgBoxStyle.Exclamation, "PERINGATAN")
        End Try
    End Sub
    'BERSIH
    Sub bersih()
        TextBox7.Text = ""
        TextBox6.Text = ""
        TextBox5.Text = ""
        TextBox16.Text = ""
        TextBox4.Text = ""
        TextBox7.Focus()
    End Sub
    'ENABEL
    Sub ENABEL()
        TextBox7.Enabled = True
        TextBox6.Enabled = True
        TextBox5.Enabled = True
        TextBox16.Enabled = True
        TextBox4.Enabled = True
        TextBox7.Focus()
    End Sub

    'GENERAL
    Private Sub BERANDA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call bersih()
        Call bukadb()
        Call tampildt()
        Call tampiltb2()
        Call tampiltb3()
        Call isicombo()
        Call isicombo2()
        Call ambilnama()
        Call ambilnama2()
        TextBox13.Text = Format(Now,"dd-MM-yyyy")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call ENABEL()
        Call bersih()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call bukadb()
        If TextBox7.Text = "" Then
            MsgBox("LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN", MsgBoxStyle.Exclamation, "PERINGATAN")
        Else
            Try
                Call bukadb()
                cmd = New MySqlCommand("select kode_buku from buku where kode_buku= '" & TextBox7.Text & "'", conn)
                dr = cmd.ExecuteReader
                dr.Read()
                If dr.HasRows Then
                    MsgBox("DATA YANG ANDA INPUTKAN MUNGKIN SUDAH ADA", MsgBoxStyle.Exclamation, "PERINGATAN")
                Else
                    Call bukadb()
                    simpan = "INSERT INTO buku(kode_buku,judul,tahun_terbit,exemplar,pengarang)VALUES('" & TextBox7.Text & "','" & TextBox6.Text & "','" & TextBox5.Text & "','" & TextBox16.Text & "','" & TextBox4.Text & "')"
                    cmd = New MySqlCommand(simpan, conn)
                    cmd.ExecuteNonQuery()
                    Call tampildt()
                    Call bersih()
                    Call isicombo()
                    Call isicombo2()
                    Call ambilnama()
                    Call ambilnama2()
                End If
            Catch ex As Exception
                MsgBox("TERJADI KESALAHAN SAAT INPUT DATA", MsgBoxStyle.Exclamation, "PERINGATAN")
            End Try

        End If
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        TextBox7.Text = DataGridView1.CurrentRow.Cells(0).Value
        TextBox6.Text = DataGridView1.CurrentRow.Cells(1).Value
        TextBox5.Text = DataGridView1.CurrentRow.Cells(2).Value
        TextBox16.Text = DataGridView1.CurrentRow.Cells(3).Value
        TextBox4.Text = DataGridView1.CurrentRow.Cells(4).Value
        TextBox7.Enabled = False
        TextBox6.Enabled = True
        TextBox5.Enabled = True
        TextBox16.Enabled = True
        TextBox4.Enabled = True
        TextBox7.Focus()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox7.Text = "" Then
            MsgBox("PILIH DAHULU DATA YANG AKAN ANDA EDIT", MsgBoxStyle.Exclamation, "PERINGATAN!!!")
        Else
            Try
                Call bukadb()
                edit = "UPDATE buku SET judul='" & TextBox6.Text & "',tahun_terbit='" & TextBox5.Text & "',exemplar='" & TextBox16.Text & "',pengarang='" & TextBox4.Text & "' WHERE kode_buku='" & TextBox7.Text & "'"
                cmd = New MySqlCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call bersih()
                Call tampildt()
            Catch ex As Exception
                MsgBox("TERJADI KESALAHAN SAAT EDIT DATA", MsgBoxStyle.Exclamation, "PERINGATAN!!!")
            End Try
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'MessageBox.Show("yakin mau dihapus ?", "perhatian", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        If TextBox7.Text = "" Then
            MsgBox("PILIH DAHULU DATA YANG AKAN ANDA HAPUS", MsgBoxStyle.Exclamation, "PERINGATAN!!!")
        Else
            Try
                Call bukadb()
                hapus = "DELETE FROM buku WHERE kode_buku='" & TextBox7.Text & "'"
                cmd = New MySqlCommand(hapus, conn)
                cmd.ExecuteNonQuery()
                Call bersih()
                Call tampildt()

            Catch ex As Exception
                MsgBox("TERJADI KESALAHAN SAAT HAPUS DATA", MsgBoxStyle.Exclamation, "PERINGATAN!!!")
            End Try
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If RadioButton1.Checked = True Then
            Call bukadb()
            TextBox1.Focus()
            da = New MySqlDataAdapter("select * from buku where kode_buku like '%" & TextBox1.Text & "%'", conn)
            ds = New DataSet
            da.Fill(ds, "KETEMU")
            DataGridView1.DataSource = ds.Tables("KETEMU")
            DataGridView1.ReadOnly = True
        End If
        If RadioButton2.Checked = True Then
            Call bukadb()
            TextBox1.Focus()
            da = New MySqlDataAdapter("select * from buku where judul like '%" & TextBox1.Text & "%'", conn)
            ds = New DataSet
            da.Fill(ds, "DAPAT")
            DataGridView1.DataSource = ds.Tables("DAPAT")
            DataGridView1.ReadOnly = True
        End If
    End Sub



    '(=========================================BAGIAN DATA ANGGOTA========================================================)

    'ENABEL
    Sub enabel2()
        TextBox11.Enabled = True
        TextBox10.Enabled = True
        TextBox9.Enabled = True
        TextBox8.Enabled = True
        TextBox11.Focus()
    End Sub

    'BERSIH
    Sub bersih2()
        TextBox11.Text = ""
        TextBox10.Text = ""
        TextBox9.Text = ""
        TextBox8.Text = ""
        TextBox11.Focus()
    End Sub
    'TAMPILTB
    Sub tampiltb2()
        Try
            Call bukadb()
            da = New MySqlDataAdapter("SELECT * FROM tb_anggota", conn)
            ds = New DataSet
            da.Fill(ds, "tb_anggota")
            DataGridView2.DataSource = ds.Tables("tb_anggota")
            DataGridView2.ReadOnly = True
            With DataGridView2
                .Columns(0).HeaderText = "KODE ANGGOTA"
                .Columns(1).HeaderText = "NAMA ANGGOTA"
                .Columns(2).HeaderText = "ALAMAT"
                .Columns(3).HeaderText = "TELPON"
                .Columns(0).Width = 100
                .Columns(1).Width = 150
                .Columns(2).Width = 150
                .Columns(3).Width = 100
            End With
            DataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.Aqua
        Catch ex As Exception
            MsgBox("MENAMPILKAN DATA GAGAL, CEK KEMBALI KONEKSI ANDA !!!", MsgBoxStyle.Exclamation, "PERINGATAN")
        End Try
    End Sub
    'TAMBAH
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call enabel2()
        Call bersih2()
    End Sub
    'SIMPAN
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call bukadb()
        If TextBox11.Text = "" Then
            MsgBox("LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN", MsgBoxStyle.Exclamation, "PERINGATAN")
        Else
            Try
                Call bukadb()
                cmd = New MySqlCommand("select kode_anggota from tb_anggota where kode_anggota= '" & TextBox11.Text & "'", conn)
                dr = cmd.ExecuteReader
                dr.Read()
                If dr.HasRows Then
                    MsgBox("DATA YANG ANDA INPUTKAN MUNGKIN SUDAH ADA", MsgBoxStyle.Exclamation, "PERINGATAN")
                Else
                    Call bukadb()
                    simpan = "INSERT INTO tb_anggota(kode_anggota,nama,alamat,telpon)VALUES('" & TextBox11.Text & "','" & TextBox10.Text & "','" & TextBox9.Text & "','" & TextBox8.Text & "')"
                    cmd = New MySqlCommand(simpan, conn)
                    cmd.ExecuteNonQuery()
                    Call tampiltb2()
                    Call bersih2()
                    Call isicombo()
                    Call isicombo2()
                    Call ambilnama()
                    Call ambilnama2()
                End If
            Catch ex As Exception
                MsgBox("TERJADI KESALAHAN SAAT INPUT DATA", MsgBoxStyle.Exclamation, "PERINGATAN")
            End Try

        End If
    End Sub
    'EDIT
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox11.Text = "" Then
            MsgBox("PILIH DAHULU DAT YANG AKAN DIEDIT !!!", MsgBoxStyle.Exclamation, "PERINGATAN")
        Else
            Try
                Call bukadb()
                edit = "UPDATE tb_anggota SET nama='" & TextBox10.Text & "',alamat='" & TextBox9.Text & "',telpon='" & TextBox8.Text & "' WHERE kode_anggota='" & TextBox11.Text & "'"
                cmd = New MySqlCommand(edit, conn)
                cmd.ExecuteNonQuery()
                Call bersih2()
                Call tampiltb2()
            Catch ex As Exception
                MsgBox("EDIT DATA GAGAL, PRIKSA DATA ANDA KEMBALI !!!", MsgBoxStyle.Exclamation, "PERINGATAN")
            End Try
        End If
    End Sub
    'MUNCULKAN
    Private Sub DataGridView2_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseDoubleClick
        TextBox11.Text = DataGridView2.CurrentRow.Cells(0).Value
        TextBox10.Text = DataGridView2.CurrentRow.Cells(1).Value
        TextBox9.Text = DataGridView2.CurrentRow.Cells(2).Value
        TextBox8.Text = DataGridView2.CurrentRow.Cells(3).Value
        TextBox11.Enabled = False
        TextBox10.Enabled = True
        TextBox9.Enabled = True
        TextBox8.Enabled = True
        TextBox11.Focus()

    End Sub
    'HAPUS
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox11.Text = "" Then
            MsgBox("PILIH DAHULU DAT YANG AKAN DIHAPUS !!!", MsgBoxStyle.Exclamation, "PERINGATAN")
        End If
        Try
            Call bukadb()
            hapus = "DELETE from tb_anggota where kode_anggota='" & TextBox11.Text & "'"
            cmd = New MySqlCommand(hapus, conn)
            cmd.ExecuteNonQuery()
            Call tampiltb2()
            Call bersih2()
        Catch ex As Exception
            MsgBox("GAGAL MENGHAPUS DATA, PRIKSA KEMBALI KONEKSI ANDA !!!", MsgBoxStyle.Exclamation, "PERINGATAN")
        End Try
    End Sub
    'CARI
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If RadioButton3.Checked = True Then
            Call bukadb()
            TextBox2.Focus()
            da = New MySqlDataAdapter("SELECT * FROM tb_anggota WHERE kode_anggota like '%" & TextBox2.Text & "%'", conn)
            ds = New DataSet
            da.Fill(ds, "KETEMU")
            DataGridView2.DataSource = ds.Tables("KETEMU")
            DataGridView2.ReadOnly = True
        End If
        If RadioButton4.Checked = True Then
            Call bukadb()
            TextBox2.Focus()
            da = New MySqlDataAdapter("SELECT * FROM tb_anggota WHERE nama like '%" & TextBox2.Text & "%'", conn)
            ds = New DataSet
            da.Fill(ds, "KETEMU")
            DataGridView2.DataSource = ds.Tables("KETEMU")
            DataGridView2.ReadOnly = True
        End If
    End Sub


    '(=========================================BAGIAN DATA PINJAM========================================================)

    'TAMPILTB3
    Sub tampiltb3()

        Call bukadb()
        da = New MySqlDataAdapter("SELECT * FROM pinjam", conn)
        ds = New DataSet
        da.Fill(ds, "pinjam")
        DataGridView5.DataSource = ds.Tables("pinjam")
        DataGridView5.ReadOnly = True
        With DataGridView5
            .Columns(0).HeaderText = "NO PINJAM"
            .Columns(1).HeaderText = "TANGGAL"
            .Columns(2).HeaderText = "JUDUL BUKU"
            .Columns(3).HeaderText = "NAMA ANGGOTA"
            .Columns(0).Width = 100
            .Columns(1).Width = 150
            .Columns(2).Width = 150
            .Columns(3).Width = 150
        End With
        DataGridView5.AlternatingRowsDefaultCellStyle.BackColor = Color.BlanchedAlmond 
    End Sub

    'BERSIH
    Sub bersih3()
        kd_pinjam.Text = ""
        kd_anggota.Text = ""
        NAM_ANGGOTA.Text = ""
        ComboBox1.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = Format(Date.Now)
        kd_pinjam.Focus()
    End Sub

    'ENABEL
    Sub enabel3()
        kd_pinjam.Enabled = True
        kd_anggota.Enabled = True
        ComboBox1.Enabled = True
        kd_pinjam.Focus()
    End Sub

    'ISI COMBO
    Sub isicombo()
        Call bukadb()
        cmd = New MySqlCommand("select kode_anggota from tb_anggota", conn)
        dr = cmd.ExecuteReader
        kd_anggota.Items.Clear()
        Do While dr.Read
            kd_anggota.Items.Add(dr.Item(0))
        Loop
        cmd.Dispose()
        dr.Close()
        conn.Close()
    End Sub

    'ISI COMBO 2
    Sub isicombo2()
        Call bukadb()
        cmd = New MySqlCommand("select kode_buku from buku", conn)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Clear()
        Do While dr.Read
            ComboBox1.Items.Add(dr.Item(0))
        Loop
        cmd.Dispose()
        dr.Close()
        conn.Close()
    End Sub

    'AMBIL NAMA
    Sub ambilnama()
        Call bukadb()
        cmd = New MySqlCommand("SELECT nama FROM tb_anggota WHERE kode_anggota ='" & kd_anggota.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            NAM_ANGGOTA.Text = dr.Item(0)
        End If
    End Sub

    'AMBIL NAMA2
    Sub ambilnama2()
        Call bukadb()
        cmd = New MySqlCommand("SELECT judul FROM buku  WHERE kode_buku ='" & ComboBox1.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            TextBox12.Text = dr.Item(0)
        End If
    End Sub

    'TAMBAH
    Private Sub Button9_Click(sender As Object, e As EventArgs)
        Call enabel3()
    End Sub

    'BERSIH
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Call bersih3()
    End Sub

    'TAMBAHKAN
    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Call enabel3()
        Call bersih3()
    End Sub

    Private Sub kd_anggota_SelectedIndexChanged(sender As Object, e As EventArgs) Handles kd_anggota.SelectedIndexChanged
        Call ambilnama()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call ambilnama2()
    End Sub
    'SIMPAN
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If kd_pinjam.Text = "" Then
            MsgBox("LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN", MsgBoxStyle.Exclamation, "PERINGATAN")
        Else
            Try
                Call bukadb()
                cmd = New MySqlCommand("SELECT * FROM pinjam WHERE no_pinjam= '" & kd_pinjam.Text & "'", conn)
                dr = cmd.ExecuteReader
                dr.Read()
                If dr.HasRows Then
                    MsgBox("DATA YANG ANDA INPUTKAN MUNGKIN SUDAH ADA", MsgBoxStyle.Exclamation, "PERINGATAN")
                Else
                    Call bukadb()
                    simpan = "INSERT INTO pinjam(no_pinjam,tanggal,kode_buku,kode_anggota) VALUES (@p1,@p2,@p3,@p4)"
                    cmd = conn.CreateCommand
                    With cmd
                        .CommandText = simpan
                        .Connection = conn
                        .Parameters.Clear()
                        .Parameters.AddWithValue("p1", (kd_pinjam.Text))
                        .Parameters.AddWithValue("p2", (Format(Now.Date, "yyyy-MM-dd")))
                        .Parameters.AddWithValue("p3", (ComboBox1.Text))
                        .Parameters.AddWithValue("p4", (kd_anggota.Text))
                        .ExecuteNonQuery()
                    End With
                    Call tampiltb3()
                    Call bersih3()
                End If
            Catch ex As Exception
                MsgBox("TERJADI KESALAHAN SAAT INPUT DATA", MsgBoxStyle.Exclamation, "PERINGATAN")
            End Try
        End If
    End Sub

    'taro di textbox
    Private Sub DataGridView5_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView5.CellMouseDoubleClick
        kd_pinjam.Text = DataGridView5.CurrentRow.Cells(0).Value
        TextBox13.Text = DataGridView5.CurrentRow.Cells(1).Value
        ComboBox1.Text = DataGridView5.CurrentRow.Cells(2).Value
        kd_anggota.Text = DataGridView5.CurrentRow.Cells(3).Value
    End Sub

    'HAPUS
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If kd_pinjam.Text = "" Then
            MsgBox("PILIH DAHULU DATA DARI DATABASE", MsgBoxStyle.Information, "PERHATIAN")
        Else
            Try
                Call bukadb()
                hapus = "DELETE FROM pinjam WHERE no_pinjam='" & kd_pinjam.Text & "'"
                cmd = New MySqlCommand(hapus, conn)
                cmd.ExecuteNonQuery()
                Call bersih3()
                Call tampiltb3()
            Catch ex As Exception
                MsgBox("HAPUS TELAH GAGAL CEK KEMBALI KONEKSI KONEKSI ANDA", MsgBoxStyle.Information, "PERHATIAN")
            End Try
        End If
    End Sub

    'CARI
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        'If RadioButton5.Checked = True Then
        Call bukadb()
        TextBox3.Focus()
        da = New MySqlDataAdapter("SELECT * FROM pinjam WHERE kode_buku like '%" & TextBox3.Text & "%'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView5.DataSource = ds.Tables
        DataGridView5.ReadOnly = True
        'End If
        'If RadioButton6.Checked = True Then
        Call bukadb()
        TextBox3.Focus()
        da = New MySqlDataAdapter("SELECT * FROM pinjam WHERE kode_anggota like '%" & TextBox3.Text & "%'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView5.DataSource = ds.Tables
        DataGridView5.ReadOnly = True
        ' End If
    End Sub
    'MEMFILTER REPORT
    'Private Sub Button11_Click(sender As Object, e As EventArgs)
    'Dim reportku As New ReportDocument
    ' reportku.Load("..\..\VINJAM.rpt")
    'reportku.SetParameterValue("mulai", DateTimePicker1.Text)
    'reportku.SetParameterValue("selesai", DateTimePicker2.Text)
    'CrystalReportViewer3.ReportSource = reportku
    'CrystalReportViewer3.Refresh()
    'End Sub
End Class