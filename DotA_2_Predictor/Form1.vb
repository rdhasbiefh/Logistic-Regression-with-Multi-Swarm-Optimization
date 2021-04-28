Imports System.Text

Public Class FrmMain
    Private Function GetDataTableFromDGV(ByVal dgv As DataGridView) As DataTable
        Dim dt = New DataTable()

        For Each column As DataGridViewColumn In dgv.Columns
            If column.Visible Then
                dt.Columns.Add(column.HeaderText)
            End If
        Next

        Dim cellValues As Object() = New Object(dgv.Columns.Count - 1) {}

        For Each row As DataGridViewRow In dgv.Rows
            For i As Integer = 0 To row.Cells.Count - 1
                cellValues(i) = row.Cells(i).Value
            Next
            dt.Rows.Add(cellValues)
        Next
        Return dt
    End Function

    Private Sub BtnNormalisasi_Click(sender As Object, e As EventArgs) Handles btnNormalisasi.Click
        pnlBtnPosition.Height = btnNormalisasi.Height
        pnlBtnPosition.Top = btnNormalisasi.Top
        pnlNormalisasi.Visible = True
        pnlOptimasi.Visible = False
        pnlPrediksi.Visible = False
    End Sub

    Private Sub BtnOptimasi_Click(sender As Object, e As EventArgs) Handles btnOptimasi.Click
        pnlBtnPosition.Height = btnOptimasi.Height
        pnlBtnPosition.Top = btnOptimasi.Top
        pnlNormalisasi.Visible = False
        pnlOptimasi.Visible = True
        pnlPrediksi.Visible = False
    End Sub

    Private Sub BtnPrediksi_Click(sender As Object, e As EventArgs) Handles btnPrediksi.Click
        pnlBtnPosition.Height = btnPrediksi.Height
        pnlBtnPosition.Top = btnPrediksi.Top
        pnlNormalisasi.Visible = False
        pnlOptimasi.Visible = False
        pnlPrediksi.Visible = True
    End Sub

    ''' <summary>
    ''' ---------------------------------------- NORMALISASI ----------------------------------------
    ''' </summary>  
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnPilihDataNorm.Click
        txtboxPilihDataNorm.Clear()
        OpenFileDialog1.Reset()
        dgvNormalisasi.Rows.Clear()
        txtboxJmlData.Clear()
        OpenFileDialog1.Filter = "CSV File (*.csv*)|*.csv"

        Try
            If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                txtboxPilihDataNorm.Text = OpenFileDialog1.FileName
            End If
            Dim fName As String = OpenFileDialog1.FileName
            Dim TextLine As String = ""
            Dim SplitLine() As String

            If System.IO.File.Exists(fName) = True Then
                Dim objReader As New System.IO.StreamReader(fName, Encoding.Default)
                Do While objReader.Peek() <> -1
                    TextLine = objReader.ReadLine()
                    SplitLine = Split(TextLine, ",")
                    Me.dgvNormalisasi.Rows.Add(SplitLine)
                Loop
                txtboxJmlData.Text = dgvNormalisasi.Rows.Count.ToString()
            Else
                MsgBox("Tidak ada file yang dipilih.")
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnNormNorm.Click
        If txtboxPilihDataNorm.Text.Length = 0 Then
            MsgBox("Pilih data yang akan dinormalisasi.")
        Else

            ' Normalisasi data 
            Dim kolom() As Integer = {2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18} ' Kolom yang di normalisasi
            Dim radWin() As Integer = {19} ' Kolom hasil --> true / false
            Dim jumlahBaris As Integer = dgvNormalisasi.Rows.Count
            Dim jumlahKolom As Integer = dgvNormalisasi.Columns.Count
            Dim hasil(1)() As Double ' [0] = Rata2 data, [1] = StdDev data

            For i = 0 To 1
                hasil(i) = New Double(jumlahKolom - 1) {}
            Next i

            For c = 0 To kolom.Length - 1
                Dim j As Integer = kolom(c)

                ' Nilai rata2 setiap kolom
                Dim total As Double = 0.0
                For r = 0 To jumlahBaris - 1
                    total += dgvNormalisasi.Rows(r).Cells(j).Value
                Next r
                Dim rata2 As Double = total / jumlahBaris
                hasil(0)(c) = rata2

                ' Nilai stdDev setiap kolom
                Dim totalKuadrat As Double = 0.0
                For r = 0 To jumlahBaris - 1
                    totalKuadrat += (dgvNormalisasi.Rows(r).Cells(j).Value - rata2) ^ 2
                Next r
                Dim stdDev As Double = Math.Sqrt(totalKuadrat / jumlahBaris)
                hasil(1)(c) = stdDev
            Next c

            For c = 0 To kolom.Length - 1
                Dim j As Integer = kolom(c) ' Kolom yang dinormalisasi
                Dim rata2 As Double = hasil(0)(c) ' Nilai rata-rata untuk setiap kolom 
                Dim stdDev As Double = hasil(1)(c) ' Nilai stdev untuk  setiap kolom

                For i = 0 To jumlahBaris - 1
                    dgvNormalisasi.Rows(i).Cells(j).Value = Math.Round((dgvNormalisasi.Rows(i).Cells(j).Value - rata2) / stdDev, 4)
                Next i
            Next c

            For x = 0 To radWin.Length - 1
                Dim k As Integer = radWin(x)
                For i = 0 To jumlahBaris - 1
                    Dim temp As String = dgvNormalisasi.Rows(i).Cells(k).Value
                    Dim str1 As String = "TRUE"

                    If String.Compare(temp, str1) = 0 Then
                        dgvNormalisasi.Rows(i).Cells(k).Value = 1
                    Else
                        dgvNormalisasi.Rows(i).Cells(k).Value = 0
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub BtnSimpanNorm_Click(sender As Object, e As EventArgs) Handles btnSimpanNorm.Click
        If txtboxPilihDataNorm.Text.Length = 0 Then
            MsgBox("Tidak ada file yang dapat disimpan.")
        Else
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            SaveFileDialog1.Filter = "CSV File (*.csv*)|*.csv"
            Try
                If (SaveFileDialog1.ShowDialog() = DialogResult.OK) Then
                    If (SaveFileDialog1.FileName IsNot Nothing) Then
                        Dim fname As String = SaveFileDialog1.FileName
                        Dim thecsvfile As String = String.Empty

                        ' get baris
                        For Each row As DataGridViewRow In dgvNormalisasi.Rows
                            ' get kolom
                            For Each cell As DataGridViewCell In row.Cells
                                thecsvfile = thecsvfile & cell.FormattedValue.replace(",", "") & ","
                            Next
                            thecsvfile = thecsvfile.TrimEnd(",")
                            thecsvfile = thecsvfile & vbCr & vbLf
                        Next
                        My.Computer.FileSystem.WriteAllText(fname, thecsvfile, False)
                        MessageBox.Show("Data berhasil disimpan.")
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            stopwatch.[Stop]()
            'Console.WriteLine(stopwatch.ElapsedMilliseconds)
        End If
    End Sub

    ''' <summary>
    ''' ---------------------------------------- OPTIMASI ----------------------------------------
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnPilihDataOptm_Click(sender As Object, e As EventArgs) Handles btnPilihDataOptm.Click
        txtboxPilihDataOptm.Clear()
        OpenFileDialog2.Reset()
        dgvOptimasi.Rows.Clear()
        txtBoxJmlData2.Clear()
        txtBoxJmlSwarm.Clear()
        txtBoxJmlPartikel.Clear()
        txtBoxJmlIterasi.Clear()
        dgvInisialisasi.Rows.Clear()
        TextboxPmax.Clear()
        TextboxPmin.Clear()
        TextboxKmax.Clear()
        TextboxKmin.Clear()
        dgvTerbaikSementara.Rows.Clear()
        dgvProsesPencarian.Rows.Clear()
        dgvFungsiTerbaik.Rows.Clear()
        dgvBobotTerbaik.Rows.Clear()
        OpenFileDialog2.Filter = "CSV File (*.csv*)|*.csv"

        Try
            If (OpenFileDialog2.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                txtboxPilihDataOptm.Text = OpenFileDialog2.FileName
            End If
            Dim fName As String = OpenFileDialog2.FileName
            Dim TextLine As String = ""
            Dim SplitLine() As String

            If System.IO.File.Exists(fName) = True Then
                Dim objReader As New System.IO.StreamReader(fName, Encoding.Default)
                Do While objReader.Peek() <> -1
                    TextLine = objReader.ReadLine()
                    SplitLine = Split(TextLine, ",")
                    Me.dgvOptimasi.Rows.Add(SplitLine)
                Loop
                txtBoxJmlData2.Text = dgvOptimasi.Rows.Count.ToString()
            Else
                MsgBox("Tidak ada file yang dipilih.")
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub TxtBoxJmlSwarm_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtBoxJmlSwarm.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TxtBoxJmlPartikel_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtBoxJmlPartikel.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TxtBoxJmlIterasi_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtBoxJmlIterasi.KeyPress
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub BtnOptmOptm_Click(sender As Object, e As EventArgs) Handles btnOptmOptm.Click
        If txtBoxJmlData2.Text.Length = 0 Then
            MsgBox("Pilih data yang akan dioptimasi.")
        ElseIf txtBoxJmlSwarm.Text.Length = 0 Then
            MsgBox("Kolom jumlah swarm belum di isi.")
        ElseIf txtBoxJmlPartikel.Text.Length = 0 Then
            MsgBox("Kolom jumlah partikel belum di isi.")
        ElseIf txtBoxJmlIterasi.Text.Length = 0 Then
            MsgBox("Kolom jumlah iterasi belum di isi.")
        ElseIf TextboxPmin.Text.Length = 0 Then
            MsgBox("Kolom P.Min belum di isi.")
        ElseIf TextboxPmax.Text.Length = 0 Then
            MsgBox("Kolom P.Max belum di isi.")
        ElseIf TextboxKmin.Text.Length = 0 Then
            MsgBox("Kolom K.Min belum di isi.")
        ElseIf TextboxKmax.Text.Length = 0 Then
            MsgBox("Kolom K.Max belum di isi.")
        Else
            Dim stopwatch As Stopwatch = Stopwatch.StartNew()

            dgvInisialisasi.Rows.Clear()
            dgvTerbaikSementara.Rows.Clear()
            dgvProsesPencarian.Rows.Clear()
            dgvFungsiTerbaik.Rows.Clear()
            dgvBobotTerbaik.Rows.Clear()

            Dim baris As Integer = dgvOptimasi.Rows.Count
            Dim kolomOptm() As Integer = {2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 19}

            ' move data dgv to array
            Dim data(baris - 1)() As Double
            For r As Integer = 0 To baris - 1
                Dim temp(kolomOptm.Length - 1) As Double
                For c As Integer = 0 To kolomOptm.Length - 1
                    Dim x As Integer = kolomOptm(c)
                    temp(c) = dgvOptimasi.Rows(r).Cells(x).Value
                Next
                data(r) = temp
            Next

            Dim rl As New RegresiLogistikMSO(16)
            Dim jumlahSwarm As Integer = CInt(txtBoxJmlSwarm.Text)
            Dim jumlahPartikel As Integer = CInt(txtBoxJmlPartikel.Text)
            Dim maksEpoch As Integer = CInt(txtBoxJmlIterasi.Text)

            Dim bobotTerbaik() As Double = rl.ProsesPerhitungan(data, jumlahSwarm, jumlahPartikel, maksEpoch)
            Dim nilaiFungsi As Double = rl.HitungNilaiKesalahan(data, bobotTerbaik)

            dgvFungsiTerbaik.Rows.Add()
            dgvFungsiTerbaik.Rows(0).Cells(0).Value = nilaiFungsi.ToString("F6")

            dgvBobotTerbaik.Rows.Add()
            For i = 0 To bobotTerbaik.Length - 1
                dgvBobotTerbaik.Rows(0).Cells(i).Value = bobotTerbaik(i).ToString("F4")
            Next i
            stopwatch.[Stop]()
            'Console.WriteLine(stopwatch.ElapsedMilliseconds)
        End If
    End Sub

    Private Sub BtnSimpanBobot_Click(sender As Object, e As EventArgs) Handles btnSimpanBobot.Click
        If txtboxPilihDataOptm.Text.Length = 0 Then
            MsgBox("Tidak ada file yang dapat disimpan.")
        Else
            SaveFileDialog2.Filter = "XML files(.xml)|*.xml"
            Try
                If (SaveFileDialog2.ShowDialog() = DialogResult.OK) Then
                    If (SaveFileDialog2.FileName IsNot Nothing) Then
                        Dim fname As String = SaveFileDialog2.FileName
                        Dim dT As DataTable = GetDataTableFromDGV(dgvBobotTerbaik)
                        dT.TableName = "TableBobotTerbaik"

                        Dim dS As DataSet = New DataSet("DataSetBobotTerbaik")
                        dS.Tables.Add(dT)

                        Dim streamWrite As System.IO.FileStream = New System.IO.FileStream(fname, System.IO.FileMode.Create)
                        dS.WriteXml(streamWrite, System.Data.XmlWriteMode.WriteSchema)
                        streamWrite.Close()
                        MessageBox.Show("Data berhasil disimpan.")
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' ---------------------------------------- PREDIKSI ----------------------------------------
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnPilihDataXml_Click(sender As Object, e As EventArgs) Handles btnPilihDataXml.Click
        txtboxPilihDataXml.Clear()
        OpenFileDialog3.Reset()
        dgvPerbandingan.Rows.Clear()
        textboxJmlDataPredPred.Clear()
        textboxJmlSamaPred.Clear()
        textboxJmlBedaPred.Clear()
        textboxAkurasi.Clear()
        OpenFileDialog3.Filter = "XML files(.xml)|*.xml"

        Try
            If (OpenFileDialog3.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                txtboxPilihDataXml.Text = OpenFileDialog3.FileName
            End If
            Dim fName As String = OpenFileDialog3.FileName

            If System.IO.File.Exists(fName) = True Then
                Dim ds As New DataSet
                ds.ReadXml(fName)
                dgvXml.Columns.Clear()
                dgvXml.DataSource = ds.Tables(0)
            Else
                MsgBox("Tidak ada file yang dipilih.")
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub BtnPilihDataNormPred_Click(sender As Object, e As EventArgs) Handles btnPilihDataNormPred.Click
        textboxPilihDataNormPred.Clear()
        OpenFileDialog4.Reset()
        dgvNormPred.Rows.Clear()
        textboxJmlDataNormPred.Clear()
        dgvPerbandingan.Rows.Clear()
        textboxJmlDataPredPred.Clear()
        textboxJmlSamaPred.Clear()
        textboxJmlBedaPred.Clear()
        textboxAkurasi.Clear()
        OpenFileDialog4.Filter = "CSV File (*.csv*)|*.csv"

        Try
            If (OpenFileDialog4.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                textboxPilihDataNormPred.Text = OpenFileDialog4.FileName
            End If
            Dim fName As String = OpenFileDialog4.FileName
            Dim TextLine As String = ""
            Dim SplitLine() As String

            If System.IO.File.Exists(fName) = True Then
                Dim objReader As New System.IO.StreamReader(fName, Encoding.Default)
                Do While objReader.Peek() <> -1
                    TextLine = objReader.ReadLine()
                    SplitLine = Split(TextLine, ",")
                    Me.dgvNormPred.Rows.Add(SplitLine)
                Loop
                textboxJmlDataNormPred.Text = dgvNormPred.Rows.Count.ToString()
            Else
                MsgBox("Tidak ada file yang dipilih.")
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub BtnPrediksiPred_Click(sender As Object, e As EventArgs) Handles btnPrediksiPred.Click
        If txtboxPilihDataXml.Text.Length = 0 Then
            MsgBox("Pilih data bobot terbaik.")
        ElseIf textboxPilihDataNormPred.Text.Length = 0 Then
            MsgBox("Pilih hasil normalisasi dari data yang akan di prediksi.")
        Else

            Dim barisBobot As Integer = dgvXml.Rows.Count
            Dim barisDataNorm As Integer = dgvNormPred.Rows.Count
            Dim kolomOptm() As Integer = {2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 19}
            Dim kolomBobot() As Integer = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}
            Dim kolomBanding() As Integer = {0, 1, 10, 19}

            ' move data bobot from dgv to array
            Dim bobotTerbaik(barisBobot - 1) As Double
            For r As Integer = 0 To barisBobot - 1
                Dim temp(kolomBobot.Length - 1) As Double
                For c As Integer = 0 To kolomBobot.Length - 1
                    Dim x As Integer = kolomBobot(c)
                    temp(c) = dgvXml.Rows(r).Cells(x).Value
                Next
                bobotTerbaik = temp
            Next

            ' move data normalized from dgv to array
            Dim data(barisDataNorm - 1)() As Double
            For r As Integer = 0 To barisDataNorm - 1
                Dim temp(kolomOptm.Length - 1) As Double
                For c As Integer = 0 To kolomOptm.Length - 1
                    Dim x As Integer = kolomOptm(c)
                    temp(c) = dgvNormPred.Rows(r).Cells(x).Value
                Next
                data(r) = temp
            Next

            Dim rl As New RegresiLogistikMSO(16)
            Dim jumlahSama As Integer = 0, jumlahBeda As Integer = 0
            For i = 0 To data.Length - 1
                dgvPerbandingan.Rows.Add()
                For j = 0 To kolomBanding.Length - 1
                    Dim x As Integer = kolomBanding(j)
                    dgvPerbandingan.Rows(i).Cells(j).Value = dgvNormPred.Rows(i).Cells(x).Value
                Next j
                Dim output As Double = Math.Round(rl.HitungNilaiOutput(data(i), bobotTerbaik), 4)
                dgvPerbandingan.Rows(i).Cells(4).Value = output
                If output <= 0.5 Then
                    dgvPerbandingan.Rows(i).Cells(5).Value = 0 '"0 (" & output & ")"
                Else
                    dgvPerbandingan.Rows(i).Cells(5).Value = 1 '"1 (" & output & ")"
                End If

                If String.Compare(dgvPerbandingan.Rows(i).Cells(3).Value, dgvPerbandingan.Rows(i).Cells(5).Value) = 0 Then
                    dgvPerbandingan.Rows(i).Cells(6).Value = "Sama"
                    jumlahSama += 1
                Else
                    dgvPerbandingan.Rows(i).Cells(6).Value = "Beda"
                    jumlahBeda += 1
                End If
            Next i

            textboxJmlDataPredPred.Text = dgvPerbandingan.Rows.Count
            textboxJmlSamaPred.Text = jumlahSama
            textboxJmlBedaPred.Text = jumlahBeda
            Dim akurasi As Double = Math.Round(jumlahSama / (dgvPerbandingan.Rows.Count) * 100, 2)
            textboxAkurasi.Text = akurasi & "%"
        End If
    End Sub

    ''' <summary>
    ''' ---------------------------------------- CLASS MSO ----------------------------------------
    ''' </summary>
    Public Class RegresiLogistikMSO
        Private ReadOnly jumlahFitur As Integer ' Jumlah kolom fitur selain fitur radiant_win
        Private ReadOnly bobot() As Double ' bobot yang akan dicari, bobot dengan indeks 0 adalah konstanta/b0
        Private ReadOnly rnd As Random

        Public Sub New(ByVal jumlahFitur As Integer)
            Me.jumlahFitur = jumlahFitur
            Me.bobot = New Double(jumlahFitur) {}
            Me.rnd = New Random()
        End Sub

        ' pencarian posisi terbaik
        Public Function ProsesPerhitungan(ByVal data As Double()(), ByVal jumlahSwarm As Integer, ByVal jumlahPartikel As Integer, ByVal maksEpoch As Integer) As Double()
            ' Kolom dimensi ditambah 1 untuk menghitung b0
            Dim dimensi As Integer = jumlahFitur + 1

            ' batas posisi minimal dan maksimal partikel WOW
            Dim minX As Double = CInt(FrmMain.TextboxPmin.Text) '-10.0
            Dim maksX As Double = CInt(FrmMain.TextboxPmax.Text) '10.0

            ' batas kecepatan minimal dan maksimal partikel
            Dim minKecepatan As Double = CInt(FrmMain.TextboxKmin.Text) '-1.0
            Dim maksKecepatan As Double = CInt(FrmMain.TextboxKmax.Text) '1.0

            ' Inisialisasi parameter partikel, swarm, multiswarm
            ' Beri nilai posisi partikel awal dengan posisi acak
            ' beri nilai kecepatan acak pada partikel tersebut
            Dim ms As New MultiSwarm(jumlahSwarm, jumlahPartikel, dimensi, minX, maksX, minKecepatan, maksKecepatan)

            Dim indeksPosisiTerbaik(1) As Integer
            For i As Integer = 0 To indeksPosisiTerbaik.Length - 1
                indeksPosisiTerbaik(i) = -1
            Next

            ' Tentukan nilai kesalahan terendah (terbaik) sementara untuk posisi acak 
            ' Bandingkan semua nilai kesalahan pada masing-masing partikel, ambil nilai kesalahan terendah (terbaik)
            ' tentukan nilai kesalahan terendah pada masing-masing swarm
            ' tentukan nilai kesalahan terendah antar swarm untuk mendapatkan nilai kesalahan terendah pada multiswarm
            For i = 0 To jumlahSwarm - 1
                Dim rowIndex As Integer
                For j = 0 To jumlahPartikel - 1
                    Dim p As Partikel = ms.daftarSwarm(i).daftarPartikel(j)
                    p.nilaiKesalahan = HitungNilaiKesalahan(data, p.posisi) ' add error

                    p.nilaiKesalahanTerbaik = p.nilaiKesalahan
                    Array.Copy(p.posisi, p.posisiTerbaik, dimensi)
                    If p.nilaiKesalahan < ms.daftarSwarm(i).nilaiKesalahanSwarmTerbaik Then ' swarm best?
                        ms.daftarSwarm(i).nilaiKesalahanSwarmTerbaik = p.nilaiKesalahan
                        Array.Copy(p.posisi, ms.daftarSwarm(i).posisiSwarmTerbaik, dimensi)
                    End If
                    If p.nilaiKesalahan < ms.nilaiKesalahanMultiSwarmTerbaik Then ' global best?
                        ms.nilaiKesalahanMultiSwarmTerbaik = p.nilaiKesalahan
                        Array.Copy(p.posisi, ms.posisiMultiSwarmTerbaik, dimensi)

                        indeksPosisiTerbaik(0) = i
                        indeksPosisiTerbaik(1) = j
                    End If
                    FrmMain.dgvInisialisasi.Rows.Add()
                    FrmMain.dgvInisialisasi.Rows(rowIndex).Cells(0).Value = (i + 1).ToString.PadRight(2)
                    FrmMain.dgvInisialisasi.Rows(rowIndex).Cells(1).Value = (j + 1).ToString.PadRight(2)
                    FrmMain.dgvInisialisasi.Rows(rowIndex).Cells(2).Value = p.nilaiKesalahan.ToString("F4")
                    rowIndex += 1
                Next j
            Next i
            FrmMain.dgvTerbaikSementara.Rows.Add()
            FrmMain.dgvTerbaikSementara.Rows(0).Cells(0).Value = (indeksPosisiTerbaik(0) + 1).ToString.PadRight(2)
            FrmMain.dgvTerbaikSementara.Rows(0).Cells(1).Value = (indeksPosisiTerbaik(1) + 1).ToString.PadRight(2)
            FrmMain.dgvTerbaikSementara.Rows(0).Cells(2).Value = ms.nilaiKesalahanMultiSwarmTerbaik.ToString("F4")

            ' bobot inertia (w), bobot kognitif (c1), bobot sosial (c2), dan bobot global (c3)
            Const w As Double = 0.729
            Const c1 As Double = 1.4945
            Const c2 As Double = 1.4945
            Const c3 As Double = 0.3645

            Dim urutanPartikel(jumlahPartikel - 1) As Integer
            For i = 0 To urutanPartikel.Length - 1
                urutanPartikel(i) = i
            Next i

            ' proses pencarian posisi terbaik sebanyak jumlah iterasi
            Dim epoch As Integer = 0
            Do While epoch < maksEpoch
                epoch += 1

                For i = 0 To jumlahSwarm - 1
                    For j = 0 To urutanPartikel.Length - 1
                        Dim r As Integer = rnd.Next(j, urutanPartikel.Length)
                        Dim tmp As Integer = urutanPartikel(r)
                        urutanPartikel(r) = urutanPartikel(j)
                        urutanPartikel(j) = tmp
                    Next j

                    ' perhitungan setiap partikel dalam masing-masing swarm
                    For pj = 0 To jumlahPartikel - 1 ' each Partikel
                        Dim j = urutanPartikel(pj)
                        Dim p As Partikel = ms.daftarSwarm(i).daftarPartikel(j)
                        Dim rowIndex2 As Integer

                        ' perhitungan kecepatan perpindahan posisi yang baru
                        For k = 0 To dimensi - 1
                            Dim r1 As Double = rnd.NextDouble
                            Dim r2 As Double = rnd.NextDouble
                            Dim r3 As Double = rnd.NextDouble

                            p.kecepatan(k) = (w * p.kecepatan(k)) + (c1 * r1 * (p.posisiTerbaik(k) - p.posisi(k))) + (c2 * r2 * (ms.daftarSwarm(i).posisiSwarmTerbaik(k) - p.posisi(k))) + (c3 * r3 * (ms.posisiMultiSwarmTerbaik(k) - p.posisi(k)))

                            If p.kecepatan(k) < minKecepatan Then
                                p.kecepatan(k) = minKecepatan
                            ElseIf p.kecepatan(k) > maksKecepatan Then
                                p.kecepatan(k) = maksKecepatan
                            End If
                        Next k

                        ' perhitungan posisi partikel baru 
                        For k = 0 To dimensi - 1
                            p.posisi(k) += p.kecepatan(k)
                            If p.posisi(k) < minX Then
                                p.posisi(k) = minX
                            ElseIf p.posisi(k) > maksX Then
                                p.posisi(k) = maksX
                            End If
                        Next k

                        ' perhitungan nilai kesalahan untuk posisi yang baru
                        p.nilaiKesalahan = HitungNilaiKesalahan(data, p.posisi)

                        ' Jika nilai kesalahan baru lebih rendah dari nilai kesalahan yang diperoleh partikel sebelumnya, ambil posisi yang baru sebagai posisi terbaik partikel 
                        If p.nilaiKesalahan < p.nilaiKesalahanTerbaik Then
                            p.nilaiKesalahanTerbaik = p.nilaiKesalahan
                            Array.Copy(p.posisi, p.posisiTerbaik, dimensi)
                        End If

                        ' Jika nilai kesalahan baru  lebih rendah dari nilai kesalahan swarm partikel, ambil posisi yang baru sebagai posisi terbaik swarm partikel 
                        If p.nilaiKesalahan < ms.daftarSwarm(i).nilaiKesalahanSwarmTerbaik Then
                            ms.daftarSwarm(i).nilaiKesalahanSwarmTerbaik = p.nilaiKesalahan
                            Array.Copy(p.posisi, ms.daftarSwarm(i).posisiSwarmTerbaik, dimensi)
                        End If

                        ' Jika nilai kesalahan baru  lebih rendah dari nilai kesalahan multi swarm, ambil posisi yang baru sebagai posisi terbaik multiswarm
                        If p.nilaiKesalahan < ms.nilaiKesalahanMultiSwarmTerbaik Then
                            ms.nilaiKesalahanMultiSwarmTerbaik = p.nilaiKesalahan
                            Array.Copy(p.posisi, ms.posisiMultiSwarmTerbaik, dimensi)

                            FrmMain.dgvProsesPencarian.Rows.Add()
                            FrmMain.dgvProsesPencarian.Rows(rowIndex2).Cells(0).Value = epoch.ToString.PadRight(3)
                            FrmMain.dgvProsesPencarian.Rows(rowIndex2).Cells(1).Value = (i + 1).ToString.PadRight(2)
                            FrmMain.dgvProsesPencarian.Rows(rowIndex2).Cells(2).Value = (j + 1).ToString.PadRight(2)
                            FrmMain.dgvProsesPencarian.Rows(rowIndex2).Cells(3).Value = ms.nilaiKesalahanMultiSwarmTerbaik.ToString("F6")
                            rowIndex2 += 1
                        End If
                    Next pj
                Next i
            Loop
            Return ms.posisiMultiSwarmTerbaik
        End Function

        'perhitungan nilai kesalahan/mse
        Public Function HitungNilaiKesalahan(ByVal data As Double()(), ByVal bobot() As Double) As Double
            Dim indeksKriteriaHasil As Integer = data(0).Length - 1
            Dim hasil As Double = 0.0
            For i = 0 To data.Length - 1
                Dim hasilPerhitungan As Double = HitungNilaiOutput(data(i), bobot)
                Dim hasilData As Double = data(i)(indeksKriteriaHasil)

                hasil += (hasilPerhitungan - hasilData) * (hasilPerhitungan - hasilData)
                '' hasil += (hasilData - hasilPerhitungan) * (hasilData - hasilPerhitungan)
            Next i
            Return hasil / data.Length
        End Function

        'perhitungan fungsi sigmoid
        Public Function HitungNilaiOutput(ByVal dataItem() As Double, ByVal bobot() As Double) As Double
            Dim z As Double = bobot(0) 'b0 = konstanta

            For i = 0 To bobot.Length - 2
                z += (bobot(i + 1) * dataItem(i))
            Next i
            Return 1.0 / (1.0 + Math.Exp(-z))
        End Function

        Private Class Partikel
            Private Shared ReadOnly rnd As New Random()
            Public posisi() As Double
            Public kecepatan() As Double
            Public nilaiKesalahan As Double
            Public posisiTerbaik() As Double
            Public nilaiKesalahanTerbaik As Double

            Public Sub New(ByVal dimensi As Integer,
                           ByVal minX As Double, ByVal maksX As Double, ByVal minKecepatan As Double, ByVal maksKecepatan As Double)
                posisi = New Double(dimensi - 1) {}
                kecepatan = New Double(dimensi - 1) {}
                posisiTerbaik = New Double(dimensi - 1) {}
                For k = 0 To dimensi - 1
                    posisi(k) = (maksX - minX) * rnd.NextDouble + minX
                    kecepatan(k) = (maksKecepatan - minKecepatan) * rnd.NextDouble + minKecepatan
                Next k
                nilaiKesalahan = Double.MaxValue
                nilaiKesalahanTerbaik = nilaiKesalahan
            End Sub
        End Class

        Private Class Swarm
            Public daftarPartikel() As Partikel
            Public posisiSwarmTerbaik() As Double
            Public nilaiKesalahanSwarmTerbaik As Double

            Public Sub New(ByVal jumlahPartikel As Integer, ByVal dimensi As Integer,
                           ByVal minX As Double, ByVal maksX As Double, ByVal minKecepatan As Double, ByVal maksKecepatan As Double)
                daftarPartikel = New Partikel(jumlahPartikel - 1) {}
                For i = 0 To jumlahPartikel - 1
                    daftarPartikel(i) = New Partikel(dimensi, minX, maksX, minKecepatan, maksKecepatan)
                Next i
                posisiSwarmTerbaik = New Double(dimensi - 1) {}
                nilaiKesalahanSwarmTerbaik = Double.MaxValue
            End Sub
        End Class

        Private Class MultiSwarm
            Public daftarSwarm() As Swarm
            Public posisiMultiSwarmTerbaik() As Double
            Public nilaiKesalahanMultiSwarmTerbaik As Double

            Public Sub New(ByVal jumlahSwarm As Integer, ByVal jumlahPartikel As Integer, ByVal dimensi As Integer,
                           ByVal minX As Double, ByVal maksX As Double, ByVal minKecepatan As Double, ByVal maksKecepatan As Double)
                daftarSwarm = New Swarm(jumlahSwarm - 1) {}
                For i = 0 To jumlahSwarm - 1
                    daftarSwarm(i) = New Swarm(jumlahPartikel, dimensi, minX, maksX, minKecepatan, maksKecepatan)
                Next i
                posisiMultiSwarmTerbaik = New Double(dimensi - 1) {}
                nilaiKesalahanMultiSwarmTerbaik = Double.MaxValue
            End Sub
        End Class
    End Class
End Class