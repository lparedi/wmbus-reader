Imports System.IO
Imports System.Text
Public Class Form1

    Dim list As New DataTable
    Dim Serials As New List(Of String)
    Dim debug As Integer = 0
    Dim simulate As Integer = 0
    Dim saved As Integer = 0


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or (Not IsNumeric(TextBox1.Text)) Then
            MsgBox("La seriale!")
            Exit Sub
        End If
        If Button1.BackColor = Color.Red Then
            Exit Sub
        End If
        Timer1.Start()
        saved = 0
        BackgroundWorker1.RunWorkerAsync()


        Button1.BackColor = Color.Red
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim MBT1COM As Integer
        Dim MBT1BLUE As Object
        Dim fout As StreamWriter
        Dim buff As String
        Dim result As String
        Dim Record As Integer
        If debug = 1 Then
            fout = File.CreateText("c:\INTERVENTI\dec.csv")
        End If

        MBT1COM = CInt(TextBox1.Text)
        If simulate = 0 Then
            result = ""
            MBT1BLUE = CreateObject("MBT1ReceiverLib.MBT1Receiver.1")
            MBT1BLUE.RadioPasskey(1) = "FFFFFFFFFFFFFFFF"            'set 64 bit radio deciphering pass key 1 (if available)
            MBT1BLUE.RadioPasskey(2) = "FFFFFFFFFFFFFFFF"            'set 64 bit radio deciphering pass key 2 (if available)
            MBT1BLUE.RadioPasskey(3) = "FFFFFFFFFFFFFFFF"            'set 64 bit radio deciphering pass key 3 (if available)
            If MaskedTextBox1.Text <> "" And MaskedTextBox1.Text.Length = 32 Then
                MBT1BLUE.RadioPasskey128(1) = MaskedTextBox1.Text   'Set 128 bit radio deciphering pass key 1 (if available)
            End If
            MBT1BLUE.RadioPasskey128(2) = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF" 'Set 128 bit radio deciphering pass key 2 (if available)
                MBT1BLUE.RadioPasskey128(3) = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF" 'Set 128 bit radio deciphering pass key 3 (if available)

                MBT1BLUE.CurrentCOMPort = MBT1COM
                MBT1BLUE.AddressInterpretationSetting = 1
                MBT1BLUE.ReadParameter
                Do
                Loop While MBT1BLUE.CommunicationThreadRuns <> 0

                Dim StartTime As Integer
                StartTime = Now().Second
                MBT1BLUE.StartRadioReading

                Do                                                                              'read out the MBT1BLUE device for 15 seconds and display all values

                    Dim TelegramStr As String
                    Dim TelValuesValid As Boolean


                    TelegramStr = MBT1BLUE.NextRadioTelegram
                    If Strings.Left(TelegramStr, 2) <> "FF" Then
                        Dim i As Integer
                        Dim Value As Integer
                        Dim svalue As String
                        Dim ovalue As String
                        Dim Pvalue As Integer



                        buff = ""

                        TelValuesValid = MBT1BLUE.RADExtractDecipherValid(TelegramStr)

                        MBT1BLUE.TelegramInterpret(Strings.Mid(TelegramStr, 17, 500), TelValuesValid)

                    If MBT1BLUE.RADManufacturer = "APA" Then
                        If RadioButton1.Checked = True And TelegramStr(35) = "8" Then
                            If TelegramStr(33) = "4" Then

                                ovalue = "&H" & TelegramStr(46) & TelegramStr(47) & TelegramStr(45) & TelegramStr(44)
                                svalue = "&H" & TelegramStr(54) & TelegramStr(55) & TelegramStr(52) & TelegramStr(53)
                            ElseIf TelegramStr(33) = "1" Then
                                ovalue = "&H" & TelegramStr(50) & TelegramStr(51) & TelegramStr(48) & TelegramStr(49)
                                svalue = "&H" & TelegramStr(58) & TelegramStr(59) & TelegramStr(56) & TelegramStr(57)
                            Else

                                svalue = "&H" & TelegramStr(52) & TelegramStr(53) & TelegramStr(50) & TelegramStr(51)
                                ovalue = "&H" & TelegramStr(56) & TelegramStr(57) & TelegramStr(54) & TelegramStr(55)
                            End If

                            Value = CInt(svalue)
                            Pvalue = CInt(ovalue)
                            For i = 31 To 24 Step -2
                                buff = buff & TelegramStr(i - 1) & TelegramStr(i)
                            Next i
                            result = buff & ";" & CStr(Value) & ";" & CStr(Pvalue) & ";" & TelegramStr & ";"
                            add_to_Serials(result)
                            If debug = 1 Then
                                fout.WriteLine(buff & ";" & CStr(Value) & ";" & TelegramStr)
                                fout.Flush()
                            End If
                        ElseIf RadioButton2.Checked = True And TelegramStr(35) = "7" Then

                            Dim sevalue As String
                            Dim oevalue As String
                            Dim pevalue As Integer
                            Dim vealue As String
                            Dim DPvalue As Single

                            sevalue = ""
                            oevalue = ""
                            vealue = ""

                            For i = 85 To 77 Step -2
                                sevalue = sevalue & TelegramStr(i - 1) & TelegramStr(i)
                            Next i
                            Dim str As String
                            str = "."
                            sevalue = sevalue.Insert(7, str)
                            vealue = sevalue
                            For i = 125 To 118 Step -2
                                oevalue = oevalue & TelegramStr(i - 1) & TelegramStr(i)
                            Next
                            oevalue = "&H" & oevalue
                            pevalue = CInt(oevalue)

                            DPvalue = pevalue / 10

                            For i = 31 To 24 Step -2
                                buff = buff & TelegramStr(i - 1) & TelegramStr(i)
                            Next i
                            If debug = 1 Then
                                fout.WriteLine(buff & ";" & CStr(Value) & ";" & TelegramStr)
                                fout.Flush()
                            End If
                            result = buff & ";" & vealue & ";" & CStr(DPvalue) & ";" & TelegramStr & ";"
                            add_to_Serials(result)
                        End If


                    ElseIf MBT1BLUE.RADManufacturer = "QDS" And RadioButton3.Checked Then

                        Dim status As String = ""
                        If MBT1BLUE.RADDatarecordUnit(1) = "HCA" Then
                            If MBT1BLUE.RADStatus = "18" Then
                                status = "OPEN"
                            Else
                                status = "OK"
                            End If
                            result = MBT1BLUE.RADDeviceAddress & ";" & CInt(MBT1BLUE.RADDatarecordValue(1)) & ";" & CInt(MBT1BLUE.RADDatarecordValue(2)) & ";" & status & "-" & TelegramStr & ";"
                            add_to_Serials(result)

                        End If
                    ElseIf MBT1BLUE.RADManufacturer = "SON" And RadioButton4.Checked Then
                        Dim status As String = ""
                        If MBT1BLUE.RADDatarecordUnit(2) = "HCA" Then
                            If MBT1BLUE.RADStatus = "18" Then
                                status = "OPEN"
                            Else
                                status = "OK"
                            End If

                            result = MBT1BLUE.RADDeviceAddress & ";" & (MBT1BLUE.RADDatarecordValue(2)) & ";" & (MBT1BLUE.RADDatarecordValue(4)) & ";" & status & "-" & TelegramStr & ";"
                            add_to_Serials(result)
                        End If
                    ElseIf MBT1BLUE.RADManufacturer = "EFE" And RadioButton6.Checked Then
                        Dim status As String = ""
                        If MBT1BLUE.RADDatarecordUnit(1) = "HCA" Then
                            If MBT1BLUE.RADStatus <> "00" Then
                                status = "OPEN"
                            Else
                                status = "OK"
                            End If

                            result = MBT1BLUE.RADDeviceAddress & ";" & (MBT1BLUE.RADDatarecordValue(1)) & ";" & (MBT1BLUE.RADDatarecordValue(2)) & ";" & status & "-" & TelegramStr
                            Record = MBT1BLUE.RADNumberOfDataRecords
                            Dim count As Integer
                            For count = 1 To Record
                                result = result & "|" & MBT1BLUE.RADDatarecordValue(count)
                            Next
                            result = result & ";"
                            add_to_Serials(result)
                        End If
                    End If


                    End If
                    System.Threading.Thread.Sleep(10)
                Loop While BackgroundWorker1.CancellationPending = False 'Now.Second < StartTime + 15         'read the MBT1BLUE device for 15 seconds
                MBT1BLUE.CommunicationThreadBreak = 1

                If debug = 1 Then
                    fout.Close()
                End If
            Else
                If RadioButton1.Checked = True Then
                Dim TelegramStr As String


                TelegramStr = "251302060E27135F1C440186093209360308A0EE461794341C2B99EBDA5568D9AACB3DDFEE101112131415161718191A1B1C1D1E1F202122232425262728292A2B2C2D2E2F303132333435363738393A3B3C3D3E3F4041424344454647"
                If Strings.Left(TelegramStr, 2) <> "FF" Then
                    Dim i As Integer
                    Dim Value As Integer
                    Dim svalue As String
                    Dim ovalue As String
                    Dim Pvalue As Integer



                    buff = ""


                    svalue = "&H" & TelegramStr(52) & TelegramStr(53) & TelegramStr(50) & TelegramStr(51)
                    ovalue = "&H" & TelegramStr(56) & TelegramStr(57) & TelegramStr(54) & TelegramStr(55)

                    Value = CInt(svalue)
                    Pvalue = CInt(ovalue)
                    For i = 31 To 24 Step -2
                        buff = buff & TelegramStr(i - 1) & TelegramStr(i)
                    Next i
                    If debug = 1 Then
                        fout.WriteLine(buff & ";" & CStr(Value) & ";" & TelegramStr)
                        fout.Flush()
                    End If
                    result = buff & ";" & CStr(Value) & ";" & CStr(Pvalue) & ";" & TelegramStr & ";"
                    add_to_Serials(result)
                End If
            ElseIf RadioButton2.Checked = True Then
                Dim TelegramStr As String

                TelegramStr = "56120C1110040A464D440186356202001537728909610101064004A1000020026C512C0E014664478916000C13438874010A5A25050A5E08054405DC7B020001FD0C010A6577330AFD4737"
                If Strings.Left(TelegramStr, 2) <> "FF" Then
                    Dim i As Integer
                    Dim Value As String
                    Dim svalue As String
                    Dim ovalue As String
                    Dim Pvalue As Integer

                    Value = ""
                    svalue = ""
                    ovalue = ""

                    buff = ""

                    For i = 85 To 77 Step -2
                        svalue = svalue & TelegramStr(i - 1) & TelegramStr(i)
                    Next i
                    Dim str As String
                    str = "."
                    svalue = svalue.Insert(7, str)
                    Value = svalue
                    For i = 125 To 118 Step -2
                        ovalue = ovalue & TelegramStr(i - 1) & TelegramStr(i)
                    Next
                    ovalue = "&H" & ovalue
                    Pvalue = CInt(ovalue)
                    Dim DPvalue As Single
                    DPvalue = Pvalue / 10

                    For i = 31 To 24 Step -2
                        buff = buff & TelegramStr(i - 1) & TelegramStr(i)
                    Next i
                    If debug = 1 Then
                        fout.WriteLine(buff & ";" & Value & ";" & TelegramStr)
                        fout.Flush()
                    End If
                    result = buff & ";" & CStr(Value) & ";" & CStr(DPvalue) & ";" & TelegramStr & ";"
                    add_to_Serials(result)
                End If
            Else

            End If





        End If



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button1.BackColor = Color.Green
        BackgroundWorker1.CancelAsync()
        Timer1.Stop()
        Me.Refresh()

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim fin As StreamReader

        Dim i As Integer
        saved = 1
        Serials.Clear()
        DataGridView1.Rows.Clear()
        OpenFileDialog1.ShowDialog()
        '
        If OpenFileDialog1.FileName = "" Then
            Exit Sub
        End If
        fin = File.OpenText(OpenFileDialog1.FileName)

        While Not fin.EndOfStream


            DataGridView1.Rows.Add(fin.ReadLine.Split(";"))
        End While
        fin.Close()



        For i = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(i).Cells(2).Value <> "" Then

                Serials.Add(DataGridView1.Rows(i).Cells(0).Value & ";" & DataGridView1.Rows(i).Cells(2).Value & ";" & DataGridView1.Rows(i).Cells(3).Value & ";" & DataGridView1.Rows(i).Cells(4).Value)
            End If

        Next

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim out As StreamWriter
        Dim i As Integer
        SaveFileDialog1.ShowDialog()
        out = File.CreateText(SaveFileDialog1.FileName)
        For i = 0 To DataGridView1.Rows.Count - 2
            out.WriteLine(DataGridView1.Rows(i).Cells(0).Value & ";" & DataGridView1.Rows(i).Cells(1).Value & ";" & DataGridView1.Rows(i).Cells(2).Value & ";" & DataGridView1.Rows(i).Cells(3).Value & ";" & DataGridView1.Rows(i).Cells(4).Value & ";" & DataGridView1.Rows(i).Cells(5).Value)
        Next
        out.Close()
        MsgBox("Salvato!")
        saved = 1
    End Sub
    Private Sub add_to_Serials(imp As String)
        Dim buff As String
        Dim inn() As String = imp.Split(";")
        Dim innn() As String
        Dim found As Integer = 0

        For Each buff In Serials
            innn = buff.Split(";")
            If inn(0) = innn(0) Then
                found = 1
                Exit For

            End If
        Next
        If found = 0 Then
            Serials.Add(imp)
            saved = 0
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim buff As String
        Dim innn() As String
        Dim enc As String
        Dim dec As String
        Dim svalue As String
        Dim ovalue As String
        Dim value As Integer
        Dim pvalue As Integer

        Dim found As Integer = 0
        On Error Resume Next
        If Serials.Count <> 0 Then
            For Each buff In Serials
                innn = buff.Split(";")
                If innn(3).Substring(0, 2) = "25" Then


                    If MaskedTextBox1.Text <> "" And MaskedTextBox1.Text.Length = 32 Then
                        enc = innn(3).Substring(38, 32)


                        WebBrowser1.Document.GetElementById("ciphertext").SetAttribute("value", enc)
                        WebBrowser1.Document.GetElementById("key").SetAttribute("value", MaskedTextBox1.Text)
                        WebBrowser1.Document.InvokeScript("doDecryption")
                        dec = WebBrowser1.Document.GetElementById("plaintext").GetAttribute("value")
                        svalue = "&H" & dec(18) & dec(19) & dec(16) & dec(17)
                        ovalue = "&H" & dec(22) & dec(23) & dec(20) & dec(21)

                        value = CInt(svalue)
                        pvalue = CInt(ovalue)
                        innn(1) = value
                        innn(2) = pvalue
                    Else
                        innn(1) = "Criptato"
                        innn(2) = "Criptato"
                        dec = "000001111111111111111111111111111"
                    End If

                End If


                For i = 0 To DataGridView1.Rows.Count - 1
                    found = 0

                    If DataGridView1.Rows(i).Cells(0).Value = innn(0) Then

                        If RadioButton1.Checked Then
                            Dim code As String
                            Dim ddmm As UInt32
                            Dim dd As UInt32
                            Dim mm As UInt32
                            Dim bs As String
                            Dim bbs(35) As Char
                            Dim ii As Integer
                            Dim st As String
                            dd = 0
                            mm = 0
                            ddmm = 0
                            If innn(3).Substring(0, 2) = "25" Then
                                bs = dec.Substring(6, 2) & dec.Substring(4, 2)
                            Else
                                bs = innn(3).Substring(40, 2) & innn(3).Substring(38, 2)
                            End If

                            ddmm = CInt("&H" & bs)
                            code = Convert.ToString(ddmm, 2).PadLeft(32, "0"c)
                            bbs = code.ToCharArray
                            'MsgBox(bbs)
                            For ii = 32 To 28 Step -1
                                If bbs(ii - 1) = "1" Then
                                    dd = dd + 2 ^ (32 - ii)
                                End If

                            Next
                            For ii = 0 To 3
                                If bbs(26 - ii) = "1" Then
                                    mm = mm + 2 ^ (ii)
                                End If
                            Next
                            If innn(3).Substring(0, 2) = "25" Then
                                bs = dec.Substring(14, 2) & dec.Substring(12, 2)
                            Else
                                bs = innn(3).Substring(48, 2) & innn(3).Substring(46, 2)
                            End If

                            ddmm = CInt("&H" & bs)
                            If ddmm <> 0 Then
                                st = "KO"
                            Else
                                st = "OK"
                            End If

                            innn(4) = dd & "/" & mm & "-" & st & "=" & ddmm
                        End If
                        DataGridView1.Rows(i).Cells(2).Value = innn(1)
                        DataGridView1.Rows(i).Cells(3).Value = innn(2)
                        DataGridView1.Rows(i).Cells(4).Value = innn(3)
                        DataGridView1.Rows(i).Cells(5).Value = innn(4)
                        found = 1
                        Exit For
                    End If

                Next
                'found = 0
                If found = 0 Then
                    If RadioButton1.Checked Then
                        Dim code As String
                        Dim ddmm As UInt32
                        Dim dd As UInt32
                        Dim mm As UInt32
                        Dim bs As String
                        Dim bbs(35) As Char
                        Dim ii As Integer
                        Dim st As String
                        dd = 0
                        ddmm = 0
                        mm = 0
                        If innn(3).Substring(0, 2) = "25" Then
                            bs = dec.Substring(6, 2) & dec.Substring(4, 2)
                        Else
                            bs = innn(3).Substring(40, 2) & innn(3).Substring(38, 2)
                        End If
                        ddmm = CInt("&H" & bs)
                        code = Convert.ToString(ddmm, 2).PadLeft(32, "0"c)
                        bbs = code.ToCharArray
                        For ii = 32 To 28 Step -1
                            If bbs(ii - 1) = "1" Then
                                dd = dd + 2 ^ (32 - ii)
                            End If

                        Next
                        For ii = 0 To 3
                            If bbs(26 - ii) = "1" Then
                                mm = mm + 2 ^ (ii)
                            End If
                        Next
                        If innn(3).Substring(0, 2) = "25" Then
                            bs = dec.Substring(14, 2) & dec.Substring(12, 2)
                        Else
                            bs = innn(3).Substring(48, 2) & innn(3).Substring(46, 2)
                        End If
                        ddmm = CInt("&H" & bs)
                        If ddmm <> 0 Then
                            st = "KO"
                        Else
                            st = "OK"
                        End If

                        innn(4) = dd & "/" & mm & "-" & st & "=" & ddmm
                    End If

                    DataGridView1.Rows.Add(innn(0), "", innn(1), innn(2), innn(3), innn(4))
                End If
            Next
        End If


            Me.Refresh()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim url As New Uri("file://C:\AES\aes.html")
        WebBrowser1.Url = url

    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dim R As Integer
        If saved = 0 Then
            R = MsgBox("Vuoi Salvare?", vbYesNo)
        End If
        If R = vbYes Then
            Dim out As StreamWriter
            Dim i As Integer
            SaveFileDialog1.ShowDialog()
            If SaveFileDialog1.FileName <> "" Then
                out = File.CreateText(SaveFileDialog1.FileName)
                For i = 0 To DataGridView1.Rows.Count - 2
                    out.WriteLine(DataGridView1.Rows(i).Cells(0).Value & ";" & DataGridView1.Rows(i).Cells(1).Value & ";" & DataGridView1.Rows(i).Cells(2).Value & ";" & DataGridView1.Rows(i).Cells(3).Value & ";" & DataGridView1.Rows(i).Cells(4).Value)
                Next
                out.Close()
                MsgBox("Salvato!")
                saved = 1
            End If
        End If
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub MaskedTextBox1_Leave(sender As Object, e As EventArgs) Handles MaskedTextBox1.Leave
        If MaskedTextBox1.Text.Length > 0 And MaskedTextBox1.Text.Length <> 32 Then
            MsgBox("Dimensione della chiave errata!")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim TelegramStr As String
        Dim svalue As String
        Dim ovalue As String
        Dim value As Integer
        Dim Pvalue As Integer
        Dim buff As String
        Dim result As String
        Dim i As Integer
        Dim k As Integer
        buff = ""
        result = ""
        For i = 0 To DataGridView1.Rows.Count - 1
            buff = ""

            TelegramStr = DataGridView1.Rows(i).Cells(4).Value

            If TelegramStr <> "" Then
                If RadioButton1.Checked = True And TelegramStr(35) = "8" Then
                    If TelegramStr(33) = "4" Then
                        ovalue = "&H" & TelegramStr(46) & TelegramStr(47) & TelegramStr(44) & TelegramStr(45)
                        svalue = "&H" & TelegramStr(54) & TelegramStr(55) & TelegramStr(52) & TelegramStr(53)
                    ElseIf TelegramStr(33) = "1" Then
                        ovalue = "&H" & TelegramStr(50) & TelegramStr(51) & TelegramStr(48) & TelegramStr(49)
                        svalue = "&H" & TelegramStr(58) & TelegramStr(59) & TelegramStr(56) & TelegramStr(57)

                    Else

                        svalue = "&H" & TelegramStr(52) & TelegramStr(53) & TelegramStr(50) & TelegramStr(51)
                        ovalue = "&H" & TelegramStr(56) & TelegramStr(57) & TelegramStr(54) & TelegramStr(55)
                    End If

                    value = CInt(svalue)
                    Pvalue = CInt(ovalue)
                    For k = 31 To 24 Step -2
                        buff = buff & TelegramStr(k - 1) & TelegramStr(k)
                    Next k
                    result = "dec-" & buff & ";" & CStr(value) & ";" & CStr(Pvalue) & ";" & TelegramStr & ";"

                    add_to_Serials(result)

                End If
            End If
            result = ""
        Next
        Timer1.Start()

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged

    End Sub
End Class
