' Calculator 001 National Cancer Intitute - Maharagama
' This program was developed by L.Swarnajith for the free use of physicists at NCISL, demonstrating
' heartfelt support for cancer research and treatment at the hospital.
' L.Swarnajith JUN/2023

Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class frmMain

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        loadMachines()
        loadNdw()

        cmbV1.Text = "300"
        cmbV2.Text = "100"
    End Sub

    Private Sub lblTitle_MouseDown(sender As Object, e As MouseEventArgs) Handles lblTitle.MouseDown
        On Error Resume Next

        isDragging = True
        mouseOffset = New Point(-e.X, -e.Y)

    End Sub

    Private Sub lblTitle_MouseMove(sender As Object, e As MouseEventArgs) Handles lblTitle.MouseMove
        On Error Resume Next

        If isDragging Then
            Dim mousePos = Control.MousePosition
            mousePos.Offset(mouseOffset.X, mouseOffset.Y)
            Me.Location = mousePos
        End If

    End Sub

    Private Sub lblTitle_MouseUp(sender As Object, e As MouseEventArgs) Handles lblTitle.MouseUp
        On Error Resume Next
        isDragging = False
    End Sub

    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click
        Dim usr = MsgBox("Confirm Exit", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "Exit")
        If usr = MsgBoxResult.Yes Then End
    End Sub

    Private Sub btnHlp_Click(sender As Object, e As EventArgs) Handles btnHlp.Click
        MsgBox("Please contact +9477-591-7771 or lakmilaswarnajith@gmail.com", vbInformation)
    End Sub

    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click

        On Error GoTo ErrHandler

        Dim AP100, AP33, AN100 As Single

        Dim T As Single = Convert.ToSingle(txtTemp.Text)
        Dim P As Single = Convert.ToSingle(txtPressure.Text)

        removeEmptyLines(txt33)
        removeEmptyLines(txtP100)
        removeEmptyLines(txtN100)
        Dim l As Byte = txtP100.Lines.Length

        If (l <> txtN100.Lines.Length) Or (l <> txt33.Lines.Length) Then
            AP100 = Convert.ToSingle(txtP100.Lines(0))
            AN100 = Convert.ToSingle(txtP100.Lines(0))
            AP33 = Convert.ToSingle(txtP100.Lines(0))
        Else
            For i = 0 To (l - 1)
                AP100 += Convert.ToSingle(txtP100.Lines(i))
                AN100 += Convert.ToSingle(txtN100.Lines(i))
                AP33 += Convert.ToSingle(txt33.Lines(i))
            Next

            AP100 /= l
            AN100 /= l
            AP33 /= l
        End If

        a0 = Convert.ToSingle(cmbA0.Text)
        a1 = Convert.ToSingle(cmbA1.Text)
        a2 = Convert.ToSingle(cmbA2.Text)

        Dnw = Convert.ToSingle(cmbDNW.Text) * 10000000
        Kq = Convert.ToSingle(cmbKq.Text)
        PDD = Convert.ToSingle(cmbPDD.Text)

        M1 = findAbsSingle(AP100)
        M2 = findAbsSingle(AP33)

        Ktp = (273.2 + T) * P0 / ((273.2 + T0) * P)
        Kpol = (AP100 + AN100) / (2 * AP100)
        Ks = a0 + a1 * (AP100 / AP33) + a2 * (AP100 / AP33) * (AP100 / AP33)

        M = Ktp * Kpol * Ks * Ke * Hp * Kq * Dnw
        DW = (M * AP100) / (PDD * 100000000)

        lblSumP100.Text = AP100
        lblSumN100.Text = AN100
        lblSum33.Text = AP33
        lblKPOL.Text = Kpol
        lblKS.Text = Ks
        lblKTP.Text = Ktp
        lblMU.Text = DW

        Dim usr0 = MsgBox("Calculation Completed. Do you want to export the report as a PDF ?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Done")
        If usr0 = MsgBoxResult.Yes Then generateAndExport()

ErrHandler:
        If Err.Number <> 0 Then
            Dim usr1 = MsgBox(Err.Number & " : " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.AbortRetryIgnore, "Error")

            If usr1 = MsgBoxResult.Abort Then
                Exit Sub
            ElseIf usr1 = MsgBoxResult.Retry Then
                Resume
            ElseIf usr1 = MsgBoxResult.Ignore Then
                Resume Next
            End If
        End If

    End Sub

    Private Sub cmbV1_TextChanged(sender As Object, e As EventArgs) Handles cmbV1.TextChanged
        If cmbV1.Text <> Nothing And cmbV2.Text <> Nothing Then filterAConst()
        lblV1P.Text = "+" + cmbV1.Text + " V"
        lblV1N.Text = "-" + cmbV1.Text + " V"
    End Sub

    Private Sub cmbV2_TextChanged(sender As Object, e As EventArgs) Handles cmbV2.TextChanged
        If cmbV1.Text <> Nothing And cmbV2.Text <> Nothing Then filterAConst()
        lblV2P.Text = "+" + cmbV2.Text + " V"
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txt33.Clear()
        txtN100.Clear()
        txtP100.Clear()
        txtPhysicist.Clear()
        txtPressure.Clear()
        txtTemp.Clear()

        cmbA0.ResetText()
        cmbA1.ResetText()
        cmbA2.ResetText()
        cmbDNW.ResetText()
        cmbKq.ResetText()
        cmbMachine.ResetText()
        cmbPDD.ResetText()
        cmbV1.ResetText()
        cmbV2.ResetText()
    End Sub

    Private Sub removeEmptyLines(txtBox As TextBox)
        On Error Resume Next
        Dim lines As String() = txtBox.Text.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
        txtBox.Text = String.Join(vbCrLf, lines)
    End Sub

    Private Function findAbsSingle(input As Single) As Single
        On Error Resume Next
        If input < 0 Then
            Return (-1 * input)
        Else
            Return input
        End If

    End Function

    Private Sub filterAConst()

        On Error Resume Next
        Dim R As Single = Math.Round(Convert.ToSingle(cmbV1.Text) / Convert.ToSingle(cmbV2.Text), 1)

        cmbA0.Items.Clear()
        cmbA1.Items.Clear()
        cmbA2.Items.Clear()

        If R = 2.0 Then
            cmbA0.Items.Add("2.337")
            cmbA1.Items.Add("-3.636")
            cmbA2.Items.Add("2.299")

        ElseIf R = 2.5 Then
            cmbA0.Items.Add("1.474")
            cmbA1.Items.Add("-1.587")
            cmbA2.Items.Add("1.114")

        ElseIf R = 3.0 Then
            cmbA0.Items.Add("1.198")
            cmbA1.Items.Add("-0.875")
            cmbA2.Items.Add("0.677")

        ElseIf R = 3.5 Then
            cmbA0.Items.Add("1.080")
            cmbA1.Items.Add("-0.542")
            cmbA2.Items.Add("0.463")

        ElseIf R = 4.0 Then
            cmbA0.Items.Add("1.022")
            cmbA1.Items.Add("-0.363")
            cmbA2.Items.Add("0.341")

        ElseIf R = 5.0 Then
            cmbA0.Items.Add("0.975")
            cmbA1.Items.Add("-0.188")
            cmbA2.Items.Add("0.214")

        Else
            cmbA0.ResetText()
            cmbA1.ResetText()
            cmbA2.ResetText()

        End If

        cmbA0.Text = cmbA0.Items(0)
        cmbA1.Text = cmbA1.Items(0)
        cmbA2.Text = cmbA2.Items(0)

    End Sub

    Public Sub generateAndExport()

        On Error GoTo ErrHandler

        Dim fileName As String = vbNullString
        Using sfd As New SaveFileDialog()
            sfd.Filter = "PDF Files (*.pdf) | *.pdf"
            sfd.ShowDialog()
            fileName = sfd.FileName
        End Using

        If fileName = vbNullString Then Exit Sub

        Dim doc As New Document()
        Dim writer As PdfWriter = PdfWriter.GetInstance(doc, New FileStream(fileName, FileMode.Create))
        Dim pageEventHandler As New CustomPageEventHandler()
        writer.PageEvent = pageEventHandler

        doc.Open()

        Dim titleFont As Font = FontFactory.GetFont(FontFactory.COURIER_BOLD, 16, BaseColor.BLACK)
        Dim contentFont As Font = FontFactory.GetFont(FontFactory.COURIER, 12, BaseColor.BLACK)
        Dim outputFont As Font = FontFactory.GetFont(FontFactory.COURIER_BOLD, 12, BaseColor.RED)
        Dim footingFont As Font = FontFactory.GetFont(FontFactory.COURIER_BOLD, 8, BaseColor.BLACK)
        titleFont.SetStyle(FontStyle.Underline)

        Dim title As New Paragraph("Output Report" + vbNewLine, titleFont)
        title.Alignment = Element.ALIGN_CENTER


        Dim content1 As New Paragraph(vbNewLine + "Timestamp                   : " + Date.Today.ToString("dd/MM/yyyy") + vbNewLine +
                                                  "Machine                     : " + cmbMachine.Text + " MV", contentFont)

        Dim content2 As New Paragraph(vbNewLine & "Nominal Voltage             : " & cmbV1.Text & " V" & vbNewLine &
                                      "Temperature                 : " & txtTemp.Text & " °C" & vbNewLine &
                                      "Pressure                    : " & txtPressure.Text & " hPa" & vbNewLine &
                                      "Beam Quality Factor         : " & Kq & vbNewLine &
                                      "Ion Recombination Factor    : " & Ks & vbNewLine &
                                      "Polarity Correction Factor  : " & Kpol & vbNewLine &
                                      "T/P Correction Factor       : " & Ktp & vbNewLine &
                                      "Dosimeter Calibration Factor: " & Dnw & vbNewLine &
                                      "PDD                         : " & PDD & vbNewLine & vbNewLine, contentFont)

        Dim content3 As New Paragraph("Output                      : " & DW.ToString & " cGy/MU" & vbNewLine & vbNewLine, outputFont)
        Dim content4 As New Paragraph("Physicist(s)-in-Charge: " & vbNewLine & vbNewLine & String.Join((vbNewLine & vbNewLine), txtPhysicist.Text.Split(",")), contentFont)

        Dim line1 As PdfContentByte = writer.DirectContent
        line1.SetLineWidth(1)
        line1.SetRGBColorStroke(0, 0, 0)
        line1.MoveTo(doc.LeftMargin, 720)
        line1.LineTo(doc.PageSize.Width - doc.RightMargin, 720)
        line1.Stroke()

        Dim line2 As PdfContentByte = writer.DirectContent
        line2.SetLineWidth(1)
        line2.SetRGBColorStroke(0, 0, 0)
        line2.MoveTo(doc.LeftMargin, 525)
        line2.LineTo(doc.PageSize.Width - doc.RightMargin, 525)
        line2.Stroke()

        Dim line3 As PdfContentByte = writer.DirectContent
        line3.SetLineWidth(1)
        line3.SetRGBColorStroke(0, 0, 0)
        line3.MoveTo(doc.LeftMargin, 500)
        line3.LineTo(doc.PageSize.Width - doc.RightMargin, 500)
        line3.Stroke()

        doc.Add(title)
        doc.Add(content1)
        doc.Add(content2)
        doc.Add(content3)
        doc.Add(content4)

        Dim contentByte As PdfContentByte = writer.DirectContent
        Dim xPosition As Single = doc.LeftMargin
        Dim yPosition As Single = doc.BottomMargin - 10
        Dim footnoteTable As New PdfPTable(1)
        footnoteTable.DefaultCell.Border = Rectangle.TOP_BORDER
        footnoteTable.DefaultCell.Padding = 0
        Dim footnoteCell As New PdfPCell(New Phrase(fiveLines & "This is an auto-generated report. Contact +9477-591-7771 or lakmilaswarnajith@gmail.com for feedback.", footingFont))
        footnoteCell.Border = Rectangle.NO_BORDER
        footnoteCell.Padding = 0
        footnoteTable.AddCell(footnoteCell)
        footnoteTable.TotalWidth = doc.PageSize.Width - doc.LeftMargin - doc.RightMargin
        footnoteTable.WriteSelectedRows(0, -1, xPosition, yPosition, contentByte)
        doc.Add(footnoteTable)

        doc.Close()

        MsgBox("File saved: " & fileName, MsgBoxStyle.Information, "Done!")

ErrHandler:
        If Err.Number <> 0 Then
            Dim usr1 = MsgBox(Err.Number & " : " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.AbortRetryIgnore, "Error")

            If usr1 = MsgBoxResult.Abort Then
                Exit Sub
            ElseIf usr1 = MsgBoxResult.Retry Then
                Resume
            ElseIf usr1 = MsgBoxResult.Ignore Then
                Resume Next
            End If
        End If

    End Sub

    Private Sub loadMachines()

        On Error GoTo ErrHandler
        Dim files As String() = Directory.GetFiles(machinePath)

        For Each f In files
            cmbMachine.Items.Add(Path.GetFileName(f).Split(".")(0))
        Next

ErrHandler:
        If Err.Number <> 0 Then
            If Err.Number <> 9 Then
                Dim usr1 = MsgBox(Err.Number & " : " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.AbortRetryIgnore, "Error")

                If usr1 = MsgBoxResult.Abort Then
                    Exit Sub
                ElseIf usr1 = MsgBoxResult.Retry Then
                    Resume
                ElseIf usr1 = MsgBoxResult.Ignore Then
                    Resume Next
                End If
            End If
        End If
    End Sub

    Private Sub loadParameters()

        On Error GoTo ErrHandler
        Dim parameters As New List(Of String)
        Dim files As String() = Directory.GetFiles(machinePath)
        currentMachine = Path.GetFullPath(files(cmbMachine.SelectedIndex))
        Dim fileContent As String = vbNullString

        Using reader As New StreamReader(currentMachine)
            fileContent = reader.ReadToEnd()
        End Using

        Dim fileLines As String() = fileContent.Split(New String() {vbNewLine}, StringSplitOptions.RemoveEmptyEntries)

        cmbKq.Items.Clear()
        cmbPDD.Items.Clear()

        For i = 1 To (fileLines.Length - 1)
            cmbKq.Items.Add(fileLines(i).Split(",")(0))
            cmbPDD.Items.Add(fileLines(i).Split(",")(1))
        Next

ErrHandler:
        If Err.Number <> 0 Then
            If Err.Number <> 9 Then
                Dim usr1 = MsgBox(Err.Number & " : " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.AbortRetryIgnore, "Error")

                If usr1 = MsgBoxResult.Abort Then
                    Exit Sub
                ElseIf usr1 = MsgBoxResult.Retry Then
                    Resume
                ElseIf usr1 = MsgBoxResult.Ignore Then
                    Resume Next
                End If
            End If
        End If

    End Sub

    Private Sub loadNdw()

        On Error GoTo ErrHandler

        Dim fileContent As String = vbNullString

        Using reader As New StreamReader("Ndw.dat")
            fileContent = reader.ReadToEnd
        End Using

        For Each line In fileContent.Split(New String() {vbNewLine}, StringSplitOptions.RemoveEmptyEntries)
            cmbDNW.Items.Add(line)
        Next

ErrHandler:
        If Err.Number <> 0 Then
            If Err.Number <> 9 Then
                Dim usr1 = MsgBox(Err.Number & " : " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.AbortRetryIgnore, "Error")

                If usr1 = MsgBoxResult.Abort Then
                    Exit Sub
                ElseIf usr1 = MsgBoxResult.Retry Then
                    Resume
                ElseIf usr1 = MsgBoxResult.Ignore Then
                    Resume Next
                End If
            End If
        End If
    End Sub

    Private Sub cmbMachine_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMachine.SelectedIndexChanged
        loadParameters()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        On Error GoTo ErrHandler

        If cmbMachine.Text = vbNullString Then
            MsgBox("Please select a machine", MsgBoxStyle.Exclamation, "No machine")
            Exit Sub
        End If

        Dim process As New Process()
        process.StartInfo.FileName = "excel.exe"
        process.StartInfo.Arguments = $" ""{currentMachine}"""
        process.Start()
        process.WaitForExit()
        process.Close()
        process.Dispose()

        loadParameters()

ErrHandler:
        If Err.Number <> 0 Then
            If Err.Number <> 9 Then
                Dim usr1 = MsgBox(Err.Number & " : " & Err.Description, MsgBoxStyle.Critical + MsgBoxStyle.AbortRetryIgnore, "Error")

                If usr1 = MsgBoxResult.Abort Then
                    Exit Sub
                ElseIf usr1 = MsgBoxResult.Retry Then
                    Resume
                ElseIf usr1 = MsgBoxResult.Ignore Then
                    Resume Next
                End If
            End If
        End If

    End Sub

End Class