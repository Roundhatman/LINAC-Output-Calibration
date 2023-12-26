Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class CustomPageEventHandler
    Inherits PdfPageEventHelper

    Public Overrides Sub OnEndPage(ByVal writer As PdfWriter, ByVal document As Document)

        Dim contentByte As PdfContentByte = writer.DirectContent
        Dim pageRect As Rectangle = document.PageSize

        Dim borderWidth As Single = 1
        Dim borderColor As BaseColor = BaseColor.BLACK

        contentByte.SetLineWidth(borderWidth)
        contentByte.SetRGBColorStroke(borderColor.R, borderColor.G, borderColor.B)
        contentByte.Rectangle(pageRect.Left + 20, pageRect.Bottom + 20, pageRect.Width - 40, pageRect.Height - 40)
        contentByte.Stroke()

    End Sub

End Class
