Imports System.Drawing.Printing
Imports System.Drawing
Imports System.Drawing.Drawing2D


Public Class Display
    Public printDoc As System.Drawing.Printing.PrintDocument

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.Close()
        Me.Text = " Summary Report By Rep"
        Me.Dispose()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dlg As New PrintPreviewDialog
        With dlg
            .Document = printDoc
            .ShowDialog()
        End With
    End Sub


    Private Sub Display_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        printDoc = New PrintDocument

        printDoc.DocumentName = "Test Print Object Model"

        AddHandler printDoc.BeginPrint, New PrintEventHandler(AddressOf Me._beginPrint)
        AddHandler printDoc.EndPrint, New PrintEventHandler(AddressOf Me._endPrint)
        AddHandler printDoc.PrintPage, New PrintPageEventHandler(AddressOf Me._printPage)

    End Sub
    Private Sub _beginPrint(ByVal sender As Object, ByVal e As PrintEventArgs)

    End Sub
    Private Sub _endPrint(ByVal sender As Object, ByVal e As PrintEventArgs)

    End Sub
    Private Sub _printPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)

        '' actual method that prints
        '' graphics = drawimage
        '' if you draw text, g.drawtext and so on....
        ''

        Dim g As Graphics = e.Graphics ' used to pass whatever you are printing to the GDI engine for printing

        If Me.PictureBox1.Image IsNot Nothing Then
            '' we wanted to print and image of the picture box so
            '' g.drawimage() method was used.
            '' similar to what we are trying to achieve from the AxMappointControl
            '' 1) set print options
            '' 2) send object with options to printer
            '' 3) print AxMappointControl.ActiveMap w/ set configurable options.
            '' 
            g.DrawImage(PictureBox1.Image, 20, 60, PictureBox1.Size.Width, PictureBox1.Size.Height)
        End If

    End Sub
End Class
