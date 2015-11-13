Public Class TruncateNotes
    Private NewSTRG As String
    Public Property NewSTRING() As String
        Get
            Return NewSTRG

        End Get
        Set(ByVal value As String)
            NewSTRG = value
        End Set
    End Property
    Public Sub Truncate(ByVal notes As String, ByVal ctr As Control)
        Dim i As Integer
        If notes <> Nothing Then
            Dim h As Integer = notes.Length / 35
            Dim newstr As String = notes

            If notes.Length > 35 Then
                For i = 1 To h
                    If i = 1 Then
                        newstr = Microsoft.VisualBasic.Left(newstr, i * 35)
                        Dim tr = InStrRev(newstr, " ")
                        newstr = Microsoft.VisualBasic.Left(newstr, tr)
                        newstr = RTrim(newstr)
                        newstr = newstr & vbCr & Microsoft.VisualBasic.Mid(notes, tr + 1)
                    Else

                        Dim tr = InStrRev(newstr, Chr(13))
                        If tr + 35 > notes.Length Then
                            Exit For
                        End If
                        newstr = Microsoft.VisualBasic.Left(newstr, tr + 35)
                        Dim t = InStrRev(newstr, " ")
                        newstr = Microsoft.VisualBasic.Left(newstr, t)
                        newstr = RTrim(newstr)
                        newstr = newstr & vbCr & Microsoft.VisualBasic.Mid(notes, t + 1)
                    End If
                Next
                Me.NewSTRING = newstr
            Else
                Me.NewSTRING = notes
            End If
        End If
    End Sub
    
End Class
