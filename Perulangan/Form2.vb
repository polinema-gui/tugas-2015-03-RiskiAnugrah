Public Class Form2
  
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Select Case ComboBox1.Text
            Case "Pyramid"
                pyramid()
            Case "Hollow Pyramid"
                hollowpyramid()
            Case "Inverted Pyramid"
                invertedpyramid()
            Case "Hollow Inverted Pyramid"
                hollowinvertedpyramid()
            Case Else
                TextBox1.Text = "Pilih pola yang anda inginkan"
        End Select
    End Sub
    Private Sub pyramid()
        Dim Cetak As String = ""
        Dim batas As Integer = TextBox1.Text
        Dim batasKolom1 As Integer = batas
        Dim batasKolom2 As Integer = batas
        For baris As Integer = 1 To batas
            For kolom As Integer = 1 To batas * 2 - 1
                If (kolom < batasKolom1 Or kolom > batasKolom2) Then
                    Cetak &= "  "
                Else
                    Cetak &= " *"
                End If
            Next
            batasKolom1 -= 1
            batasKolom2 += 1
            Cetak &= vbCrLf
        Next
        TextBox2.Text = Cetak
        Cetak = ""
    End Sub
    Private Sub hollowpyramid()
        Dim cetak As String = ""
        Dim batas As Integer = TextBox1.Text
        Dim batasKolom1 As Integer = batas
        Dim batasKolom2 As Integer = batas
        For baris As Integer = 1 To batas
            For kolom As Integer = 1 To batas * 2 - 1
                If (baris < batas) Then
                    If (kolom = batasKolom1 Or kolom = batasKolom2) Then
                        cetak &= " *"
                    Else
                        cetak &= "  "
                    End If
                Else
                    cetak &= " *"
                End If
            Next
            batasKolom1 -= 1
            batasKolom2 += 1
            cetak &= vbCrLf
        Next
        TextBox2.Text = cetak
        cetak = ""
    End Sub
    Private Sub invertedpyramid()
        Dim cetak As String = ""
        Dim batas As Integer = TextBox1.Text
        Dim batasKolom1 As Integer = 0
        Dim batasKolom2 As Integer = batas * 2
        For baris As Integer = 1 To batas
            For kolom As Integer = 1 To batas * 2 - 1
                If (kolom > batasKolom1 And kolom < batasKolom2) Then
                    cetak &= " *"
                Else
                    cetak &= "  "
                End If
            Next
            cetak &= vbCrLf
            batasKolom1 += 1
            batasKolom2 -= 1
        Next
        TextBox2.Text = cetak
        cetak = ""
    End Sub
    Private Sub hollowinvertedpyramid()
        Dim cetak As String = ""
        Dim batas As Integer = TextBox1.Text
        Dim batasKolom1 As Integer = 1
        Dim batasKolom2 As Integer = batas * 2 - 1
        For baris As Integer = 1 To batas
            For kolom As Integer = 1 To batas * 2 - 1
                If (baris = 1) Then
                    cetak &= " *"
                Else
                    If (kolom = batasKolom1 Or kolom = batasKolom2) Then
                        cetak &= " *"
                    Else
                        cetak &= "  "
                    End If
                End If
            Next
            batasKolom1 += 1
            batasKolom2 -= 1
            cetak &= vbCrLf
        Next
        TextBox2.Text = cetak
        cetak = ""
    End Sub

End Class