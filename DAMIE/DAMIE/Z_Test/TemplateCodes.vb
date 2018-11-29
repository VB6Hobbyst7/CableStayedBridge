Imports DAMIE.Miscellaneous
Public Class TemplateCodes
    Inherits System.Windows.Forms.Form

    Private Sub TemplateCodes_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim i As Integer
        For Each i In test()
            Debug.Print(i)
        Next
    End Sub
    Function test() As Integer()
        Static i As Integer
        i += 1
        Return {i, i + 1, i + 2}

    End Function
    Private Sub ModelCode()

        ' -------- Tye 语句 --------------------
        Try
        Catch ex As Exception
            MessageBox.Show("" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)


            MessageBox.Show("" & vbCrLf & ex.Message & vbCrLf & "报错位置：" &
                            ex.TargetSite.Name, "Error", _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try



    End Sub


End Class


