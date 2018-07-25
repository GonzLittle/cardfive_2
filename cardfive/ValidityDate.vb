Imports System.ComponentModel
Imports System.Data.OleDb

Imports MetroFramework
Public Class ValidityDate
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Dim _date As String = TextBox1.Text.Trim

        Try


            If e.KeyCode = Keys.Enter Then
                If _date = "" Then
                    MetroMessageBox.Show(Me, vbCrLf & "No details!" & vbCrLf & vbCrLf & vbCrLf & "©(PGLU-HRMD Series of 2017)", "Date Change", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    Form1.Button6.Text = _date
                End If
                Me.Close()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ValidityDate_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Me.Close()

    End Sub

    Private Sub ValidityDate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BringToFront()
    End Sub
End Class