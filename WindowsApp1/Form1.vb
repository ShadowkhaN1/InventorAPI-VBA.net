Imports Inventor
Imports System.Runtime.InteropServices



Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim inventorApp As Inventor.Application

        Try
            inventorApp = Marshal.GetActiveObject("Inventor.Application")
            MessageBox.Show("Connect with Inventor")

        Catch ex As Exception
            MessageBox.Show("Cannot connect to Inventor")
        End Try

        Dim oPartDoc As PartDocument

        oPartDoc = inventorApp.ActiveDocument

        oPartDoc.ComponentDefinition.Parameters.Item("d0").Value = Convert.ToInt32(TextBox1.Text)
        oPartDoc.ComponentDefinition.Parameters.Item("d1").Value = Convert.ToInt32(TextBox2.Text)
        oPartDoc.ComponentDefinition.Parameters.Item("d2").Value = Convert.ToInt32(TextBox3.Text)

        oPartDoc.Update()

    End Sub
End Class
