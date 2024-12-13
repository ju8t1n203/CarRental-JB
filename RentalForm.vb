Option Explicit On
Option Strict On
Option Compare Text
'Justin Bell
'RCET0265
'Car Rental
'link

Imports System.Web

Public Class RentalForm


    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        'prompts the user if they want to close the form
        Dim result As DialogResult = MessageBox.Show(
            "Are you sure you want to close the form?",
            "Confirm Exit",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Me.Close()
        Else
            Return
        End If

    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim incorrect As String = ""
        Dim zip As Integer
        If NameTextBox.Text IsNot Text Then
            incorrect &= "Enter a valid name." & vbCrLf
        End If

        If CBool(AddressTextBox.Text) = False Then
            incorrect &= "Enter an address." & vbCrLf
        End If

        If StateTextBox.Text IsNot Text Then
            incorrect &= "Enter a valid state." & vbCrLf
        End If

        If Integer.TryParse(ZipCodeTextBox.Text, zip) Then
            If zip > 99999 Then
                incorrect &= "Enter a 5-digit zip code." & vbCrLf
            End If
        End If



    End Sub
End Class
