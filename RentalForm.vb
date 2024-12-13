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
        Dim startOD As Integer
        Dim endOD As Integer
        Dim days As Integer

        If String.IsNullOrWhiteSpace(NameTextBox.Text) Then
            incorrect &= "Enter a valid name." & vbCrLf
        End If

        If String.IsNullOrWhiteSpace(AddressTextBox.Text) Then
            incorrect &= "Enter an address." & vbCrLf
        End If

        If String.IsNullOrWhiteSpace(StateTextBox.Text) Then
            incorrect &= "Enter a valid state." & vbCrLf
        End If

        If Integer.TryParse(ZipCodeTextBox.Text, zip) Then
            If zip > 99999 Or zip < 0 Then
                incorrect &= "Enter a 5-digit zip code." & vbCrLf
            End If
        Else
            If String.IsNullOrWhiteSpace(ZipCodeTextBox.Text) Then
                incorrect &= "Enter a zip code." & vbCrLf
            Else
                incorrect &= "Enter a numeric zip code." & vbCrLf
            End If
        End If

        If Integer.TryParse(BeginOdometerTextBox.Text, startOD) Then
            If startOD < 0 Then
                incorrect &= "Enter a valid odometer reading." & vbCrLf
            End If
        Else
            If String.IsNullOrWhiteSpace(BeginOdometerTextBox.Text) Then
                incorrect &= "Enter valid odometer reading." & vbCrLf
            Else
                incorrect &= "Enter a valid odometer reading." & vbCrLf
            End If
        End If

        If Integer.TryParse(EndOdometerTextBox.Text, endOD) Then
            If startOD < 0 Then
                incorrect &= "Enter a valid odometer reading." & vbCrLf
            End If
        Else
            If String.IsNullOrWhiteSpace(BeginOdometerTextBox.Text) Then
                incorrect &= "Enter valid odometer reading." & vbCrLf
            Else
                incorrect &= "Enter a valid odometer reading." & vbCrLf
            End If
        End If

        If startOD > endOD Then
            incorrect &= "End odometer reading must be larger than beginning odometer reading." & vbCrLf
        End If

        If Integer.TryParse(DaysTextBox.Text, days) Then
            If days > 45 Or days < 0 Then
                incorrect &= "Rental duration exceeds maxium allowed time" & vbCrLf
            End If
        Else
            If String.IsNullOrWhiteSpace(DaysTextBox.Text) Then
                incorrect &= "Enter a rental duration." & vbCrLf
            Else
                incorrect &= "Enter a number of days to rent." & vbCrLf
            End If
        End If

        If incorrect <> "" Then
            MessageBox.Show(incorrect, "Missing Fields", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

End Class
