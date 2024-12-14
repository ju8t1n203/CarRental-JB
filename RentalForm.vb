'Justin Bell
'RCET0265
'Car Rental
'https://github.com/ju8t1n203/CarRental-JB

Option Explicit On
Option Strict On
Option Compare Text

Imports System.Runtime.CompilerServices
Imports System.Web

Public Class RentalForm
    Private customers As Integer = 1
    'buttons-------------------------------
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
        Else
            CostCalculator()
        End If

    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

        NameTextBox.Text = Nothing
        AddressTextBox.Text = Nothing
        CityTextBox.Text = Nothing
        StateTextBox.Text = Nothing
        ZipCodeTextBox.Text = Nothing
        BeginOdometerTextBox.Text = Nothing
        EndOdometerTextBox.Text = Nothing
        DaysTextBox.Text = Nothing
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
        Seniorcheckbox.Checked = False
        AAAcheckbox.Checked = False
        TotalMilesTextBox.Text = Nothing
        MileageChargeTextBox.Text = Nothing
        DayChargeTextBox.Text = Nothing
        TotalDiscountTextBox.Text = Nothing
        TotalChargeTextBox.Text = Nothing

        customers = customers + 1

    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Dim miles As Double
        Dim profits As Double

        miles = CDbl(CStr(TotalMilesTextBox.Text).Replace(" mi", "").Trim()) + miles
        profits = CDbl(CStr(TotalChargeTextBox.Text).Replace("$", "").Trim()) + profits

        MessageBox.Show($"Total Customers Served: {customers}
Total Miles Driven: {miles}
Total Profit: {profits}")

        ClearButton.PerformClick()
    End Sub


    'top menu operations------------------------
    Private Sub CalculateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculateToolStripMenuItem.Click
        CalculateButton.PerformClick()
    End Sub

    Private Sub ClearToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem1.Click
        ClearButton.PerformClick()
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        ExitButton.PerformClick()
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        MsgBox("This form is used to calculate the cost of renting a car from ACME Car Rentals. 
The zip code must be 5 digits. 
The end odometer reading must exceed the beginning reading.
The days must be a whole number.")
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("my name is justin bell, I created the software for this form. I think I deserve an A.")
    End Sub

    'to be called upon---------------------------
    Sub CostCalculator()
        Dim distance As Double
        Dim diCharge As Double
        Dim days As Integer
        Dim daCharge As Integer
        Dim discount As Double
        Dim total As Double

        If MilesradioButton.Checked = True Then
            distance = CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)
        Else
            distance = Math.Round((CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)) / 1.609, 3)
        End If

        TotalMilesTextBox.Text = $"{distance} mi"

        Select Case distance
            Case 0 To 200
                diCharge = 0
            Case 201 To 500
                diCharge = distance * 0.12
            Case > 500
                diCharge = distance * 0.1
        End Select

        MileageChargeTextBox.Text = diCharge.ToString("C")

        days = CInt(DaysTextBox.Text)

        daCharge = days * 15

        DayChargeTextBox.Text = daCharge.ToString("C")

        If Seniorcheckbox.Checked = True Then
            discount = discount + 0.03
        End If

        If AAAcheckbox.Checked = True Then
            discount = discount + 0.05
        End If

        total = (1 - discount) * (diCharge + daCharge)

        TotalDiscountTextBox.Text = (discount * total).ToString("C")

        TotalChargeTextBox.Text = total.ToString("C")

        SummaryButton.Enabled = True

    End Sub

End Class