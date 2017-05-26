'***************************************************************
' Programmer: Afonin, Anthony
' Chemeketa Community College
' Date: 3/12/2016
' Class: CIS133VB
' Assignment: Final Project, Health Calculator
' File Name: frmInput.vb
' Description: The input form that stores the data from the user
'
'***************************************************************

Public Class frmInput
    'Declare general variable used for calculating.
    Protected Friend strGender As String
    Protected Friend intAge As Integer
    Protected Friend dblHeight As Double
    Protected Friend dblWeight As Double
    Protected Friend intActivity As Integer
    Protected Friend blnSmoke As Boolean
    Protected Friend intSmokes As Integer
    Protected Friend intMonths As Integer
    Protected Friend blnEng As Boolean
    Protected Friend blnMet As Boolean

    'Clears the form
    Function ClearForm()
        radMale.Checked = True
        radNoSmoke.Checked = True
        cmbActivity.SelectedIndex = -1
        txtAge.Clear()
        txtHeight.Clear()
        txtWeight.Clear()
        txtSmokes.Clear()
        txtSmokes.Clear()
    End Function

    'Declares variables based on user input
    Function StoreInput()
        'Initialize variables with the data given by the user.
        Try
            'Checks gender
            If radMale.Checked = True Then
                strGender = "Male"
            Else
                strGender = "Female"
            End If

            'Age, height, weight
            intAge = CInt(txtAge.Text)
            dblHeight = CDec(txtHeight.Text)
            dblWeight = CDec(txtWeight.Text)

            'Activity index level
            intActivity = cmbActivity.SelectedIndex()

            'Checks if user smokes
            If radYesSmoke.Checked = True Then
                blnSmoke = True
                intSmokes = CInt(txtSmokes.Text)
                intMonths = CInt(txtMonths.Text)
            Else
                blnSmoke = False
            End If

            'Checks if any numbers are negative then throw exception
            If intAge < 0 Or dblHeight < 0 Or dblWeight < 0 Or
                intSmokes < 0 Or intMonths < 0 Then
                Throw New Exception()
            End If

            If (cmbActivity.SelectedIndex = -1) = True Then
                Throw New ArgumentException()
            End If

            'Checks if a measurement is not selected then throw
            If (mnuMeasureEng.Checked = False) And (mnuMeasureMet.Checked = False) Then
                Throw New EvaluateException("Error.")
            End If

            'Opens the output form
            frmOutput.ShowDialog()

        Catch ex As EvaluateException
            'Error message for not selecting measurement type
            MessageBox.Show("Please select a Measurement Type.")

        Catch ex As ArgumentException
            'Error Message for not selecting activity level
            MessageBox.Show("Please Select an Activity Level")

        Catch
            'Error message for invalid input types
            MessageBox.Show("Please Enter Valid Numeric Values.")
        End Try
    End Function

    Private Sub frmInput_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Clears the form on load
        blnEng = True
        radMale.Checked = True
        radNoSmoke.Checked = True
        cmbActivity.SelectedIndex = -1
        txtAge.Clear()
        txtHeight.Clear()
        txtWeight.Clear()
        txtMonths.Clear()
        txtSmokes.Clear()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'calls the clearform function
        ClearForm()
    End Sub

    Protected Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        'Calls the calculateform function
        StoreInput()
    End Sub

    Private Sub radYesSmoke_CheckedChanged(sender As Object, e As EventArgs) Handles radYesSmoke.CheckedChanged
        'Enables the smoking option textboxes
        txtMonths.Enabled = True
        txtSmokes.Enabled = True
    End Sub

    Private Sub radNoSmoke_CheckedChanged(sender As Object, e As EventArgs) Handles radNoSmoke.CheckedChanged
        'Disables the smoking option textboxes
        txtMonths.Enabled = False
        txtSmokes.Enabled = False
    End Sub

    Protected Sub mnuMeasureEng_Click(sender As Object, e As EventArgs) Handles mnuMeasureEng.Click
        'Checks if English is selected, unchecks Metric
        'Can also convert current metric measures into english scale.
        If mnuMeasureEng.Checked = True Then
            blnEng = True
            blnMet = False
            mnuMeasureMet.Checked = False

            'Change the measurement labels
            lblHeight.Text = "Height (in):"
            lblWeight.Text = "Weight (lb):"

            Try
                'Converts text values to metric values
                dblHeight = (CDec(txtHeight.Text) / 2.54)
                dblWeight = (CDec(txtWeight.Text) / 0.453592)
                txtHeight.Text = dblHeight.ToString()
                txtWeight.Text = dblWeight.ToString()
            Catch

            End Try
        End If
    End Sub

    Protected Sub mnuMeasureMet_Click(sender As Object, e As EventArgs) Handles mnuMeasureMet.Click
        'Checks if Metric is selected, unchecks English
        'Can also convert current english measures into metric scale.
        If mnuMeasureMet.Checked = True Then
            blnMet = True
            blnEng = False
            mnuMeasureEng.Checked = False

            lblHeight.Text = "Height (cm):"
            lblWeight.Text = "Weight (kg):"

            Try
                'Converts text values to metric values
                dblHeight = (CDec(txtHeight.Text) * 2.54)
                dblWeight = (CDec(txtWeight.Text) * 0.453592)
                txtHeight.Text = dblHeight.ToString()
                txtWeight.Text = dblWeight.ToString()
            Catch

            End Try
        End If
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        'A helpful about message
        MessageBox.Show("Enter data about yourself. Be wary of the measurements.")
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        'Close form.
        Me.Close()
    End Sub
End Class
