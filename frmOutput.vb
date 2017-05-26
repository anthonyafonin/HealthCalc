'***************************************************************
' Programmer: Afonin, Anthony
' Chemeketa Community College
' Date: 3/12/2016
' Class: CIS133VB
' Assignment: Final Project, Health Calculator
' File Name: frmOutput.vb
' Description: The output form that calculates and displays various
' sorts of health information using data from the user input.
'
'***************************************************************

Public Class frmOutput
    'Declare global variables
    Dim dblCurrentBMI As Double
    Dim strWeightType As String
    Dim dblHbr As Double
    Dim dblHrate As Double
    Dim dblBmr As Double
    Dim dblLbm As Double
    Dim intSmokes As Integer
    Dim dblCost As Double

    Function ChangeLabels()
        'Assign general information to labels
        lblGender.Text = frmInput.strGender
        lblAge.Text = frmInput.intAge
        lblHeight.Text = frmInput.dblHeight
        lblWeight.Text = frmInput.dblWeight

        'Changes measurement labels accordingly
        If frmInput.blnEng = True Then
            lblHeight1.Text = "Height (in):"
            lblWeight1.Text = "Weight (lb):"
            lblLbm1.Text = "LBM (lb):"
        ElseIf frmInput.blnMet = True Then
            lblHeight1.Text = "Height (cm):"
            lblWeight1.Text = "Weight (kg):"
            lblLbm1.Text = "LBM (kg):"
        End If
    End Function

    Private Function CalculateSmoking()
        'Checks Smoking conditions and calculates output
        If frmInput.blnSmoke = True Then

            'Calculates total smokes the user has had 
            lblSmokes.Text = frmInput.intSmokes
            intSmokes = (CInt(frmInput.intSmokes * (frmInput.intMonths * 30)))
            lblTotalSmokes.Text = intSmokes.ToString()

            'Calculates hours shaved off life expectancy from total smokes.
            'Each smoke is said to cut 10 minutes off a person's life.
            lblHours.Text = CDec((intSmokes * 10) / 60)

            'Calculates financial costs of smoking total smokes.
            'There are aon average 25 cigarettes per pack and on average cost $5 a pack.
            dblCost = CDec((intSmokes / 25) * 5)
            lblMoney.Text = dblCost.ToString("c")

        ElseIf frmInput.blnSmoke = False Then
            lblSmokes.Text = "0"
            lblTotalSmokes.Text = "0"
            lblHours.Text = "0"
            lblMoney.Text = "0"
        End If
    End Function

    Private Function CalculateMeasures()
        'Calculates bmr, bmi, and lbm that use either enlglish or metric scale.

        '-------------------------------------------
        '           English Measurements
        '-------------------------------------------
        If frmInput.blnEng = True Then

            'Basal Metabolic Rate (BMR) and Lean Body Mass 
            '---------------------------------------------
            'Checks if gender is Male or Female
            If (frmInput.strGender = "Male") Then

                'BMR formula for Males in English Scale
                dblBmr = (66 + (6.23 * frmInput.dblWeight) +
                    (12.7 * frmInput.dblHeight) - (6.8 * frmInput.intAge))

                'LBM formula for Males in English Scale
                dblLbm = ((0.3281 * (frmInput.dblWeight * 0.453592)) +
                    (0.33929 * (frmInput.dblHeight * 2.54)) - 29.5336)
                dblLbm = (dblLbm / 0.453592)
            ElseIf (frmInput.strGender = "Female") Then

                'BMR formula for Females in English Scale
                dblBmr = (655 + (4.35 * frmInput.dblWeight) +
                    (4.7 * frmInput.dblHeight) - (4.7 * frmInput.intAge))

                'LBM formula for Females in English Scale
                dblLbm = ((0.29569 * (frmInput.dblWeight * 0.453592)) +
                    (0.41813 * (frmInput.dblHeight * 2.54)) - 43.2933)
                dblLbm = (dblLbm / 0.453592)
            End If

            'Current English BMI
            dblCurrentBMI = ((frmInput.dblWeight * 703) / (frmInput.dblHeight ^ 2))

            '--------------------------------------
            '         Metric Measurements
            '--------------------------------------
        ElseIf frmInput.blnMet = True Then

            'Basal Metabolic Rate (BMR) and Lean Body Mass 
            '---------------------------------------------
            'Checks if gender is Male or Female
            If (frmInput.strGender = "Male") Then

                'BMR formula for Males in Metric Scale
                dblBmr = (66 + (13.7 * frmInput.dblWeight) +
                    (5 * frmInput.dblHeight) - (6.8 * frmInput.intAge))

                'LBM formula for Males in Metric Scale
                dblLbm = ((0.3281 * frmInput.dblWeight) +
                    (0.33929 * frmInput.dblHeight) - 29.5336)

            ElseIf (frmInput.strGender = "Female") Then

                'BMR formula for Females in Metric Scale
                dblBmr = (655 + (9.6 * frmInput.dblWeight) +
                    (1.8 * frmInput.dblHeight) - (4.7 * frmInput.intAge))

                'LBM formula for Females in Metric Scale
                dblLbm = ((0.29569 * frmInput.dblWeight) +
                    (0.41813 * frmInput.dblHeight) - 43.2933)
            End If

            'Current Metric BMI
            dblCurrentBMI = ((frmInput.dblWeight) / ((frmInput.dblHeight / 100) ^ 2))
        End If
    End Function

    Function WeightType()
        'Body Weight Type 
        '-------------------------------------------------------------
        'Checks what range the BMI is in and selects body weight type
        If (dblCurrentBMI <= 18.5) Then
            strWeightType = "Underweight"
        ElseIf (dblCurrentBMI > 18.5) And (dblCurrentBMI <= 24.99) Then
            strWeightType = "Normal Weight"
        ElseIf (dblCurrentBMI >= 25) And (dblCurrentBMI <= 29.99) Then
            strWeightType = "Overweight"
        ElseIf (dblCurrentBMI >= 30) Then
            strWeightType = "Obese"
        End If
    End Function

    Function HarrisB()
        'Harris Benedict Formula 
        '--------------------------------------------
        'Checks the selected combo box activity level 
        'and selects equation accordingly.
        If (frmInput.cmbActivity.SelectedIndex = 0) Then
            dblHbr = (dblBmr * 1.2)
        ElseIf (frmInput.cmbActivity.SelectedIndex = 1) Then
            dblHbr = (dblBmr * 1.375)
        ElseIf (frmInput.cmbActivity.SelectedIndex = 2) Then
            dblHbr = (dblBmr * 1.55)
        ElseIf (frmInput.cmbActivity.SelectedIndex = 3) Then
            dblHbr = (dblBmr * 1.725)
        ElseIf (frmInput.cmbActivity.SelectedIndex = 4) Then
            dblHbr = (dblBmr * 1.9)
        End If
    End Function

    Function HeartRate()
        'Target Heart Rate 
        dblHrate = (220 - frmInput.intAge)
    End Function

    Function DisplayResults()
        Try
            'Format results
            dblCurrentBMI = FormatNumber(dblCurrentBMI, 2)
            dblBmr = FormatNumber(dblBmr, 2)
            dblHbr = FormatNumber(dblHbr, 2)
            dblHrate = FormatNumber(dblHrate, 2)
            dblLbm = FormatNumber(dblLbm, 2)

            'Displays Results
            lblCurrentBMI.Text = dblCurrentBMI.ToString()
            lblWeightType.Text = strWeightType
            lblBMR.Text = dblBmr.ToString()
            lblCalories.Text = dblHbr.ToString()
            lblRate.Text = dblHrate.ToString()
            lblLbm.Text = dblLbm.ToString()

        Catch ex As Exception
            'Error message if any error occurs
            MessageBox.Show("There was an error calculating the results.")
        End Try
    End Function

    Protected Sub frmOutput_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Calls ChangeLabels Function
        ChangeLabels()

        'Calls CalculateSmoking function
        CalculateSmoking()

        'Calls Measurement calculation function
        CalculateMeasures()

        'Calls the body weight type function
        WeightType()

        'Calls the HarrisB function
        HarrisB()

        'Calls the HeartRate function
        HeartRate()

        'Calls the displayresults function
        DisplayResults()
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        'Close the form
        Me.Close()
    End Sub

    '------------------------------------------------
    ' Shows more about information of each result
    '------------------------------------------------
    Private Sub mnuBMICurrent_Click_1(sender As Object, e As EventArgs) Handles mnuBMICurrent.Click
        MessageBox.Show("Displays the Body Mass Index of the user.")
    End Sub

    Private Sub mnuBMIBodytype_Click(sender As Object, e As EventArgs) Handles mnuBMIBodytype.Click
        MessageBox.Show("Displays the weight type of the user." & vbNewLine &
                        "The weight type is determined by the BMI.")
    End Sub

    Private Sub mnuBMRRate_Click(sender As Object, e As EventArgs) Handles mnuBMRRate.Click
        MessageBox.Show("Displays the Basal Matabolic Rate of the user.")
    End Sub

    Private Sub mnuBMRCal_Click(sender As Object, e As EventArgs) Handles mnuBMRCal.Click
        MessageBox.Show("Displays the number of calories the user should intake per day.")
    End Sub

    Private Sub mnuHeartTarget_Click(sender As Object, e As EventArgs) Handles mnuHeartTarget.Click
        MessageBox.Show("Displays the Max Target Heart Rate of the user.")
    End Sub

    Private Sub mnuLBM_Click(sender As Object, e As EventArgs) Handles mnuLBM.Click
        MessageBox.Show("Displays the user's Lean Body Mass." & vbNewLine &
                        "The LBM is the total amount of weight that is not fat.")
    End Sub

    Private Sub mnuSmokingLife_Click(sender As Object, e As EventArgs) Handles mnuSmokingLife.Click
        MessageBox.Show("Displays the user how many hours they lost by smoking." & vbNewLine &
                        "A cigerette, on average, takes away 10 minutes from a life.")
    End Sub

    Private Sub mnuSmokingCost_Click(sender As Object, e As EventArgs) Handles mnuSmokingCost.Click
        MessageBox.Show("Displays the user how much money they spent by smoking." & vbNewLine &
                        "There are usually 25 cigerettes in a pack and a pack costs $5.")
    End Sub

    Private Sub mnuFilePrint_Click(sender As Object, e As EventArgs) Handles mnuFilePrint.Click
        'Prints the form
        pdPrint.Print()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        'About Message
        MessageBox.Show("This form displays the results of various calculations.")
    End Sub
End Class