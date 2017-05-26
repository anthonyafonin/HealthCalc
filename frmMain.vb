'***************************************************************
' Programmer: Afonin, Anthony
' Chemeketa Community College
' Date: 3/12/2016
' Class: CIS133VB
' Assignment: Final Project, Health Calculator
' File Name: frmMain.vb
' Description: A starting point of the Final Project, 
' takes the user to the input form.
'
'***************************************************************

Public Class frmMain

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        'Opens the input form
        frmInput.ShowDialog()
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        'Closes program
        Me.Close()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        'About message
        MessageBox.Show("Begin the program by pressing the 'Start' button!")
    End Sub
End Class