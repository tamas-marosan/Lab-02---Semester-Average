'        Name: Tamas Marosan
'     Program: Lab 02 - Semester Average
'        Date: 2018-02-20
' Description: Calculate the grade of a semester and assign a letter grade
Option Strict On
Public Class SemesterAverageForm

    Const minGrade As Double = 0.0   ' Declare a constant for the minimum grade
    Const maxGrade As Double = 100.0 ' Declare a constant for the maximum grade

    ' When any textbox loses focus
    Private Sub txtPercent_LostFocus() Handles txtPercentA.LostFocus, txtPercentB.LostFocus, txtPercentC.LostFocus, txtPercentD.LostFocus, txtPercentE.LostFocus, txtPercentF.LostFocus
        lblOutput.ResetText() ' Reset the text for the output label

        ' For each grade's textbox, and its corresponding label, perform input validation and assign a letter grade
        validEntry(txtPercentA, lblGradeA)
        validEntry(txtPercentB, lblGradeB)
        validEntry(txtPercentC, lblGradeC)
        validEntry(txtPercentD, lblGradeD)
        validEntry(txtPercentE, lblGradeE)
        validEntry(txtPercentF, lblGradeF)
    End Sub

    ' When the calculate button is used
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Dim gradeArray(5) As Double ' Declare the array for all given grades with a length of 6
        Dim i As Integer = 0 ' Declare an index for the array
        Dim gradeAverage As Double = 0.0 ' Declare a variable for the average grade
        Dim allValid As Boolean = True ' Declare a boolean as true to continue the method if it's not set to false at some point
        lblOutput.Text = $"ERROR(S):{vbCrLf}" ' Set the text of the output label to let the user know it's displaying errors (if there are any)

        ' Revalidate each textbox before putting their values into the array
        For Each txtBox As Object In Me.Controls ' Search through all objects in the form
            If TypeOf (txtBox) Is TextBox Then ' If the object is a textbox
                If CType(txtBox, TextBox).Text = String.Empty Or                   ' If any of the textboxes are empty 
                Not Double.TryParse(CType(txtBox, TextBox).Text, gradeArray(i)) Or ' ...or not numeric
                gradeArray(i) < minGrade Or gradeArray(i) > maxGrade Then          ' ...or not within range then...
                    lblOutput.Text += $"Please ensure input in Course {i + 1} is a number between 0 and 100!{vbCrLf}" ' give an error for that textbox
                    allValid = False ' ...and set a flag that not all textboxes are valid
                Else ' If the textbox has valid input then...
                    Double.TryParse(CType(txtBox, TextBox).Text, gradeArray(i)) ' insert the value of the textbox into the array
                End If
                i += 1 ' Use the next index for the next textbox
            End If
        Next ' Move onto the next object

        If allValid = True Then ' If all of the textboxes had valid input, then...
            For Each item As Double In gradeArray
                gradeAverage += item / 6 ' obtain an average of all the grades
            Next

            lblPercentSem.Text = $"{Math.Round(gradeAverage, 2)}" ' change the semester percentage label to the average of the grades
            lblGradeSem.Text = letterGrade(gradeAverage) ' Assign the corresponding grade letter to the semester letter grade label using the letterGrade function

            ' Disable all textboxes on the form
            For Each formObject As Object In Me.Controls ' Search all objects on the form
                If TypeOf (formObject) Is TextBox Then ' If the object is a textbox
                    CType(formObject, TextBox).Enabled = False ' Disable the textbox
                End If
            Next ' Move onto the next object
            btnCalculate.Enabled = False ' Disable the Calculate button
            lblOutput.ResetText() ' Reset the text for the output label
        End If
    End Sub

    ' When the reset button is used
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        formReset() ' Call the formReset method
    End Sub

    ' When the exit button is used
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ' Close this application if the Exit button is pressed
        Application.Exit()
    End Sub

    ' Method to validate textbox input
    Private Sub validEntry(ByVal currentTextBox As Object, ByVal currentLabel As Object)
        Dim percentDouble As Double = 0.0 ' Declare a double for the textbox to be parsed into
        CType(currentLabel, Label).Text = String.Empty ' reset the corresponding label to empty

        If Not CType(currentTextBox, TextBox).Text = String.Empty Then ' If there is text in the textbox
            If Not Double.TryParse(CType(currentTextBox, TextBox).Text, percentDouble) Then ' If the textbox isn't numeric
                lblOutput.Text = $"Please ensure your input is numeric!{vbCrLf}" ' Give an approriate error

                CType(currentTextBox, TextBox).Focus() ' Set the cursor back onto the textbox
                CType(currentTextBox, TextBox).SelectAll() ' Highlight the text in the textbox to make it more streamlined for the user
            ElseIf percentDouble < minGrade Or percentDouble > maxGrade Then ' If the input is not in range
                lblOutput.Text = $"Please ensure your input is a number between 0 and 100!{vbCrLf}" ' Give an approriate error

                CType(currentTextBox, TextBox).Focus() ' Set the cursor back onto the textbox
                CType(currentTextBox, TextBox).SelectAll() ' Highlight the text in the textbox to make it more streamlined for the user
            Else ' If the entry in the textbox is valid
                CType(currentLabel, Label).Text = letterGrade(percentDouble) ' Use the letterGrade function to assign a letter grade to the corresponding label
            End If
        End If
    End Sub

    ' Function to assign a letter grade given a percentage grade
    Private Function letterGrade(ByVal percentValue As Double) As String
        Dim gradeLetter As String = "A+" ' Define a return variable as a string and initialize it to A+

        ' Set the return variable to an appropriate grade depending on the percentage
        If percentValue < 90 Then gradeLetter = "A"
        If percentValue < 85 Then gradeLetter = "A-"
        If percentValue < 80 Then gradeLetter = "B+"
        If percentValue < 77 Then gradeLetter = "B"
        If percentValue < 73 Then gradeLetter = "B-"
        If percentValue < 70 Then gradeLetter = "C+"
        If percentValue < 67 Then gradeLetter = "C"
        If percentValue < 63 Then gradeLetter = "C-"
        If percentValue < 60 Then gradeLetter = "D+"
        If percentValue < 57 Then gradeLetter = "D"
        If percentValue < 53 Then gradeLetter = "D-"
        If percentValue < 50 Then gradeLetter = "F"

        Return gradeLetter ' return the grade assigned when this function is called
    End Function

    ' Method to reset the form to its original state
    Private Sub formReset()
        For Each formObject As Object In Me.Controls ' Search through all objects in the form
            If TypeOf (formObject) Is TextBox Then ' If the object is a textbox
                CType(formObject, TextBox).ResetText() ' Reset the textbox's text
                CType(formObject, TextBox).Enabled = True ' Enable the textbox
            ElseIf TypeOf (formObject) Is Label AndAlso CType(formObject, Label).BorderStyle = BorderStyle.Fixed3D Then ' If the object is a label with a Fixed3D border
                CType(formObject, Label).ResetText() ' Reset the label's text
            End If
        Next ' Move onto the next object
        btnCalculate.Enabled = True ' Enable the calculate button
        txtPercentA.Focus() ' Set the focus to the first textbox
    End Sub
End Class