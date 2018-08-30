Public Class Form1


    'Heading string is first line of csv document, contains all fields needed for gmail compatibility
    Dim headingString As String = "Time&Date,Matriculation Number,Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Group Membership,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value"
    Dim ShowPersonalEmail As Boolean = False


    'Initiasise program for use and enter group name'
    Private Sub button2_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        TextBox2.Text = headingString & vbCrLf

        My.Computer.FileSystem.CreateDirectory("C:\NapierArchery\")
        My.Computer.FileSystem.WriteAllText("C:\NapierArchery\Mailinglist.csv", TextBox2.Text, True)

        If CheckBox1.Checked Then
            ShowPersonalEmail = True
        End If

        hideInitFields()
        readyDataRecording()

    End Sub



    ' Reveals personal email field if checkbox checked in Initialisation
    Private Sub PersonalEmail(ByVal ShowPersonalEmail As Boolean)
        If ShowPersonalEmail = True Then
            Label3.Show()   'Show label Personal Email
            InputBox3.Show() 'Personal email field
        End If
    End Sub



    'User entered data saved as gmail compatible csv file
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        OutputBox1.Text = Date.Now & "," & InputBox2.Text & "," & InputBox1.Text & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & "," & InputBox2.Text & "," & TextBox1.Text & "," & "* Home" & "," & InputBox3.Text & "," & "* Uni" & "," & InputBox2.Text & "@live.napier.ac.uk" & vbCrLf


        My.Computer.FileSystem.WriteAllText("C:\NapierArchery\Mailinglist.csv", OutputBox1.Text, True)


        'Display info submitted for user to review
        If ShowPersonalEmail = True Then

            MessageBox.Show("Information Submitted" & vbCrLf & vbCrLf & "Name : " & vbTab & InputBox1.Text & vbCrLf &
                        "Matric : " & vbTab & InputBox2.Text & vbCrLf &
                        "Email : " & vbTab & InputBox3.Text & vbCrLf & vbCrLf &
                        "If you have made a mistake, just re-enter your details")

        Else
            MessageBox.Show("Information Submitted" & vbCrLf & vbCrLf & "Name : " & vbTab & InputBox1.Text & vbCrLf &
                        "Matric : " & vbTab & InputBox2.Text & vbCrLf & vbCrLf &
                        "If you have made a mistake, just re-enter your details")

        End If


        'ready fields for next user
        readyDataRecording()


    End Sub



    ' Prepare form for user input
    Private Sub readyDataRecording()
        InputBox1.Text = ""
        InputBox2.Text = ""
        InputBox3.Text = ""

        PictureBox1.Show() 'Shows logo
        Label1.Show()
        Label2.Show()
        Label7.Show()
        InputBox1.Show()
        InputBox2.Show()
        PersonalEmail(ShowPersonalEmail)

        Button1.Show()  'Submit Details Button
        InputBox1.Focus() 'Select Name Field for next user

    End Sub


    Private Sub hideInitFields()
        Label6.Hide()   'Hide Instructions
        Button2.Hide()  'Init Button
        Label4.Hide()   'Info text for Initialisation
        TextBox1.Hide() 'MailingList Group Name field
        CheckBox1.Hide() 'Hide Personal Email checkbox

    End Sub

End Class
