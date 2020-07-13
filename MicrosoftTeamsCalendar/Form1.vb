Imports Microsoft.Identity.Client
Imports Microsoft.Graph
Imports Microsoft.Graph.Auth
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim startTime As DateTime = Convert.ToDateTime(TextBox1.Text)
        Dim endTime As DateTime = Convert.ToDateTime(TextBox2.Text)

        Dim objOutlook As Outlook.Application
        objOutlook = CreateObject("Outlook.Application")

        Dim newAppointment As Outlook.AppointmentItem =
        objOutlook.CreateItem(Outlook.OlItemType.olAppointmentItem)
        Try
            With newAppointment
                .Start = startTime
                .End = endTime
                .Location = TextBox4.Text
                .Body = RichTextBox1.Text
                .Subject = TextBox3.Text
                .ReminderSet = True

                '.AllDayEvent = False
                '.Recipients.Add("Roger Harui")
                'Dim sentTo As Outlook.Recipients = .Recipients
                'Dim sentInvite As Outlook.Recipient
                'sentInvite = sentTo.Add("Holly Holt")
                'sentInvite.Type = Outlook.OlMeetingRecipientType.olRequired
                'sentInvite = sentTo.Add("David Junca")
                'sentInvite.Type = Outlook.OlMeetingRecipientType.olOptional
                'sentTo.ResolveAll()
                .Save()
            End With
            MessageBox.Show("You have appointed correctly")
        Catch ex As Exception
            MessageBox.Show("The following error occurred: " & ex.Message)
        Finally
            If Not IsNothing(newAppointment) Then Marshal.ReleaseComObject(newAppointment)
            If Not IsNothing(objOutlook) Then Marshal.ReleaseComObject(objOutlook)
        End Try

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Date.Now.AddHours(2).ToString("MM/dd/yyyy H:mm")
        TextBox2.Text = Date.Now.AddHours(3).ToString("MM/dd/yyyy H:mm")
    End Sub
End Class

