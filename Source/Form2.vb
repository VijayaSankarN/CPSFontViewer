Public Class Form2
    Dim FC As Color = Color.Black    'Foreground/ text colour
    Dim BC As Color = Color.White    'Background Colour
    Dim FontName(10) As String
    Dim ListBoxCountNo As Integer

    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        MsgBox("Close main window to close")
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ListBoxCount As Integer = ListBox1.SelectedItems.Count
        If ListBoxCount > 0 And ListBox1.Enabled Then
            For ListBoxCount = 1 To ListBoxCount
                ListBox1.Items.Remove(ListBox1.SelectedItem)
            Next
        ElseIf ListBox1.Enabled = False Then
            MsgBox("Refresh the List first")
        Else
            MsgBox("Nothing Selected")
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim ListBoxCount As Integer = ListBox1.SelectedItems.Count
        If ListBoxCount > 9 Then
            ListBox1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ListBoxCount As Integer = ListBox1.SelectedItems.Count
        If ListBoxCount > 0 Then
            For ListBoxCount = 1 To ListBoxCount
                ListBox1.SetSelected(ListBox1.SelectedIndex, False)
            Next
            ListBox1.Enabled = True
        ElseIf ListBox1.Enabled = False Then
            ListBox1.Enabled = True
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ListBoxCountNo = ListBox1.SelectedItems.Count
        Dim y As Integer = 10
        Dim FontSize As Integer = ComboBox1.Text
        'Dim FontName(10) As String
        Dim ListBoxCount As Integer = ListBox1.SelectedItems.Count

        For ListBoxCount = 1 To ListBoxCount
            FontName(ListBoxCount) = ListBox1.SelectedItem
            ListBox1.SetSelected(ListBox1.SelectedIndex, False)
        Next

        If ListBoxCount > 1 Then
            Panel1.Controls.Clear()
            For ListBoxCount = 1 To ListBoxCount - 1
                Dim l As New Label()
                Dim lf As New Label()
                Dim Rbut As New Button()
                Dim Lbut As New Button()
                With l              ' label to display font style
                    .Name = FontName(ListBoxCount) & "FonSty"
                    .Location = New Point(0, y)
                    .ForeColor = FC
                    .Text = TextBox1.Text
                    .BackColor = BC
                    .TextAlign = ContentAlignment.MiddleCenter
                    .Height = 80
                    .Width = 584
                    If CheckBox1.Checked And CheckBox2.Checked And CheckBox3.Checked Then
                        .Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Bold Or FontStyle.Italic Or FontStyle.Underline)

                    ElseIf CheckBox1.Checked And CheckBox2.Checked Then
                        .Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Bold Or FontStyle.Italic)

                    ElseIf CheckBox1.Checked And CheckBox3.Checked Then
                        .Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Bold Or FontStyle.Underline)

                    ElseIf CheckBox2.Checked And CheckBox3.Checked Then
                        .Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Italic Or FontStyle.Underline)

                    ElseIf CheckBox1.Checked Then
                        .Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Bold)

                    ElseIf CheckBox2.Checked Then
                        .Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Italic)

                    ElseIf CheckBox3.Checked Then
                        .Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Underline)

                    Else
                        .Font = New Font(FontName(ListBoxCount), FontSize)
                    End If

                End With
                With lf     'label to display font name
                    .Name = FontName(ListBoxCount) & "FonNam"
                    .Location = New Point(0, y + 65)
                    .Text = FontName(ListBoxCount)
                    .BackColor = Color.White
                    .TextAlign = ContentAlignment.MiddleCenter
                    .Width = 584
                    .Height = 15
                    .Font = New Font("Microsoft Sans Serif", 7)
                End With
                With Lbut
                    .Location = New Point(590, y + 45)
                    .Text = "Lock"
                    .Name = FontName(ListBoxCount)
                End With
                With Rbut
                    .Location = New Point(590, y + 15)
                    .Text = "Remove"
                    .Name = FontName(ListBoxCount)
                End With
                Me.Panel1.Controls.Add(l)
                Me.Panel1.Controls.Add(lf)
                Me.Panel1.Controls.Add(Lbut)
                Me.Panel1.Controls.Add(Rbut)
                AddHandler Rbut.Click, AddressOf RemButton_Click
                AddHandler Lbut.Click, AddressOf LocButton_Click
                lf.BringToFront()
                y = y + 90
            Next
        ElseIf ListBoxCount > 0 Then
            MsgBox("Select atleast 2 fonts from the list to compare")
        Else
            MsgBox("Nothing Selected in the list to compare")
        End If
    End Sub
    Private Sub RemButton_Click(ByVal Sender As Object, ByVal e As System.EventArgs)
        Dim ListBoxCount As Integer = ListBox1.SelectedItems.Count

        Dim FonNam As Label = Panel1.Controls.Item(Sender.Name & "FonNam")
        FonNam.Enabled = False
        Dim FonSty As Label = Panel1.Controls.Item(Sender.Name & "FonSty")
        FonSty.Enabled = False
        Dim LBut As Button = Panel1.Controls.Item(Sender.Name)
        LBut.Enabled = False

        ListBox1.Items.Remove(Sender.Name)
        Sender.Enabled = False
    End Sub
    Private Sub LocButton_Click(ByVal Sender As Object, ByVal e As System.EventArgs)
        If ListBox1.Items.Contains(Sender.Name) And Sender.Text = "Lock" Then
            ListBox1.SelectedItem = Sender.Name
            Sender.Text = "Locked"
        ElseIf ListBox1.Items.Contains(Sender.Name) And Sender.Text = "Locked" Then
            MsgBox("Font has been already locked for your next comparison")
        Else
            MsgBox("The font has been already removed from your choice")
            Sender.Enabled = False
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label3.Text = "Fonts in your wishlist : " & ListBox1.Items.Count
        Label4.Text = "Fonts selected for comparison : " & ListBox1.SelectedItems.Count & "/10"
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        FontStyleInFilter()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ColorDialog1.ShowDialog = DialogResult.OK Then
            FC = ColorDialog1.Color
        End If
        Button4.BackColor = FC
        FontStyleInFilter()
    End Sub

    Private  Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If ColorDialog2.ShowDialog = DialogResult.OK Then
            BC = ColorDialog2.Color
        End If
        Button5.BackColor = BC
        FontStyleInFilter()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        FontStyleInFilter()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        FontStyleInFilter()
    End Sub
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        FontStyleInFilter()
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        FontStyleInFilter()
    End Sub

    Private Sub FontStyleInFilter()
        Dim FontSize = ComboBox1.SelectedItem
        For ListBoxCount = 1 To ListBoxCountNo
            Dim FonSty As Label = Panel1.Controls.Item(FontName(ListBoxCount) & "FonSty")
            If CheckBox1.Checked And CheckBox2.Checked And CheckBox3.Checked Then
                FonSty.Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Bold Or FontStyle.Italic Or FontStyle.Underline)

            ElseIf CheckBox1.Checked And CheckBox2.Checked Then
                FonSty.Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Bold Or FontStyle.Italic)

            ElseIf CheckBox1.Checked And CheckBox3.Checked Then
                FonSty.Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Bold Or FontStyle.Underline)

            ElseIf CheckBox2.Checked And CheckBox3.Checked Then
                FonSty.Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Italic Or FontStyle.Underline)

            ElseIf CheckBox1.Checked Then
                FonSty.Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Bold)

            ElseIf CheckBox2.Checked Then
                FonSty.Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Italic)

            ElseIf CheckBox3.Checked Then
                FonSty.Font = New Font(FontName(ListBoxCount), FontSize, FontStyle.Underline)

            Else
                FonSty.Font = New Font(FontName(ListBoxCount), FontSize)
            End If
            FonSty.ForeColor = FC
            FonSty.BackColor = BC
            FonSty.Text = TextBox1.Text
        Next
    End Sub

    
End Class
