Public Class Form1
    Dim ProgBarVal As Integer
    Dim SelectionCount As Integer   'For activating Filter Button
    Dim ScreenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Size.Width
    Dim ScreenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Size.Height

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        SplashScreen1.Close()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim totFonts As Integer
        Me.WindowState = FormWindowState.Maximized
        Panel1.Size = New Point(ScreenWidth - 20, ScreenHeight - 90)
        Button2.Location = New Point(ScreenWidth - 52, 7)

        Dim allFonts As New Drawing.Text.InstalledFontCollection
        Dim fontFamilies() As FontFamily = allFonts.Families()
        For Each myFont As FontFamily In fontFamilies
            totFonts = totFonts + 1
        Next
        Label1.Text = "Total number of Fonts installed in your system : " & totFonts

        ToolStripStatusLabel2.Text = Now

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Panel1.Controls.Clear()
        Dim y As Integer = 10
        Dim x As Integer = 10
        Dim allFonts As New Drawing.Text.InstalledFontCollection
        Dim fontFamilies() As FontFamily = allFonts.Families()
        MsgBox("Please wait till the fonts are loaded.." + vbNewLine + vbNewLine + "Time taken depends upon the number of fonts installed in your" + vbNewLine + "default fonts folder and configuration of your computer.")
        Me.WindowState = FormWindowState.Minimized
        System.Diagnostics.Process.Start("CPSloading.exe")
        For Each myFont As FontFamily In fontFamilies
            Dim l As New Label()
            Dim lf As New Label()
            Dim chk As New CheckBox()
            Try
                With l
                    .Location = New Point(x, y)
                    .Font = New Font(myFont, 20)
                    .Text = TextBox1.Text
                    .BackColor = Color.White
                    .TextAlign = ContentAlignment.MiddleCenter
                    .Height = 80
                    .Width = 200
                End With
                With lf
                    .Location = New Point(x, y + 70)
                    .Text = myFont.Name
                    .BackColor = Color.White
                    .TextAlign = ContentAlignment.MiddleCenter
                    .Width = 200
                    .Height = 30
                    .Font = New Font("Microsoft Sans Serif", 7)
                End With
                With chk
                    .Location = New Point(x, y)
                    .AutoSize = False
                    .Width = 15
                    .Height = 15
                    .Text = myFont.Name
                End With
                If y < ScreenHeight - 300 Then
                    y = y + 110
                Else
                    x = x + 210
                    y = 10
                End If
                Panel1.Controls.Add(l)
                Panel1.Controls.Add(lf)
                Panel1.Controls.Add(chk)
                AddHandler chk.CheckedChanged, AddressOf MyClickHandler
                lf.BringToFront()
                chk.BringToFront()
            Catch ex As Exception

            End Try
        Next
        Dim pProcess() As Process = System.Diagnostics.Process.GetProcessesByName("CPSloading")
        For Each p As Process In pProcess
            p.Kill()
        Next
        MessageBox.Show("Fonts Loaded succesfully", "CPS Font Viewer", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub MyClickHandler(ByVal Sender As Object, ByVal e As System.EventArgs)
        If Sender.Checked Then
            ' MsgBox("You Clicked " & Sender.Text)
            Form2.ListBox1.Items.Add(Sender.Text)
            SelectionCount = SelectionCount + 1
        Else
            ' MsgBox("You Unclicked " & Sender.Text)
            Form2.ListBox1.Items.Remove(Sender.Text)
            SelectionCount = SelectionCount - 1
        End If

        If SelectionCount > 1 Then
            Button3.Enabled = True
        Else
            Button3.Enabled = False
        End If
    End Sub



    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form2.Show()
        Form2.BringToFront()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel2.Text = Now
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If My.Computer.Keyboard.CapsLock Then
            ToolStripStatusLabel3.Text = "CAPS"
        Else
            ToolStripStatusLabel3.Text = ""
        End If
        If My.Computer.Keyboard.NumLock Then
            ToolStripStatusLabel4.Text = "NUM"
        Else
            ToolStripStatusLabel4.Text = ""
        End If
        If TSSL5.Text = " " Then
            TSSL5.Text = "Ready"
        End If
    End Sub

#Region "Mouse Events"
    Private Sub Button3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.MouseEnter
        TSSL5.Text = "Filter"
    End Sub

    Private Sub Button3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.MouseLeave
        TSSL5.Text = " "
    End Sub
#End Region
    

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form3.Show()
    End Sub
End Class
