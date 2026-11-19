' ============================================================
' frmWelcome.vb : start form. Log in tab, Create account tab, Exit.
' Nobody gets past this form without an account: log in, create an account, or sign in with Google or Facebook.
' The admin logs in here too, with the special admin email (AdminEmail in modLink).
' Errors show twice: red text on the form AND a message box that says what went wrong.
' ============================================================
Imports System.IO
Imports System.Text.RegularExpressions

Public Class frmWelcome
    Inherits Form

    ' ---------- CONTROLS (built in code, so you do not design anything) ----------
    Private WithEvents pnlHeader As Panel
    Private WithEvents picLogo As PictureBox
    Private WithEvents lblTitle As Label
    Private WithEvents lblTagline As Label
    Private WithEvents tcAuth As TabControl
    Private WithEvents tpLogin As TabPage
    Private WithEvents tpCreate As TabPage
    Private WithEvents lblLoginEmail As Label
    Private WithEvents txtLoginEmail As TextBox
    Private WithEvents lblLoginPass As Label
    Private WithEvents txtLoginPass As TextBox
    Private WithEvents chkShowLogin As CheckBox
    Private WithEvents chkRemember As CheckBox
    Private WithEvents lblLoginMsg As Label
    Private WithEvents btnLogin As Button
    Private WithEvents lnkCreate As LinkLabel
    Private WithEvents btnGoogle As Button
    Private WithEvents btnFacebook As Button
    Private WithEvents lblAdminHint As Label
    Private WithEvents lblRole As Label
    Private WithEvents rdoResident As RadioButton
    Private WithEvents rdoCollector As RadioButton
    Private WithEvents lblName As Label
    Private WithEvents txtName As TextBox
    Private WithEvents lblEmail As Label
    Private WithEvents txtEmail As TextBox
    Private WithEvents lblPhone As Label
    Private WithEvents txtPhone As TextBox
    Private WithEvents lblArea As Label
    Private WithEvents cboCreateArea As ComboBox
    Private WithEvents lblPass As Label
    Private WithEvents txtPass As TextBox
    Private WithEvents chkShowCreate As CheckBox
    Private WithEvents lblCreateMsg As Label
    Private WithEvents btnCreate As Button
    Private WithEvents lblCreateNote As Label
    Private WithEvents btnExit As Button
    Private WithEvents ssMain As StatusStrip
    Private WithEvents tslHint As ToolStripStatusLabel

    Public Sub New()
        MyBase.New()
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Me.SuspendLayout()
        Me.Text = "TidyAccra Connect"
        Me.ClientSize = New Size(900, 690)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.AutoScaleMode = AutoScaleMode.None
        pnlHeader = NewPanel(Me, 0, 0, 900, 90)
        picLogo = NewPicture(pnlHeader, 14, 20, 50, 50)
        lblTitle = NewLabel(pnlHeader, "TidyAccra Connect", 76, 12, 0, 0, size:=22, bold:=True, fore:="1B4332")
        lblTagline = NewLabel(pnlHeader, "Clean streets. Green Accra. One pickup at a time.", 78, 54, 0, 0, size:=11, italic:=True)
        tcAuth = NewTabs(Me, 240, 104, 420, 540, 200, 36)
        tpLogin = NewPage(tcAuth, "Log in")
        tpCreate = NewPage(tcAuth, "Create account")
        lblLoginEmail = NewLabel(tpLogin, "Email", 24, 20, 0, 0, bold:=True)
        txtLoginEmail = NewText(tpLogin, 24, 44, 372, 30)
        lblLoginPass = NewLabel(tpLogin, "Password", 24, 90, 0, 0, bold:=True)
        txtLoginPass = NewText(tpLogin, 24, 114, 300, 30, password:=True)
        chkShowLogin = NewCheck(tpLogin, "&Show", 334, 118)
        chkRemember = NewCheck(tpLogin, "Remember my email", 24, 160)
        lblLoginMsg = NewLabel(tpLogin, "", 24, 190, 372, 24, fore:="A8400F")
        btnLogin = NewButton(tpLogin, "&Log in", 24, 218, 372, 52)
        lnkCreate = NewLink(tpLogin, "New here? Create an account", 24, 288, 372, 24)
        btnGoogle = NewButton(tpLogin, "Continue with &Google", 24, 320, 372, 44, "secondary")
        btnFacebook = NewButton(tpLogin, "Continue with &Facebook", 24, 372, 372, 44, "secondary")
        lblAdminHint = NewLabel(tpLogin, "Admin? Log in with your admin email.", 24, 426, 372, 22, italic:=True, fore:="555555", align:=ContentAlignment.MiddleCenter)
        lblRole = NewLabel(tpCreate, "I am a", 24, 14, 0, 0, bold:=True)
        rdoResident = NewRadio(tpCreate, "&Resident", 100, 12, 0, 0, isChecked:=True)
        rdoCollector = NewRadio(tpCreate, "C&ollector", 230, 12, 0, 0)
        lblName = NewLabel(tpCreate, "Full name", 24, 52, 0, 0, bold:=True)
        txtName = NewText(tpCreate, 24, 74, 372, 30)
        lblEmail = NewLabel(tpCreate, "Email", 24, 114, 0, 0, bold:=True)
        txtEmail = NewText(tpCreate, 24, 136, 372, 30)
        lblPhone = NewLabel(tpCreate, "Phone", 24, 176, 0, 0, bold:=True)
        txtPhone = NewText(tpCreate, 24, 198, 372, 30, maxLen:=12)
        lblArea = NewLabel(tpCreate, "Area (residents only)", 24, 238, 0, 0, bold:=True)
        cboCreateArea = NewCombo(tpCreate, 24, 260, 372)
        lblPass = NewLabel(tpCreate, "Password (at least 6 characters)", 24, 300, 0, 0, bold:=True)
        txtPass = NewText(tpCreate, 24, 322, 300, 30, password:=True)
        chkShowCreate = NewCheck(tpCreate, "&Show", 334, 326)
        lblCreateMsg = NewLabel(tpCreate, "", 24, 364, 372, 24, fore:="A8400F")
        btnCreate = NewButton(tpCreate, "&Create account and log in", 24, 394, 372, 52)
        lblCreateNote = NewLabel(tpCreate, "Admin accounts cannot be created here.", 24, 454, 372, 22, italic:=True, fore:="555555", align:=ContentAlignment.MiddleCenter)
        btnExit = NewButton(Me, "E&xit", 752, 616, 120, 38, "danger")
        ssMain = New StatusStrip()
        tslHint = New ToolStripStatusLabel("Tip: press Enter to log in")
        ssMain.Items.Add(tslHint)
        Me.Controls.Add(ssMain)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    ' the email is remembered in this small text file next to the program
    Private ReadOnly RememberFile As String = Path.Combine(Application.StartupPath, "remember.txt")

    ' Only residents choose an area. For a collector the area row hides and the controls below it move up.
    Private rowsBelowArea() As Control
    Private areaGap As Integer
    Private rowsMoved As Boolean = False

    ' ---------- form opens ----------
    Private Sub frmWelcome_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplyTheme(Me)
        WelcomeForm = Me                             ' modLink needs to find this form again
        SeedDemo()                                   ' adds the demo users once
        cboCreateArea.Items.AddRange(AreaCoords.Keys.ToArray())
        cboCreateArea.SelectedIndex = 0
        Me.AcceptButton = btnLogin                   ' Enter key = Log in
        If File.Exists(RememberFile) Then
            txtLoginEmail.Text = File.ReadAllText(RememberFile).Trim()
            chkRemember.Checked = (txtLoginEmail.Text <> "")
        End If
        If txtLoginEmail.Text <> "" Then Me.ActiveControl = txtLoginPass Else Me.ActiveControl = txtLoginEmail
        rowsBelowArea = New Control() {lblPass, txtPass, chkShowCreate, lblCreateMsg, btnCreate, lblCreateNote}
        areaGap = lblPass.Top - lblArea.Top          ' the height of the Area row
        Dim lnkForgot As LinkLabel = NewLink(tpLogin, "Forgot your password?", 24, 288, 170, 24)
        lnkCreate.SetBounds(194, 288, 202, 24)
        lnkForgot.Font = New Font("Segoe UI", 9)
        lnkForgot.TextAlign = ContentAlignment.MiddleLeft
        lnkCreate.TextAlign = ContentAlignment.MiddleRight
        lnkCreate.Font = New Font("Segoe UI", 9)
        AddHandler lnkForgot.LinkClicked, AddressOf ForgotPassword
        FitToScreen(Me)
    End Sub

    ' Called by modLink.ReturnToWelcome when somebody logs out
    Public Sub ShowAgain()
        txtLoginPass.Clear()
        tcAuth.SelectedTab = tpLogin
        Me.Show()
        txtLoginPass.Focus()
    End Sub

    Private Sub pnlHeader_Paint(sender As Object, e As PaintEventArgs) Handles pnlHeader.Paint
        PaintHeader(sender, e)
    End Sub

    ' Enter key does the job of the tab that is open
    Private Sub tcAuth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tcAuth.SelectedIndexChanged
        If tcAuth.SelectedTab Is tpLogin Then
            Me.AcceptButton = btnLogin
            tslHint.Text = "Tip: press Enter to log in"
        Else
            Me.AcceptButton = btnCreate
            tslHint.Text = "Tip: Tab moves to the next box"
        End If
    End Sub

    ' ---------- show or hide the password ----------
    Private Sub chkShowLogin_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowLogin.CheckedChanged
        txtLoginPass.UseSystemPasswordChar = Not chkShowLogin.Checked
    End Sub

    Private Sub chkShowCreate_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowCreate.CheckedChanged
        txtPass.UseSystemPasswordChar = Not chkShowCreate.Checked
    End Sub

    Private Sub lnkCreate_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkCreate.LinkClicked
        tcAuth.SelectedTab = tpCreate
        txtName.Focus()
    End Sub

    ' ---------- FORGOT PASSWORD ----------
    ' Demo rule: the person proves who they are with the email AND phone number used at sign-up.
    Private Sub ForgotPassword(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Using dlg As New Form()
            dlg.Text = "Reset your password"
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.StartPosition = FormStartPosition.CenterScreen
            dlg.MaximizeBox = False
            dlg.MinimizeBox = False
            dlg.ClientSize = New Size(380, 400)
            dlg.BackColor = Color.White
            dlg.Font = New Font("Segoe UI", 10)

            Dim lblHead As New Label With {.Text = "Forgot your password?", .Font = New Font("Segoe UI", 13, FontStyle.Bold), .ForeColor = DarkGreen, .AutoSize = True, .Location = New Point(20, 16)}
            Dim lblNote As New Label With {.Text = "Type your email and the phone number you signed up with.", .ForeColor = Color.DimGray, .AutoSize = False, .Size = New Size(340, 36), .Location = New Point(20, 48)}
            Dim lblMail As New Label With {.Text = "Email", .AutoSize = True, .Location = New Point(20, 90)}
            Dim txtMail As New TextBox With {.Location = New Point(20, 112), .Width = 340, .Text = txtLoginEmail.Text}
            Dim lblPh As New Label With {.Text = "Phone number", .AutoSize = True, .Location = New Point(20, 150)}
            Dim txtPh As New TextBox With {.Location = New Point(20, 172), .Width = 340, .MaxLength = 12}
            Dim lblN1 As New Label With {.Text = "New password (at least 6 characters)", .AutoSize = True, .Location = New Point(20, 210)}
            Dim txtN1 As New TextBox With {.Location = New Point(20, 232), .Width = 340, .UseSystemPasswordChar = True}
            Dim lblN2 As New Label With {.Text = "Type the new password again", .AutoSize = True, .Location = New Point(20, 270)}
            Dim txtN2 As New TextBox With {.Location = New Point(20, 292), .Width = 340, .UseSystemPasswordChar = True}
            Dim btnGo As New Button With {.Text = "Change password", .Location = New Point(20, 340), .Size = New Size(170, 40)}
            Dim btnNo As New Button With {.Text = "Cancel", .Tag = "secondary", .Location = New Point(200, 340), .Size = New Size(160, 40), .DialogResult = DialogResult.Cancel}
            StyleButton(btnGo)
            StyleButton(btnNo)
            dlg.Controls.AddRange(New Control() {lblHead, lblNote, lblMail, txtMail, lblPh, txtPh, lblN1, txtN1, lblN2, txtN2, btnGo, btnNo})
            dlg.AcceptButton = btnGo
            dlg.CancelButton = btnNo

            AddHandler btnGo.Click, Sub(s As Object, a As EventArgs)
                                        Dim mail As String = txtMail.Text.Trim().ToLower()
                                        If mail = AdminEmail Then
                                            ShowError("The admin password cannot be reset here. It is set in the code (AdminPassword in modLink).")
                                            Return
                                        End If
                                        Dim u As User = Users.FirstOrDefault(Function(x) x.Email.ToLower() = mail)
                                        If u Is Nothing OrElse u.Phone <> txtPh.Text.Trim() Then
                                            ShowError("We could not match that email and phone number. Please check and try again.")
                                            Return
                                        End If
                                        If u.Provider <> "" Then
                                            ShowError("This account uses " & u.Provider & " sign-in. Use the Continue with " & u.Provider & " button.")
                                            Return
                                        End If
                                        If txtN1.Text.Length < 6 Then
                                            ShowError("The new password must have at least 6 characters.")
                                            Return
                                        End If
                                        If txtN1.Text <> txtN2.Text Then
                                            ShowError("The two passwords are not the same.")
                                            Return
                                        End If
                                        u.Password = txtN1.Text
                                        ShowInfo("Your password has been changed. You can now log in with the new password.")
                                        dlg.DialogResult = DialogResult.OK
                                    End Sub

            If dlg.ShowDialog() = DialogResult.OK Then
                txtLoginEmail.Text = txtMail.Text.Trim()
                txtLoginPass.Clear()
                txtLoginPass.Focus()
            End If
        End Using
    End Sub

    ' ---------- LOG IN ----------
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        lblLoginMsg.Text = ""
        Dim email As String = txtLoginEmail.Text.Trim()
        Dim pass As String = txtLoginPass.Text

        If email = "" OrElse pass = "" Then
            lblLoginMsg.Text = "Please enter your email and password."
            ShowError("Please enter your email and password.")
            Exit Sub
        End If

        Dim found As User
        If email.ToLower() = AdminEmail AndAlso pass = AdminPassword Then
            ' the admin has one special built-in account. Nobody can sign up with this email.
            found = New User With {.Name = "Admin", .Email = AdminEmail, .Role = "admin"}
        Else
            ' everybody else needs an account made on Create account or by Google / Facebook sign-in
            found = Users.FirstOrDefault(Function(u) u.Email.ToLower() = email.ToLower() AndAlso u.Password = pass)
        End If

        If found Is Nothing Then
            lblLoginMsg.Text = "Wrong email or password."
            ShowError("Wrong email or password. Please try again.")
            txtLoginPass.Clear()
            txtLoginPass.Focus()
            Exit Sub
        End If

        ' remember (or forget) the email
        If chkRemember.Checked Then
            File.WriteAllText(RememberFile, email)
        ElseIf File.Exists(RememberFile) Then
            File.Delete(RememberFile)
        End If

        GoToHomeForm(found)
    End Sub

    ' ---------- CREATE ACCOUNT ----------
    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        lblCreateMsg.Text = ""
        Dim name As String = txtName.Text.Trim()
        Dim email As String = txtEmail.Text.Trim()
        Dim phone As String = txtPhone.Text.Replace(" ", "")
        Dim pass As String = txtPass.Text

        If name = "" Then Fail("Please enter your full name.", txtName) : Exit Sub
        If Not Regex.IsMatch(email, "^[^@\s]+@[^@\s]+\.[^@\s]+$") Then
            Fail("Please enter a valid email, for example ama@gmail.com.", txtEmail) : Exit Sub
        End If
        If email.ToLower() = AdminEmail OrElse Users.Any(Function(u) u.Email.ToLower() = email.ToLower()) Then
            Fail("That email is already registered. Try logging in.", txtEmail) : Exit Sub
        End If
        If Not Regex.IsMatch(phone, "^0\d{9}$") Then
            Fail("Phone must be 10 digits, for example 0241234567.", txtPhone) : Exit Sub
        End If
        If pass.Length < 6 Then Fail("Password must be at least 6 characters.", txtPass) : Exit Sub

        ' only a resident has an area. A collector chooses the area later, when pressing Go online.
        Dim newUser As New User With {.Name = name, .Email = email, .Phone = phone, .Password = pass}
        If rdoCollector.Checked Then
            newUser.Role = "collector"
            newUser.Area = ""
        Else
            newUser.Role = "resident"
            newUser.Area = cboCreateArea.Text
            newUser.Lat = AreaCoords(newUser.Area)(0)
            newUser.Lon = AreaCoords(newUser.Area)(1)
        End If
        Users.Add(newUser)

        ' clear the boxes and go straight to the user's page
        txtName.Clear() : txtEmail.Clear() : txtPhone.Clear() : txtPass.Clear()
        ShowInfo("Account created. Welcome, " & newUser.Name & "!")
        GoToHomeForm(newUser)
    End Sub

    ' Resident: the Area row shows. Collector: the Area row hides and everything below moves up.
    Private Sub RoleChanged(sender As Object, e As EventArgs) Handles rdoResident.CheckedChanged, rdoCollector.CheckedChanged
        Dim showArea As Boolean = rdoResident.Checked
        lblArea.Visible = showArea
        cboCreateArea.Visible = showArea
        If showArea AndAlso rowsMoved Then
            For Each c As Control In rowsBelowArea : c.Top += areaGap : Next
            rowsMoved = False
        ElseIf (Not showArea) AndAlso (Not rowsMoved) Then
            For Each c As Control In rowsBelowArea : c.Top -= areaGap : Next
            rowsMoved = True
        End If
    End Sub

    ' ---------- SIGN IN WITH GOOGLE OR FACEBOOK ----------
    ' Social accounts are residents only. The first time, the person agrees to create a linked
    ' TidyAccra account. After that, one click signs them in.
    ' (AskSocialEmail in modLink is a demo window. A real app would get the email from Google or Facebook.)
    Private Sub btnGoogle_Click(sender As Object, e As EventArgs) Handles btnGoogle.Click
        SocialSignIn("Google")
    End Sub

    Private Sub btnFacebook_Click(sender As Object, e As EventArgs) Handles btnFacebook.Click
        SocialSignIn("Facebook")
    End Sub

    Private Sub SocialSignIn(provider As String)
        Dim email As String = AskSocialEmail(provider)
        If email = "" Then Exit Sub                  ' Cancel
        If Not Regex.IsMatch(email, "^[^@\s]+@[^@\s]+\.[^@\s]+$") Then
            ShowError("Please type a valid " & provider & " email, for example ama@gmail.com.") : Exit Sub
        End If

        Dim found As User = Users.FirstOrDefault(Function(u) u.Email.ToLower() = email.ToLower())
        If email.ToLower() = AdminEmail OrElse (found IsNot Nothing AndAlso found.Provider <> provider) Then
            ShowError("That email already has a TidyAccra account with a password. Please log in with your password.")
            Exit Sub
        End If

        If found Is Nothing Then
            If Not Confirm("No TidyAccra account is linked to " & email & " yet." & vbCrLf &
                           "Create one with " & provider & "?") Then Exit Sub
            found = New User With {.Name = email.Substring(0, email.IndexOf("@")), .Email = email, .Phone = "Not given",
                                   .Role = "resident", .Provider = provider, .Password = Guid.NewGuid().ToString(),
                                   .Area = AreaCoords.Keys.First()}
            found.Lat = AreaCoords(found.Area)(0)
            found.Lon = AreaCoords(found.Area)(1)
            Users.Add(found)
            ShowInfo("Account created and linked to " & provider & ". Next time one click is enough. You can set your address on the Request a pickup page.")
        End If
        GoToHomeForm(found)
    End Sub

    ' Shows the problem on the form and puts the cursor in the box to fix
    Private Sub Fail(message As String, target As Control)
        lblCreateMsg.Text = message
        ShowError(message)
        target.Focus()
    End Sub

    ' ---------- EXIT ----------
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    ' ---------- LINKING THE FORMS ----------
    ' Everyone lands on frmHome. The navigation bar there leads to the pages for their role.
    Private Sub GoToHomeForm(who As User)
        CurrentUser = who
        txtLoginPass.Clear()
        OpenForm(New frmHome())
    End Sub

    ' Shows the new page, then hides this form (it comes back when somebody logs out)
    Private Sub OpenForm(target As Form)
        WelcomeForm = Me
        target.Show()
        Me.Hide()
    End Sub

End Class
