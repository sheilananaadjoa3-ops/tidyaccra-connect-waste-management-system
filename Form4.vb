' ============================================================
' frmHome.vb : the home page everybody sees first after logging in.
' The navigation bar (built by BuildNav in modLink) leads to the other pages.
' ============================================================
Public Class frmHome
    Inherits Form

    ' ---------- CONTROLS (built in code, so you do not design anything) ----------
    Private WithEvents pnlHeader As Panel
    Private WithEvents picLogo As PictureBox
    Private WithEvents lblTitle As Label
    Private WithEvents lblTagline As Label
    Private WithEvents lblHello As Label
    Private WithEvents lblHeroTitle As Label
    Private WithEvents lblHeroText As Label
    Private WithEvents btnMyPage As Button
    Private WithEvents lblStep1 As Label
    Private WithEvents lblStep2 As Label
    Private WithEvents lblStep3 As Label
    Private WithEvents lblStep4 As Label
    Private WithEvents btnLogout As Button
    Private WithEvents ssMain As StatusStrip
    Private WithEvents tslUser As ToolStripStatusLabel

    Public Sub New()
        MyBase.New()
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Me.SuspendLayout()
        Me.Text = "TidyAccra Connect: Home"
        Me.ClientSize = New Size(900, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.AutoScaleMode = AutoScaleMode.None
        pnlHeader = NewPanel(Me, 0, 0, 900, 90)
        picLogo = NewPicture(pnlHeader, 14, 20, 50, 50)
        lblTitle = NewLabel(pnlHeader, "TidyAccra Connect", 76, 12, 0, 0, size:=22, bold:=True, fore:="1B4332")
        lblTagline = NewLabel(pnlHeader, "Clean streets. Green Accra. One pickup at a time.", 78, 54, 0, 0, size:=11, italic:=True)
        lblHello = NewLabel(Me, "", 24, 110, 0, 0, size:=18, bold:=True, fore:="1B4332")
        lblHeroTitle = NewLabel(Me, "", 24, 150, 852, 80, size:=26, bold:=True, fore:="1B4332")
        lblHeroText = NewLabel(Me, "", 24, 236, 852, 50, size:=12, fore:="2D6A4F")
        btnMyPage = NewButton(Me, "", 24, 296, 260, 48)
        lblStep1 = NewLabel(Me, "", 24, 376, 204, 120, size:=10, back:="FFFFFF", pad:=10)
        lblStep2 = NewLabel(Me, "", 240, 376, 204, 120, size:=10, back:="FFFFFF", pad:=10)
        lblStep3 = NewLabel(Me, "", 456, 376, 204, 120, size:=10, back:="FFFFFF", pad:=10)
        lblStep4 = NewLabel(Me, "", 672, 376, 204, 120, size:=10, back:="FFFFFF", pad:=10)
        btnLogout = NewButton(Me, "Log &out", 756, 512, 120, 38, "danger")
        ssMain = New StatusStrip()
        tslUser = New ToolStripStatusLabel("")
        ssMain.Items.Add(tslUser)
        Me.Controls.Add(ssMain)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    ' ---------- form opens ----------
    Private Sub frmHome_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplyTheme(Me)
        BuildNav(Me, "home")                         ' navigation bar at the top
        Me.Text = "TidyAccra Connect: Home"
        tslUser.Text = "Logged in as " & CurrentUser.Email

        lblHello.Text = "Welcome, " & CurrentUser.Name
        lblHeroTitle.Text = "Clean streets. Green Accra. One pickup at a time."

        ' the text and the main button depend on who is logged in
        Select Case CurrentUser.Role
            Case "resident"
                lblHeroText.Text = "Request a pickup in a few clicks. The nearest collector sees it, accepts it and comes to your door."
                btnMyPage.Text = "&Request a pickup"
            Case "collector"
                lblHeroText.Text = "Go online to see pickup requests near you. The first collector to accept gets the job."
                btnMyPage.Text = "Go to my &requests"
            Case Else
                lblHeroText.Text = "See every request, collector and resident. Assign only collectors who are close enough."
                btnMyPage.Text = "Open the &dashboard"
        End Select

        lblStep1.Text = "STEP 1" & vbCrLf & "Log in" & vbCrLf & "Use your email and password, or Google or Facebook."
        lblStep2.Text = "STEP 2" & vbCrLf & "Request" & vbCrLf & "Choose the waste type and date. Your saved address is used."
        lblStep3.Text = "STEP 3" & vbCrLf & "Nearest collector" & vbCrLf & "Online collectors within " & BaseRadius & " km see it. The first to accept gets the job."
        lblStep4.Text = "STEP 4" & vbCrLf & "Collected" & vbCrLf & "The status updates by itself. No refresh button."
    End Sub

    Private Sub pnlHeader_Paint(sender As Object, e As PaintEventArgs) Handles pnlHeader.Paint
        PaintHeader(sender, e)
    End Sub

    ' ---------- main button: go to the page for this role ----------
    Private Sub btnMyPage_Click(sender As Object, e As EventArgs) Handles btnMyPage.Click
        Select Case CurrentUser.Role
            Case "resident" : NavigateTo(Me, "resident")
            Case "collector" : NavigateTo(Me, "collector")
            Case Else : NavigateTo(Me, "admin")
        End Select
    End Sub

    ' ---------- log out (closing this form brings frmWelcome back) ----------
    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        If Not Confirm("Do you want to log out?") Then Exit Sub
        CurrentUser = Nothing
        Me.Close()
    End Sub

End Class
