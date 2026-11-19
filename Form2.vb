' ============================================================
' frmSubscribe.vb : optional plans. A normal page: reached from the Subscription link in the bar
' or from the Subscribe / Change plan button on frmResident.
' ============================================================
Public Class frmSubscribe
    Inherits Form

    ' ---------- CONTROLS (built in code, so you do not design anything) ----------
    Private WithEvents pnlHeader As Panel
    Private WithEvents picLogo As PictureBox
    Private WithEvents lblTitle As Label
    Private WithEvents btnLogout As Button
    Private WithEvents lblChoose As Label
    Private WithEvents lblOptional As Label
    Private WithEvents lblPopular As Label
    Private WithEvents grpBasic As GroupBox
    Private WithEvents lblBasicName As Label
    Private WithEvents lblBasicPrice As Label
    Private WithEvents lblBasicNote As Label
    Private WithEvents lblBasicPerk1 As Label
    Private WithEvents lblBasicPerk2 As Label
    Private WithEvents lblBasicPerk3 As Label
    Private WithEvents btnBasic As Button
    Private WithEvents grpStandard As GroupBox
    Private WithEvents lblStandardName As Label
    Private WithEvents lblStandardPrice As Label
    Private WithEvents lblStandardNote As Label
    Private WithEvents lblStandardPerk1 As Label
    Private WithEvents lblStandardPerk2 As Label
    Private WithEvents lblStandardPerk3 As Label
    Private WithEvents btnStandard As Button
    Private WithEvents grpBusiness As GroupBox
    Private WithEvents lblBusinessName As Label
    Private WithEvents lblBusinessPrice As Label
    Private WithEvents lblBusinessNote As Label
    Private WithEvents lblBusinessPerk1 As Label
    Private WithEvents lblBusinessPerk2 As Label
    Private WithEvents lblBusinessPerk3 As Label
    Private WithEvents btnBusiness As Button
    Private WithEvents lblPay As Label
    Private WithEvents cboPay As ComboBox
    Private WithEvents lblSim As Label
    Private WithEvents btnBack As Button
    Private WithEvents ssMain As StatusStrip
    Private WithEvents tslUser As ToolStripStatusLabel

    Public Sub New()
        MyBase.New()
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Me.SuspendLayout()
        Me.Text = "Subscribe"
        Me.ClientSize = New Size(1000, 560)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.AutoScaleMode = AutoScaleMode.None
        pnlHeader = NewPanel(Me, 0, 0, 1000, 70)
        picLogo = NewPicture(pnlHeader, 14, 10, 50, 50)
        lblTitle = NewLabel(pnlHeader, "Subscribe and save", 76, 16, 0, 0, size:=18, bold:=True)
        btnLogout = NewButton(pnlHeader, "Log &out", 870, 17, 110, 36, "danger")
        lblChoose = NewLabel(Me, "Choose a plan", 20, 84, 0, 0, size:=18, bold:=True)
        lblOptional = NewLabel(Me, "", 20, 120, 960, 22)
        lblPopular = NewLabel(Me, "Most popular", 355, 150, 290, 24, bold:=True, back:="FFD166", align:=ContentAlignment.MiddleCenter)
        grpBasic = NewGroup(Me, "", 20, 174, 290, 276)
        lblBasicName = NewLabel(grpBasic, "Basic", 20, 22, 0, 0, size:=16, bold:=True)
        lblBasicPrice = NewLabel(grpBasic, "", 20, 58, 0, 0, size:=22, bold:=True, fore:="1E7A3C")
        lblBasicNote = NewLabel(grpBasic, "per month (estimate)", 20, 98, 0, 0, fore:="555555")
        lblBasicPerk1 = NewLabel(grpBasic, "", 20, 130, 0, 0)
        lblBasicPerk2 = NewLabel(grpBasic, "Track every pickup", 20, 156, 0, 0)
        lblBasicPerk3 = NewLabel(grpBasic, "No payment box when booking", 20, 182, 0, 0)
        btnBasic = NewButton(grpBasic, "Subscribe to Basic", 20, 218, 250, 42, "secondary")
        grpStandard = NewGroup(Me, "", 355, 174, 290, 276)
        lblStandardName = NewLabel(grpStandard, "Standard", 20, 22, 0, 0, size:=16, bold:=True)
        lblStandardPrice = NewLabel(grpStandard, "", 20, 58, 0, 0, size:=22, bold:=True, fore:="1E7A3C")
        lblStandardNote = NewLabel(grpStandard, "per month (estimate)", 20, 98, 0, 0, fore:="555555")
        lblStandardPerk1 = NewLabel(grpStandard, "", 20, 130, 0, 0)
        lblStandardPerk2 = NewLabel(grpStandard, "Track every pickup", 20, 156, 0, 0)
        lblStandardPerk3 = NewLabel(grpStandard, "No payment box when booking", 20, 182, 0, 0)
        btnStandard = NewButton(grpStandard, "Subscribe to Standard", 20, 218, 250, 42)
        grpBusiness = NewGroup(Me, "", 690, 174, 290, 276)
        lblBusinessName = NewLabel(grpBusiness, "Business", 20, 22, 0, 0, size:=16, bold:=True)
        lblBusinessPrice = NewLabel(grpBusiness, "", 20, 58, 0, 0, size:=22, bold:=True, fore:="1E7A3C")
        lblBusinessNote = NewLabel(grpBusiness, "per month (estimate)", 20, 98, 0, 0, fore:="555555")
        lblBusinessPerk1 = NewLabel(grpBusiness, "", 20, 130, 0, 0)
        lblBusinessPerk2 = NewLabel(grpBusiness, "Track every pickup", 20, 156, 0, 0)
        lblBusinessPerk3 = NewLabel(grpBusiness, "No payment box when booking", 20, 182, 0, 0)
        btnBusiness = NewButton(grpBusiness, "Subscribe to Business", 20, 218, 250, 42, "secondary")
        lblPay = NewLabel(Me, "Pay with", 20, 488, 0, 0, bold:=True)
        cboPay = NewCombo(Me, 90, 484, 230)
        lblSim = NewLabel(Me, "Payment is simulated. Prices are illustrative.", 350, 490, 0, 0, fore:="555555")
        btnBack = NewButton(Me, "&Not now, go back", 800, 478, 170, 38, "secondary")
        ssMain = New StatusStrip()
        tslUser = New ToolStripStatusLabel("")
        ssMain.Items.Add(tslUser)
        Me.Controls.Add(ssMain)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    Private Sub frmSubscribe_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplyTheme(Me)
        BuildNav(Me, "subscribe")                  ' navigation bar at the top
        Me.Text = "Subscribe: " & CurrentUser.Name
        tslUser.Text = "Logged in as " & CurrentUser.Email & "   |   Subscription: " & If(CurrentUser.Plan = "", "none", CurrentUser.Plan)
        Me.CancelButton = btnBack                    ' Esc = go back

        lblOptional.Text = "Subscribing is optional. You can always book one pickup at a time for " & Cedi & PerPickupPrice & " (estimate)."
        lblBasicPrice.Text = Cedi & PlanPrice("Basic")
        lblStandardPrice.Text = Cedi & PlanPrice("Standard")
        lblBusinessPrice.Text = Cedi & PlanPrice("Business")
        lblBasicPerk1.Text = ChrW(10004) & " " & PlanPickups("Basic") & " pickups a month"
        lblStandardPerk1.Text = ChrW(10004) & " " & PlanPickups("Standard") & " pickups a month"
        lblBusinessPerk1.Text = ChrW(10004) & " " & PlanPickups("Business") & " pickups a month"

        cboPay.Items.AddRange(PayMethods)
        cboPay.SelectedIndex = 0

        ' the plan you already have cannot be chosen again
        MarkCurrentPlan(btnBasic, "Basic")
        MarkCurrentPlan(btnStandard, "Standard")
        MarkCurrentPlan(btnBusiness, "Business")
    End Sub

    Private Sub pnlHeader_Paint(sender As Object, e As PaintEventArgs) Handles pnlHeader.Paint
        PaintHeader(sender, e)
    End Sub

    Private Sub MarkCurrentPlan(b As Button, planName As String)
        If CurrentUser.Plan = planName Then
            b.Text = "Your current plan"
            b.Enabled = False
        Else
            b.Text = "Subscribe to " & planName
        End If
    End Sub

    ' ---------- one handler for the three plan buttons ----------
    Private Sub PlanButtons_Click(sender As Object, e As EventArgs) Handles btnBasic.Click, btnStandard.Click, btnBusiness.Click
        Dim planName As String = DirectCast(sender, Button).Name.Substring(3)   ' btnBasic -> Basic
        Dim u As User = CurrentUser

        ' mobile money payment (simulated): ask first, then treat it as paid
        Dim number As String = If(u.Phone = "Not given", "your mobile money number", u.Phone)
        If Not Confirm("Pay " & Cedi & PlanPrice(planName) & " (estimate) for the " & planName & " plan with " & cboPay.Text & "?" & vbCrLf &
                       "A payment request is sent to " & number & ". (Payment is simulated.)") Then Exit Sub

        u.Plan = planName
        u.PickupsLeft = PlanPickups(planName)
        u.RenewsOn = Date.Today.AddMonths(1)
        ShowInfo("Payment received: " & Cedi & PlanPrice(planName) & " by " & cboPay.Text & " (simulated)." & vbCrLf &
                 "You are now on the " & planName & " plan (" & PlanPickups(planName) & " pickups).")
        NavigateTo(Me, "resident")                   ' back to the resident page, which shows the new plan
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        NavigateTo(Me, "resident")
    End Sub

    ' Log out from here
    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        If Not Confirm("Do you want to log out?") Then Exit Sub
        CurrentUser = Nothing
        Me.Close()                                   ' closing sends the person back to Welcome
    End Sub

End Class
