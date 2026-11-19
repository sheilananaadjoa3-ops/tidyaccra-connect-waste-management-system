' ============================================================
' frmResident.vb : residents and businesses request pickups.
' Address is saved once. Subscription is optional.
' ============================================================
Imports System.Text.RegularExpressions

Public Class frmResident
    Inherits Form

    ' ---------- CONTROLS (built in code, so you do not design anything) ----------
    Private WithEvents pnlHeader As Panel
    Private WithEvents picLogo As PictureBox
    Private WithEvents lblHello As Label
    Private WithEvents btnLogout As Button
    Private WithEvents grpSub As GroupBox
    Private WithEvents lblPlan As Label
    Private WithEvents lblPlanState As Label
    Private WithEvents lblPrice As Label
    Private WithEvents lblLeft As Label
    Private WithEvents prgLeft As ProgressBar
    Private WithEvents lblRenew As Label
    Private WithEvents btnSubscribe As Button
    Private WithEvents grpAddress As GroupBox
    Private WithEvents lblAddr1 As Label
    Private WithEvents lblAddr2 As Label
    Private WithEvents lblAddr3 As Label
    Private WithEvents lblAddr4 As Label
    Private WithEvents btnEditAddress As Button
    Private WithEvents pnlEditAddress As Panel
    Private WithEvents lblHouse As Label
    Private WithEvents txtHouseNo As TextBox
    Private WithEvents lblStreet As Label
    Private WithEvents txtStreet As TextBox
    Private WithEvents lblLandmark As Label
    Private WithEvents txtLandmark As TextBox
    Private WithEvents lblEditArea As Label
    Private WithEvents cboArea As ComboBox
    Private WithEvents lblDigital As Label
    Private WithEvents txtDigital As TextBox
    Private WithEvents lblDirections As Label
    Private WithEvents txtDirections As TextBox
    Private WithEvents lblEditMsg As Label
    Private WithEvents btnSaveAddress As Button
    Private WithEvents btnCancelAddress As Button
    Private WithEvents grpRequest As GroupBox
    Private WithEvents lblWhere As Label
    Private WithEvents pnlWhere As Panel
    Private WithEvents rdoUseLocation As RadioButton
    Private WithEvents rdoSavedAddress As RadioButton
    Private WithEvents lblLocState As Label
    Private WithEvents lblWhen As Label
    Private WithEvents pnlWhen As Panel
    Private WithEvents rdoToday As RadioButton
    Private WithEvents rdoTomorrow As RadioButton
    Private WithEvents rdoPickDate As RadioButton
    Private WithEvents dtpDate As DateTimePicker
    Private WithEvents lblWaste As Label
    Private WithEvents cboWaste As ComboBox
    Private WithEvents lblNote As Label
    Private WithEvents txtNote As TextBox
    Private WithEvents lblCovered As Label
    Private WithEvents lblPay As Label
    Private WithEvents lblPayAmount As Label
    Private WithEvents cboPay As ComboBox
    Private WithEvents btnRequest As Button
    Private WithEvents lblRequestMsg As Label
    Private WithEvents grpLive As GroupBox
    Private WithEvents lblLiveRef As Label
    Private WithEvents lblSteps As Label
    Private WithEvents lblLiveInfo As Label
    Private WithEvents grpMine As GroupBox
    Private WithEvents lvMine As ListView
    Private WithEvents ssMain As StatusStrip
    Private WithEvents tslUser As ToolStripStatusLabel
    Private WithEvents tmrLive As System.Windows.Forms.Timer
    Private WithEvents ttHelp As ToolTip

    Public Sub New()
        MyBase.New()
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Me.SuspendLayout()
        Me.Text = "My pickups"
        Me.ClientSize = New Size(1000, 710)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.AutoScaleMode = AutoScaleMode.None
        pnlHeader = NewPanel(Me, 0, 0, 1000, 70)
        picLogo = NewPicture(pnlHeader, 14, 10, 50, 50)
        lblHello = NewLabel(pnlHeader, "Hello", 76, 16, 0, 0, size:=18, bold:=True)
        btnLogout = NewButton(pnlHeader, "Log &out", 870, 17, 110, 36, "danger")
        grpSub = NewGroup(Me, "Subscription", 20, 84, 300, 236)
        lblPlan = NewLabel(grpSub, "", 16, 28, 270, 26, size:=14, bold:=True)
        lblPlanState = NewLabel(grpSub, "", 196, 30, 76, 24, bold:=True, align:=ContentAlignment.MiddleCenter)
        lblPrice = NewLabel(grpSub, "", 16, 60, 270, 40)
        lblLeft = NewLabel(grpSub, "", 16, 102, 270, 24, bold:=True)
        prgLeft = NewProgress(grpSub, 16, 130, 268, 16)
        lblRenew = NewLabel(grpSub, "", 16, 154, 270, 22, fore:="555555")
        btnSubscribe = NewButton(grpSub, "&Subscribe", 16, 184, 268, 44)
        grpAddress = NewGroup(Me, "My pickup address", 20, 334, 300, 236)
        lblAddr1 = NewLabel(grpAddress, "", 16, 28, 270, 24, size:=12, bold:=True)
        lblAddr2 = NewLabel(grpAddress, "", 16, 58, 270, 22)
        lblAddr3 = NewLabel(grpAddress, "", 16, 82, 270, 22)
        lblAddr4 = NewLabel(grpAddress, "", 16, 106, 270, 48)
        btnEditAddress = NewButton(grpAddress, "&Edit address", 16, 184, 130, 36, "secondary")
        pnlEditAddress = NewPanel(Me, 20, 84, 300, 486, back:="FFFFFF", border:=True)
        pnlEditAddress.Visible = False
        lblHouse = NewLabel(pnlEditAddress, "House / building no.", 10, 10, 0, 0, bold:=True)
        txtHouseNo = NewText(pnlEditAddress, 10, 32, 276, 26)
        lblStreet = NewLabel(pnlEditAddress, "Street name", 10, 66, 0, 0, bold:=True)
        txtStreet = NewText(pnlEditAddress, 10, 88, 276, 26)
        lblLandmark = NewLabel(pnlEditAddress, "Nearest landmark", 10, 122, 0, 0, bold:=True)
        txtLandmark = NewText(pnlEditAddress, 10, 144, 276, 26)
        lblEditArea = NewLabel(pnlEditAddress, "Area", 10, 178, 0, 0, bold:=True)
        cboArea = NewCombo(pnlEditAddress, 10, 200, 276)
        lblDigital = NewLabel(pnlEditAddress, "Digital address (optional)", 10, 234, 0, 0, bold:=True)
        txtDigital = NewText(pnlEditAddress, 10, 256, 276, 26)
        lblDirections = NewLabel(pnlEditAddress, "Directions for the collector", 10, 290, 0, 0, bold:=True)
        txtDirections = NewText(pnlEditAddress, 10, 312, 276, 56, multiline:=True)
        lblEditMsg = NewLabel(pnlEditAddress, "", 10, 372, 276, 22, fore:="A8400F")
        btnSaveAddress = NewButton(pnlEditAddress, "&Save", 10, 400, 134, 40)
        btnCancelAddress = NewButton(pnlEditAddress, "Cancel", 152, 400, 134, 40, "secondary")
        grpRequest = NewGroup(Me, "Request a pickup", 340, 84, 640, 290)
        lblWhere = NewLabel(grpRequest, "Where?", 16, 32, 0, 0, bold:=True)
        pnlWhere = NewPanel(grpRequest, 86, 22, 170, 60)
        rdoUseLocation = NewRadio(pnlWhere, "Use my &location", 4, 2, 160, 24, isChecked:=True)
        rdoSavedAddress = NewRadio(pnlWhere, "My saved a&ddress", 4, 30, 160, 24)
        lblLocState = NewLabel(grpRequest, "", 262, 26, 360, 26, back:="CFEECF", align:=ContentAlignment.MiddleLeft)
        lblWhen = NewLabel(grpRequest, "When?", 16, 100, 0, 0, bold:=True)
        pnlWhen = NewPanel(grpRequest, 86, 92, 320, 34)
        rdoToday = NewRadio(pnlWhen, "&Today", 4, 6, 0, 0, isChecked:=True)
        rdoTomorrow = NewRadio(pnlWhen, "Tomo&rrow", 94, 6, 0, 0)
        rdoPickDate = NewRadio(pnlWhen, "Pick a &date", 204, 6, 0, 0)
        dtpDate = NewDate(grpRequest, 410, 94, 206)
        lblWaste = NewLabel(grpRequest, "Type of waste", 16, 142, 0, 0, bold:=True)
        cboWaste = NewCombo(grpRequest, 130, 138, 230)
        lblNote = NewLabel(grpRequest, "Note", 380, 142, 0, 0, bold:=True)
        txtNote = NewText(grpRequest, 430, 138, 186, 28)
        lblCovered = NewLabel(grpRequest, "", 16, 184, 600, 24, bold:=True, fore:="1E7A3C")
        lblPay = NewLabel(grpRequest, "Pay for this pickup", 16, 184, 0, 0, bold:=True)
        lblPayAmount = NewLabel(grpRequest, "", 170, 184, 0, 0, bold:=True, fore:="1E7A3C")
        cboPay = NewCombo(grpRequest, 338, 180, 270)
        btnRequest = NewButton(grpRequest, "&Request pickup", 16, 220, 250, 46)
        lblRequestMsg = NewLabel(grpRequest, "", 282, 234, 340, 24, bold:=True)
        grpLive = NewGroup(Me, "Live status", 340, 386, 640, 118)
        lblLiveRef = NewLabel(grpLive, "", 16, 26, 610, 24, size:=11, bold:=True)
        lblSteps = NewLabel(grpLive, "", 16, 52, 610, 24)
        lblLiveInfo = NewLabel(grpLive, "", 16, 80, 610, 24)
        grpMine = NewGroup(Me, "My pickups", 340, 516, 640, 150)
        lvMine = NewList(grpMine, 14, 22, 612, 118, "Reference:130|Date:120|Status:120|Collector:220")
        ssMain = New StatusStrip()
        tslUser = New ToolStripStatusLabel("")
        ssMain.Items.Add(tslUser)
        Me.Controls.Add(ssMain)
        tmrLive = New System.Windows.Forms.Timer()
        ttHelp = New ToolTip()
        AddHandler Me.FormClosed, Sub(s As Object, a As FormClosedEventArgs)
                                      tmrLive.Stop()
                                      tmrLive.Dispose()
                                  End Sub
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    ' ---------- form opens ----------
    Private Sub frmResident_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplyTheme(Me)
        BuildNav(Me, "resident")                   ' navigation bar at the top
        Me.Text = "My pickups: " & CurrentUser.Name
        lblHello.Text = "Hello, " & CurrentUser.Name
        tslUser.Text = "Logged in as " & CurrentUser.Email

        cboWaste.Items.AddRange(WasteTypes)
        cboPay.Items.AddRange(PayMethods)
        cboArea.Items.AddRange(AreaCoords.Keys.ToArray())
        cboWaste.SelectedIndex = 0
        cboPay.SelectedIndex = 0
        dtpDate.MinDate = Date.Today                 ' a past date cannot be chosen
        dtpDate.Value = Date.Today
        dtpDate.Enabled = False                      ' only used when "Pick a date" is chosen

        ttHelp.SetToolTip(txtNote, "Anything that helps the collector, for example: please knock")
        ttHelp.SetToolTip(txtDigital, "GhanaPostGPS address, for example GA-032-1458 (optional)")
        ttHelp.SetToolTip(btnSubscribe, "Subscribe and save, or change your plan")

        tmrLive.Interval = 3000                      ' check for news every 3 seconds
        tmrLive.Enabled = True
        RefreshAll()
    End Sub

    Private Sub pnlHeader_Paint(sender As Object, e As PaintEventArgs) Handles pnlHeader.Paint
        PaintHeader(sender, e)
    End Sub

    Private Sub RefreshAll()
        RefreshSubscription()
        RefreshAddress()
        UpdateLocState()
        RefreshList()
        RefreshLive()
    End Sub

    ' ---------- subscription card ----------
    Private Sub RefreshSubscription()
        Dim u As User = CurrentUser
        If u.Plan = "" Then
            lblPlan.Text = "No subscription yet"
            lblPlanState.Visible = False
            lblPrice.Text = "You can still book. Subscribe and save: 4 to 12 pickups a month."
            lblLeft.Text = "From " & Cedi & PlanPrice("Basic") & " per month (estimate)"
            prgLeft.Visible = False
            lblRenew.Visible = False
            btnSubscribe.Text = "&Subscribe"
            btnSubscribe.Tag = ""                    ' main green button
        Else
            Dim st As String = SubStatus(u)
            lblPlan.Text = u.Plan & " plan"
            lblPlanState.Text = st
            lblPlanState.BackColor = StatusColor(st)
            lblPlanState.Visible = True
            lblPrice.Text = Cedi & PlanPrice(u.Plan) & " per month (estimate)"
            lblLeft.Text = u.PickupsLeft & " of " & PlanPickups(u.Plan) & " pickups left this month"
            prgLeft.Maximum = PlanPickups(u.Plan)
            prgLeft.Value = Math.Min(u.PickupsLeft, prgLeft.Maximum)
            prgLeft.Visible = True
            lblRenew.Text = "Renews on " & u.RenewsOn.ToString("d MMM yyyy")
            lblRenew.Visible = True
            btnSubscribe.Text = "Change &plan"
            btnSubscribe.Tag = "secondary"
        End If
        StyleButton(btnSubscribe)

        ' the payment row only shows when a subscription does not cover this pickup
        Dim covered As Boolean = UsesSubscription(u)
        lblCovered.Text = ChrW(10004) & " Covered by your " & u.Plan & " plan. No payment needed."
        lblCovered.Visible = covered
        lblPay.Visible = Not covered
        lblPayAmount.Text = Cedi & PerPickupPrice & " (estimate) with"
        lblPayAmount.Visible = Not covered
        cboPay.Visible = Not covered
        btnRequest.Text = If(covered, "&Request pickup", "&Request pickup (" & Cedi & PerPickupPrice & ")")
    End Sub

    Private Sub btnSubscribe_Click(sender As Object, e As EventArgs) Handles btnSubscribe.Click
        NavigateTo(Me, "subscribe")                  ' same as clicking Subscription in the bar
    End Sub

    ' ---------- saved address ----------
    Private Sub RefreshAddress()
        Dim u As User = CurrentUser
        If u.HasAddress() Then
            lblAddr1.Text = u.HouseNo & " " & u.Street & ", " & u.Area
            lblAddr2.Text = u.Landmark
            lblAddr3.Text = If(u.DigitalAddress = "", "", "Digital address: " & u.DigitalAddress)
            lblAddr4.Text = u.Directions
        Else
            lblAddr1.Text = "No address saved yet"
            lblAddr2.Text = "Press Edit address to add it once."
            lblAddr3.Text = ""
            lblAddr4.Text = ""
        End If
    End Sub

    Private Sub btnEditAddress_Click(sender As Object, e As EventArgs) Handles btnEditAddress.Click
        Dim u As User = CurrentUser
        txtHouseNo.Text = u.HouseNo
        txtStreet.Text = u.Street
        txtLandmark.Text = u.Landmark
        cboArea.Text = u.Area
        txtDigital.Text = u.DigitalAddress
        txtDirections.Text = u.Directions
        lblEditMsg.Text = ""
        pnlEditAddress.Visible = True                ' covers the two cards on the left
        pnlEditAddress.BringToFront()
        txtHouseNo.Focus()
    End Sub

    Private Sub btnSaveAddress_Click(sender As Object, e As EventArgs) Handles btnSaveAddress.Click
        Dim digital As String = txtDigital.Text.Trim().ToUpper()
        If txtHouseNo.Text.Trim() = "" Then ShowEditProblem("Please enter the house or building number.", txtHouseNo) : Exit Sub
        If txtStreet.Text.Trim() = "" Then ShowEditProblem("Please enter the street name.", txtStreet) : Exit Sub
        If txtLandmark.Text.Trim() = "" Then ShowEditProblem("Please enter a nearest landmark.", txtLandmark) : Exit Sub
        If cboArea.SelectedIndex = -1 Then ShowEditProblem("Please choose your area.", cboArea) : Exit Sub
        If digital <> "" AndAlso Not Regex.IsMatch(digital, "^[A-Z]{2}-\d{3,4}-\d{4}$") Then
            ShowEditProblem("A digital address looks like GA-032-1458.", txtDigital) : Exit Sub
        End If

        Dim u As User = CurrentUser
        u.HouseNo = txtHouseNo.Text.Trim()
        u.Street = txtStreet.Text.Trim()
        u.Landmark = txtLandmark.Text.Trim()
        u.Area = cboArea.Text
        u.DigitalAddress = digital
        u.Directions = txtDirections.Text.Trim()
        pnlEditAddress.Visible = False
        RefreshAddress()
        UpdateLocState()
        ShowInfo("Your address has been saved.")
    End Sub

    Private Sub ShowEditProblem(message As String, target As Control)
        lblEditMsg.Text = message
        ShowError(message)
        target.Focus()
    End Sub

    Private Sub btnCancelAddress_Click(sender As Object, e As EventArgs) Handles btnCancelAddress.Click
        pnlEditAddress.Visible = False
    End Sub

    ' ---------- where / when ----------
    Private Sub WhereChanged(sender As Object, e As EventArgs) Handles rdoUseLocation.CheckedChanged, rdoSavedAddress.CheckedChanged
        UpdateLocState()
    End Sub

    Private Sub rdoPickDate_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPickDate.CheckedChanged
        dtpDate.Enabled = rdoPickDate.Checked
    End Sub

    ' Shows where the pickup will be (green box next to "Use my location")
    Private Sub UpdateLocState()
        Dim lat As Double, lon As Double
        Dim src As String = GetCoords(CurrentUser, rdoUseLocation.Checked, lat, lon)
        lblLocState.Text = ChrW(9679) & " Location on: " & src
    End Sub

    ' ---------- REQUEST A PICKUP ----------
    Private Sub btnRequest_Click(sender As Object, e As EventArgs) Handles btnRequest.Click
        lblRequestMsg.ForeColor = Terracotta
        Dim u As User = CurrentUser

        If Not u.HasAddress() Then
            lblRequestMsg.Text = "Please add your address first (Edit address)."
            ShowError("Please add your address first (click Edit address).")
            Exit Sub
        End If

        ' 1. where is the pickup?
        Dim lat As Double, lon As Double
        GetCoords(u, rdoUseLocation.Checked, lat, lon)

        ' 2. when?
        Dim pickDate As Date = Date.Today
        If rdoTomorrow.Checked Then pickDate = Date.Today.AddDays(1)
        If rdoPickDate.Checked Then pickDate = dtpDate.Value.Date

        ' 3. how is it paid? (a subscription covers it, otherwise mobile money, simulated)
        Dim payText As String
        Dim payNote As String = ""
        If UsesSubscription(u) Then
            payText = "Subscription (" & u.Plan & ")"
        Else
            ' mobile money payment (simulated): ask, then treat it as paid
            Dim number As String = If(u.Phone = "Not given", "your mobile money number", u.Phone)
            If Not Confirm("Pay " & Cedi & PerPickupPrice & " (estimate) with " & cboPay.Text & "?" & vbCrLf &
                           "A payment request is sent to " & number & ". (Payment is simulated.)") Then
                lblRequestMsg.Text = "Payment cancelled. No request was made."
                Exit Sub
            End If
            payText = cboPay.Text & " (" & Cedi & PerPickupPrice & " paid, simulated)"
            payNote = vbCrLf & "Payment received: " & Cedi & PerPickupPrice & " by " & cboPay.Text & " (simulated)."
        End If
        If UsesSubscription(u) Then u.PickupsLeft -= 1     ' a pickup is used from the plan

        ' 4. save the booking. Collectors near the pickup will see it.
        Dim b As New Booking With {.Ref = NewReference(), .Owner = u, .HouseNo = u.HouseNo, .Street = u.Street,
                                   .Landmark = u.Landmark, .Area = u.Area, .DigitalAddress = u.DigitalAddress,
                                   .Directions = u.Directions, .Waste = cboWaste.Text, .Note = txtNote.Text.Trim(),
                                   .Pay = payText, .PickupDate = pickDate, .Lat = lat, .Lon = lon}
        Bookings.Add(b)

        txtNote.Clear()
        lblRequestMsg.ForeColor = Green
        lblRequestMsg.Text = ChrW(10004) & " Requested! Reference " & b.Ref
        ShowInfo("Pickup requested." & vbCrLf & "Reference: " & b.Ref & payNote & vbCrLf & "Nearby collectors have been notified.")
        lvMine.SelectedItems.Clear()                 ' so Live status shows the new request
        RefreshAll()
    End Sub

    ' ---------- my pickups list ----------
    Private Sub RefreshList()
        Dim keep As String = If(lvMine.SelectedItems.Count > 0, lvMine.SelectedItems(0).Text, "")
        lvMine.BeginUpdate()
        lvMine.Items.Clear()
        Dim mine = Bookings.Where(Function(x) x.Owner Is CurrentUser).OrderByDescending(Function(x) x.CreatedAt)
        For Each b As Booking In mine
            Dim item As New ListViewItem(b.Ref)
            item.SubItems.Add(If(b.PickupDate = Date.Today, "Today", b.PickupDate.ToString("d MMM")))
            item.SubItems.Add(b.Status)
            item.SubItems.Add(If(b.Collector Is Nothing, "-", b.Collector.Name))
            item.BackColor = StatusColor(b.Status)
            lvMine.Items.Add(item)
            If b.Ref = keep Then item.Selected = True
        Next
        lvMine.EndUpdate()
    End Sub

    Private Sub lvMine_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvMine.SelectedIndexChanged
        RefreshLive()
    End Sub

    ' ---------- live status (the row you clicked, or your newest open request) ----------
    Private Sub RefreshLive()
        Dim b As Booking = Nothing
        If lvMine.SelectedItems.Count > 0 Then
            Dim wanted As String = lvMine.SelectedItems(0).Text
            b = Bookings.FirstOrDefault(Function(x) x.Ref = wanted)
        Else
            Dim openOnes = Bookings.Where(Function(x) x.Owner Is CurrentUser AndAlso x.Status <> "Collected")
            b = openOnes.OrderByDescending(Function(x) x.CreatedAt).FirstOrDefault()
        End If

        If b Is Nothing Then
            lblLiveRef.Text = "No active pickup right now."
            lblSteps.Text = ""
            lblLiveInfo.Text = "Press Request pickup and nearby collectors will be alerted."
            Exit Sub
        End If

        Dim dot As String = ChrW(9679)               ' filled circle
        Dim ring As String = ChrW(9675)              ' empty circle
        Select Case b.Status
            Case "Pending"
                Dim radius As Double = CurrentRadius(b)
                Dim n As Integer = CollectorsNear(b).Where(Function(c) BookingDistance(b, c) <= radius).Count()
                lblLiveRef.Text = b.Ref & "   Looking for a collector near you..."
                lblSteps.Text = dot & " Requested  --  " & ring & " Accepted  --  " & ring & " Collected"
                lblLiveInfo.Text = n & " collector(s) online within " & radius & " km. The first to accept gets your pickup."
            Case "Assigned"
                lblLiveRef.Text = b.Ref & "   " & b.Collector.Name & " accepted your pickup"
                lblSteps.Text = dot & " Requested  --  " & dot & " Accepted  --  " & ring & " Collected"
                lblLiveInfo.Text = "About " & BookingDistance(b, b.Collector).ToString("0.0") & " km away. Phone: " & b.Collector.Phone
            Case Else
                lblLiveRef.Text = b.Ref & "   Collected. Thank you!"
                lblSteps.Text = dot & " Requested  --  " & dot & " Accepted  --  " & dot & " Collected"
                lblLiveInfo.Text = "Collected by " & b.Collector.Name
        End Select
    End Sub

    ' every 3 seconds: pick up news from the collector
    Private Sub tmrLive_Tick(sender As Object, e As EventArgs) Handles tmrLive.Tick
        RefreshList()
        RefreshLive()
    End Sub

    ' ---------- log out (closing this form brings frmWelcome back) ----------
    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        If Not Confirm("Do you want to log out?") Then Exit Sub
        CurrentUser = Nothing
        Me.Close()
    End Sub

End Class
