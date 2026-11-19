' ============================================================
' frmCollector.vb : collectors go Online, see requests near them and accept.
' ============================================================
Public Class frmCollector
    Inherits Form

    ' ---------- CONTROLS (built in code, so you do not design anything) ----------
    Private WithEvents pnlHeader As Panel
    Private WithEvents picLogo As PictureBox
    Private WithEvents lblHello As Label
    Private WithEvents btnLogout As Button
    Private WithEvents lblBanner As Label
    Private WithEvents lblOnline As Label
    Private WithEvents lblMyArea As Label
    Private WithEvents cboMyArea As ComboBox
    Private WithEvents btnOnline As Button
    Private WithEvents grpNear As GroupBox
    Private WithEvents lvNear As ListView
    Private WithEvents btnAccept As Button
    Private WithEvents btnDecline As Button
    Private WithEvents lblRule As Label
    Private WithEvents lblDone As Label
    Private WithEvents prgDone As ProgressBar
    Private WithEvents lblDoneCount As Label
    Private WithEvents grpJob As GroupBox
    Private WithEvents lblJobRef As Label
    Private WithEvents lblJobAddr As Label
    Private WithEvents lblJobLandmark As Label
    Private WithEvents lblJobDigital As Label
    Private WithEvents lblJobDirTitle As Label
    Private WithEvents lblJobDir As Label
    Private WithEvents lblJobPhone As Label
    Private WithEvents btnCopy As Button
    Private WithEvents btnCollected As Button
    Private WithEvents ssMain As StatusStrip
    Private WithEvents tslUser As ToolStripStatusLabel
    Private WithEvents tmrRefresh As System.Windows.Forms.Timer
    Private WithEvents ttHelp As ToolTip

    Public Sub New()
        MyBase.New()
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Me.SuspendLayout()
        Me.Text = "Collector"
        Me.ClientSize = New Size(1000, 640)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.AutoScaleMode = AutoScaleMode.None
        pnlHeader = NewPanel(Me, 0, 0, 1000, 70)
        picLogo = NewPicture(pnlHeader, 14, 10, 50, 50)
        lblHello = NewLabel(pnlHeader, "Hello", 76, 16, 0, 0, size:=18, bold:=True)
        btnLogout = NewButton(pnlHeader, "Log &out", 870, 17, 110, 36, "danger")
        lblBanner = NewLabel(Me, "", 20, 80, 960, 40, size:=11, bold:=True, back:="FFF3C4", align:=ContentAlignment.MiddleLeft)
        lblBanner.Visible = False
        lblOnline = NewLabel(Me, "", 20, 128, 300, 30, bold:=True, align:=ContentAlignment.MiddleLeft)
        lblMyArea = NewLabel(Me, "My area", 334, 134, 0, 0, bold:=True)
        cboMyArea = NewCombo(Me, 400, 129, 160)
        btnOnline = NewButton(Me, "Go &online", 576, 128, 150, 30, "secondary")
        grpNear = NewGroup(Me, "Requests near you (nearest first)", 20, 170, 590, 414)
        lvNear = NewList(grpNear, 14, 24, 562, 170, "Reference:100|Area:110|Distance:90|Waste:140|When:90")
        btnAccept = NewButton(grpNear, "&Accept", 14, 206, 150, 50)
        btnDecline = NewButton(grpNear, "&Decline", 176, 206, 150, 50, "secondary")
        lblRule = NewLabel(grpNear, "", 14, 268, 560, 50, fore:="555555")
        lblDone = NewLabel(grpNear, "My pickups done", 14, 332, 0, 0, bold:=True)
        prgDone = NewProgress(grpNear, 14, 358, 300, 22)
        lblDoneCount = NewLabel(grpNear, "", 330, 358, 0, 0, bold:=True)
        grpJob = NewGroup(Me, "My current job", 630, 170, 350, 414)
        lblJobRef = NewLabel(grpJob, "", 16, 26, 318, 26, size:=12, bold:=True)
        lblJobAddr = NewLabel(grpJob, "", 16, 60, 318, 22)
        lblJobLandmark = NewLabel(grpJob, "", 16, 84, 318, 22)
        lblJobDigital = NewLabel(grpJob, "", 16, 108, 318, 22)
        lblJobDirTitle = NewLabel(grpJob, "Directions", 16, 142, 0, 0, bold:=True)
        lblJobDir = NewLabel(grpJob, "", 16, 164, 318, 48)
        lblJobPhone = NewLabel(grpJob, "", 16, 218, 318, 24, bold:=True)
        btnCopy = NewButton(grpJob, "Copy &number", 16, 248, 140, 34, "secondary")
        btnCollected = NewButton(grpJob, "&Mark as collected", 16, 330, 318, 60)
        ssMain = New StatusStrip()
        tslUser = New ToolStripStatusLabel("")
        ssMain.Items.Add(tslUser)
        Me.Controls.Add(ssMain)
        tmrRefresh = New System.Windows.Forms.Timer()
        ttHelp = New ToolTip()
        AddHandler Me.FormClosed, Sub(s As Object, a As FormClosedEventArgs)
                                      tmrRefresh.Stop()
                                      tmrRefresh.Dispose()
                                  End Sub
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    Private currentJob As Booking = Nothing          ' the job being shown on the right
    Private sorryTime As DateTime = DateTime.MinValue ' when the "already accepted" message was shown

    ' ---------- form opens ----------
    Private Sub frmCollector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplyTheme(Me)
        BuildNav(Me, "collector")                  ' navigation bar at the top
        Me.Text = "Collector: " & CurrentUser.Name
        lblHello.Text = "Hello, " & CurrentUser.Name
        tslUser.Text = "Logged in as " & CurrentUser.Email

        cboMyArea.Items.AddRange(AreaCoords.Keys.ToArray())
        cboMyArea.SelectedItem = CurrentUser.Area
        If cboMyArea.SelectedIndex = -1 Then cboMyArea.SelectedIndex = 0

        ttHelp.SetToolTip(btnOnline, "Online = you share your location and receive requests near you")
        ttHelp.SetToolTip(cboMyArea, "Used when GPS is off")

        tmrRefresh.Interval = 3000                   ' look for new requests every 3 seconds
        tmrRefresh.Enabled = True
        UpdateOnlineUI()
        RefreshAll()
    End Sub

    Private Sub pnlHeader_Paint(sender As Object, e As PaintEventArgs) Handles pnlHeader.Paint
        PaintHeader(sender, e)
    End Sub

    ' ---------- online / offline ----------
    Private Sub btnOnline_Click(sender As Object, e As EventArgs) Handles btnOnline.Click
        Dim c As User = CurrentUser
        If c.Online Then
            c.Online = False
        Else
            c.Area = cboMyArea.Text
            GetCoords(c, True, c.Lat, c.Lon)         ' GPS if turned on, otherwise the area
            c.Online = True
        End If
        UpdateOnlineUI()
        RefreshAll()
    End Sub

    ' changing the area while Online moves the collector
    Private Sub cboMyArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMyArea.SelectedIndexChanged
        If CurrentUser Is Nothing OrElse Not CurrentUser.Online Then Exit Sub
        CurrentUser.Area = cboMyArea.Text
        GetCoords(CurrentUser, True, CurrentUser.Lat, CurrentUser.Lon)
        UpdateOnlineUI()
        RefreshAll()
    End Sub

    Private Sub UpdateOnlineUI()
        If CurrentUser.Online Then
            lblOnline.Text = ChrW(9679) & " You are Online  |  " & CurrentUser.Area
            lblOnline.BackColor = StatusColor("Online")
            btnOnline.Text = "Go &offline"
        Else
            lblOnline.Text = "You are Offline. You will not get requests."
            lblOnline.BackColor = StatusColor("Pending")
            btnOnline.Text = "Go &online"
        End If
    End Sub

    ' ---------- requests near me ----------
    Private Sub RefreshAll()
        RefreshNear()
        RefreshJob()
    End Sub

    Private Sub RefreshNear()
        Dim keep As String = If(lvNear.SelectedItems.Count > 0, lvNear.SelectedItems(0).Text, "")
        Dim list As List(Of Booking) = NearbyRequests(CurrentUser)

        lvNear.BeginUpdate()
        lvNear.Items.Clear()
        For Each b As Booking In list
            Dim item As New ListViewItem(b.Ref)
            item.SubItems.Add(b.Area)
            item.SubItems.Add(BookingDistance(b, CurrentUser).ToString("0.0") & " km")
            item.SubItems.Add(b.Waste)
            item.SubItems.Add(If(b.PickupDate = Date.Today, "Today", b.PickupDate.ToString("d MMM")))
            lvNear.Items.Add(item)
            If b.Ref = keep Then item.Selected = True
        Next
        If list.Count > 0 AndAlso lvNear.SelectedItems.Count = 0 Then lvNear.Items(0).Selected = True
        lvNear.EndUpdate()

        btnAccept.Enabled = (list.Count > 0)
        btnDecline.Enabled = (list.Count > 0)
        If list.Count > 0 Then
            Dim first As Booking = list(0)
            lblBanner.Text = "  New request " & BookingDistance(first, CurrentUser).ToString("0.0") & " km away: " &
                             first.Ref & ", " & first.Area & ", " & first.Waste.ToLower() & ". Press Accept to take it."
            lblBanner.Visible = True
        ElseIf (DateTime.Now - sorryTime).TotalSeconds > 6 Then
            lblBanner.Visible = False                ' the "Sorry" message stays for 6 seconds
        End If
        lblRule.Text = "The first collector to accept gets the job. Then the other collectors no longer see it. " &
                       "Requests are shown only when they are within " & BaseRadius & " km (" & WideRadius & " km if nobody accepts)."
    End Sub

    Private Function SelectedRequest() As Booking
        If lvNear.SelectedItems.Count = 0 Then Return Nothing
        Dim wanted As String = lvNear.SelectedItems(0).Text
        Return Bookings.FirstOrDefault(Function(x) x.Ref = wanted)
    End Function

    Private Sub btnAccept_Click(sender As Object, e As EventArgs) Handles btnAccept.Click
        Dim b As Booking = SelectedRequest()
        If b Is Nothing Then Exit Sub
        If AcceptJob(b, CurrentUser) Then
            lblBanner.Visible = False
        Else
            lblBanner.Text = "  Sorry, another collector already accepted that request."
            lblBanner.Visible = True
            sorryTime = DateTime.Now
            ShowError("Sorry, another collector already accepted that request.")
        End If
        RefreshAll()
    End Sub

    Private Sub btnDecline_Click(sender As Object, e As EventArgs) Handles btnDecline.Click
        Dim b As Booking = SelectedRequest()
        If b Is Nothing Then Exit Sub
        b.Declined.Add(CurrentUser)                  ' it will not show for this collector again
        RefreshAll()
    End Sub

    ' ---------- my current job ----------
    Private Sub RefreshJob()
        Dim mine = Bookings.Where(Function(x) x.Collector Is CurrentUser)
        Dim openJobs = mine.Where(Function(x) x.Status = "Assigned").OrderBy(Function(x) x.CreatedAt)
        currentJob = openJobs.FirstOrDefault()

        If currentJob Is Nothing Then
            lblJobRef.Text = "No current job"
            lblJobAddr.Text = "Accept a request to see the details here."
            lblJobLandmark.Text = ""
            lblJobDigital.Text = ""
            lblJobDir.Text = ""
            lblJobPhone.Text = ""
        Else
            lblJobRef.Text = currentJob.Ref & "   " & currentJob.Owner.Name
            lblJobAddr.Text = currentJob.HouseNo & " " & currentJob.Street & ", " & currentJob.Area
            lblJobLandmark.Text = currentJob.Landmark
            lblJobDigital.Text = If(currentJob.DigitalAddress = "", "", "Digital address: " & currentJob.DigitalAddress)
            lblJobDir.Text = currentJob.Directions & If(currentJob.Note = "", "", "  Note: " & currentJob.Note)
            lblJobPhone.Text = "Phone   " & currentJob.Owner.Phone
        End If
        btnCopy.Enabled = (currentJob IsNot Nothing)
        btnCollected.Enabled = (currentJob IsNot Nothing)

        ' progress: collected jobs out of all my jobs
        Dim total As Integer = mine.Count()
        Dim done As Integer = mine.Where(Function(x) x.Status = "Collected").Count()
        prgDone.Maximum = Math.Max(total, 1)
        prgDone.Value = Math.Min(done, prgDone.Maximum)
        lblDoneCount.Text = done & " of " & total
    End Sub

    Private Sub btnCollected_Click(sender As Object, e As EventArgs) Handles btnCollected.Click
        If currentJob Is Nothing Then Exit Sub
        If Not Confirm("Mark this pickup as collected?") Then Exit Sub
        currentJob.Status = "Collected"
        RefreshAll()
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        If currentJob IsNot Nothing Then Clipboard.SetText(currentJob.Owner.Phone)
    End Sub

    ' every 3 seconds: new requests appear by themselves (no Refresh button)
    Private Sub tmrRefresh_Tick(sender As Object, e As EventArgs) Handles tmrRefresh.Tick
        RefreshAll()
    End Sub

    ' ---------- log out ----------
    ' Logging out does NOT take you offline. Press "Go offline" first when your work is finished.
    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        If Not Confirm("Do you want to log out?") Then Exit Sub
        CurrentUser = Nothing
        Me.Close()
    End Sub

End Class
