' ============================================================
' frmAdmin.vb : the admin sees requests, collectors and residents.
' The admin can only assign a collector who is inside the search radius.
' ============================================================
Public Class frmAdmin
    Inherits Form

    ' ---------- CONTROLS (built in code, so you do not design anything) ----------
    Private WithEvents pnlHeader As Panel
    Private WithEvents picLogo As PictureBox
    Private WithEvents lblTitle As Label
    Private WithEvents btnLogout As Button
    Private WithEvents pnlWaiting As Panel
    Private WithEvents lblWaitingCap As Label
    Private WithEvents lblWaitingNum As Label
    Private WithEvents pnlAccepted As Panel
    Private WithEvents lblAcceptedCap As Label
    Private WithEvents lblAcceptedNum As Label
    Private WithEvents pnlCollected As Panel
    Private WithEvents lblCollectedCap As Label
    Private WithEvents lblCollectedNum As Label
    Private WithEvents pnlSubs As Panel
    Private WithEvents lblSubsCap As Label
    Private WithEvents lblSubsNum As Label
    Private WithEvents tcAdmin As TabControl
    Private WithEvents tpRequests As TabPage
    Private WithEvents tpCollectors As TabPage
    Private WithEvents tpResidents As TabPage
    Private WithEvents lvRequests As ListView
    Private WithEvents pnlAssign As Panel
    Private WithEvents lblSelected As Label
    Private WithEvents lblWarn As Label
    Private WithEvents btnWiden As Button
    Private WithEvents lblNear As Label
    Private WithEvents lvNearCollectors As ListView
    Private WithEvents btnAssign As Button
    Private WithEvents lblAssignRule As Label
    Private WithEvents lvCollectors As ListView
    Private WithEvents lblCollJobs As Label
    Private WithEvents lvCollJobs As ListView
    Private WithEvents rdoAll As RadioButton
    Private WithEvents rdoSubscribed As RadioButton
    Private WithEvents rdoNotSubscribed As RadioButton
    Private WithEvents lvResidents As ListView
    Private WithEvents lblSubSummary As Label
    Private WithEvents lblPrices As Label
    Private WithEvents ssMain As StatusStrip
    Private WithEvents tslUser As ToolStripStatusLabel
    Private WithEvents tmrAdmin As System.Windows.Forms.Timer

    Public Sub New()
        MyBase.New()
        BuildUi()
    End Sub

    Private Sub BuildUi()
        Me.SuspendLayout()
        Me.Text = "Admin"
        Me.ClientSize = New Size(1150, 660)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.AutoScaleMode = AutoScaleMode.None
        pnlHeader = NewPanel(Me, 0, 0, 1150, 70)
        picLogo = NewPicture(pnlHeader, 14, 10, 50, 50)
        lblTitle = NewLabel(pnlHeader, "Admin dashboard", 76, 16, 0, 0, size:=18, bold:=True, fore:="FFFFFF")
        btnLogout = NewButton(pnlHeader, "Log &out", 1022, 17, 110, 36, "danger")
        pnlWaiting = NewPanel(Me, 20, 84, 265, 80, back:="FFE3D8")
        lblWaitingCap = NewLabel(pnlWaiting, "Waiting for a collector", 14, 10, 0, 0)
        lblWaitingNum = NewLabel(pnlWaiting, "0", 14, 30, 0, 0, size:=22, bold:=True)
        pnlAccepted = NewPanel(Me, 300, 84, 265, 80, back:="FFF3C4")
        lblAcceptedCap = NewLabel(pnlAccepted, "Accepted", 14, 10, 0, 0)
        lblAcceptedNum = NewLabel(pnlAccepted, "0", 14, 30, 0, 0, size:=22, bold:=True)
        pnlCollected = NewPanel(Me, 580, 84, 265, 80, back:="CFEECF")
        lblCollectedCap = NewLabel(pnlCollected, "Collected", 14, 10, 0, 0)
        lblCollectedNum = NewLabel(pnlCollected, "0", 14, 30, 0, 0, size:=22, bold:=True)
        pnlSubs = NewPanel(Me, 860, 84, 270, 80, back:="FFFFFF")
        lblSubsCap = NewLabel(pnlSubs, "Active subscriptions", 14, 10, 0, 0)
        lblSubsNum = NewLabel(pnlSubs, "0", 14, 30, 0, 0, size:=22, bold:=True)
        tcAdmin = NewTabs(Me, 20, 180, 1110, 430, 150, 34)
        tpRequests = NewPage(tcAdmin, "Requests (0)")
        tpCollectors = NewPage(tcAdmin, "Collectors (0)")
        tpResidents = NewPage(tcAdmin, "Residents (0)")
        lvRequests = NewList(tpRequests, 10, 10, 1086, 140, "Reference:110|Resident:150|Area:130|Date:90|Status:110|Collector:160|Note:300")
        pnlAssign = NewPanel(tpRequests, 10, 158, 1086, 206, back:="E8F5E9")
        lblSelected = NewLabel(pnlAssign, "Select a request above.", 14, 10, 380, 24, bold:=True)
        lblWarn = NewLabel(pnlAssign, "", 400, 6, 270, 30, bold:=True, fore:="A8400F", back:="FFE3D8", align:=ContentAlignment.MiddleCenter)
        lblWarn.Visible = False
        btnWiden = NewButton(pnlAssign, "&Widen search to 5 km", 686, 6, 210, 30)
        lblNear = NewLabel(pnlAssign, "", 14, 46, 600, 22, bold:=True)
        lvNearCollectors = NewList(pnlAssign, 14, 70, 1056, 84, "Collector:200|Area:150|Distance:120|Status:110|Can take this job?:300")
        btnAssign = NewButton(pnlAssign, "A&ssign", 14, 164, 150, 32)
        lblAssignRule = NewLabel(pnlAssign, "Assign works only for a collector inside the search radius.", 180, 170, 880, 22, fore:="A8400F")
        lvCollectors = NewList(tpCollectors, 10, 10, 1086, 150, "Name:170|Phone:130|Email:230|Area:150|Assigned:90|Collected:90|Status:100")
        lblCollJobs = NewLabel(tpCollectors, "", 10, 170, 600, 22, bold:=True)
        lvCollJobs = NewList(tpCollectors, 10, 196, 1086, 130, "Reference:110|Resident:160|Address:320|Date:90|Status:110")
        rdoAll = NewRadio(tpResidents, "&All", 10, 12, 0, 0, isChecked:=True)
        rdoSubscribed = NewRadio(tpResidents, "S&ubscribed", 90, 12, 0, 0)
        rdoNotSubscribed = NewRadio(tpResidents, "&Not subscribed", 230, 12, 0, 0)
        lvResidents = NewList(tpResidents, 10, 44, 1086, 200, "Name:160|Phone:130|Area:110|Plan:100|Subscription:140|Pickups left:110|Renews:130")
        lblSubSummary = NewLabel(tpResidents, "", 10, 254, 900, 22)
        lblPrices = NewLabel(tpResidents, "", 10, 280, 900, 22, fore:="555555")
        ssMain = New StatusStrip()
        tslUser = New ToolStripStatusLabel("")
        ssMain.Items.Add(tslUser)
        Me.Controls.Add(ssMain)
        tmrAdmin = New System.Windows.Forms.Timer()
        AddHandler Me.FormClosed, Sub(s As Object, a As FormClosedEventArgs)
                                      tmrAdmin.Stop()
                                      tmrAdmin.Dispose()
                                  End Sub
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub


    ' ---------- form opens ----------
    Private Sub frmAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplyTheme(Me)
        BuildNav(Me, "admin")                      ' navigation bar at the top
        pnlHeader.Tag = "dark"                       ' darker green header for the admin
        Me.Text = "Admin"
        tslUser.Text = "Logged in as " & CurrentUser.Email & " (admin)"
        rdoAll.Checked = True
        tmrAdmin.Interval = 3000
        tmrAdmin.Enabled = True
        RefreshAll()
    End Sub

    Private Sub pnlHeader_Paint(sender As Object, e As PaintEventArgs) Handles pnlHeader.Paint
        PaintHeader(sender, e)
    End Sub

    Private Sub RefreshAll()
        RefreshCounts()
        RefreshRequests()
        RefreshCollectors()
        RefreshResidents()
    End Sub

    Private Sub tmrAdmin_Tick(sender As Object, e As EventArgs) Handles tmrAdmin.Tick
        RefreshAll()
    End Sub

    ' ---------- the four number boxes at the top ----------
    Private Sub RefreshCounts()
        lblWaitingNum.Text = Bookings.Where(Function(b) b.Status = "Pending").Count().ToString()
        lblAcceptedNum.Text = Bookings.Where(Function(b) b.Status = "Assigned").Count().ToString()
        lblCollectedNum.Text = Bookings.Where(Function(b) b.Status = "Collected").Count().ToString()
        Dim residents = Users.Where(Function(u) u.Role = "resident").ToList()
        Dim subscribed As Integer = residents.Where(Function(u) SubStatus(u) = "Active" OrElse SubStatus(u) = "Expires soon").Count()
        lblSubsNum.Text = subscribed & " of " & residents.Count
        tcAdmin.TabPages(0).Text = "Requests (" & Bookings.Count & ")"
        tcAdmin.TabPages(1).Text = "Collectors (" & Users.Where(Function(u) u.Role = "collector").Count() & ")"
        tcAdmin.TabPages(2).Text = "Residents (" & residents.Count & ")"
    End Sub

    ' ===================== TAB 1: REQUESTS =====================
    Private Sub RefreshRequests()
        Dim keep As String = If(lvRequests.SelectedItems.Count > 0, lvRequests.SelectedItems(0).Text, "")
        lvRequests.BeginUpdate()
        lvRequests.Items.Clear()
        For Each b As Booking In Bookings.OrderByDescending(Function(x) x.CreatedAt)
            Dim item As New ListViewItem(b.Ref)
            item.SubItems.Add(b.Owner.Name)
            item.SubItems.Add(b.Area)
            item.SubItems.Add(If(b.PickupDate = Date.Today, "Today", b.PickupDate.ToString("d MMM")))
            item.SubItems.Add(b.Status)
            item.SubItems.Add(If(b.Collector Is Nothing, "-", b.Collector.Name))
            item.SubItems.Add(NoteFor(b))
            item.BackColor = StatusColor(b.Status)
            lvRequests.Items.Add(item)
            If b.Ref = keep Then item.Selected = True
        Next
        lvRequests.EndUpdate()
        ShowAssign()
    End Sub

    Private Function NoteFor(b As Booking) As String
        Select Case b.Status
            Case "Collected"
                Return "Done"
            Case "Assigned"
                Return "Accepted by " & b.Collector.Name
            Case Else
                Dim radius As Double = CurrentRadius(b)
                Dim n As Integer = CollectorsNear(b).Where(Function(c) BookingDistance(b, c) <= radius).Count()
                If n = 0 Then Return "No collector within " & radius & " km"
                Return "Offered to " & n & " collector(s) near " & b.Area
        End Select
    End Function

    Private Function SelectedRequest() As Booking
        If lvRequests.SelectedItems.Count = 0 Then Return Nothing
        Dim wanted As String = lvRequests.SelectedItems(0).Text
        Return Bookings.FirstOrDefault(Function(x) x.Ref = wanted)
    End Function

    Private Sub lvRequests_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvRequests.SelectedIndexChanged
        ShowAssign()
    End Sub

    ' ---------- the assign panel under the list ----------
    Private Sub ShowAssign()
        Dim keep As String = If(lvNearCollectors.SelectedItems.Count > 0, lvNearCollectors.SelectedItems(0).Text, "")
        lvNearCollectors.Items.Clear()
        lblWarn.Visible = False
        btnWiden.Enabled = False
        btnAssign.Enabled = False

        Dim b As Booking = SelectedRequest()
        If b Is Nothing Then
            lblSelected.Text = "Select a request above."
            lblNear.Text = ""
            Exit Sub
        End If

        lblSelected.Text = "Selected: " & b.Ref & " (" & b.Owner.Name & ", " & b.Area & ")"
        If b.Status <> "Pending" Then
            lblNear.Text = "This request is already " & b.Status.ToLower() & "."
            Exit Sub
        End If

        Dim radius As Double = CurrentRadius(b)
        lblNear.Text = "Collectors online, nearest to " & b.Area & " first"
        Dim anyInside As Boolean = False
        For Each c As User In CollectorsNear(b)
            Dim d As Double = BookingDistance(b, c)
            Dim item As New ListViewItem(c.Name)
            item.Tag = c
            item.SubItems.Add(c.Area)
            item.SubItems.Add(d.ToString("0.0") & " km")
            item.SubItems.Add("Online")
            If d <= radius Then
                item.SubItems.Add("Yes (within " & radius & " km)")
                anyInside = True
            ElseIf d <= WideRadius Then
                item.SubItems.Add("Outside " & radius & " km (inside " & WideRadius & " km)")
            Else
                item.SubItems.Add("Too far")
            End If
            item.BackColor = If(d <= radius, StatusColor("Online"), If(d <= WideRadius, StatusColor("Assigned"), StatusColor("Pending")))
            lvNearCollectors.Items.Add(item)
            If c.Name = keep Then item.Selected = True
        Next

        lblWarn.Text = ChrW(9888) & " No collector within " & radius & " km"
        lblWarn.Visible = Not anyInside
        btnWiden.Enabled = (radius < WideRadius)     ' can only widen once, to 5 km
        UpdateAssignButton()
    End Sub

    Private Sub btnWiden_Click(sender As Object, e As EventArgs) Handles btnWiden.Click
        Dim b As Booking = SelectedRequest()
        If b Is Nothing Then Exit Sub
        b.RadiusKm = WideRadius
        RefreshAll()
    End Sub

    Private Sub lvNearCollectors_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvNearCollectors.SelectedIndexChanged
        UpdateAssignButton()
    End Sub

    ' Assign only works for a collector who is inside the search radius
    Private Sub UpdateAssignButton()
        btnAssign.Enabled = False
        Dim b As Booking = SelectedRequest()
        If b Is Nothing OrElse lvNearCollectors.SelectedItems.Count = 0 Then Exit Sub
        Dim c As User = DirectCast(lvNearCollectors.SelectedItems(0).Tag, User)
        btnAssign.Enabled = (b.Status = "Pending" AndAlso BookingDistance(b, c) <= CurrentRadius(b))
    End Sub

    Private Sub btnAssign_Click(sender As Object, e As EventArgs) Handles btnAssign.Click
        Dim b As Booking = SelectedRequest()
        If b Is Nothing OrElse lvNearCollectors.SelectedItems.Count = 0 Then Exit Sub
        Dim c As User = DirectCast(lvNearCollectors.SelectedItems(0).Tag, User)
        If BookingDistance(b, c) > CurrentRadius(b) Then Exit Sub    ' the rule: no far assignments
        AcceptJob(b, c)                              ' same rule as a collector pressing Accept
        ShowInfo("Request " & b.Ref & " has been assigned to " & c.Name & ".")
        RefreshAll()
    End Sub

    ' ===================== TAB 2: COLLECTORS =====================
    Private Sub RefreshCollectors()
        Dim keep As String = If(lvCollectors.SelectedItems.Count > 0, lvCollectors.SelectedItems(0).Text, "")
        lvCollectors.BeginUpdate()
        lvCollectors.Items.Clear()
        For Each c As User In Users.Where(Function(u) u.Role = "collector")
            Dim item As New ListViewItem(c.Name)
            item.SubItems.Add(c.Phone)
            item.SubItems.Add(c.Email)
            item.SubItems.Add(c.Area)
            item.SubItems.Add(Bookings.Where(Function(b) b.Collector Is c AndAlso b.Status = "Assigned").Count().ToString())
            item.SubItems.Add(Bookings.Where(Function(b) b.Collector Is c AndAlso b.Status = "Collected").Count().ToString())
            item.SubItems.Add(If(c.Online, "Online", "Offline"))
            If c.Online Then item.BackColor = StatusColor("Online")
            lvCollectors.Items.Add(item)
            If c.Name = keep Then item.Selected = True
        Next
        lvCollectors.EndUpdate()
        ShowCollectorJobs()
    End Sub

    Private Sub lvCollectors_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvCollectors.SelectedIndexChanged
        ShowCollectorJobs()
    End Sub

    Private Sub ShowCollectorJobs()
        lvCollJobs.Items.Clear()
        If lvCollectors.SelectedItems.Count = 0 Then
            lblCollJobs.Text = "Select a collector to see their pickups."
            Exit Sub
        End If
        Dim name As String = lvCollectors.SelectedItems(0).Text
        lblCollJobs.Text = "Pickups of " & name
        For Each b As Booking In Bookings.Where(Function(x) x.Collector IsNot Nothing AndAlso x.Collector.Name = name)
            Dim item As New ListViewItem(b.Ref)
            item.SubItems.Add(b.Owner.Name)
            item.SubItems.Add(b.HouseNo & " " & b.Street & ", " & b.Area)
            item.SubItems.Add(If(b.PickupDate = Date.Today, "Today", b.PickupDate.ToString("d MMM")))
            item.SubItems.Add(b.Status)
            item.BackColor = StatusColor(b.Status)
            lvCollJobs.Items.Add(item)
        Next
    End Sub

    ' ===================== TAB 3: RESIDENTS =====================
    Private Sub FilterChanged(sender As Object, e As EventArgs) Handles rdoAll.CheckedChanged, rdoSubscribed.CheckedChanged, rdoNotSubscribed.CheckedChanged
        RefreshResidents()
    End Sub

    Private Sub RefreshResidents()
        Dim keep As String = If(lvResidents.SelectedItems.Count > 0, lvResidents.SelectedItems(0).Text, "")
        lvResidents.BeginUpdate()
        lvResidents.Items.Clear()
        Dim everyone = Users.Where(Function(u) u.Role = "resident").ToList()
        Dim shown As Integer = 0
        For Each r As User In everyone
            Dim st As String = SubStatus(r)
            Dim isSub As Boolean = (r.Plan <> "")
            If rdoSubscribed.Checked AndAlso Not isSub Then Continue For
            If rdoNotSubscribed.Checked AndAlso isSub Then Continue For
            Dim item As New ListViewItem(r.Name)
            item.SubItems.Add(r.Phone)
            item.SubItems.Add(r.Area)
            item.SubItems.Add(If(isSub, r.Plan, "-"))
            item.SubItems.Add(st)
            item.SubItems.Add(If(isSub, r.PickupsLeft & " of " & PlanPickups(r.Plan), "-"))
            item.SubItems.Add(If(isSub, r.RenewsOn.ToString("d MMM yyyy"), "-"))
            item.BackColor = StatusColor(st)
            lvResidents.Items.Add(item)
            If r.Name = keep Then item.Selected = True
            shown += 1
        Next
        lvResidents.EndUpdate()
        Dim subs As Integer = everyone.Where(Function(u) u.Plan <> "").Count()
        lblSubSummary.Text = "Subscribed residents: " & subs & " of " & everyone.Count & "   |   Showing " & shown
        lblPrices.Text = "Prices are illustrative: Basic " & Cedi & PlanPrice("Basic") & ", Standard " & Cedi & PlanPrice("Standard") &
                         ", Business " & Cedi & PlanPrice("Business") & " per month (estimates)."
    End Sub

    ' ---------- log out ----------
    Private Sub btnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        If Not Confirm("Do you want to log out?") Then Exit Sub
        CurrentUser = Nothing
        Me.Close()
    End Sub

End Class
