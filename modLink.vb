' ============================================================
' modLink.vb : data, colours and helper code shared by ALL forms
' ============================================================
Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions

' ---------- DATA CLASSES (outside the module, so every form can use User and Booking) ----------
Public Class User
    Public Name As String
    Public Email As String
    Public Phone As String
    Public Password As String
    Public Role As String                 ' "resident", "collector" or "admin"
    Public Provider As String = ""        ' "Google" or "Facebook" for accounts made by social sign-in, otherwise empty
    Public Area As String = "Osu"         ' the area of the address (residents) or where the collector is
    ' saved pickup address (residents)
    Public HouseNo As String = ""
    Public Street As String = ""
    Public Landmark As String = ""
    Public DigitalAddress As String = ""
    Public Directions As String = ""
    ' subscription (residents). Plan = "" means no subscription
    Public Plan As String = ""
    Public PickupsLeft As Integer = 0
    Public RenewsOn As Date = Date.MinValue
    ' location (collectors)
    Public Online As Boolean = False
    Public Lat As Double
    Public Lon As Double

    Public Function HasAddress() As Boolean
        Return HouseNo <> "" AndAlso Street <> ""
    End Function
End Class

Public Class Booking
    Public Ref As String
    Public Owner As User                  ' the resident who booked
    Public Collector As User              ' Nothing until a collector accepts
    Public HouseNo As String
    Public Street As String
    Public Landmark As String
    Public Area As String
    Public DigitalAddress As String
    Public Directions As String
    Public Waste As String
    Public Note As String
    Public Pay As String                  ' "Subscription (Standard)" or the mobile money used
    Public PickupDate As Date
    Public CreatedAt As DateTime = DateTime.Now
    Public Lat As Double                  ' where the pickup is
    Public Lon As Double
    Public RadiusKm As Double = BaseRadius
    Public Status As String = "Pending"   ' Pending, Assigned or Collected
    Public Declined As New List(Of User)  ' collectors who pressed Decline
End Class


Public Module modLink

    ' ---------- 1. SETTINGS (change these if you like) ----------
    Public Const AdminEmail As String = "admin@tidyaccra.com"   ' the special admin email (nobody can sign up with it)
    Public Const AdminPassword As String = "admin123"           ' change before you present
    Public Const NavHeight As Integer = 46                ' height of the navigation bar at the top of every page
    Public Const BaseRadius As Double = 3                 ' km: first search around the resident
    Public Const WideRadius As Double = 5                 ' km: wider search
    Public Const WaitMinutes As Double = 2                ' minutes before the search widens by itself
    Public Const PerPickupPrice As Integer = 15           ' GH cedis, an estimate
    Public UseGps As Boolean = False                      ' False = use the area of the address (best for a class demo)
    '                                                       True  = try the real GPS / Wi-Fi location first
    Public ReadOnly Cedi As String = "GH" & ChrW(162)     ' GH followed by the cedi sign

    ' ---------- 2. COLOURS ----------
    Public ReadOnly Green As Color = ColorTranslator.FromHtml("#1E7A3C")       ' main buttons
    Public ReadOnly DarkGreen As Color = ColorTranslator.FromHtml("#1B4332")   ' text, admin header
    Public ReadOnly Terracotta As Color = ColorTranslator.FromHtml("#A8400F")  ' Exit, Log out, warnings
    Public ReadOnly FormBack As Color = ColorTranslator.FromHtml("#F0F4F0")

    Public Users As New List(Of User)
    Public Bookings As New List(Of Booking)
    Public CurrentUser As User                ' who is logged in right now

    ' ---------- 4. LISTS FOR THE DROP-DOWNS ----------
    Public ReadOnly WasteTypes As String() = {"General waste", "Recyclables (plastic, paper, metal)", "Bulky waste"}
    Public ReadOnly PayMethods As String() = {"MTN Mobile Money", "AirtelTigo Money", "Telecel Cash"}
    Public ReadOnly PlanPrice As New Dictionary(Of String, Integer) From {{"Basic", 40}, {"Standard", 70}, {"Business", 150}}
    Public ReadOnly PlanPickups As New Dictionary(Of String, Integer) From {{"Basic", 4}, {"Standard", 8}, {"Business", 12}}

    ' Approximate centre of each area: latitude, longitude (for the project only)
    Public ReadOnly AreaCoords As New Dictionary(Of String, Double()) From {
        {"Osu", New Double() {5.556, -0.182}},
        {"Labone", New Double() {5.562, -0.175}},
        {"Cantonments", New Double() {5.575, -0.176}},
        {"Madina", New Double() {5.668, -0.168}},
        {"Adenta", New Double() {5.706, -0.167}},
        {"Nima", New Double() {5.583, -0.204}},
        {"Kaneshie", New Double() {5.565, -0.238}},
        {"Teshie", New Double() {5.58, -0.105}},
        {"Dansoman", New Double() {5.538, -0.267}},
        {"Achimota", New Double() {5.614, -0.23}},
        {"Lapaz", New Double() {5.609, -0.248}}}

    ' ---------- 5. DISTANCE AND LOCATION ----------
    ' Straight-line distance in km between two points (the haversine formula)
    Public Function DistanceKm(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
        Const R As Double = 6371                      ' radius of the Earth in km
        Dim dLat As Double = (lat2 - lat1) * Math.PI / 180
        Dim dLon As Double = (lon2 - lon1) * Math.PI / 180
        Dim a As Double = Math.Sin(dLat / 2) ^ 2 +
                          Math.Cos(lat1 * Math.PI / 180) * Math.Cos(lat2 * Math.PI / 180) * Math.Sin(dLon / 2) ^ 2
        Return R * 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a))
    End Function

    ' Asks Windows for the real location. Returns False if it is off or unavailable.
    Public Function TryGetGps(ByRef lat As Double, ByRef lon As Double) As Boolean
        Return False    ' GPS is switched off: the app uses the area of the address
    End Function

    ' Fills lat and lon for this person and says where they came from.
    Public Function GetCoords(u As User, useMyLocation As Boolean, ByRef lat As Double, ByRef lon As Double) As String
        If useMyLocation AndAlso UseGps AndAlso TryGetGps(lat, lon) Then Return "your GPS location"
        Dim xy As Double() = AreaCoords(u.Area)
        lat = xy(0)
        lon = xy(1)
        Return u.Area & " (area)"
    End Function

    ' ---------- 6. MATCHING A REQUEST WITH COLLECTORS ----------
    Public Function BookingDistance(b As Booking, c As User) As Double
        Return DistanceKm(c.Lat, c.Lon, b.Lat, b.Lon)
    End Function

    ' How far we search for this request right now (3 km, then 5 km)
    Public Function CurrentRadius(b As Booking) As Double
        Dim r As Double = Math.Max(b.RadiusKm, BaseRadius)
        If (DateTime.Now - b.CreatedAt).TotalMinutes >= WaitMinutes Then r = Math.Max(r, WideRadius)
        Return r
    End Function

    ' Requests a collector can see: Online, Pending, not declined, inside the radius. Nearest first.
    Public Function NearbyRequests(c As User) As List(Of Booking)
        If Not c.Online Then Return New List(Of Booking)()
        Dim found = Bookings.Where(Function(b) b.Status = "Pending" AndAlso Not b.Declined.Contains(c) AndAlso BookingDistance(b, c) <= CurrentRadius(b))
        Return found.OrderBy(Function(b) BookingDistance(b, c)).ToList()
    End Function

    ' All Online collectors, nearest to the request first
    Public Function CollectorsNear(b As Booking) As List(Of User)
        Dim online = Users.Where(Function(u) u.Role = "collector" AndAlso u.Online)
        Return online.OrderBy(Function(u) BookingDistance(b, u)).ToList()
    End Function

    ' The first collector to accept gets the job
    Public Function AcceptJob(b As Booking, c As User) As Boolean
        If b.Status <> "Pending" Then Return False
        b.Collector = c
        b.Status = "Assigned"
        Return True
    End Function

    ' ---------- 7. SUBSCRIPTION ----------
    Public Function SubStatus(u As User) As String
        If u.Plan = "" Then Return "Not subscribed"
        If u.RenewsOn < Date.Today Then Return "Expired"
        If (u.RenewsOn - Date.Today).TotalDays <= 3 Then Return "Expires soon"
        Return "Active"
    End Function

    ' True when the next pickup is covered by the plan (so no payment box)
    Public Function UsesSubscription(u As User) As Boolean
        Dim s As String = SubStatus(u)
        Return (s = "Active" OrElse s = "Expires soon") AndAlso u.PickupsLeft > 0
    End Function

    ' ---------- 8. SMALL HELPERS ----------
    Private ReadOnly Dice As New Random()
    Public Function NewReference() As String
        Return "TA-" & Dice.Next(100000, 999999).ToString()
    End Function

    Public Function StatusColor(status As String) As Color
        Select Case status
            Case "Pending", "Not subscribed", "Expired" : Return ColorTranslator.FromHtml("#FFE3D8")
            Case "Assigned", "Expires soon" : Return ColorTranslator.FromHtml("#FFF3C4")
            Case "Collected", "Active", "Online" : Return ColorTranslator.FromHtml("#CFEECF")
            Case Else : Return Color.White
        End Select
    End Function

    ' ---------- MESSAGE BOXES ----------
    ' Red X: something went wrong
    Public Sub ShowError(msg As String)
        MessageBox.Show(msg, "TidyAccra Connect", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ' Blue "i": something worked
    Public Sub ShowInfo(msg As String)
        MessageBox.Show(msg, "TidyAccra Connect", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' Yes/No question: returns True when the user clicks Yes
    Public Function Confirm(msg As String) As Boolean
        Return MessageBox.Show(msg, "TidyAccra Connect", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function

    ' ---------- LOGO (drawn in code, so no logo.png file is needed) ----------
    Public Function MakeLogo() As Bitmap
        Dim bmp As New Bitmap(100, 100)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.Clear(Color.Transparent)
            Using br As New SolidBrush(ColorTranslator.FromHtml("#1E7A3C"))
                g.FillEllipse(br, 2, 2, 96, 96)               ' green circle
            End Using
            Using f As New Font("Segoe UI", 44, FontStyle.Bold)
                Dim sz As SizeF = g.MeasureString("T", f)
                g.DrawString("T", f, Brushes.White, (100 - sz.Width) / 2, (100 - sz.Height) / 2)   ' white T
            End Using
        End Using
        Return bmp
    End Function

    ' ---------- SOCIAL SIGN-IN WINDOW (Google or Facebook) ----------
    ' This window is built in code (no designer). It is a DEMO of the sign-in step:
    ' a real app would open the Google or Facebook page here and receive the email from them.
    ' Returns the email typed, or "" if the person pressed Cancel.
    Public Function AskSocialEmail(provider As String) As String
        Using dlg As New Form()
            dlg.Text = "Sign in with " & provider
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.StartPosition = FormStartPosition.CenterScreen
            dlg.MaximizeBox = False
            dlg.MinimizeBox = False
            dlg.ClientSize = New Size(380, 210)
            dlg.BackColor = Color.White
            dlg.Font = New Font("Segoe UI", 10)

            Dim lblHead As New Label With {.Text = "Choose your " & provider & " account", .Font = New Font("Segoe UI", 13, FontStyle.Bold),
                                           .ForeColor = DarkGreen, .AutoSize = True, .Location = New Point(20, 16)}
            Dim lblNote As New Label With {.Text = "Demo sign-in. A real app would open the " & provider & " page here.",
                                           .ForeColor = Color.DimGray, .AutoSize = False, .Size = New Size(340, 36), .Location = New Point(20, 48)}
            Dim lblMail As New Label With {.Text = provider & " email", .AutoSize = True, .Location = New Point(20, 92)}
            Dim txtMail As New TextBox With {.Location = New Point(20, 114), .Width = 340}
            Dim btnGo As New Button With {.Text = "Continue", .Location = New Point(20, 156), .Size = New Size(160, 38), .DialogResult = DialogResult.OK}
            Dim btnNo As New Button With {.Text = "Cancel", .Tag = "secondary", .Location = New Point(200, 156), .Size = New Size(160, 38), .DialogResult = DialogResult.Cancel}
            StyleButton(btnGo)
            StyleButton(btnNo)
            dlg.Controls.AddRange(New Control() {lblHead, lblNote, lblMail, txtMail, btnGo, btnNo})
            dlg.AcceptButton = btnGo
            dlg.CancelButton = btnNo

            If dlg.ShowDialog() = DialogResult.OK Then Return txtMail.Text.Trim()
            Return ""
        End Using
    End Function

    ' ---------- NAVIGATION BAR AND ACCESS CONTROL ----------
    Public WelcomeForm As Form                 ' the start form (hidden while someone is logged in)
    Public Navigating As Boolean = False       ' True for a moment while one page is swapped for another

    ' The rule that keeps people out: nobody gets a page without being logged in,
    ' and each role only gets its own pages.
    Public Function PageAllowed(page As String) As Boolean
        If CurrentUser Is Nothing Then Return False
        Select Case page
            Case "home" : Return True
            Case "resident", "subscribe" : Return CurrentUser.Role = "resident"
            Case "collector" : Return CurrentUser.Role = "collector"
            Case "admin" : Return CurrentUser.Role = "admin"
            Case Else : Return False
        End Select
    End Function

    Private Function CreatePage(page As String) As Form
        Select Case page
            Case "resident" : Return New frmResident()
            Case "subscribe" : Return New frmSubscribe()
            Case "collector" : Return New frmCollector()
            Case "admin" : Return New frmAdmin()
            Case Else : Return New frmHome()
        End Select
    End Function

    ' Opens another page in the same place and closes the current one
    Public Sub NavigateTo(current As Form, page As String)
        If CurrentUser Is Nothing Then
            ShowError("Please log in first. You need an account to use TidyAccra Connect.")
            current.Close()                        ' this sends the person back to the Welcome form
            Return
        End If
        If Not PageAllowed(page) Then
            ShowError("Your account cannot open that page.")
            Return
        End If
        Dim target As Form = CreatePage(page)
        target.StartPosition = FormStartPosition.Manual
        target.Location = current.Location
        Navigating = True                          ' so closing the old page does not log the person out
        Try
            target.Show()
            current.Close()
        Finally
            Navigating = False
        End Try
    End Sub

    ' Brings back the Welcome form (log in screen) and forgets who was logged in
    Public Sub ReturnToWelcome()
        CurrentUser = Nothing
        Dim w As frmWelcome = TryCast(WelcomeForm, frmWelcome)
        If w IsNot Nothing Then w.ShowAgain()
    End Sub

    ' The list of links in the bar. Each role sees only its own pages.
    Private Function NavItems() As List(Of String())
        Dim items As New List(Of String())
        items.Add({"Home", "home"})
        Select Case CurrentUser.Role
            Case "resident"
                items.Add({"Request a pickup", "resident"})
                items.Add({"Subscription", "subscribe"})
            Case "collector"
                items.Add({"Collector", "collector"})
            Case "admin"
                items.Add({"Admin", "admin"})
        End Select
        Return items
    End Function

    ' Call this FIRST THING in the Load event of every page (after ApplyTheme).
    ' It builds the bar in code, so you do not draw it in the designer.
    Public Sub BuildNav(f As Form, active As String)
        ' 1. make room: push every control down and make the form taller
        For Each c As Control In f.Controls.Cast(Of Control).ToList()
            c.Top += NavHeight
        Next
        f.ClientSize = New Size(f.ClientSize.Width, f.ClientSize.Height + NavHeight)

        ' 2. the bar with one link (a Label) for each page
        Dim bar As New Panel With {.Dock = DockStyle.Top, .Height = NavHeight, .BackColor = Color.White}
        Dim x As Integer = 12
        For Each item As String() In NavItems()
            Dim isActive As Boolean = (item(1) = active)
            Dim lbl As New Label With {
                .Text = item(0), .Tag = item(1), .AutoSize = True, .Cursor = Cursors.Hand,
                .Font = New Font("Segoe UI", 10, If(isActive, FontStyle.Bold, FontStyle.Regular)),
                .ForeColor = DarkGreen, .Padding = New Padding(12, 7, 12, 7),
                .BackColor = If(isActive, ColorTranslator.FromHtml("#CFEECF"), Color.White),
                .Left = x, .Top = 6}
            AddHandler lbl.Click, AddressOf NavClick
            bar.Controls.Add(lbl)
            x += lbl.GetPreferredSize(Size.Empty).Width + 6
        Next
        f.Controls.Add(bar)

        ' 3. closing a page with the X button (or Log out) goes back to Welcome,
        '    but not while we are only switching pages
        AddHandler f.FormClosed, Sub(s As Object, a As FormClosedEventArgs)
                                     If Not Navigating Then ReturnToWelcome()
                                 End Sub
        FitToScreen(f)
    End Sub

    Private Sub NavClick(sender As Object, e As EventArgs)
        Dim lbl As Label = DirectCast(sender, Label)
        If lbl.Font.Bold Then Return               ' already on this page
        NavigateTo(lbl.FindForm(), CStr(lbl.Tag))
    End Sub


    ' ============================================================
    ' BUILDING CONTROLS IN CODE
    ' Every form builds its own controls with these helpers, so you never open the designer.
    ' ============================================================
    Private Function Hx(code As String) As Color
        Return ColorTranslator.FromHtml("#" & code)
    End Function

    Public Function NewLabel(parent As Control, text As String, x As Integer, y As Integer, w As Integer, h As Integer,
                             Optional size As Single = 10, Optional bold As Boolean = False, Optional italic As Boolean = False,
                             Optional fore As String = "1B4332", Optional back As String = "",
                             Optional align As ContentAlignment = ContentAlignment.TopLeft,
                             Optional pad As Integer = 0) As Label
        Dim l As New Label()
        l.Text = text
        l.Location = New Point(x, y)
        If w <= 0 Then
            l.AutoSize = True
        Else
            l.AutoSize = False
            l.Size = New Size(w, h)
        End If
        Dim st As FontStyle = FontStyle.Regular
        If bold Then st = st Or FontStyle.Bold
        If italic Then st = st Or FontStyle.Italic
        l.Font = New Font("Segoe UI", size, st)
        l.ForeColor = Hx(fore)
        If back = "" Then l.BackColor = Color.Transparent Else l.BackColor = Hx(back)
        l.TextAlign = align
        If pad > 0 Then l.Padding = New Padding(pad)
        parent.Controls.Add(l)
        Return l
    End Function

    Public Function NewLink(parent As Control, text As String, x As Integer, y As Integer, w As Integer, h As Integer) As LinkLabel
        Dim l As New LinkLabel()
        l.Text = text
        l.Location = New Point(x, y)
        l.AutoSize = False
        l.Size = New Size(w, h)
        l.TextAlign = ContentAlignment.MiddleCenter
        l.LinkColor = Hx("1E7A3C")
        l.ActiveLinkColor = Hx("1B4332")
        l.BackColor = Color.Transparent
        l.Font = New Font("Segoe UI", 10)
        parent.Controls.Add(l)
        Return l
    End Function

    Public Function NewButton(parent As Control, text As String, x As Integer, y As Integer, w As Integer, h As Integer,
                              Optional tag As String = "") As Button
        Dim b As New Button()
        b.Text = text
        b.Location = New Point(x, y)
        b.Size = New Size(w, h)
        If tag <> "" Then b.Tag = tag          ' "danger" or "secondary" (ApplyTheme reads it)
        StyleButton(b)
        parent.Controls.Add(b)
        Return b
    End Function

    Public Function NewText(parent As Control, x As Integer, y As Integer, w As Integer, h As Integer,
                            Optional password As Boolean = False, Optional multiline As Boolean = False,
                            Optional maxLen As Integer = 0) As TextBox
        Dim t As New TextBox()
        t.Font = New Font("Segoe UI", 10)
        t.Location = New Point(x, y)
        t.Multiline = multiline
        If multiline Then t.Size = New Size(w, h) Else t.Width = w
        t.UseSystemPasswordChar = password
        If maxLen > 0 Then t.MaxLength = maxLen
        parent.Controls.Add(t)
        Return t
    End Function

    Public Function NewCombo(parent As Control, x As Integer, y As Integer, w As Integer) As ComboBox
        Dim c As New ComboBox()
        c.Font = New Font("Segoe UI", 10)
        c.DropDownStyle = ComboBoxStyle.DropDownList
        c.Location = New Point(x, y)
        c.Width = w
        parent.Controls.Add(c)
        Return c
    End Function

    Public Function NewCheck(parent As Control, text As String, x As Integer, y As Integer) As CheckBox
        Dim c As New CheckBox()
        c.Text = text
        c.AutoSize = True
        c.Font = New Font("Segoe UI", 10)
        c.ForeColor = Hx("1B4332")
        c.BackColor = Color.Transparent
        c.Location = New Point(x, y)
        parent.Controls.Add(c)
        Return c
    End Function

    Public Function NewRadio(parent As Control, text As String, x As Integer, y As Integer, w As Integer, h As Integer,
                             Optional isChecked As Boolean = False) As RadioButton
        Dim r As New RadioButton()
        r.Text = text
        r.Font = New Font("Segoe UI", 10)
        r.ForeColor = Hx("1B4332")
        r.BackColor = Color.Transparent
        r.Location = New Point(x, y)
        If w <= 0 Then
            r.AutoSize = True
        Else
            r.AutoSize = False
            r.Size = New Size(w, h)
        End If
        r.Checked = isChecked
        parent.Controls.Add(r)
        Return r
    End Function

    Public Function NewPanel(parent As Control, x As Integer, y As Integer, w As Integer, h As Integer,
                             Optional back As String = "", Optional border As Boolean = False) As Panel
        Dim p As New Panel()
        p.Location = New Point(x, y)
        p.Size = New Size(w, h)
        If back <> "" Then p.BackColor = Hx(back)
        If border Then p.BorderStyle = BorderStyle.FixedSingle
        parent.Controls.Add(p)
        Return p
    End Function

    Public Function NewGroup(parent As Control, text As String, x As Integer, y As Integer, w As Integer, h As Integer) As GroupBox
        Dim g As New GroupBox()
        g.Text = text
        g.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        g.ForeColor = Hx("1B4332")
        g.Location = New Point(x, y)
        g.Size = New Size(w, h)
        parent.Controls.Add(g)
        Return g
    End Function

    Public Function NewProgress(parent As Control, x As Integer, y As Integer, w As Integer, h As Integer) As ProgressBar
        Dim p As New ProgressBar()
        p.Location = New Point(x, y)
        p.Size = New Size(w, h)
        parent.Controls.Add(p)
        Return p
    End Function

    Public Function NewDate(parent As Control, x As Integer, y As Integer, w As Integer) As DateTimePicker
        Dim d As New DateTimePicker()
        d.Font = New Font("Segoe UI", 10)
        d.Format = DateTimePickerFormat.Custom
        d.CustomFormat = "ddd d MMM yyyy"
        d.Location = New Point(x, y)
        d.Width = w
        parent.Controls.Add(d)
        Return d
    End Function

    Public Function NewPicture(parent As Control, x As Integer, y As Integer, w As Integer, h As Integer) As PictureBox
        Dim p As New PictureBox()
        p.Location = New Point(x, y)
        p.Size = New Size(w, h)
        p.SizeMode = PictureBoxSizeMode.Zoom
        p.BackColor = Color.Transparent
        p.Image = MakeLogo()
        parent.Controls.Add(p)
        Return p
    End Function

    ' columns looks like "Reference:130|Date:120|Status:120"
    Public Function NewList(parent As Control, x As Integer, y As Integer, w As Integer, h As Integer, columns As String) As ListView
        Dim lv As New ListView()
        lv.Font = New Font("Segoe UI", 10)
        lv.Location = New Point(x, y)
        lv.Size = New Size(w, h)
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.HideSelection = False
        lv.MultiSelect = False
        lv.GridLines = True
        For Each part As String In columns.Split("|"c)
            Dim bits() As String = part.Split(":"c)
            lv.Columns.Add(bits(0), CInt(bits(1)))
        Next
        parent.Controls.Add(lv)
        Return lv
    End Function

    Public Function NewTabs(parent As Control, x As Integer, y As Integer, w As Integer, h As Integer,
                            tabW As Integer, tabH As Integer) As TabControl
        Dim t As New TabControl()
        t.Font = New Font("Segoe UI", 10)
        t.Location = New Point(x, y)
        t.Size = New Size(w, h)
        t.SizeMode = TabSizeMode.Fixed
        t.ItemSize = New Size(tabW, tabH)
        parent.Controls.Add(t)
        Return t
    End Function

    Public Function NewPage(tabs As TabControl, text As String) As TabPage
        Dim p As New TabPage(text)
        p.BackColor = Color.White
        tabs.TabPages.Add(p)
        Return p
    End Function

    ' ---------- 9. LOOK AND FEEL ----------
    ' Call ApplyTheme(Me) in every form's Load event
    ' ---------- FULL SCREEN ----------
    ' Welcome: the content stays centred.  Every other page: the content stretches to fill the screen.
    Public Sub FitToScreen(f As Form)
        Dim designW As Integer = f.ClientSize.Width
        Dim designH As Integer = f.ClientSize.Height
        Dim centred As Boolean = TypeOf f Is frmWelcome
        Dim offX As Integer = 0
        Dim header As Control = f.Controls.OfType(Of Panel)().FirstOrDefault(Function(pn) pn.Controls.OfType(Of PictureBox)().Any())   ' the green header is the panel that holds the logo
        Dim headerW As Integer = 0                        ' the design width of the header
        Dim logoutRight As Integer = 0                    ' gap between Log out and the right edge
        Dim logout As Control = Nothing
        If header IsNot Nothing Then
            headerW = header.Width
            For Each c As Control In header.Controls
                If TypeOf c Is Button Then logout = c
            Next
            If logout IsNot Nothing Then logoutRight = header.Width - logout.Right
        End If
        Dim baseBounds As New Dictionary(Of Control, Rectangle)()
        If Not centred Then
            RememberBounds(f, header, baseBounds)         ' the design size and place of every control
            MakeCardsTransparent(f, f)                    ' the group boxes become see-through cards over the background
        End If
        AddHandler f.Resize, Sub(s As Object, e As EventArgs)
                                 If f.WindowState = FormWindowState.Minimized Then Return
                                 If header IsNot Nothing Then
                                     header.Left = 0                                       ' the green header always fills the whole width
                                     header.Width = Math.Max(headerW, f.ClientSize.Width)
                                     If logout IsNot Nothing Then logout.Left = header.Width - logoutRight - logout.Width   ' Log out stays at the right edge
                                     header.Invalidate()
                                 End If
                                 If centred Then
                                     Dim nx As Integer = Math.Max(0, (f.ClientSize.Width - designW) \ 2)
                                     Dim dx As Integer = nx - offX
                                     If dx <> 0 Then
                                         For Each c As Control In f.Controls
                                             If c.Dock = DockStyle.None AndAlso c IsNot header Then c.Left += dx
                                         Next
                                         offX = nx
                                     End If
                                 Else
                                     Dim fx As Double = Math.Max(1.0, f.ClientSize.Width / designW)     ' how much bigger the window is than the design
                                     Dim fy As Double = Math.Max(1.0, f.ClientSize.Height / designH)
                                     For Each kv As KeyValuePair(Of Control, Rectangle) In baseBounds
                                         StretchControl(kv.Key, kv.Value, fx, fy)
                                     Next
                                 End If
                                 f.Invalidate()
                             End Sub
        f.AutoScroll = True                      ' scroll bars appear on a small screen
        f.FormBorderStyle = FormBorderStyle.Sizable
        f.MaximizeBox = True
        f.WindowState = FormWindowState.Maximized
    End Sub

    ' remembers where every control was designed to be (the header and the docked bars are left alone)
    Private Sub RememberBounds(parent As Control, header As Control, store As Dictionary(Of Control, Rectangle))
        For Each c As Control In parent.Controls
            If c Is header OrElse c.Dock <> DockStyle.None Then Continue For
            If Not TypeOf c Is TabPage Then store(c) = c.Bounds
            RememberBounds(c, header, store)
        Next
    End Sub

    ' puts one control at its stretched place: positions follow the window, big boxes and lists grow, buttons keep their size
    Private Sub StretchControl(c As Control, b As Rectangle, fx As Double, fy As Double)
        Dim w As Integer = b.Width
        Dim h As Integer = b.Height
        Dim growsWide As Boolean = TypeOf c Is GroupBox OrElse TypeOf c Is Panel OrElse TypeOf c Is TabControl OrElse TypeOf c Is ListView _
            OrElse TypeOf c Is TextBox OrElse TypeOf c Is ComboBox OrElse TypeOf c Is ProgressBar OrElse TypeOf c Is DateTimePicker _
            OrElse (TypeOf c Is Label AndAlso Not DirectCast(c, Label).AutoSize AndAlso c.BackColor <> Color.Transparent)
        Dim growsTall As Boolean = TypeOf c Is GroupBox OrElse TypeOf c Is Panel OrElse TypeOf c Is TabControl OrElse TypeOf c Is ListView
        If growsWide Then w = CInt(b.Width * fx)
        If growsTall Then h = CInt(b.Height * fy)
        c.SetBounds(CInt(b.Left * fx), CInt(b.Top * fy), w, h)
        If TypeOf c Is ListView Then
            Dim lv As ListView = DirectCast(c, ListView)
            If lv.View = View.Details AndAlso lv.Columns.Count > 0 Then lv.Columns(lv.Columns.Count - 1).Width = -2   ' the last column takes the free space
        End If
    End Sub

    ' group boxes that use the page colour become see-through; the form paints a white card behind them (see PaintBackground)
    Private Sub MakeCardsTransparent(f As Form, parent As Control)
        For Each c As Control In parent.Controls
            If TypeOf c Is GroupBox AndAlso c.BackColor = f.BackColor Then c.BackColor = Color.Transparent
            If TypeOf c Is GroupBox OrElse TypeOf c Is Panel Then MakeCardsTransparent(f, c)
        Next
    End Sub

    ' ---------- BACKGROUND: a sunny Accra street, drawn in code (no image file) ----------
    Public Sub AddBackground(f As Form)
        Try   ' stops the background from flickering
            GetType(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic).SetValue(f, True, Nothing)
        Catch
        End Try
        AddHandler f.Paint, AddressOf PaintBackground
        AddHandler f.Resize, Sub(s, e) f.Invalidate()
    End Sub

    Private Sub Box(g As Graphics, c As Color, a As Integer, x As Single, y As Single, w As Single, h As Single)
        Using b As New SolidBrush(Color.FromArgb(a, c))
            g.FillRectangle(b, x, y, w, h)
        End Using
    End Sub

    Private Sub Oval(g As Graphics, c As Color, a As Integer, x As Single, y As Single, w As Single, h As Single)
        Using b As New SolidBrush(Color.FromArgb(a, c))
            g.FillEllipse(b, x, y, w, h)
        End Using
    End Sub

    Private Sub DrawHouse(g As Graphics, x As Single, ground As Single, w As Single, h As Single, body As String)
        Box(g, Hx(body), 120, x, ground - h, w, h)                                     ' walls
        Using b As New SolidBrush(Color.FromArgb(130, Hx("A8400F")))                   ' roof
            g.FillPolygon(b, New PointF() {New PointF(x - 8, ground - h), New PointF(x + w + 8, ground - h), New PointF(x + w / 2, ground - h - h * 0.4F)})
        End Using
        Box(g, Hx("DDEFFF"), 200, x + w * 0.15F, ground - h * 0.7F, w * 0.25F, h * 0.25F)   ' window
        Box(g, Hx("8D6E63"), 170, x + w * 0.6F, ground - h * 0.5F, w * 0.22F, h * 0.5F)     ' door
    End Sub

    Private Sub DrawTree(g As Graphics, x As Single, ground As Single, size As Single)
        Box(g, Hx("8D6E63"), 150, x - size * 0.06F, ground - size * 0.5F, size * 0.12F, size * 0.5F)   ' trunk
        Oval(g, Hx("66BB6A"), 150, x - size * 0.35F, ground - size * 1.0F, size * 0.7F, size * 0.7F)   ' leaves
    End Sub

    Private Sub DrawPerson(g As Graphics, x As Single, ground As Single)
        Box(g, Hx("546E7A"), 150, x - 4, ground - 30, 4, 30)                           ' legs
        Box(g, Hx("546E7A"), 150, x + 3, ground - 30, 4, 30)
        Using b As New SolidBrush(Color.FromArgb(170, Hx("FFD166")))                   ' yellow dress
            g.FillRectangle(b, x - 8, ground - 62, 22, 34)
        End Using
        Oval(g, Hx("8D5524"), 170, x - 6, ground - 78, 18, 18)                         ' head
    End Sub

    Private Sub DrawTruck(g As Graphics, x As Single, ground As Single)
        Box(g, Hx("2D6A4F"), 150, x, ground - 52, 96, 42)                              ' body
        Box(g, Hx("1E7A3C"), 160, x + 96, ground - 38, 40, 28)                         ' cab
        Box(g, Hx("DDEFFF"), 200, x + 106, ground - 34, 22, 14)                        ' windscreen
        Oval(g, Hx("263238"), 200, x + 14, ground - 20, 22, 22)                        ' wheels
        Oval(g, Hx("263238"), 200, x + 104, ground - 20, 22, 22)
        Oval(g, Color.White, 220, x + 38, ground - 42, 20, 20)                         ' recycle disc
    End Sub

    Private Sub PaintBackground(sender As Object, e As PaintEventArgs)
        Dim f As Form = DirectCast(sender, Form)
        Dim w As Integer = f.ClientSize.Width
        Dim h As Integer = f.ClientSize.Height
        If w < 100 OrElse h < 200 Then Return
        Dim g As Graphics = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        Using br As New LinearGradientBrush(New Rectangle(0, 0, w, h), Hx("F4FBF4"), Hx("CFEECF"), 90.0F)
            g.FillRectangle(br, 0, 0, w, h)                                            ' soft green sky
        End Using
        Oval(g, Hx("FFD166"), 130, w - 170, 170, 90, 90)                               ' sun
        Dim ground As Single = h - 70                                                  ' where the street starts
        Box(g, Hx("B0BEC5"), 120, 0, ground, w, 70)                                    ' road
        Box(g, Hx("FFFFFF"), 150, 0, ground + 33, w, 4)                                ' road line
        Dim gap As Single = w / 6.0F
        DrawHouse(g, gap * 0.3F, ground, 110, 80, "F5E6CA")
        DrawTree(g, gap * 1.45F, ground, 90)
        DrawHouse(g, gap * 2.0F, ground, 120, 100, "CFEECF")
        DrawPerson(g, gap * 3.0F, ground + 20)
        DrawTree(g, gap * 3.4F, ground, 110)
        DrawHouse(g, gap * 3.8F, ground, 110, 85, "FFE3D8")
        DrawTruck(g, gap * 4.7F, ground + 24)
        DrawTree(g, gap * 5.8F, ground, 80)
        If TypeOf f Is frmWelcome Then Return
        ' the other pages: the street is softened, and every group box gets a white card on top of it
        Box(g, FormBack, 150, 0, 0, w, h)
        For Each c As Control In f.Controls
            If TypeOf c Is GroupBox AndAlso c.Visible AndAlso c.BackColor = Color.Transparent Then
                Box(g, Color.White, 225, c.Left, c.Top, c.Width, c.Height)
            End If
        Next
    End Sub

    Public Sub ApplyTheme(f As Form)
        f.BackColor = FormBack
        f.Font = New Font("Segoe UI", 10)
        AddBackground(f)
        StyleControls(f)
    End Sub

    Private Sub StyleControls(parent As Control)
        For Each c As Control In parent.Controls
            If TypeOf c Is Button Then StyleButton(DirectCast(c, Button))
            StyleControls(c)
        Next
    End Sub

    ' Button Tag decides the style: "danger" (Exit, Log out), "secondary" (outline) or nothing (main green)
    Public Sub StyleButton(b As Button)
        b.FlatStyle = FlatStyle.Flat
        b.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        b.Cursor = Cursors.Hand
        Select Case CStr(b.Tag)
            Case "danger"
                b.BackColor = Color.White : b.ForeColor = Terracotta : b.FlatAppearance.BorderColor = Terracotta
            Case "secondary"
                b.BackColor = Color.White : b.ForeColor = Green : b.FlatAppearance.BorderColor = Green
            Case Else
                b.BackColor = Green : b.ForeColor = Color.White : b.FlatAppearance.BorderColor = DarkGreen
        End Select
        RemoveHandler b.EnabledChanged, AddressOf ButtonEnabledChanged
        AddHandler b.EnabledChanged, AddressOf ButtonEnabledChanged
        RemoveHandler b.Click, AddressOf SaveAfterClick
        AddHandler b.Click, AddressOf SaveAfterClick          ' every button press saves the data
        If Not b.Enabled Then GreyOut(b)
    End Sub

    ' a button that is switched off turns light grey, so it is clear it cannot be pressed
    Private Sub SaveAfterClick(sender As Object, e As EventArgs)
        SaveData()
    End Sub
    Private Sub ButtonEnabledChanged(sender As Object, e As EventArgs)
        Dim b As Button = DirectCast(sender, Button)
        If b.Enabled Then StyleButton(b) Else GreyOut(b)
    End Sub

    Private Sub GreyOut(b As Button)
        b.BackColor = Hx("E3E8E3") : b.ForeColor = Hx("7A857A") : b.FlatAppearance.BorderColor = Hx("C8D0C8")
    End Sub

    ' Paints the green header. If the panel's Tag is "dark" (admin) it uses the darker green.
    Public Sub PaintHeader(sender As Object, e As PaintEventArgs)
        Dim p As Panel = DirectCast(sender, Panel)
        Dim c1 As Color, c2 As Color
        If CStr(p.Tag) = "dark" Then
            c1 = ColorTranslator.FromHtml("#2D6A4F") : c2 = DarkGreen
        Else
            c1 = ColorTranslator.FromHtml("#A8E6CF") : c2 = ColorTranslator.FromHtml("#66BB6A")
        End If
        Using br As New LinearGradientBrush(p.ClientRectangle, c1, c2, LinearGradientMode.Vertical)
            e.Graphics.FillRectangle(br, p.ClientRectangle)
        End Using
    End Sub

    ' ---------- 10. DEMO DATA (so you have people to log in with) ----------
    Private Seeded As Boolean = False

    Public Sub SeedDemo()
        If Seeded OrElse Users.Count > 0 Then Return      ' skip the demo users when saved data was loaded
        Seeded = True
        AddSeedUser("Kofi Collector", "kofi@example.com", "0201112222", "collector", "Labone")
        AddSeedUser("Akua Mensima", "akua@example.com", "0553334444", "collector", "Adenta")

        Dim ama As User = AddSeedUser("Ama Mensah", "ama@example.com", "0241234567", "resident", "Osu")
        ama.HouseNo = "14" : ama.Street = "Oxford Street" : ama.Landmark = "Opposite Osu Children's Park"
        ama.DigitalAddress = "GA-032-1458" : ama.Directions = "Blue gate, 2nd house after the pharmacy"
        ama.Plan = "Standard" : ama.PickupsLeft = 6 : ama.RenewsOn = Date.Today.AddDays(27)

        Dim yaw As User = AddSeedUser("Yaw Boateng", "yaw@example.com", "0207654321", "resident", "Madina")
        yaw.HouseNo = "7" : yaw.Street = "Zongo Junction Road" : yaw.Landmark = "Near the mosque"
        yaw.DigitalAddress = "GM-045-2210" : yaw.Directions = "Green gate, first house"
    End Sub

    Private Function AddSeedUser(name As String, email As String, phone As String, role As String, area As String) As User
        Dim u As New User With {.Name = name, .Email = email, .Phone = phone, .Password = "pass123", .Role = role, .Area = area}
        u.Lat = AreaCoords(area)(0)
        u.Lon = AreaCoords(area)(1)
        Users.Add(u)
        Return u
    End Function

End Module
