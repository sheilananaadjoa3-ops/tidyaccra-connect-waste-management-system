' ============================================================
' modData.vb : saves all users and pickups to a file, so nothing is lost
' when the program is closed. It loads the file again at the next start.
' ============================================================
Imports System.IO
Imports System.Globalization
Imports System.Xml.Linq

Public Module modData

    ' The database file. It lives in the Windows user folder, so rebuilding the project never deletes it.
    ' (Delete this file to go back to the demo users.)
    Public ReadOnly DataFile As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TidyAccraConnect", "tidyaccra_data.xml")

    Private lastSaved As String = ""
    Private ReadOnly Inv As CultureInfo = CultureInfo.InvariantCulture

    ' ---------- helpers to write and read single values ----------
    Private Function D(value As Date) As String
        Return value.ToString("o", Inv)
    End Function
    Private Function N(value As Double) As String
        Return value.ToString("R", Inv)
    End Function
    Private Function GetS(x As XElement, name As String) As String
        Dim a As XAttribute = x.Attribute(name)
        If a Is Nothing Then Return ""
        Return a.Value
    End Function
    Private Function GetD(x As XElement, name As String) As Date
        Dim r As Date
        If Date.TryParse(GetS(x, name), Inv, DateTimeStyles.RoundtripKind, r) Then Return r
        Return Date.MinValue
    End Function
    Private Function GetN(x As XElement, name As String) As Double
        Dim r As Double
        If Double.TryParse(GetS(x, name), NumberStyles.Float, Inv, r) Then Return r
        Return 0
    End Function
    Private Function FindUser(email As String) As User
        If email = "" Then Return Nothing
        Return Users.FirstOrDefault(Function(u) u.Email.ToLower() = email.ToLower())
    End Function

    ' ---------- SAVE ----------
    Public Sub SaveData()
        Try
            Dim root As New XElement("TidyAccra")
            Dim xu As New XElement("Users")
            For Each u As User In Users
                xu.Add(New XElement("User",
                    New XAttribute("Name", u.Name), New XAttribute("Email", u.Email), New XAttribute("Phone", u.Phone),
                    New XAttribute("Password", u.Password), New XAttribute("Role", u.Role), New XAttribute("Provider", u.Provider),
                    New XAttribute("Area", u.Area), New XAttribute("HouseNo", u.HouseNo), New XAttribute("Street", u.Street),
                    New XAttribute("Landmark", u.Landmark), New XAttribute("DigitalAddress", u.DigitalAddress),
                    New XAttribute("Directions", u.Directions), New XAttribute("Plan", u.Plan),
                    New XAttribute("PickupsLeft", u.PickupsLeft), New XAttribute("RenewsOn", D(u.RenewsOn)),
                    New XAttribute("Lat", N(u.Lat)), New XAttribute("Lon", N(u.Lon))))
            Next
            Dim xb As New XElement("Bookings")
            For Each b As Booking In Bookings
                Dim declined As String = String.Join(";", b.Declined.Select(Function(c) c.Email))
                xb.Add(New XElement("Booking",
                    New XAttribute("Ref", b.Ref),
                    New XAttribute("Owner", If(b.Owner Is Nothing, "", b.Owner.Email)),
                    New XAttribute("Collector", If(b.Collector Is Nothing, "", b.Collector.Email)),
                    New XAttribute("HouseNo", b.HouseNo), New XAttribute("Street", b.Street), New XAttribute("Landmark", b.Landmark),
                    New XAttribute("Area", b.Area), New XAttribute("DigitalAddress", b.DigitalAddress),
                    New XAttribute("Directions", b.Directions), New XAttribute("Waste", b.Waste), New XAttribute("Note", b.Note),
                    New XAttribute("Pay", b.Pay), New XAttribute("PickupDate", D(b.PickupDate)),
                    New XAttribute("CreatedAt", D(b.CreatedAt)), New XAttribute("Lat", N(b.Lat)), New XAttribute("Lon", N(b.Lon)),
                    New XAttribute("RadiusKm", N(b.RadiusKm)), New XAttribute("Status", b.Status),
                    New XAttribute("Declined", declined)))
            Next
            root.Add(xu)
            root.Add(xb)

            Dim text As String = root.ToString()
            If text = lastSaved Then Return                       ' nothing changed
            Directory.CreateDirectory(Path.GetDirectoryName(DataFile))
            Dim tmp As String = DataFile & ".tmp"
            File.WriteAllText(tmp, text)                          ' write to a temporary file first,
            If File.Exists(DataFile) Then File.Delete(DataFile)
            File.Move(tmp, DataFile)                              ' then swap it in (safe if the power goes off)
            lastSaved = text
        Catch ex As Exception
            ' if saving fails, the reason is written next to the data file
            Try
                File.WriteAllText(DataFile & ".error.txt", ex.ToString())
            Catch
            End Try
        End Try
    End Sub

    ' ---------- LOAD ----------
    ' Returns True if a saved file was found and loaded.
    Public Function LoadData() As Boolean
        Try
            If Not File.Exists(DataFile) Then Return False
            Dim root As XElement = XDocument.Load(DataFile).Root
            Users.Clear()
            Bookings.Clear()
            For Each x As XElement In root.Element("Users").Elements("User")
                Users.Add(New User With {
                    .Name = GetS(x, "Name"), .Email = GetS(x, "Email"), .Phone = GetS(x, "Phone"),
                    .Password = GetS(x, "Password"), .Role = GetS(x, "Role"), .Provider = GetS(x, "Provider"),
                    .Area = GetS(x, "Area"), .HouseNo = GetS(x, "HouseNo"), .Street = GetS(x, "Street"),
                    .Landmark = GetS(x, "Landmark"), .DigitalAddress = GetS(x, "DigitalAddress"),
                    .Directions = GetS(x, "Directions"), .Plan = GetS(x, "Plan"),
                    .PickupsLeft = CInt(GetN(x, "PickupsLeft")), .RenewsOn = GetD(x, "RenewsOn"),
                    .Online = False, .Lat = GetN(x, "Lat"), .Lon = GetN(x, "Lon")})        ' collectors start Offline
            Next
            For Each x As XElement In root.Element("Bookings").Elements("Booking")
                Dim b As New Booking With {
                    .Ref = GetS(x, "Ref"), .Owner = FindUser(GetS(x, "Owner")), .Collector = FindUser(GetS(x, "Collector")),
                    .HouseNo = GetS(x, "HouseNo"), .Street = GetS(x, "Street"), .Landmark = GetS(x, "Landmark"),
                    .Area = GetS(x, "Area"), .DigitalAddress = GetS(x, "DigitalAddress"), .Directions = GetS(x, "Directions"),
                    .Waste = GetS(x, "Waste"), .Note = GetS(x, "Note"), .Pay = GetS(x, "Pay"),
                    .PickupDate = GetD(x, "PickupDate"), .CreatedAt = GetD(x, "CreatedAt"),
                    .Lat = GetN(x, "Lat"), .Lon = GetN(x, "Lon"), .RadiusKm = GetN(x, "RadiusKm"), .Status = GetS(x, "Status")}
                For Each email As String In GetS(x, "Declined").Split(";"c)
                    Dim c As User = FindUser(email)
                    If c IsNot Nothing Then b.Declined.Add(c)
                Next
                If b.Owner IsNot Nothing Then Bookings.Add(b)
            Next
            Return Users.Count > 0
        Catch
            Users.Clear()
            Bookings.Clear()
            Return False
        End Try
    End Function

End Module

