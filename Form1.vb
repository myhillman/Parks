' Program to extract VKFF park details from CAPAD database, and make KML files for each state.
' $ = string interpolation
Option Explicit On
Imports System.Data.SQLite
Imports System.Diagnostics.Contracts
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Math
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Security
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Web.HttpUtility
Imports System.Xml
Imports System.Xml.XPath.Extensions
Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.AspNetCore.WebUtilities
Imports Microsoft.Office.Interop
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Form1
    Const GOOGLE_API_KEY As String = "AIzaSyDTv6-UZY1VNYb0rElEovb-7RiSY49qrU0"
    Const GOOGLE_STREET_VIEW_KEY As String = "AIzaSyD2AkQboX5YMF9FiBd0qxpRBIWbH2hqa9Y"
    Const PnPurl As String = "https://parksnpeaks.org/files/"        ' root url of PnP files
    Const YMDHMS As String = "yyyy/MM/dd HH:mm:ss"      ' date/time format
    Public Const PARKSdb = "data Source=parks.db"      ' SQLite database containing PARKS data
    Const SOTAdb = "data Source=SOTA.db"        ' SQLite database containing SOTA data
    Const SILOSdb = "data Source=silos.db"        ' SQLite database containing silos data
    Const MetersperFoot = 0.3048                ' conversion factor for feet to meters
    Const KMLheader = "<?xml version='1.0' encoding='UTF-8'?><kml xmlns='http://www.opengis.net/kml/2.2'><Document>"    ' standard header for kml file
    Const KMLfooter = "</Document></kml>"       ' standard footer for kml file
    Const PolyTransparency = 0.2                ' transparency of polygons (fraction of 1.0)
    Const PolyAlpha = 255 * PolyTransparency    ' Polygon alpha value
    Const GEOSERVER_SA4 = "https://censusdata.abs.gov.au/arcgis/rest/services/ASGS2016/SA4/MapServer/0/query"    ' Statistical Areas Level 4
    Const GEOSERVER_SA3 = "https://censusdata.abs.gov.au/arcgis/rest/services/ASGS2016/SA3/MapServer/0/query"    ' Statistical Areas Level 3
    Const SILO_ACTIVATION_ZONE = 1000   ' Activation zone radius for silo

    Dim DataSets As New Dictionary(Of String, DataSet)
    ReadOnly DatasetDict As New Dictionary(Of String, Dictionary(Of String, Object)) From {
                    {"CAPAD_T", New Dictionary(Of String, Object) From {
                            {"Shapefile", "F:/GIS Data/CAPAD/2020/CAPAD2020_terrestrial.shp"},  ' name of shapefile
                            {"shpShapeFileTable", Nothing},         ' ShapeFileFeatureTable
                            {"shpStyle", "terrestrial"},
                            {"shpFragments", Nothing},              ' collection of features as result of query
                            {"shpCopyright", "Collaborative Australian Protected Areas Database (CAPAD) 2020, Commonwealth of Australia 2021"},
                            {"shpAuthorityField", "AUTHORITY"},
                            {"dbfTableName", "terrestrial"},
                            {"AreaField", "GIS_AREA"},              ' name of area field in shapefile
                            {"shpNameField", "NAME"},
                            {"shpIDField", "PA_ID"},                ' name of Primary Key field
                            {"shpTypeField", "TYPE_ABBR"},
                            {"dbfAreaScale", 1},
                            {"HasXYmetaData", True}
                        }
                    },
                    {"CAPAD_M", New Dictionary(Of String, Object) From {
                            {"Shapefile", "F:/GIS Data/CAPAD/2020/CAPAD2020_marine.shp"},
                            {"shpShapeFileTable", Nothing},
                            {"shpStyle", "marine"},
                            {"shpFragments", Nothing},
                            {"shpCopyright", "Collaborative Australian Protected Areas Database (CAPAD) 2020, Commonwealth of Australia 2021"},
                            {"shpAuthorityField", "AUTHORITY"},
                            {"dbfTableName", "marine"},
                            {"AreaField", "GIS_AREA"},
                            {"shpNameField", "NAME"},
                            {"shpIDField", "PA_ID"},
                            {"shpTypeField", "TYPE_ABBR"},
                            {"dbfAreaScale", 1},
                            {"HasXYmetaData", True}
                        }
                    },
                    {"VIC_PARKS", New Dictionary(Of String, Object) From {
                            {"Shapefile", "F:/GIS Data/PARKRES/parkres.shp"},
                            {"shpShapeFileTable", Nothing},
                            {"shpStyle", "terrestrial"},
                            {"shpFragments", Nothing},
                            {"shpCopyright", "Victorian Department of Environment, Land, Water and Planning (DELWP)"},
                            {"shpAuthorityField", "MANAGER"},
                            {"dbfTableName", "parkres"},
                            {"AreaField", "HECTARES"},
                            {"shpNameField", "NAME"},
                            {"shpIDField", "PRIMS_ID"},
                            {"shpTypeField", "AREA_TYPE"},
                            {"dbfAreaScale", 1},
                            {"HasXYmetaData", False}
                        }
                    },
                    {"ZL", New Dictionary(Of String, Object) From {
                            {"Shapefile", "F:/GIS Data/NZ/DOC_PublicConservationAreas_2017_06_01.shp"},
                            {"shpShapeFileTable", Nothing},
                            {"shpStyle", "terrestrial"},
                            {"shpFragments", Nothing},
                            {"shpCopyright", "New Zealand Department of Conservation"},
                            {"shpAuthorityField", "Legislatio"},
                            {"dbfTableName", "DOC_PublicConservationAreas_2017_06_01"},
                            {"AreaField", "Shape_Area"},
                            {"shpNameField", "Name"},
                            {"shpIDField", "Conservati"},
                            {"shpTypeField", "Type"},
                            {"dbfAreaScale", 10000},
                            {"HasXYmetaData", False}
                    }
                }
            }

    ' All TYPE_ABBR from CAPAD (and a few others)
    ' https://www.environment.gov.au/fed/catalog/search/resource/details.page?uuid=%7B4448CACD-9DA8-43D1-A48F-48149FD5FCFD%7D
    ReadOnly LongNames As New Dictionary(Of String, String) From
            {
                {"AA", "Aboriginal Area"},
                {"ACCP", "Conservation Covenant"},
                {"ASMA", "Antarctic Specially Managed Areas"},
                {"ASPA", "Antarctic Specially Protected Area"},
                {"BG", "Botanic Gardens"},
                {"CA", "Conservation Area"},
                {"CCA", "Coordinated Conservation Area"},
                {"CCAZ1", "CCA Zone 1 NP"},
                {"CCAZ3", "CCA Zone 3 SCA"},
                {"CMP", "Coastal Marine Park"},
                {"CMR", "Commonwealth Marine Reserve"},
                {"COP", "Coastal Park"},
                {"COR", "Coastal Reserve"},
                {"CP", "Conservation Park"},
                {"CR", "Conservation Reserve"},
                {"DS", "Dolphin Sanctuary"},
                {"FFR", "Flora & Fauna Reserve"},
                {"FLR", "Flora Reserve"},
                {"FP", "Forest Park"},
                {"FR", "Forest Reserve"},
                {"GR", "Game Reserve"},
                {"HA", "Heritage Agreement"},
                {"HIR", "Historical Reserve"},
                {"HR", "Heritage River"},
                {"HS", "Historic Site"},
                {"HTR", "Hunting Reserve"},
                {"IPA", "Indigenous Protected Area"},
                {"KCR", "Karst Conservation Reserve"},
                {"MA", "Management Area"},
                {"MNP", "Marine National Park"},
                {"MP", "Marine Park"},
                {"NAP", "Nature Park"},
                {"NCR", "Nature Conservation Reserve"},
                {"NFR", "Natural Features Reserve"},
                {"NP", "National Park"},
                {"NPA", "National Park Aboriginal"},
                {"NPC", "Nature Conservation Reserve"},
                {"NR", "Nature Reserve"},
                {"NRA", "Nature Recreation Area"},
                {"NRS", "NRS Addition - Gazettal in Progress"},
                {"NREF", "Nature Refuge"},
                {"NS", "National Park (Scientific)"},
                {"OCA", "Other Conservation Area"},
                {"PNP", "Proposed National Park"},
                {"PPP", "Permanent Park Preserve"},
                {"PS", "Private Sanctuary"},
                {"R", "Reserve"},
                {"RCP", "Recreation Park"},
                {"REP", "Regional Park"},
                {"RR", "Regional Reserve"},
                {"RSR", "Resources Reserve"},
                {"S5G", "5(1)(g) Reserve"},
                {"S5H", "5(1)(h) Reserve"},
                {"SCA", "State Conservation Area"},
                {"SCR", "Scenic Reserve"},
                {"SP", "State Park"},
                {"SR", "State Reserve"},
                {"WP", "Wilderness Park"},
                {"WPA", "Wilderness Protection Area"},
                {"WR", "Wildlife Reserve"},
                {"WZ", "Wilderness Zone"}
            }

    ' Define winding direction for rings
    Enum Winding As Integer
        None
        CW              ' Clockwise
        Outer = CW      ' alias for clockwise
        CCW             ' Counter clockwise
        Inner = CCW     ' alias for counter clockwise
    End Enum

    Private ParksAdded As Integer   ' Count of parks added
    Private PointsBefore As Long     ' Count of polygon points before compression
    Private PointsAfter As Long  ' Count of polygon points after compression
    Public ReadOnly states As String() = {"VK0", "VK1", "VK2", "VK3", "VK4", "VK5", "VK6", "VK7", "VK8", "VK9", "ZL"} ' list of all states
    Dim count As Integer
    Dim found As Integer
    Dim Cookies As String = ""

    Private Delegate Sub SetTextCallback(tb As TextBox, ByVal text As String)

    Private Sub SetText(tb As TextBox, ByVal text As String)

        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If tb.InvokeRequired Then
            tb.Invoke(New SetTextCallback(AddressOf SetText), New Object() {tb, text})
        Else
            tb.Text = text
        End If
        Application.DoEvents()
    End Sub

    Private Sub AppendText(tb As TextBox, ByVal text As String)

        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If tb.InvokeRequired Then
            tb.Invoke(New SetTextCallback(AddressOf AppendText), New Object() {tb, text})
        Else
            tb.AppendText(text)
        End If
        Application.DoEvents()
    End Sub

    '=======================================================================================================
    ' Items on PARKS menu
    '=======================================================================================================
    Private Async Sub GenerateKMLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateKMLToolStripMenuItem.Click
        ' Generate KML files for selected states
        Dim count As Integer = 0, total As Integer = 0   ' count of parks
        Dim sql As SQLiteCommand, SQLdr As SQLiteDataReader, SOTAsql As SQLiteCommand
        Dim i As Integer, started As DateTime
        Dim state As String

        Try
            With Form3
                If .ShowDialog() = DialogResult.OK Then
                    Using connect As New SQLiteConnection(PARKSdb), SOTAconnect As New SQLiteConnection(SOTAdb), logWriter As New System.IO.StreamWriter("log.txt", True)
                        Dim htmlFile As String = CType(logWriter.BaseStream, FileStream).Name
                        ' String interpolation
                        logWriter.WriteLine($"{Now.ToString(YMDHMS) } - KML generation started")
                        PointsBefore = 0
                        PointsAfter = 0
                        ParksAdded = 0
                        Application.UseWaitCursor = True

                        connect.Open()  ' open database
                        sql = connect.CreateCommand
                        SOTAconnect.Open()  ' open database
                        SOTAsql = SOTAconnect.CreateCommand
                        For i = 0 To .ctrls.Count - 2
                            Dim box As KeyValuePair(Of String, CheckBox) = .ctrls.Item(i)    ' extract the checkbox
                            If box.Value.Checked Or .ctrls.Item(.ctrls.Count - 1).Value.Checked Then
                                ' Create kml file for this state
                                state = box.Key
                                count = 0
                                ' Get count of all parks in this state
                                sql.CommandText = $"SELECT COUNT(*) as count FROM parks WHERE State='{state }' AND lower(Status) IN ('active', 'pending')"
                                SQLdr = sql.ExecuteReader()
                                SQLdr.Read()
                                total = CInt(SQLdr("count"))
                                ' select all parks in this state
                                SQLdr.Close()
                                sql.CommandText = $"SELECT * FROM parks WHERE State='{state }' AND lower(Status) IN ('active', 'pending') ORDER BY WWFFID"
                                SQLdr = sql.ExecuteReader()
                                started = Now()
                                While SQLdr.Read()
                                    count += 1
                                    Await ParkToKML(SQLdr("WWFFID").ToString, SOTAsql).ConfigureAwait(False)
                                    SetText(TextBox1, $"Processed {count }/{total } in {state }. Finish {TogoFormat(started, count, total) }")
                                End While
                                SQLdr.Close()
                                logWriter.WriteLine($"{Now.ToString(YMDHMS) } - KML for {count } parks generated in {state }")
                            End If
                        Next
                        Application.UseWaitCursor = False
                        Dim msg As String = $"{Now.ToString(YMDHMS) } - Parks added {ParksAdded }. Polygons compressed to {PointsAfter * 100 / PointsBefore:f0}%"
                        SetText(TextBox1, msg)
                        logWriter.WriteLine(msg)
                        logWriter.Close()
                        connect.Close()
                        SOTAconnect.Close()
                        Process.Start(htmlFile)
                    End Using
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
        End Try
    End Sub

    Private Async Function ParkToKML(WWFFID As String, SOTAsql As SQLiteCommand, Optional debug As Boolean = False) As Task(Of Integer)
        ' terrestrial parks have a single PA_ID containing multiple polygons
        ' marine parks have multiple PA_ID containing a single polygon. They are of the form <PA_ID_GRP>_nnnn_M
        ' Why? Who knows
        Const IconScale = 0.7        ' scale of GE icons
        Const CoordinateGroupSize = 8     ' number of coordinates per line
        Dim ThisWinding As Winding = Winding.None, where As String = ""
        Dim myQueryFilter As New QueryParameters  ' create query
        Dim Centroid As MapPoint, GIS_AREA As Double, Style As String = "", ds As String, Authority As String = ""
        Dim env As Envelope = Nothing
        Dim GIS_Name As String = ""       ' park name in GIS data
        Dim Copyright As String
        Dim tagStack As New Stack(Of String)    ' FIFO for end tags
        Dim Windings As New List(Of Winding)        ' list of winding order for polygons
        Dim Rings As New List(Of Geometry)           ' list of rings in part. Each ring is a polygon with only 1 ring
        Dim extent As Envelope = Nothing   ' extent of all geometry
        Dim LabelPoint As MapPoint      ' position for park label
        Dim Polygons As Integer = 0       ' count of total polygons
        Dim Holes As Integer = 0       ' count of total polygons
        Dim Areas As New List(Of Double)
        Dim ParkData = New NameValueCollection, GISIDListQuoted As String, GeoData As (Area As Double, Center As MapPoint, Env As Envelope)
        Dim TotalWatch As New Stopwatch     ' measures total execution time

        Try     ' In an Async function you need a try block to catch line number of error
            TotalWatch.Restart()      ' start timing
            ParkData = GetParkData(WWFFID)       ' list of all data for this park
            If ParkData.Count = 0 Then Return 1     ' Nothing to show
            ds = ParkData("DataSet")
            If ds = Nothing Then GoTo done      ' Nothing to show

            Using logWriter As New System.IO.StreamWriter("debuglog.txt", True), PolygonLog As New System.IO.StreamWriter("polygonlog.txt", True)

                Dim logWriterName As String = CType(logWriter.BaseStream, FileStream).Name
                If PolygonLog.BaseStream.Position = 0 Then PolygonLog.WriteLine("Time,WWFFID,Name,State,Polygons,Holes")    ' header

                GISIDListQuoted = ParkData("GISIDListQuoted")
                where = DataSets(ds).BuildWhere(GISIDListQuoted)           ' build query statement
                With myQueryFilter
                    .WhereClause = where    ' query parameters
                    .OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
                    .ReturnGeometry = True
                End With
                DataSets(ds).shpFragments = Await DataSets(ds).shpShapeFileTable.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)           ' run query

                If DataSets(ds).shpFragments.Count = 0 Then
                    AppendText(TextBox2, $"Park {WWFFID } has missing GISID {GISIDListQuoted } in Dataset {ds }{vbCrLf }")
                Else
                    Style = DataSets(ds).shpStyle
                    Copyright = DataSets(ds).shpCopyright
                    GeoData = DataSets(ds).GeoData          ' get area, centroid, extent
                    GIS_AREA = GeoData.Area
                    Centroid = GeoData.Center
                    extent = GeoData.Env
                    Authority = DataSets(ds).shpFragments(0).GetAttributeValue(DataSets(ds).shpAuthorityField)
                    GIS_Name = DataSets(ds).ParkName & " " & DataSets(ds).ParkType
                    Dim BaseFilename As String = $"{Application.StartupPath }\files\{ParkData("State") }\{WWFFID }"
                    Using kmlWriter As New System.IO.StreamWriter(BaseFilename & ".kml")
                        ' Create kml header
                        kmlWriter.WriteLine(KMLheader)
                        kmlWriter.WriteLine("<description><![CDATA[<style>table, th, td {white-space:nowrap; }</style>")
                        kmlWriter.WriteLine("<table>")
                        kmlWriter.WriteLine("<tr><td>Data produced by Marc Hillman - VK3OHM/VK3IP</td></tr>")
                        Dim utc As String = String.Format("{0:dd MMM yyyy hh:mm:ss UTC}", DateTime.UtcNow)     ' time now in UTC
                        kmlWriter.WriteLine(String.Format("<tr><td>Data extracted from '{0}' key {1} on {2}.</td></tr>", DataSets(ds).shpShapeFileTable.DisplayName, ParkData("GISIDList"), utc))
                        kmlWriter.WriteLine($"<tr><td>Polygons have been generalised with a maximum deviation of {My.Settings.TrackSmoothness }m.</tr></td>")
                        kmlWriter.WriteLine("</table>")
                        kmlWriter.WriteLine("]]>")
                        kmlWriter.WriteLine("</description>")

                        kmlWriter.WriteLine("<Style id='red'>")
                        kmlWriter.WriteLine("<LineStyle><width>2</width><color>ffffffff</color></LineStyle>")
                        kmlWriter.WriteLine($"<PolyStyle><color>{KMLColor(PolyAlpha, 0, 0, 255) }</color><fill>1</fill><outline>1</outline></PolyStyle>")
                        kmlWriter.WriteLine("</Style>")
                        kmlWriter.WriteLine("<Style id='blue'>")
                        kmlWriter.WriteLine("<LineStyle><width>2</width><color>ffffffff</color></LineStyle>")
                        kmlWriter.WriteLine($"<PolyStyle><color>{KMLColor(PolyAlpha, 255, &H99, 0) }</color><fill>1</fill><outline>1</outline></PolyStyle>")
                        kmlWriter.WriteLine("</Style>")
                        kmlWriter.WriteLine("<Style id='pink'>")
                        kmlWriter.WriteLine("<LineStyle><width>2</width><color>ffffffff</color></LineStyle>")
                        kmlWriter.WriteLine($"<PolyStyle><color>{KMLColor(PolyAlpha, 255, &H33, 255) }</color><fill>1</fill><outline>1</outline></PolyStyle>")
                        kmlWriter.WriteLine("</Style>")

                        ' Style polygons so that they are Blue/Pink normally, but Red when highlighted
                        kmlWriter.WriteLine("<StyleMap id='terrestrial'>")
                        kmlWriter.WriteLine("<Pair><key>normal</key><styleUrl>#blue</styleUrl></Pair>")
                        kmlWriter.WriteLine("<Pair><key>highlight</key><styleUrl>#red</styleUrl></Pair>")
                        kmlWriter.WriteLine("</StyleMap>")
                        kmlWriter.WriteLine("<StyleMap id='marine'>")
                        kmlWriter.WriteLine("<Pair><key>normal</key><styleUrl>#pink</styleUrl></Pair>")
                        kmlWriter.WriteLine("<Pair><key>highlight</key><styleUrl>#red</styleUrl></Pair>")
                        kmlWriter.WriteLine("</StyleMap>")

                        ParksAdded += 1
                        ' produce the placemark header
                        kmlWriter.WriteLine($"<Placemark id='{WWFFID }'>")
                        kmlWriter.WriteLine($"<styleUrl>#{Style }</styleUrl>")
                        ' Convert abbreviation to long name
                        kmlWriter.WriteLine($"<name>{HttpUtility.HtmlEncode(ParkData("LongName")) }</name>")
                        kmlWriter.WriteLine("<description><![CDATA[<style>table, th, td {white-space:nowrap; }</style>")
                        kmlWriter.WriteLine("<table border='1'><tr><th>Item</th><th>Value</th></tr>")
                        kmlWriter.WriteLine($"<tr><td><b>WWFF</b></td><td><b>{WWFFID }</b></td></tr>")
                        kmlWriter.WriteLine($"<tr><td><b>Name</b></td><td><b>{HttpUtility.HtmlEncode(ParkData("LongName")) }</b></td></tr>")
                        If ParkData("DXCC").Length > 0 Then kmlWriter.WriteLine($"<tr><td>DXCC</td><td>{ParkData("DXCC") }</td></tr>")
                        kmlWriter.WriteLine($"<tr><td>IOTA</td><td>{ParkData("IOTAID") }</td></tr>")
                        ' kmlWriter.WriteLine($"<tr><td>POTA</td><td>{ParkData("POTAID") }</td></tr>")
                        If ParkData("SANPCPAID").Length > 0 Then kmlWriter.WriteLine($"<tr><td>SANPCPA</td><td>{ParkData("SANPCPAID") }</td></tr>")
                        ' Get Shire data, if any
                        If Not String.IsNullOrEmpty(ParkData("ShireID")) Then
                            kmlWriter.WriteLine($"<tr><td>Local Government Area</td><td>{ParkData("ShireName") } ({ParkData("ShireID") })</td></tr>")
                        End If
                        If ParkData("Region").Length > 0 Then kmlWriter.WriteLine($"<tr><td>Region</td><td>{HttpUtility.HtmlEncode(ParkData("Region")) }</td></tr>")
                        If ParkData("District").Length > 0 Then kmlWriter.WriteLine($"<tr><td>District</td><td>{HttpUtility.HtmlEncode(ParkData("District")) }</td></tr>")
                        If ParkData("KRMNPAID").Length > 0 Then kmlWriter.WriteLine($"<tr><td>KRMNPA</td><td>{ParkData("KRMNPAID") }</td></tr>")
                        If ParkData("HTTPLink").Length > 0 Then kmlWriter.WriteLine($"<tr><td>Hyperlink</td><td>{ParkData("HTTPLink") }</td></tr>")
                        If ParkData("Notes").Length > 0 Then kmlWriter.WriteLine($"<tr><td>Notes</td><td>{HttpUtility.HtmlEncode(ParkData("Notes")) }</td></tr>")
                        kmlWriter.WriteLine($"<tr><td>AREA (hectares)</td><td>{GIS_AREA:f0}</td></tr>")
                        kmlWriter.WriteLine($"<tr><td>CENTROID</td><td>{Centroid.Y:f5},{Centroid.X:f5}</td></tr>")
                        kmlWriter.WriteLine($"<tr><td>AUTHORITY</td><td>{HttpUtility.HtmlEncode(Authority) }</td></tr>")
                        kmlWriter.WriteLine($"<tr><td>GIS ID</td><td>Key {DataSets(ds).shpIDField }={ParkData("GISIDList") } in shapefile '{DataSets(ds).shpShapeFileTable.DisplayName }'</td></tr>")
                        kmlWriter.WriteLine($"<tr><td>GIS Name</td><td>{GIS_Name }</td></tr>")
                        If Copyright.Length > 0 Then kmlWriter.WriteLine($"<tr><td>Copyright</td><td>{HttpUtility.HtmlEncode(Copyright) }</td></tr>")
                        kmlWriter.WriteLine("</table>]]></description>")

                        ' if a terrestrial park there will be a handful of fragments fragment, if marine or VIC_PARKS there could be many
                        kmlWriter.WriteLine("<MultiGeometry>")
                        For Each fragment As Feature In DataSets(ds).shpFragments
                            Dim original = CType(fragment.Geometry, Esri.ArcGISRuntime.Geometry.Polygon)
                            ' Calculate number of points before generalization
                            For Each part In original.Parts
                                PointsBefore += part.Count
                            Next
                            fragment.Geometry = GeometryEngine.Project(fragment.Geometry, SpatialReferences.WebMercator)  ' convert fragment to datum with linear units
                            fragment.Geometry = GeometryEngine.Simplify(fragment.Geometry)              ' ensure polygons are correct

                            Dim poly = CType(fragment.Geometry, Esri.ArcGISRuntime.Geometry.Polygon)    ' cast geometry to polygon to calculate winding direction
                            If poly.Parts.Any Then                                ' sometimes we get a degenerate part
                                Polygons += poly.Parts.Count
                                ' Calculate number of points after generalization

                                Rings.Clear()
                                Windings.Clear()
                                Areas.Clear()
                                If debug Then
                                    logWriter.WriteLine($"Date/time: {Now() }, Park: {WWFFID } - {ParkData("Name") }")
                                End If
                                Dim pindx As Integer = 0
                                For Each part As ReadOnlyPart In poly.Parts
                                    If part.Count >= 3 Then     ' need 3 or more segments to be a polygon
                                        Dim PntsBefore As Integer = part.Count
                                        Dim ring As Geometry = New Polygon(part)
                                        ' Calculate area of all rings. <0 = CW, >=0 = CCW
                                        Dim Area As Double = PolygonArea(part)
                                        ' Calculate the required smoothness.
                                        ' "TrackSmoothness" is the default. If a polygon is smaller than "SmallPark" then use a smaller smoothness value
                                        ' "SmallPark" is in square meters. 1 hectacre = 10,000 sq. m, i.e. 100m x 100m
                                        Dim Smoothness As Integer = My.Settings.TrackSmoothness
                                        If Abs(Area) < My.Settings.SmallPark Then Smoothness /= 10
                                        ring = GeometryEngine.Generalize(ring, Smoothness, True) ' remove extraneous points
                                        Dim p = CType(ring, Esri.ArcGISRuntime.Geometry.Polygon)
                                        If p.Parts.Any Then   ' sometimes the generalize seems to knock out polygons
                                            Rings.Add(ring)
                                            Areas.Add(Area)
                                            If Area < 0 Then ThisWinding = Winding.Outer Else ThisWinding = Winding.Inner
                                            Windings.Add(ThisWinding)
                                            p = CType(ring, Esri.ArcGISRuntime.Geometry.Polygon)
                                            Dim PntsAfter = p.ToPolyline.Parts(0).PointCount
                                            PointsAfter += PntsAfter
                                            If debug Then
                                                logWriter.WriteLine(String.Format("Fragment: {0,2},  Winding: {1,5}, Area(ha): {2,12:n2}, Points before: {3,6}, Points after: {4,6}, Smoothness: {5,2}", pindx, ThisWinding.ToString, Area / 10000, PntsBefore, PntsAfter, Smoothness))
                                            End If
                                            pindx += 1
                                        End If
                                    End If
                                Next

                                ' ------------------------------------------------------------------
                                ' Marine parks have a 100m littoral area
                                '  If a marine park, then add 100m buffer zone where it intersects with land
                                ' ------------------------------------------------------------------

                                If ds = "CAPAD_M" Then
                                    Dim fs As FeatureCollection
                                    Using MyWebClient As New WebClient
                                        Try
                                            For i = 0 To Rings.Count - 1
                                                If Windings(i) = Winding.Outer Then
                                                    Dim ring As Polygon = CType(Rings(i), Esri.ArcGISRuntime.Geometry.Polygon)
                                                    Dim buffer As Geometry = GeometryEngine.BufferGeodetic(ring, 100, LinearUnits.Meters)    ' create 100m buffer around ring
                                                    buffer = GeometryEngine.Generalize(buffer, My.Settings.TrackSmoothness, True) ' make buffer sensibly small

                                                    ' Prepare some POST fields for a request to the map server. We use POST because the requests are too large for a GET
                                                    Dim POSTfields = New NameValueCollection From {
                                                    {"f", "json"},
                                                    {"geometryType", "esriGeometryPolygon"},
                                                    {"inSR", $"{{'wkid':{buffer.SpatialReference.Wkid }}}"},
                                                    {"spatialRel", "esriSpatialRelIntersects"},
                                                    {"returnGeometry", "true"},          '  need the geometry
                                                    {"geometry", buffer.ToJson},
                                                    {"outFields", "*"}
                                                }
                                                    ' return all fields, even though we only use the default
                                                    Dim resp As Byte() = MyWebClient.UploadValues(GEOSERVER_SA4, "POST", POSTfields)
                                                    Dim responseStr = System.Text.Encoding.UTF8.GetString(resp)
                                                    fs = FeatureCollection.FromJson(responseStr)

                                                    If fs IsNot Nothing And fs.Tables.Any Then
                                                        ' Intersection(s) found. Produce a littoral polygon
                                                        For Each feature In fs.Tables(0)
                                                            Dim coastline As Polygon = feature.Geometry         ' polygon of intersecting coastline
                                                            coastline = GeometryEngine.Project(coastline, Rings(i).SpatialReference)  ' convert coastline
                                                            Dim littoral = GeometryEngine.Intersection(buffer, coastline)  ' the intersection of the buffer and the coastline
                                                            If Not littoral.IsEmpty Then
                                                                littoral = GeometryEngine.Project(littoral, Rings(i).SpatialReference)  ' convert intersection
                                                                ' Remove bits of intersect that overlap the ring
                                                                Dim overlap As Geometry = GeometryEngine.Intersection(littoral, Rings(i))   ' generate overlaps
                                                                littoral = GeometryEngine.Difference(littoral, overlap)     ' remove areas of overlap
                                                                ' Now fill in gaps. Join polygons together and remove holes
                                                                Dim joined As Geometry = GeometryEngine.Union(littoral, Rings(i))
                                                                Dim j As Polygon = CType(joined, Polygon)
                                                                For Each jp As IReadOnlyList(Of ReadOnlyPart) In j.Parts
                                                                    If PolygonArea(jp) > 0 Then
                                                                        Dim jpp As Geometry = New Polygon(jp)
                                                                        If GeometryEngine.Touches(littoral, jpp) Then
                                                                            littoral = GeometryEngine.Union(littoral, jpp)
                                                                        End If
                                                                    End If
                                                                Next
                                                                ' Add littoral piece to other rings
                                                                Dim p = CType(littoral, Polygon)
                                                                For Each pp As IReadOnlyList(Of ReadOnlyPart) In p.Parts
                                                                    Dim rng As Geometry = New Polygon(pp)
                                                                    Windings.Add(Winding.CW)
                                                                    Rings.Add(rng)
                                                                    Areas.Add(PolygonArea(pp))
                                                                    If debug Then logWriter.WriteLine("Adding a littoral ring")
                                                                Next
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            Next
                                        Catch ex As WebException
                                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                                        End Try
                                    End Using
                                End If
                                ' ------------------------------------------------------------------

                                SetText(TextBox1, $"Generating {ParkData("State") } - {WWFFID } {ParkData("Name") } : Polygons {poly.Parts.Count:d}")
                                Application.DoEvents()
                                tagStack.Clear()
                                ' convert all rings to GE datum
                                For RingIndex As Integer = 0 To Rings.Count - 1
                                    Rings(RingIndex) = GeometryEngine.Project(Rings(RingIndex), SpatialReferences.Wgs84)
                                Next

                                Dim Coordinates As New List(Of String)      ' list of polygon vertices
                                Dim PartIndex As Integer = 0
                                For Each ring In Rings      ' create all polygons
                                    poly = CType(ring, Polygon)
                                    If poly.Parts.Count > 0 Then    ' Count should always be 1, but sometimes Project adds rings
                                        ' In ESRI shapefile, an outer boundary is CW, and inner is CCW
                                        ' In KML all polygon must be CCW. Type is determined by <innerBoundaryIs> and <outerBoundaryIs>

                                        ' Create a list of coordinates, in case we have to reverse it
                                        Coordinates.Clear()
                                        For Each p As Segment In poly.Parts(0)
                                            Coordinates.Add($"{p.StartPoint.X:f5},{p.StartPoint.Y:f5}")
                                        Next
                                        Coordinates.Add($"{poly.Parts(0).StartPoint.X:f5},{poly.Parts(0).StartPoint.Y:f5}")   ' Close the polygon

                                        If Windings(PartIndex) = Winding.Outer Then
                                            ' start of a new outer boundary - may need to close preceding one
                                            While (tagStack.Any)
                                                kmlWriter.WriteLine(tagStack.Pop)
                                            End While
                                            kmlWriter.WriteLine("<Polygon>")    ' open new one
                                            tagStack.Push("</Polygon>")         ' push end tag
                                            kmlWriter.WriteLine("<tessellate>1</tessellate>")
                                            kmlWriter.WriteLine("<outerBoundaryIs>")
                                            tagStack.Push("</outerBoundaryIs>")
                                        Else
                                            kmlWriter.WriteLine("<innerBoundaryIs>")
                                            Holes += 1          '  inner boundary is a hole
                                            tagStack.Push("</innerBoundaryIs>")
                                        End If
                                        ' Now produce the polygon as a linear ring
                                        kmlWriter.WriteLine("<LinearRing><coordinates>")
                                        If Windings(PartIndex) = Winding.CW Then Coordinates.Reverse()           ' make coordinates CCW
                                        ' produce coordinates in a group per line
                                        Dim i As Integer, GroupStart As Integer, GroupEnd As Integer
                                        Dim group As New List(Of String)
                                        For GroupStart = 0 To Coordinates.Count - 1 Step CoordinateGroupSize  ' do coordinates in groups
                                            group.Clear()
                                            GroupEnd = Min(GroupStart + CoordinateGroupSize - 1, Coordinates.Count - 1)     ' last group may be short
                                            For i = GroupStart To GroupEnd
                                                group.Add(Coordinates(i))
                                            Next
                                            kmlWriter.WriteLine(Join(group.ToArray, " "))       ' write group of coordinates with spaces
                                        Next
                                        kmlWriter.WriteLine("</coordinates></LinearRing>")
                                        kmlWriter.WriteLine(tagStack.Pop)    ' close of boundaryIs
                                    End If
                                    PartIndex += 1
                                Next
                                ' Close any open polygon
                                While (tagStack.Any)
                                    kmlWriter.WriteLine(tagStack.Pop)
                                End While
                            End If
                        Next

                        kmlWriter.WriteLine("</MultiGeometry>")
                        kmlWriter.WriteLine("</Placemark>")
                        GoTo NoName         ' name is too big when zoomed out
                        ' Add park name label
                        kmlWriter.WriteLine("<Style id='label'>")
                        kmlWriter.WriteLine("<IconStyle><Icon/></IconStyle>")  ' no icon
                        kmlWriter.WriteLine("<LabelStyle><color>ff000000</color><scale>2</scale></LabelStyle>")
                        kmlWriter.WriteLine("</Style>")
                        kmlWriter.WriteLine("<Placemark>")
                        kmlWriter.WriteLine("<styleUrl>#label</styleUrl>")
                        kmlWriter.WriteLine($"<name>{WWFFID }</name>")
                        kmlWriter.WriteLine($"<Point><coordinates>{LabelPoint.X },{LabelPoint.Y }</coordinates></Point>")
                        kmlWriter.WriteLine("</Placemark>")
NoName:
                        ' Now add any SOTA placemarks
                        Dim SOTASQLdr As SQLiteDataReader
                        SOTAsql.CommandText = $"SELECT * FROM SummitsInParks JOIN SOTA on SummitsInParks.SummitCode=SOTA.SummitCode WHERE WWFFID='{WWFFID }'"
                        SOTASQLdr = SOTAsql.ExecuteReader()
                        Dim StylesReqd As Boolean = True        ' need to add pin styles if we have any SOTA summits
                        While SOTASQLdr.Read()
                            If StylesReqd Then
                                ' Styles for SOTA pins
                                Dim styles() As String = {"_park", "_caution"}      ' only 2 summit styles required for parks
                                Dim url As String = PnPurl & "pins/"
                                ' url = "F:\Users\Marc\Documents\Visual Studio 2017\Projects\Parks\Parks\bin\Debug\files\pins\"   ' local testing
                                For pin As Integer = 1 To 10
                                    For Each s In styles
                                        kmlWriter.WriteLine($"<Style id='{pin }{s }_normal'>")
                                        kmlWriter.WriteLine("<IconStyle>")
                                        kmlWriter.WriteLine($"<Icon><href>{url }{pin }{s }.png</href></Icon>")
                                        kmlWriter.WriteLine($"<scale>{IconScale }</scale>")
                                        kmlWriter.WriteLine("</IconStyle>")
                                        kmlWriter.WriteLine("<LabelStyle><scale>0</scale></LabelStyle>")
                                        kmlWriter.WriteLine("</Style>")
                                        kmlWriter.WriteLine($"<Style id='{pin }{s }_highlight'>")
                                        kmlWriter.WriteLine("<IconStyle>")
                                        kmlWriter.WriteLine($"<Icon><href>{url }{pin }{s }.png</href></Icon>")
                                        kmlWriter.WriteLine($"<scale>{IconScale }</scale>")
                                        kmlWriter.WriteLine("</IconStyle>")
                                        kmlWriter.WriteLine("<LabelStyle><scale>1</scale></LabelStyle>")
                                        kmlWriter.WriteLine("</Style>")

                                        kmlWriter.WriteLine($"<StyleMap id ='{pin }{s }'>")
                                        kmlWriter.WriteLine($"<Pair><key>normal</key><styleUrl>#{pin }{s }_normal</styleUrl></Pair>")
                                        kmlWriter.WriteLine($"<Pair><key>highlight</key><styleUrl>#{pin }{s }_highlight</styleUrl></Pair>")
                                        kmlWriter.WriteLine("</StyleMap>")
                                    Next
                                Next
                                StylesReqd = False
                            End If
                            kmlWriter.WriteLine("<Placemark>")
                            kmlWriter.WriteLine($"<name>{SOTASQLdr.Item("SummitCode") } - {SOTASQLdr.Item("SummitName") }</name>")
                            kmlWriter.WriteLine("<description><![CDATA[<style>table, th, td {white-space:nowrap; }</style>")
                            kmlWriter.WriteLine("<table border='1'><tr><th>Item</th><th>Value</th></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Summit Code" }</td><td>{SOTASQLdr.Item("SummitCode") }</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Summit Name" }</td><td>{SOTASQLdr.Item("SummitName") }</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Association Name" }</td><td>{SOTASQLdr.Item("AssociationName") }</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Region Name" }</td><td>{SOTASQLdr.Item("RegionName") }</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Altitude (m)" }</td><td>{SOTASQLdr.Item("AltM") }</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Longitude" }</td><td>{SOTASQLdr.Item("GridRef1"):f5}</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Latitude" }</td><td>{SOTASQLdr.Item("GridRef2"):f5}</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Points" }</td><td>{SOTASQLdr.Item("Points") }</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Bonus Points" }</td><td>{SOTASQLdr.Item("BonusPoints") }</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Valid from" }</td><td>{SOTASQLdr.Item("ValidFrom") }</td></tr>")
                            kmlWriter.WriteLine($"<tr><td>{"Valid to" }</td><td>{SOTASQLdr.Item("ValidTo") }</td></tr>")
                            Dim points As Integer = SOTASQLdr.Item("Points").ToString
                            Dim PinStyle As String
                            If SOTASQLdr.Item("WithIn").ToString = "N" Then
                                PinStyle = $"{points }_caution"
                                kmlWriter.WriteLine($"<tr><td>{"<b>WARNING</b>" }</td><td>{"<b>Summit is within 100m of park boundary</b>" }</td></tr>")
                            Else
                                PinStyle = $"{points }_park"
                            End If
                            kmlWriter.WriteLine("</table>]]></description>")
                            kmlWriter.WriteLine($"<styleUrl>#{PinStyle }</styleUrl>")
                            kmlWriter.WriteLine($"<Point><coordinates>{CDbl(SOTASQLdr.Item("GridRef1")):f5},{CDbl(SOTASQLdr.Item("GridRef2")):f5}</coordinates></Point>")
                            kmlWriter.WriteLine("</Placemark>")
                        End While
                        SOTASQLdr.Close()
                        ' finish kml document
                        ' Create an appropriate LookAt
                        Dim Range As Double = CalcRange(extent)
                        Dim LookAtPoint As MapPoint = extent.GetCenter      ' center of extent
                        kmlWriter.WriteLine($"<LookAt><longitude>{LookAtPoint.X:f5}</longitude><latitude>{LookAtPoint.Y:f5}</latitude><range>{Range:f0}</range><heading>0</heading><tilt>0</tilt></LookAt>")
                        ' Record extents of this park for index routine
                        kmlWriter.WriteLine($"<ExtendedData><Data name=""north""><value>{extent.YMax:f5}</value></Data><Data name=""south""><value>{extent.YMin:f5}</value></Data><Data name=""east""><value>{extent.XMax:f5}</value></Data><Data name=""west""><value>{extent.XMin:f5}</value></Data></ExtendedData>")
                        kmlWriter.WriteLine(KMLfooter) ' terminate kml
                        kmlWriter.Close()
                    End Using
                    ' compress to zip file
                    System.IO.File.Delete(BaseFilename & ".kmz")
                    Dim zip As ZipArchive = ZipFile.Open(BaseFilename & ".kmz", ZipArchiveMode.Create)    ' create new archive file
                    zip.CreateEntryFromFile(BaseFilename & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
                    zip.Dispose()
                End If
                logWriter.Close()
                PolygonLog.WriteLine($"{Now() },{WWFFID },""{Name }"",{ParkData("State") },{Polygons },{Holes }")
                PolygonLog.Close()
                If debug Then Process.Start(logWriterName)      ' display log file
            End Using
done:
            AppendText(TextBox2, $"ParkToKML {WWFFID} - {TotalWatch.ElapsedMilliseconds / 1000:f1}s{vbCrLf }")
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
        End Try
        Return 1
    End Function
    Private Sub UpdateTextBox(ByVal text As String)
        If Me.InvokeRequired Then
            Dim args() As String = {text}
            Me.Invoke(New Action(Of String)(AddressOf UpdateTextBox), args)
            Return
        End If
        TextBox2.AppendText(text)
    End Sub

    Private Sub ImportParkscsvToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportParkscsvToolStripMenuItem.Click
        ' Import the parks data file into SQLite
        Dim count As Integer = 0    ' count of lines read
        Dim sql As SQLiteCommand

        Using connect As New SQLiteConnection(PARKSdb)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("PARKS.csv")
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                MyReader.HasFieldsEnclosedInQuotes = True
                Dim currentRow As String()
                Dim cmd As String
                Dim fields As String = "", values As String

                sql.CommandText = "begin"
                sql.ExecuteNonQuery()       ' start transaction for speed
                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        count += 1
                        If count = 1 Then
                            ' remove existing data
                            sql.CommandText = "DELETE FROM parks"   ' remove existing data
                            sql.ExecuteNonQuery()
                            fields = Join(currentRow, ",")
                        Else
                            currentRow(18) = currentRow(18).Replace("ha", "")      ' remove hectacres
                            For i = 0 To currentRow.Length - 1
                                currentRow(i) = currentRow(i).Replace("'", "''")   ' escape single quotes
                                currentRow(i) = currentRow(i).Replace("NULL", "").Replace("NOTHING", "").Replace("UNKNOWN", "")
                                currentRow(i) = "'" & currentRow(i) & "'"
                            Next
                            values = Join(currentRow, ",")
                            cmd = "INSERT INTO parks (" & fields & ") VALUES (" & values & ")"
                            sql.CommandText = cmd
                            sql.ExecuteNonQuery()
                        End If
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped." & vbCrLf & "Import aborted")
                        sql.CommandText = "rollback"
                        sql.ExecuteNonQuery()
                        GoTo done
                    End Try
                End While
            End Using
            sql.CommandText = "end"
            sql.ExecuteNonQuery()
done:
            connect.Close()
        End Using
        MsgBox(count - 1 & " parks imported", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Import complete")
    End Sub

    Private Sub ImportParkssqlToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ImportParkssqlToolStripMenuItem.Click
        ' Import a parks.sql file
        Dim sql As SQLiteCommand, sqlcmd As String, added As Integer, deleted As Integer, index As Integer

        Using connect As New SQLiteConnection(PARKSdb)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            Application.UseWaitCursor = True : Application.DoEvents()
            Try
                sqlcmd = File.ReadAllText("Parks.sql")  ' read entire sql script
                sqlcmd = Replace(sqlcmd, "\'", "''")    ' convert php escape to SQLite escape character
                ' Remove all before "INSERT INTO"
                index = sqlcmd.IndexOf("INSERT INTO `PARKS` (", StringComparison.CurrentCulture)
                If index < 0 Then index = sqlcmd.IndexOf("INSERT INTO PARKS (", StringComparison.CurrentCulture)
                If index < 0 Then
                    MsgBox("No INSERT INTO `PARKS` statement found in file", vbCritical + vbOKOnly, "Bad SQL file")
                    Exit Sub
                Else
                    sqlcmd = sqlcmd.Remove(0, index - 1)  ' remove leading stuff
                    index = sqlcmd.LastIndexOf("INSERT INTO `PARKS` (", StringComparison.CurrentCulture)             ' find last INSERT statement
                    If index < 0 Then index = sqlcmd.LastIndexOf("INSERT INTO PARKS (", StringComparison.CurrentCulture)             ' find last INSERT statement
                    index = sqlcmd.IndexOf(";", index, StringComparison.CurrentCulture)                              ' find end of INSERT
                    If index > 0 Then sqlcmd = sqlcmd.Remove(index + 1, sqlcmd.Length - index - 1)  ' remove trailing stuff
                    sql.CommandText = "begin"   ' start transaction for speed
                    sql.ExecuteNonQuery()
                    sql.CommandText = "DELETE FROM PARKS"   ' remove existing data
                    deleted = sql.ExecuteNonQuery()
                    'deleted = connect.Changes       ' number of records deleted
                    sql.CommandText = sqlcmd    ' execute the sql file
                    added = sql.ExecuteNonQuery()
                    sql.CommandText = "end"
                    sql.ExecuteNonQuery()
                End If
            Catch ex As SQLiteException
                MsgBox(ex.Message & vbCrLf & "Last row inserted: " & connect.LastInsertRowId & vbCrLf & sql.CommandText, vbCritical + vbOKOnly, "Import aborted")
                sql.CommandText = "rollback"
                sql.ExecuteNonQuery()
            End Try
            Application.UseWaitCursor = False
            MsgBox($"{deleted } parks deleted and {added } added", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Import complete")
            connect.Close()
        End Using
    End Sub

    Private Async Sub FindMissingPAIDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindMissingPAIDToolStripMenuItem.Click
        Await Findparks().ConfigureAwait(False)
    End Sub

    Private Async Function Findparks() As Task
        ' Attempt to find any missing PA_ID
        ' Two tests are performed
        ' 1. Look for matching name
        ' 2. Look for parks within specified distance

        Dim sql As SQLiteCommand
        Dim SQLdr As SQLiteDataReader
        Dim f As Boolean
        Dim distance As Double = 2000    ' radius of circle about point
        Dim myQueryFilter As New QueryParameters, line As String, lines As New List(Of String)
        Dim dsList As New List(Of String) From {"CAPAD_T", "CAPAD_M"}, count As Integer = 0, where As String
        Dim dBaseDataReader As System.Data.OleDb.OleDbDataReader = Nothing, query As String

        Dim missing As Integer

        Using logWriter_name As New System.IO.StreamWriter("found_name.html", False),
            logWriter_location As New System.IO.StreamWriter("found_location.html", False),
            connect As New SQLiteConnection(PARKSdb)
            logWriter_name.AutoFlush = True
            logWriter_location.AutoFlush = True
            Try
                Dim htmlFileName As String = CType(logWriter_name.BaseStream, FileStream).Name
                Dim htmlFileLocation As String = CType(logWriter_location.BaseStream, FileStream).Name
                connect.Open()  ' open database
                found = 0

                Application.UseWaitCursor = True

                lines.Clear()

                sql = connect.CreateCommand
                sql.CommandText = "SELECT count(*) AS COUNT FROM parks a left join GISmapping b using(WWFFID) WHERE lower(Status) IN ('active', 'pending') and b.dataset is null"
                SQLdr = sql.ExecuteReader()
                SQLdr.Read()
                missing = SQLdr.Item("COUNT")
                logWriter_name.WriteLine(DateTime.Now.ToString(YMDHMS) & " *** Searching for missing PA_ID<br>")
                logWriter_name.WriteLine($"There are {missing } parks with a missing PA_ID<br>")
                SQLdr.Close()
                logWriter_name.WriteLine("<table border=1>")
                logWriter_name.WriteLine("<tr><th colspan=3>ParksnPeaks</th><th colspan=3>CAPAD</th></tr>")
                logWriter_name.WriteLine("<tr><th>VKFF</th><th>Name</th><th>Area</th><th>PA_ID</th><th>DataSet</th><th>State</th></tr>")

                sql.CommandText = "SELECT * FROM parks a left join GISmapping b using(WWFFID) WHERE lower(Status) IN ('active', 'pending') and b.dataset is null ORDER BY Name"
                SQLdr = sql.ExecuteReader()
                While SQLdr.Read()
                    count += 1
                    f = False
                    Dim WWFFID As String = SQLdr.Item("WWFFID").ToString
                    Dim Name As String = SQLdr.Item("Name").ToString
                    Dim State As String = SQLdr.Item("State").ToString
                    Dim type As String = SQLdr.Item("Type").ToString
                    Dim Park_lat As Double = CDbl(SQLdr.Item("Latitude"))
                    Dim Park_lon As Double = CDbl(SQLdr.Item("Longitude"))
                    SetText(TextBox1, $"Searching for {Name } {type }: found {found }: Processed: {count }/{missing }")
                    ' find park with matching name and type
                    where = $"NAME='{SQLEscape(Name) }' AND TYPE_ABBR='{type }' ORDER BY PA_ID"  ' query parameters
                    For Each ds As String In dsList
                        query = $"SELECT * FROM {DataSets(ds).dbfTableName } WHERE {where }"
                        Try
                            Using dBaseCommand As New System.Data.OleDb.OleDbCommand(query, DataSets(ds).dbfConnection)
                                dBaseDataReader = dBaseCommand.ExecuteReader()
                                While dBaseDataReader.Read
                                    f = True
                                    logWriter_name.WriteLine($"<tr><td>{WWFFID }</td><td>{Name } {type }</td><td>{SQLdr.Item("State")}</td><td>{dBaseDataReader("PA_ID") }</td><td>{ds }</td><td>{dBaseDataReader("STATE") }</td></tr>")
                                    logWriter_name.Flush()
                                    GoTo done
                                End While
                                dBaseDataReader.Close()
                            End Using
                        Catch ex As OleDb.OleDbException
                            Dim msg As String = $"dBase command {query }; failed with code {ex.ErrorCode } - {ex.Message }"
                            MsgBox(msg & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                            Environment.Exit(0)
                        End Try
                    Next

                    ' park not found with name match - attempt location match
                    Dim prk As New MapPoint(Park_lon, Park_lat, SpatialReferences.Wgs84)
                    Dim buffer As Geometry = GeometryEngine.BufferGeodetic(prk, distance, LinearUnits.Meters)              ' park location with buffer
                    myQueryFilter.SpatialRelationship = SpatialRelationship.Intersects
                    myQueryFilter.ReturnGeometry = False
                    For Each ds As String In dsList
                        myQueryFilter.Geometry = GeometryEngine.Project(buffer, DataSets(ds).shpShapeFileTable.SpatialReference)
                        Dim parks = Await DataSets(ds).shpShapeFileTable.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)
                        For Each park In parks
                            ' got a park
                            Dim PA_ID As String = park.GetAttributeValue("PA_ID")
                            Dim Latitude As Double = CDbl(park.GetAttributeValue("LATITUDE"))
                            Dim Longitude As Double = CDbl(park.GetAttributeValue("LONGITUDE"))
                            Dim point = New MapPoint(Longitude, Latitude, New SpatialReference(DataSets(ds).shpShapeFileTable.SpatialReference.Wkid))
                            point = GeometryEngine.Project(point, SpatialReferences.Wgs84)
                            Dim park_point = New MapPoint(Park_lon, Park_lat, SpatialReferences.Wgs84)
                            Dim dist As Integer = GeometryEngine.DistanceGeodetic(point, park_point, LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic).Distance
                            line = $"<tr><td>{WWFFID }</td><td>{Name } {type }</td><td>{State }</td><td>{Latitude:f5}</td><td>{Longitude:f5}</td><td>{park.GetAttributeValue("NAME") } {park.GetAttributeValue("TYPE_ABBR") }</td><td>{PA_ID }</td><td>{ds }</td><td>{park.GetAttributeValue("STATE") }</td><td align='right'>{dist:d}</td></tr>"
                            lines.Add(line)
                            f = True
                        Next
                    Next
done:
                    If dBaseDataReader IsNot Nothing AndAlso Not dBaseDataReader.IsClosed Then dBaseDataReader.Close()
                    If (f) Then found += 1
                End While
                ' Display list of matches, sorted by distance
                lines.Sort(Function(a As String, b As String)
                               Return CompareList(a).CompareTo(CompareList(b))
                           End Function)
                logWriter_location.WriteLine(DateTime.Now.ToString(YMDHMS) & " *** Searching for missing PA_ID. Search radius " & distance / 1000 & "km<br><br>")
                logWriter_location.WriteLine("<table border=1>")
                logWriter_location.WriteLine("<tr><th colspan=5>ParksnPeaks</th><th colspan=4>CAPAD</th></tr>")
                logWriter_location.WriteLine("<tr><th>VKFF</th><th>Name</th><th>State</th><th>Latitude</th><th>Longitude</th><th>Name</th><th>PA_ID</th><th>DataSet</th><th>State</th><th>distance (m)</th></tr>")
                For Each line In lines
                    logWriter_location.WriteLine(line)
                Next
                logWriter_location.WriteLine("</table>")
                logWriter_location.WriteLine($"Done - found {found } parks<br>")
                logWriter_location.WriteLine(DateTime.Now.ToString(YMDHMS) & " *** Finished search")
                logWriter_location.Close()
                SQLdr.Close()
                connect.Close()
                logWriter_name.WriteLine("</table>")
                logWriter_name.WriteLine($"Done - found {found } parks<br>")
                logWriter_name.WriteLine(DateTime.Now.ToString(YMDHMS) & " *** Finished search")
                logWriter_name.Close()

                Application.UseWaitCursor = False
                Process.Start(htmlFileName)
                Process.Start(htmlFileLocation)
            Catch ex As Exception
                MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
            End Try
        End Using
    End Function

    Private Shared Function CompareList(s As String) As Integer
        ' extract the distance bit of the string
        Dim r As New Regex("^<tr>(?:(?:<td>.*?</td>){8})<td.*?>(\d+)</td></tr>$", RegexOptions.IgnoreCase + RegexOptions.Singleline)
        Dim m As String() = r.Split(s)
        Dim c As Integer = m.Length
        Return CInt(m(1))
    End Function

    Private Sub CrossCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CrossCheckToolStripMenuItem.Click
        ' Cross check Parks & Peaks data with CAPAD
        ' For speed, CAPAD data is taken from .dbf file
        Dim connect As SQLiteConnection ' declare the connection
        Dim sql As SQLiteCommand
        Dim SQLdr As SQLiteDataReader
        Dim errors As Integer, count As Integer
        Dim pp_name As String, pp_type As String, pp_state As String
        Dim capad_name As String, capad_type As String, capad_state As String, wwff As String
        Dim mismatch As Boolean, found As Boolean
        Dim states As String() = {"xxx", "ACT", "NSW", "VIC", "QLD", "SA", "WA", "TAS", "NT", "EXT"}
        Dim areas As String() = {"VK0", "VK1", "VK2", "VK3", "VK4", "VK5", "VK6", "VK7", "VK8", "VK9"}
        Dim active As Integer, noGISID As Integer
        Dim ignores() = {" National Park", " Coastal Park"}      ' park types to remove
        Dim rgx As New Regex("\s{0,5}\(.*\)")        ' regexp pattern for something in brackets
        Dim trimChars() As Char = Array.Empty(Of Char)()            ' chars to trim (all whitespace)
        Dim htmlWriter As New System.IO.StreamWriter("crosscheck.html", False)
        Dim htmlFile As String = CType(htmlWriter.BaseStream, FileStream).Name
        Dim dsList As New List(Of String) From {"CAPAD_T", "CAPAD_M"}      ' list of shapefiles to search for park
        Dim Dataset As String
        Dim ParkData As NameValueCollection, GISID As String
        Dim dBaseCommand As System.Data.OleDb.OleDbCommand, dBaseDataReader As System.Data.OleDb.OleDbDataReader

        Application.UseWaitCursor = True
        htmlWriter.WriteLine($"Cross-check of CAPAD data against Parks&Peaks data using CAPAD files run {Now }<br>")
        For Each ds In DataSets
            htmlWriter.WriteLine($"{ds.Key } contains {ds.Value.shpShapeFileTable.NumberOfFeatures } objects<br>")
        Next
        htmlWriter.WriteLine("<br>Some park names have the park type in the name. The park types below have been removed from the name<br>")
        htmlWriter.WriteLine("Also, anything in brackets, usually 'formerly . . ' is ignored<br>")
        For Each ignore In ignores
            htmlWriter.WriteLine($"{ignore }<br>")
        Next
        htmlWriter.WriteLine("<br><br>")

        Dim myQueryFilter As New QueryParameters  ' create query

        connect = New SQLiteConnection(PARKSdb)
        connect.Open()  ' open database
        sql = connect.CreateCommand

        ' Get basic statistics
        sql.CommandText = "SELECT COUNT(*) AS COUNT FROM parks WHERE lower(Status) IN ('active','Active', 'pending')"
        SQLdr = sql.ExecuteReader()
        SQLdr.Read()
        active = SQLdr.Item("COUNT")
        SQLdr.Close()

        sql.CommandText = "SELECT count(*) AS COUNT FROM parks a left join GISmapping b using(WWFFID) WHERE lower(Status) IN ('active', 'pending') and b.dataset is null"
        SQLdr = sql.ExecuteReader()
        SQLdr.Read()
        noGISID = SQLdr.Item("COUNT")
        SQLdr.Close()

        Dim percent = noGISID / active * 100
        htmlWriter.WriteLine($"There are {active } active/pending parks, with {noGISID } missing a GISID, i.e. {percent:##.00}%<br><br>")

        ' Check VKFFID unique
        errors = 0
        sql.CommandText = "SELECT WWFFID, count(WWFFID) FROM parks WHERE WWFFID AND lower(Status) IN ('active', 'pending') GROUP BY WWFFID HAVING (COUNT(WWFFID) > 1)"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            htmlWriter.WriteLine($"Duplicate WWFFID {SQLdr.Item("WWFFID")}<br>")
            htmlWriter.Flush()
            errors += 1
        End While
        SQLdr.Close()
        htmlWriter.WriteLine($"There are {errors } duplicate WWFFID<br>")

        ' Check VKFFID properly formed
        errors = 0
        sql.CommandText = "SELECT WWFFID FROM parks WHERE NOT (WWFFID REGEXP '^(VK|ZL)FF\-[0-9]{4}$') AND lower(Status) IN ('active', 'pending')"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            htmlWriter.WriteLine($"Malformed WWFFID {SQLdr.Item("WWFFID")}<br>")
            htmlWriter.Flush()
            errors += 1
        End While
        SQLdr.Close()
        htmlWriter.WriteLine($"There are {errors } malformed WWFFID<br>")

        ' Check SANPCPAID unique
        errors = 0
        sql.CommandText = "SELECT SANPCPAID, count(SANPCPAID) FROM parks WHERE WWFFID LIKE 'VKFF-%' AND lower(Status) IN ('active', 'Active', 'pending') AND SANPCPAID<>'' GROUP BY SANPCPAID HAVING (COUNT(SANPCPAID) > 1)"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            htmlWriter.WriteLine($"Duplicate SANPCPAID {SQLdr.Item("SANPCPAID")}<br>")
            htmlWriter.Flush()
            errors += 1
        End While
        SQLdr.Close()
        htmlWriter.WriteLine($"There are {errors } duplicate SANPCPAID<br>")
        ' Check SANPCPAID properly formatted
        errors = 0
        sql.CommandText = "SELECT SANPCPAID FROM parks WHERE SANPCPAID <>'' AND NOT (SANPCPAID REGEXP '^5[CN]P-\d{3}$') AND lower(Status) IN ('active', 'pending')"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            htmlWriter.WriteLine($"Malformed SANPCPAID {SQLdr.Item("SANPCPAID")}<br>")
            htmlWriter.Flush()
            errors += 1
        End While
        SQLdr.Close()
        htmlWriter.WriteLine($"There are {errors } malformed SANPCPAID<br>")

        ' Check KRMNPAID unique
        errors = 0
        sql.CommandText = "SELECT KRMNPAID, count(KRMNPAID) FROM parks WHERE WWFFID LIKE 'VKFF-%' AND lower(Status) IN ('active', 'Active', 'pending') AND KRMNPAID<>'' GROUP BY KRMNPAID HAVING (COUNT(KRMNPAID) > 1)"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            htmlWriter.WriteLine($"Duplicate KRMNPAID {SQLdr.Item("KRMNPAID")}<br>")
            htmlWriter.Flush()
            errors += 1
        End While
        SQLdr.Close()
        htmlWriter.WriteLine($"There are {errors } duplicate KRMNPAID<br>")
        ' Check KRMNPAID properly formatted
        errors = 0
        sql.CommandText = "SELECT KRMNPAID FROM parks WHERE KRMNPAID <>'' AND NOT (KRMNPAID REGEXP '^3NP-\d{3}$') AND lower(Status) IN ('active', 'pending')"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            htmlWriter.WriteLine($"Malformed KRMNPAID {SQLdr.Item("KRMNPAID")}<br>")
            htmlWriter.Flush()
            errors += 1
        End While
        SQLdr.Close()
        htmlWriter.WriteLine($"There are {errors } malformed KRMNPAID<br>")
        htmlWriter.Flush()

        htmlWriter.WriteLine("<table border=1><tr><th>Dataset</th><th>PA_ID</th><th>CAPAD</th><th>State</th><th>Parks&amp;Peaks</th><th>Area</th><th>WWFF</th><th>Status</th></tr>")
        ' Check Name, Type and State
        ' Create list of parks that have any GISID
        sql.CommandText = "SELECT * FROM GISmapping JOIN parks USING(WWFFID) WHERE parks.WWFFID LIKE 'VKFF-%' AND GISmapping.Dataset IN (""CAPAD_T"", ""CAPAD_M"") AND lower(Status) IN ('active', 'pending')  group by parks.WWFFID ORDER BY Name"
        errors = 0
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            count += 1
            ParkData = GetParkData(SQLdr)
            wwff = ParkData("WWFFID")
            pp_name = ParkData("Name")
            pp_type = ParkData("Type")
            pp_state = ParkData("State")
            GISID = ParkData("GISIDList").Split({","c})(0)      ' get first GISID only
            Dataset = ParkData("Dataset")
            Dim area_indx As Integer = Array.IndexOf(areas, pp_state)
            SetText(TextBox1, $"Checking {wwff,8} - {pp_name,-30} {pp_type,-20} : errors {errors,5}/{count,5}")

            ' find park with matching name and type
            found = False
            capad_name = ""
            capad_state = ""
            capad_type = ""
            ' Look for park data in .dbf file (for speed)
            Dim where = $"PA_ID='{GISID }'"
            dBaseCommand = New System.Data.OleDb.OleDbCommand($"SELECT * FROM {DataSets(Dataset).dbfTableName } WHERE {where }", DataSets(Dataset).dbfConnection)
            dBaseDataReader = dBaseCommand.ExecuteReader()
            While dBaseDataReader.Read
                mismatch = False
                found = True
                capad_name = dBaseDataReader.Item("NAME").ToString.Trim(trimChars)
                For Each ignore In ignores
                    capad_name = Replace(capad_name, ignore, "")      ' remove type from name
                Next
                capad_name = rgx.Replace(capad_name, "")    ' remove anything in brackets
                pp_name = rgx.Replace(pp_name, "")    ' remove anything in brackets
                capad_type = dBaseDataReader.Item("TYPE_ABBR").ToString.Trim(trimChars)
                capad_state = dBaseDataReader.Item("STATE").ToString.Trim(trimChars)
                Dim state_indx As Integer = Array.IndexOf(states, capad_state.Replace("JBT", "NSW"))
                If pp_name <> capad_name Then
                    mismatch = True
                    pp_name = "<b><font color='red'>" & pp_name & "</font></b>"
                End If
                If pp_type <> capad_type Then
                    mismatch = True
                    pp_type = "<b><font color='red'>" & pp_type & "</font></b>"
                End If
                If state_indx <> area_indx Then
                    mismatch = True
                    pp_state = "<b><font color='red'>" & pp_state & "</font></b>"
                End If
            End While
            dBaseDataReader.Close()
            dBaseCommand.Dispose()
            If Not found Then
                capad_name = "<b><font color='red'>Not Found in CAPAD</font></b>"
                capad_type = ""
                capad_state = ""
            End If
            If Not found Or mismatch Then
                errors += 1
                htmlWriter.WriteLine($"<tr><td>{Dataset }</td><td>{GISID }</td><td>{capad_name } {capad_type }</td><td>{capad_state }</td><td>{pp_name } {pp_type }</td><td>{pp_state }</td><td>{wwff }</td><td>{ParkData("Status") }</td></tr>")
                htmlWriter.Flush()
            End If
            Application.DoEvents()
        End While

        SQLdr.Close()
        connect.Close()
        htmlWriter.WriteLine($"</table><br>Total of {errors } errors in {count } parks")
        htmlWriter.Close()
        SetText(TextBox1, $"Done : errors {errors }/{count }")
        Application.UseWaitCursor = False
        Process.Start(htmlFile)
    End Sub

    Private Async Sub CreatedbfFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreatedbfFileToolStripMenuItem.Click
        Await CreateDBF("CAPAD_M").ConfigureAwait(True)
        Await CreateDBF("CAPAD_T").ConfigureAwait(True)
    End Sub

    Private Async Function CreateDBF(ds As String) As Task(Of Integer)
        ' Create a dbf file for dataset
        ' The dbf files for CAPAD 2020 seem to be broken in a way I don't understand. Every time they are accessed, it crashes with unexpected driver error
        ' This routine recreates the dbf files from the shapefile
        Dim myQueryFilter As New QueryParameters  ' create query
        Dim fieldList As New List(Of String), valueList As New List(Of String)
        Dim table As String = $"{IO.Path.GetFileNameWithoutExtension(DataSets(ds).shpFileName) }"
        Dim featureData As New Dictionary(Of String, String), value As String, fields As New List(Of String), records As Integer
        ' Can't work out type of field name, so need to hardwire them by name
        Dim FieldTypes As New Dictionary(Of String, String) From {
            {"OBJECTID", "NUMERIC(10,0)"},
            {"PA_ID", "CHAR(20)"},
            {"PA_PID", "CHAR(20)"},
            {"NAME", "CHAR(254)"},
            {"TYPE", "CHAR(60)"},
            {"TYPE_ABBR", "CHAR(10)"},
            {"IUCN", "CHAR(5)"},
            {"NRS_PA", "CHAR(5)"},
            {"GAZ_AREA", "NUMERIC(24,10)"},
            {"GIS_AREA", "NUMERIC(24,10)"},
            {"GAZ_DATE", "CHAR(20)"},
            {"LATEST_GAZ", "CHAR(20)"},
            {"STATE", "CHAR(4)"},
            {"AUTHORITY", "CHAR(15)"},
            {"DATASOURCE", "CHAR(20)"},
            {"GOVERNANCE", "CHAR(3)"},
            {"COMMENTS", "CHAR(120)"},
            {"ENVIRON", "CHAR(3)"},
            {"OVERLAP", "CHAR(3)"},
            {"MGT_PLAN", "CHAR(3)"},
            {"RES_NUMBER", "CHAR(15)"},
            {"EPBC", "CHAR(15)"},
            {"LONGITUDE", "NUMERIC(24,10)"},
            {"LATITUDE", "NUMERIC(24,10)"},
            {"SHAPE_LENG", "NUMERIC(24,10)"},
            {"SHAPE_Leng", "NUMERIC(24,10)"},
            {"SHAPE_AREA", "NUMERIC(24,10)"},
            {"SHAPE_Area", "NUMERIC(24,10)"},
            {"NRS_MPA", "CHAR(5)"},
            {"ZONE_TYPE", "CHAR(150)"}
         }
        ' Remove existing datafile
        Dim datafile As String = $"F:\GIS data\CAPAD\{table }.DBF"
        If System.IO.File.Exists(datafile) = True Then
            System.IO.File.Delete(datafile)
        End If
        Dim BaseConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBase IV"
        Dim ConnectionString = $"{BaseConnectionString };Data Source=F:\GIS data\CAPAD\;"
        Using cn As New System.Data.OleDb.OleDbConnection(ConnectionString)
            cn.Open()

            ' fetch all features in shapefile
            Dim PK As New OrderBy("OBJECTID", Esri.ArcGISRuntime.Data.SortOrder.Ascending)  ' order by primary key (OBJECTID)
            With myQueryFilter
                .ReturnGeometry = False
                .WhereClause = "1=1"
                .OrderByFields.Add(PK)
            End With
            DataSets(ds).shpFragments = Await DataSets(ds).shpShapeFileTable.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False) ' run query
            ' Create the table
            ' get list of all fields
            fieldList.Clear()
            For Each field In DataSets(ds).shpShapeFileTable.Fields
                fieldList.Add(field.Name)
            Next
            fieldList.RemoveAt(0)   ' remove unwanted entry
            ' make list of fields with type
            fields.Clear()
            For Each field In fieldList
                fields.Add($"{field } {FieldTypes(field) }")
            Next
            Dim f As String = String.Join(",", fields.ToArray)
            Dim cmd As String = $"Create Table {table } ({f });"
            Using cmdCreate As New OleDb.OleDbCommand(cmd, cn)
                cmdCreate.ExecuteNonQuery()
            End Using
            records = DataSets(ds).shpShapeFileTable.NumberOfFeatures     ' total number of features
            For Each fragment As Esri.ArcGISRuntime.Data.Feature In DataSets(ds).shpFragments
                featureData.Clear()
                For Each attr In fragment.Attributes
                    featureData.Add(attr.Key, attr.Value)
                Next
                valueList.Clear()
                For Each f In fieldList
                    value = featureData(f)
                    value = value.Replace("""", """""")    ' escape double quotes
                    value = value.Replace("''''''''", "'")    ' bizarrely, 8 single quotes are used to represent a single quote
                    If ds = "CAPAD_M" And f = "PA_ID" Then value = $"{CInt(value):d3}"      ' Maritime PA_ID is 3 digit leading zero
                    ' If datatype is CHAR, needs to be quoted
                    If FieldTypes(f).Contains("CHAR") Then
                        valueList.Add($"""{value }""")
                    Else
                        valueList.Add(value)
                    End If
                Next
                SetText(TextBox1, $"Inserting record {table } - {featureData("OBJECTID") } of {records }")
                Application.DoEvents()
                Try
                    cmd = $"INSERT INTO {table } VALUES ({String.Join(",", valueList.ToArray) });"
                    Using cmdInsert As New OleDb.OleDbCommand(cmd, cn)
                        cmdInsert.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                End Try
            Next
            cn.Close()
        End Using
        Return 1
    End Function


    '=======================================================================================================
    ' Items on SOTA menu
    '=======================================================================================================
    Private Async Sub FindParksForSOTASummitsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindParksForSOTASummitsToolStripMenuItem.Click
        ' Scan through all SOTA summits and identify those in a park
        Dim state As String

        found = 0 : count = 0
        With Form3
            If .ShowDialog() = DialogResult.OK Then
                Dim sqlWriter As New System.IO.StreamWriter("SOTAinWWFF.sql", False)
                Application.UseWaitCursor = True
                sqlWriter.WriteLine("/* Summits within parks. File created {0} */", DateTime.Now.ToString(YMDHMS))
                For i = 0 To .ctrls.Count - 2
                    Dim box As KeyValuePair(Of String, CheckBox) = .ctrls.Item(i)    ' extract the checkbox
                    If box.Value.Checked Or .ctrls.Item(.ctrls.Count - 1).Value.Checked Then
                        state = box.Key
                        Await FindParksForSOTASummitsState(sqlWriter, state).ConfigureAwait(False)
                    End If
                Next
                sqlWriter.Close()
                SetText(TextBox1, $"Done : found {found } summits in {count } parks: Results in SOTAinWWFF.sql")
                Application.UseWaitCursor = False
            End If
        End With
    End Sub

    Private Async Function FindParksForSOTASummitsState(sqlwriter As StreamWriter, state As String) As Task
        ' Scan through all SOTA summits and identify those in a park
        ' Creates SOTAinWWFF.sql for export
        ' Updates local table SummitsInParks
        Const distance As Double = 100    ' radius of circle about point
        Dim SOTA_connect As SQLiteConnection, PARKS_connect As SQLiteConnection ' declare the connections
        Dim SOTA_sqlR As SQLiteCommand, SOTA_sqlW As SQLiteCommand, PARKS_sql As SQLiteCommand
        Dim SOTA_SQLdr As SQLiteDataReader, PARKS_SQLdr As SQLiteDataReader
        Dim myQueryFilter As New QueryParameters
        Dim WithIn As String
        Dim PA_ID As String
        Dim vkff As String, name As String, Type As String, summit As MapPoint
        Dim values As New List(Of String)
        Dim dsList As New List(Of String)      ' list of shapefiles to search for park
        Dim PAIDList As New List(Of String)     ' list of PA_ID so we can filter duplicates

        Try
            values.Clear()
            ' Connect to SOTA database
            SOTA_connect = New SQLiteConnection(SOTAdb)
            SOTA_connect.Open()  ' open database
            SOTA_sqlR = SOTA_connect.CreateCommand      ' to read
            SOTA_sqlW = SOTA_connect.CreateCommand      ' to write
            SOTA_sqlW.CommandText = "BEGIN TRANSACTION"   ' start transaction
            SOTA_sqlW.ExecuteNonQuery()
            ' Delete any existing data for this state
            SOTA_sqlW.CommandText = $"DELETE FROM `SummitsInParks` WHERE `State`='{state }'"
            SOTA_sqlW.ExecuteNonQuery()
            ' Prepare insertion statement for speed
            SOTA_sqlW.CommandText = "INSERT INTO `SummitsInParks` (`SummitCode`,`WWFFID`,`State`,`WithIn`) VALUES (@SummitCode,@vkff,@state,@WithIn)"
            SOTA_sqlW.Prepare()
            ' Connect to PARKS database
            PARKS_connect = New SQLiteConnection(PARKSdb)
            PARKS_connect.Open()  ' open database
            PARKS_sql = PARKS_connect.CreateCommand

            Dim prefix As String = Strings.Left(state, 2)
            ' make list of shapefiles to look in
            Select Case prefix
                Case "VK" : dsList = New List(Of String) From {"CAPAD_T"}
                Case "ZL" : dsList = New List(Of String) From {"ZL"}
            End Select
            If state = "VK3" Then dsList.Add("VIC_PARKS")       ' add regional parks
            ' Find all valid SOTA summits in this state
            SOTA_sqlR.CommandText = $"Select * FROM SOTA where Like('{state }%',SummitCode)=1 ORDER BY SummitCode"
            SOTA_SQLdr = SOTA_sqlR.ExecuteReader()
            While SOTA_SQLdr.Read()
                ' SQLite doesn't handle dates well so have to date filter after SELECT
                Dim ValidFrom As Date = SOTA_SQLdr.Item("ValidFrom").ToString
                Dim ValidTo As Date = SOTA_SQLdr.Item("ValidTo").ToString
                If (Now() >= ValidFrom And Now() <= ValidTo) Then
                    Dim SummitCode As String = SOTA_SQLdr.Item("SummitCode").ToString
                    Dim SummitName As String = SOTA_SQLdr.Item("SummitName").ToString
                    Dim Latitude As Double = CDbl(SOTA_SQLdr.Item("GridRef2").ToString)
                    Dim Longitude As Double = CDbl(SOTA_SQLdr.Item("Gridref1").ToString)
                    count += 1
                    SetText(TextBox1, $"Looking for {SummitCode } {SummitName,-50} : found {found } summits in {count } parks")
                    Application.DoEvents()        ' update text box
                    ' Find park(s)
                    summit = New MapPoint(Longitude, Latitude, SpatialReferences.Wgs84)     ' location of summit
                    Dim buffer As Geometry = GeometryEngine.BufferGeodetic(summit, distance, LinearUnits.Meters)              ' summit location with buffer
                    PAIDList.Clear()
                    For Each ds In dsList
                        myQueryFilter.Geometry = GeometryEngine.Project(buffer, DataSets(ds).shpShapeFileTable.SpatialReference)
                        myQueryFilter.WhereClause = ""
                        myQueryFilter.ReturnGeometry = True
                        myQueryFilter.OutSpatialReference = DataSets(ds).shpShapeFileTable.SpatialReference
                        ' Now look for parks where buffer intersects
                        myQueryFilter.SpatialRelationship = SpatialRelationship.Intersects
                        DataSets(ds).shpFragments = Await DataSets(ds).shpShapeFileTable.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)
                        For Each park As Feature In DataSets(ds).shpFragments
                            PA_ID = park.GetAttributeValue(DataSets(ds).shpIDField)
                            name = park.GetAttributeValue(DataSets(ds).shpNameField)
                            Type = park.GetAttributeValue(DataSets(ds).shpTypeField)
                            If Not PAIDList.Contains(PA_ID) Then
                                PAIDList.Add(PA_ID)
                                WithIn = If(GeometryEngine.Contains(park.Geometry, myQueryFilter.Geometry), "Y", "N")
                                PARKS_sql.CommandText = $"SELECT * FROM GISmapping WHERE GISID='{PA_ID }' AND DataSet='{ds }'"  ' find matching park
                                PARKS_SQLdr = PARKS_sql.ExecuteReader()
                                If PARKS_SQLdr.Read() Then
                                    vkff = PARKS_SQLdr.Item("WWFFID").ToString
                                    ' Add summit/park to database
                                    SOTA_sqlW.Parameters.Clear()
                                    SOTA_sqlW.Parameters.AddWithValue("@SummitCode", SummitCode)
                                    SOTA_sqlW.Parameters.AddWithValue("@vkff", vkff)
                                    SOTA_sqlW.Parameters.AddWithValue("@State", state)
                                    SOTA_sqlW.Parameters.AddWithValue("@WithIn", WithIn)
                                    SOTA_sqlW.ExecuteNonQuery()
                                    values.Add($"('{SummitCode }','{vkff }','{WithIn }')")
                                    found += 1
                                End If
                                PARKS_SQLdr.Close()
                            End If
                        Next
                    Next
                End If
            End While
            SOTA_sqlW.Parameters.Clear()
            SOTA_sqlW.CommandText = "COMMIT"   ' commit the changes
            SOTA_sqlW.ExecuteNonQuery()
            SOTA_connect.Close()
            PARKS_connect.Close()
            ' Create the SQL data for this state
            If values.Any Then
                sqlwriter.WriteLine("/*")
                sqlwriter.WriteLine($"  Data for {state }")
                sqlwriter.WriteLine("*/")
                sqlwriter.WriteLine($"DELETE FROM `SOTAinWWFF` WHERE `SOTARef` LIKE '{state }%';")
                sqlwriter.WriteLine("INSERT INTO `SOTAinWWFF` (`SOTARef`,`WWFFID`,`within`)")
                sqlwriter.WriteLine("VALUES (")
                sqlwriter.WriteLine(Join(values.ToArray, "," & vbCrLf))
                sqlwriter.WriteLine(");")
                sqlwriter.Flush()
            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
        End Try
    End Function

    Private Async Sub CreateFragmentsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateFragmentsToolStripMenuItem.Click
        ' Create a list of VKFF and GISID
        Dim connect As SQLiteConnection ' declare the connection
        Dim sql As SQLiteCommand, insert As SQLiteCommand
        Dim SQLdr As SQLiteDataReader
        Dim count As Integer = 0
        Dim myQueryFilter As New QueryParameters, GISID As String, WWFFID As String, where As String, DataSet As String
        Dim GIS_AREA As Double
        Dim logWriter As New System.IO.StreamWriter("CAPAD.txt", False)

        connect = New SQLiteConnection(PARKSdb)
        connect.Open()  ' open database
        sql = connect.CreateCommand
        insert = connect.CreateCommand
        sql.CommandText = "DELETE FROM CAPAD"   ' remove existing data
        sql.ExecuteNonQuery()
        sql = connect.CreateCommand
        sql.CommandText = "Select WWFFID, GISID, DataSet FROM parks WHERE GISID <> '' AND lower(Status) IN ('active', 'pending') ORDER BY WWFFID"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            GISID = SQLdr.Item("GISID").ToString
            WWFFID = SQLdr.Item("WWFFID").ToString
            DataSet = SQLdr.Item("DataSet").ToString
            where = DataSets(DataSet).BuildWhere(GISID)
            myQueryFilter.WhereClause = where    ' query parameters
            myQueryFilter.ReturnGeometry = False
            DataSets(DataSet).shpFragments = Await DataSets(DataSet).shpShapeFileTable.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)           ' run query
            If DataSets(DataSet).shpFragments IsNot Nothing And DataSets(DataSet).shpFragments.Any Then
                SetText(TextBox1, $"Processing {WWFFID } - {count }")
                ' Aggregate all the instances of this PA_ID
                Dim centroid As MapPoint = DataSets(DataSet).Centroid
                GIS_AREA = DataSets(DataSet).Area

                ' Now insert records for fragments
                For Each fragment In DataSets(DataSet).shpFragments
                    Dim values = $"'{WWFFID }','{fragment.GetAttributeValue("PA_ID") }','{fragment.GetAttributeValue("TYPE_ABBR") }','{fragment.GetAttributeValue("IUCN") }',{centroid.X },{centroid.Y },{GIS_AREA }"
                    insert.CommandText = $"INSERT INTO CAPAD VALUES ({values })"
                    Try
                        insert.ExecuteNonQuery()
                    Catch Ex As SQLiteException
                        Dim msg As String = "SQLite command " & insert.CommandText & "; failed with code " & Ex.ErrorCode & " - " & Ex.Message
                        logWriter.WriteLine(msg)
                        logWriter.Flush()
                    End Try
                    count += 1
                Next
            End If
        End While
        SQLdr.Close()
        connect.Close()
        logWriter.Close()
        MsgBox(count & " fragments created", vbInformation + vbOKOnly, "Done")
    End Sub

    Shared Function SQLEscape(st As String) As String
        ' escape a string for SQL
        Contract.Requires(Not String.IsNullOrEmpty(st), "Illegal string")
        Return st.Replace("'", "''")
    End Function

    Private Async Sub CrossCheckCoordinatesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CrossCheckCoordinatesToolStripMenuItem.Click
        ' Cross check distance of centroids of PnP data with CAPAD data
        ' For speed, use the dbf metadata file
        Const DistThreshold = 500, AreaThreshold = 1    ' thresholds above which we produce corrections
        Dim myQueryFilter As New QueryParameters  ' create query
        Dim GISID As String, WWFFID As String
        Dim centroid As MapPoint = Nothing, name As String, Type As String, DeltaArea As Integer
        Dim Latitude As Double, Longitude As Double, Area As Double
        Dim CAPAD_latitude As Double, CAPAD_longitude As Double, CAPAD_area As Double
        Dim sql As SQLiteCommand, where As String
        Dim SQLdr As SQLiteDataReader
        Dim count As Integer = 0, updateCount As Integer = 0, dist As Integer
        Dim DataSet As String
        Dim ParkData As NameValueCollection
        Dim Updates As New List(Of String)    ' fields to update

        Using connect As New SQLiteConnection(PARKSdb),
            htmlWriter As New System.IO.StreamWriter(String.Format("centroid-{0}.html", DateTime.Now.ToString("yyMMdd-HHmm")), False),
            sqlWriter As New System.IO.StreamWriter(String.Format("centroid-{0}.sql", DateTime.Now.ToString("yyMMdd-HHmm")), False)

            Try
                Dim htmlFile As String = CType(htmlWriter.BaseStream, FileStream).Name
                htmlWriter.WriteLine($"Script started {Now:u}<br>")
                htmlWriter.WriteLine("<table border=1>")
                htmlWriter.WriteLine("<tr><th colspan=5>ParksnPeaks</th><th colspan=7>Shapefile data</th></tr>")
                htmlWriter.WriteLine("<tr><th>VKFF</th><th>Name</th><th>Latitude</th><th>Longitude</th><th>Area (ha)</th><th>PA_ID</th><th>DataSet</th><th>Latitude</th><th>Longitude</th><th>Area (ha)</th><th>Distance (m)</th><th>&Delta; Area</th></tr>")
                sqlWriter.WriteLine("/* SQL file to add corrections to park centroid as area */")
                sqlWriter.WriteLine($"/* ∆ Distance threshold {DistThreshold }m, ∆ Area threshold {AreaThreshold }ha */")
                sqlWriter.WriteLine($"/* File produced {Now:u} */")

                connect.Open()  ' open database
                sql = connect.CreateCommand
                ' can only do metadata check on CAPAD parks
                sql.CommandText = "Select * FROM parks JOIN GISmapping USING(WWFFID) WHERE lower(Status) IN ('active', 'pending') ORDER BY WWFFID"
                SQLdr = sql.ExecuteReader()
                While SQLdr.Read()
                    count += 1
                    ParkData = GetParkData(SQLdr)
                    WWFFID = ParkData("WWFFID")
                    If ParkData.AllKeys.Contains("GISIDList") Then
                        GISID = ParkData("GISIDList")
                        DataSet = ParkData("DataSet")
                        Area = ParkData("Area")
                        name = ParkData("NAME")
                        Type = ParkData("Type")
                        Longitude = CDbl(ParkData("Longitude"))
                        Latitude = CDbl(ParkData("Latitude"))
                        SetText(TextBox1, $"Processing {name } {Type } - {count }")
                        Application.DoEvents()
                        If Not DataSets.ContainsKey(DataSet) Then
                            MessageBox.Show($"Unknown dataset <{DataSet }> when processing {WWFFID }", "Unknown dataset", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Else
                            ' Now get CAPAD data from dbf file
                            where = DataSets(DataSet).BuildWhere(ParkData("GISIDListQuoted"))
                            Dim Query As String = $"SELECT * FROM {DataSets(DataSet).dbfTableName } WHERE {where }"
                            Using dBaseCommand As New System.Data.OleDb.OleDbCommand(Query, DataSets(DataSet).dbfConnection)
                                Dim dBaseDataReader = dBaseCommand.ExecuteReader()
                                If dBaseDataReader.HasRows Then
                                    dBaseDataReader.Read()
                                    If DataSets(DataSet).HasXYmetaData Then
                                        ' dbf file contains lat/lon
                                        CAPAD_latitude = dBaseDataReader("LATITUDE")
                                        CAPAD_longitude = dBaseDataReader("LONGITUDE")
                                        centroid = New MapPoint(CAPAD_longitude, CAPAD_latitude, SpatialReferences.Wgs84)
                                    Else
                                        ' dbf file does not contain lat/lon. Need to calculate it from shape
                                        myQueryFilter.WhereClause = where    ' query parameters
                                        DataSets(DataSet).shpFragments = Await DataSets(DataSet).shpShapeFileTable.QueryFeaturesAsync(myQueryFilter)           ' run query
                                        If DataSets(DataSet).shpFragments IsNot Nothing And DataSets(DataSet).shpFragments.Count > 0 Then
                                            count += 1
                                            centroid = DataSets(DataSet).Centroid
                                            centroid = GeometryEngine.Project(centroid, SpatialReferences.Wgs84)    ' convert to WGS84
                                        Else
                                            MsgBox($"No shape fragments found for {WWFFID }", vbAbort + vbOKOnly, "Missing CAPAD data")
                                        End If
                                    End If
                                    CAPAD_area = CDbl(dBaseDataReader(DataSets(DataSet).AreaField)) / CDbl(DataSets(DataSet).dbfAreaScale)
                                    dBaseDataReader.Close()
                                    dist = GeometryEngine.DistanceGeodetic(New MapPoint(Longitude, Latitude, SpatialReferences.Wgs84), centroid, LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic).Distance
                                    DeltaArea = Abs(CInt(Area - CAPAD_area))

                                    htmlWriter.WriteLine($"<tr><td>{WWFFID }</td><td>{name } {Type }</td><td>{Latitude }</td><td>{Longitude }</td><td align='right'>{Area }</td><td>{GISID }</td><td>{DataSet }</td><td>{centroid.Y:f5}</td><td>{centroid.X:f5}</td><td align='right'>{CAPAD_area:f2}</td><td align='right'>{dist }</td><td align='right'>{DeltaArea }</td></tr>")
                                    htmlWriter.Flush()
                                    ' Now do SQL corrections
                                    Updates.Clear()
                                    If dist >= DistThreshold Then
                                        Updates.Add($"`Latitude`={centroid.Y:f5}")
                                        Updates.Add($"`Longitude`={centroid.X:f5}")
                                    End If
                                    If DeltaArea >= AreaThreshold Then Updates.Add($"`Area`={CAPAD_area:f1}")
                                    ' produce SQL
                                    If Updates.Any Then
                                        updateCount += 1
                                        sqlWriter.WriteLine($"/* WWFFID {WWFFID } - ∆ Distance {dist }m, ∆ Area {DeltaArea }ha */")  ' comment
                                        sqlWriter.WriteLine($"UPDATE `PARKS` SET {Join(Updates.ToArray, ",") } WHERE `WWFFID`='{WWFFID }';")
                                        sqlWriter.Flush()
                                    End If
                                End If
                            End Using
                        End If
                    End If
                End While
done:
                SQLdr.Close()
                connect.Close()
                htmlWriter.WriteLine("</table>")
                htmlWriter.WriteLine($"Script finished {Now:u}<br>")
                htmlWriter.Close()
                sqlWriter.WriteLine($"/* Updates to {updateCount } parks")
                sqlWriter.Close()
                MsgBox(count & " parks checked", vbInformation + vbOKOnly, "Done")
                Process.Start(htmlFile)
            Catch ex As Exception
                MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
            End Try
        End Using
    End Sub

    Private Sub MakeIndexFilesToolStripMenuItem_Click(sender As Object, E As EventArgs) Handles MakeIndexFilesToolStripMenuItem.Click
        ' Make index files for vk1, vk2, ...
        Dim total As Integer = 0, count As Integer, state As String, target_local As String, target_remote As String, msg As String
        Dim BaseFilename As String, BaseFolder As String, NetworkList As New List(Of String)
        Dim fi As FileInfo, item As String, fn As String, id As String, files As New List(Of StreamWriter)
        ' Links generated can either be to production files (url) or local files for testing
        Dim extent As Envelope = Nothing        ' extent of state

        Dim logWriter As New System.IO.StreamWriter(Application.StartupPath & "\files\log.txt", False)

        For i = LBound(states) To UBound(states)
            state = states(i)
            target_local = state & "\"
            target_remote = PnPurl & state & "/"
            count = 0
            extent = Nothing

            BaseFilename = Application.StartupPath & "\files\VKFF-" & state
            BaseFolder = Application.StartupPath & "\files\" & state & "\"
            Dim RemoteindxWriter As New System.IO.StreamWriter(BaseFilename & ".kml", False)
            Dim LocalindxWriter As New System.IO.StreamWriter(BaseFilename & "_local.kml", False)

            files.Clear()
            files.Add(RemoteindxWriter)
            files.Add(LocalindxWriter)
            ' Write KML header
            For Each sr In files
                sr.WriteLine(KMLheader)
                sr.WriteLine($"<description>VKFF index file for {state } created by Marc Hillman - VK3OHM</description>")
            Next
            ' Loop through folder building index
            NetworkList.Clear()
            For Each VKFFfile In Directory.GetFiles(BaseFolder, "*.kml")
                count += 1
                Try
                    Dim doc As XDocument = XDocument.Load(VKFFfile)    ' read the XML
                    Dim ns As XNamespace = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
                    Dim nsmgr As New XmlNamespaceManager(New NameTable())
                    nsmgr.AddNamespace("x", ns.NamespaceName)
                    Name = doc.XPathSelectElement("//x:Placemark/x:name", nsmgr).Value.ToString     ' park name
                    fi = My.Computer.FileSystem.GetFileInfo(VKFFfile)   ' get file details
                    id = Path.GetFileNameWithoutExtension(fi.Name)
                    fn = id & ".kmz" ' filename
                    NetworkList.Add(String.Format("<NetworkLink id=""{0}""><name>{1}</name><Link><href>{{0}}{2}</href></Link></NetworkLink>", id, XMLencode(Replace(Replace(Name, "{", "{{"), "}", "}}")), fn))

                    Try
                        ' Extract the bounding box details, if there are any
                        Dim north As Double = doc.XPathSelectElement("//x:ExtendedData/x:Data[@name=""north""]/x:value", nsmgr).Value.ToString
                        Dim south As Double = doc.XPathSelectElement("//x:ExtendedData/x:Data[@name=""south""]/x:value", nsmgr).Value.ToString
                        Dim east As Double = doc.XPathSelectElement("//x:ExtendedData/x:Data[@name=""east""]/x:value", nsmgr).Value.ToString
                        Dim west As Double = doc.XPathSelectElement("//x:ExtendedData/x:Data[@name=""west""]/x:value", nsmgr).Value.ToString
                        Dim env As New Envelope(west, south, east, north, SpatialReferences.Wgs84)
                        extent = EnvelopeUnion(extent, env)
                    Catch ex As Exception
                        MsgBox($"Missing ExtendedData in {id }")
                    End Try
                Catch ex As Exception
                    Me.TextBox1.Text = $"Failed To open {VKFFfile }"
                End Try
            Next
            ' Create an appropriate LookAt

            If extent IsNot Nothing Then
                Dim Range As Double = CalcRange(extent)
                Dim LookAtPoint As MapPoint = extent.GetCenter      ' center of extent
                If state = "ZL" Then
                    ' the ZL envelop crosses date line, so LookAtPoint is wrong side of earth
                    LookAtPoint = New MapPoint(LookAtPoint.X + 180, LookAtPoint.Y)
                End If
                For Each f In files
                    ' Record extents of this park for index routine
                    f.WriteLine("<ExtendedData>")
                    f.WriteLine($"<Data name=""north""><value>{extent.YMax:f5}</value></Data>")
                    f.WriteLine($"<Data name=""south""><value>{extent.YMin:f5}</value></Data>")
                    f.WriteLine($"<Data name=""east""><value>{extent.XMax:f5}</value></Data>")
                    f.WriteLine($"<Data name=""west""><value>{extent.XMin:f5}</value></Data>")
                    f.WriteLine("</ExtendedData>")
                    f.WriteLine($"<LookAt><longitude>{LookAtPoint.X:f5}</longitude><latitude>{LookAtPoint.Y:f5}</latitude><range>{Range:f0}</range><heading>0</heading><tilt>0</tilt></LookAt>")
                Next
            End If

            ' Sort the list by name
            NetworkList.Sort(Function(a As String, b As String)
                                 Return CompareName(a).CompareTo(CompareName(b))
                             End Function)
            For Each item In NetworkList
                RemoteindxWriter.WriteLine(String.Format(item, target_remote))
                LocalindxWriter.WriteLine(String.Format(item, target_local))
            Next
            For Each sr In files
                sr.WriteLine(KMLfooter)
                sr.Close()
            Next
            ' compress all kml to zip file
            System.IO.File.Delete(BaseFilename & ".kmz")        ' delete any existing kmz
            Dim zip As ZipArchive = ZipFile.Open(BaseFilename & ".kmz", ZipArchiveMode.Create)    ' create new archive file
            zip.CreateEntryFromFile(BaseFilename & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
            zip.Dispose()
            msg = $"{count } files written to index for {state }"
            SetText(TextBox1, msg)
            logWriter.WriteLine(msg)
            total += count
        Next
        msg = $"{total } total files written to index"
        SetText(TextBox1, msg)
        logWriter.WriteLine(msg)
        logWriter.Close()

        ' compress all.kml to zip file
        Dim AllFolder As String = Application.StartupPath & "\files\"
        System.IO.File.Delete(AllFolder & "VKFF-all.kmz")        ' delete any existing kmz
        Dim zp As ZipArchive = ZipFile.Open(AllFolder & "VKFF-all.kmz", ZipArchiveMode.Create)    ' create new archive file
        zp.CreateEntryFromFile(AllFolder & "VKFF-all.kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
        zp.Dispose()

        ' Compress all files for release
        ' Add VKFF files
        System.IO.File.Delete("files\VKFF.zip")        ' delete any existing index
        Dim zipall As ZipArchive = ZipFile.Open("files\VKFF.zip", ZipArchiveMode.Create)    ' create new archive file
        SetText(TextBox1, "Generating zip file")
        For Each kmlFile In Directory.GetFiles("files\", "*.kmz", SearchOption.AllDirectories)
            Dim s As Integer = InStr(kmlFile, "\files")      ' find start of folder
            Name = kmlFile.Substring(s)
            SetText(TextBox1, "Adding " & Name)
            zipall.CreateEntryFromFile(kmlFile, Name, CompressionLevel.Optimal)   ' compress output file
        Next
        ' add _local files
        For Each kmlFile In Directory.GetFiles("files\", "*_local.kml", SearchOption.TopDirectoryOnly)
            Dim s As Integer = InStr(kmlFile, "\files")      ' find start of folder
            Name = kmlFile.Substring(s)
            SetText(TextBox1, "Adding " & Name)
            zipall.CreateEntryFromFile(kmlFile, Name, CompressionLevel.Optimal)   ' compress output file
        Next
        ' Now add SOTA icons
        For Each pngFile In Directory.GetFiles("files\pins", "*.png", SearchOption.AllDirectories)
            Dim s As Integer = InStr(pngFile, "\files\pins")      ' find start of folder
            Name = pngFile.Substring(s)
            SetText(TextBox1, "Adding " & Name)
            zipall.CreateEntryFromFile(pngFile, Name, CompressionLevel.Optimal)   ' compress output file
        Next
        zipall.Dispose()
        SetText(TextBox1, "Done")
        MsgBox(total & " files indexed", vbInformation + vbOKOnly, "Done")
    End Sub

    Shared Function CompareName(a As String) As String
        ' Extract the name part of the string, i.e. surrounded by <name>, </name>
        Dim m As MatchCollection = Regex.Matches(a, "<name>(.*)</name>")
        Dim aName As String = m(0).Groups(1).Value
        Return aName
    End Function

    Const EARTH_RADIUS = 6371000    ' radius of earth in meters
    Shared Function CalcRange(bounds As Envelope) As Double
        ' Calculate a LookAt range based on bounding box
        ' Based on earth API utility library - createBoundsView
        ' https://code.google.com/archive/p/earth-api-utility-library/wikis/GEarthExtensionsViewReference.wiki#createBoundsView(bounds,_options)
        Dim DistEW As Double, DistNS As Double
        Contract.Requires(bounds IsNot Nothing AndAlso Not bounds.IsEmpty, "Bounds is empty")

        Dim center As MapPoint = bounds.GetCenter()

        With bounds
            ' Calculate width of envelope at center
            DistEW = GeometryEngine.DistanceGeodetic(New MapPoint(.XMin, center.Y, .SpatialReference), New MapPoint(.XMax, center.Y, .SpatialReference), LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic).Distance
            DistNS = GeometryEngine.DistanceGeodetic(New MapPoint(center.X, .YMin, .SpatialReference), New MapPoint(center.X, .YMax, .SpatialReference), LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic).Distance
        End With
        Dim expandToDistance As Double = Max(DistNS, DistEW)
        Dim aspectRatio As Double = Min(Max(1.5, DistEW / DistNS), 1)
        Dim scaleRange As Double = 1.5
        Dim alpha As Double = Radians(45.0 / (aspectRatio + 0.4) - 2.0)
        Dim beta = Min(Radians(90), alpha + expandToDistance / (2 * EARTH_RADIUS))
        Dim lookAtRange As Double = scaleRange * EARTH_RADIUS * (Sin(beta) * Sqrt(1 + 1 / Pow(Tan(alpha), 2)) - 1)
        Return lookAtRange
    End Function

    Shared Function Radians(degree As Double) As Double
        ' Convert degrees to radians
        Return degree * PI / 180.0
    End Function

    Private Shared Function XMLencode(ByVal entry As String) As String
        Dim returnValue As String = entry

        ' Replace the special characters
        returnValue = returnValue.Replace("&", "&amp;").Replace("""", "&quot;").Replace("'", "&apos;").Replace("<", "&lt;").Replace(">", "&gt;")

        ' return the escaped string
        Return returnValue
    End Function

    Private Async Sub ZLParksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZLParksToolStripMenuItem.Click
        ' Extract candidate parks from ZL data, and calculate centroid
        Dim Conservati As String, count As Integer = 0, skipped As Integer = 0, line As String
        Dim name As String, centroid As MapPoint, values() As String, DistrictName As String
        Dim env As Envelope
        Dim logWriter As New System.IO.StreamWriter("zl_parks.csv", False)
        Dim connect As SQLiteConnection ' declare the connection
        Dim sql As SQLiteCommand
        Dim SQLdr As SQLiteDataReader
        Dim Existing As New List(Of String)
        Dim Coastline As ShapefileFeatureTable         '  Shapefiletable for NZ coastline
        Dim RegionalCouncils As ShapefileFeatureTable         '  Shapefiletable for NZ regional councils
        Dim Districts As ShapefileFeatureTable         '  Shapefiletable for NZ regional councils
        Dim myQueryFilter As New QueryParameters, region As String, RegionalCouncil As String

        Coastline = Await ShapefileFeatureTable.OpenAsync("F:\GIS Data\NZ\Coastlines\nz-coastlines-and-islands-polygons-topo-150k.shp").ConfigureAwait(False)  ' open shape file
        RegionalCouncils = Await ShapefileFeatureTable.OpenAsync("F:\GIS Data\NZ\Boundaries\REGC2017_GV_Clipped.shp").ConfigureAwait(False)  ' open shape file
        Districts = Await ShapefileFeatureTable.OpenAsync("F:\GIS Data\NZ\Boundaries\TA2017_GV_Clipped.shp").ConfigureAwait(False)  ' open shape file

        ' Get a list of existing parks so we don't repeat them in the extracted list
        connect = New SQLiteConnection(PARKSdb)
        connect.Open()  ' open database
        sql = connect.CreateCommand
        sql.CommandText = "Select a.GISID FROM GISmapping a join parks b WHERE a.WWFFID=b.WWFFID AND State='ZL'"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            Existing.Add(SQLdr.Item("GISID").ToString)
        End While
        SQLdr.Close()
        connect.Close()

        myQueryFilter.WhereClause = "Conservati<>'' ORDER BY Shape_Area desc"
        myQueryFilter.ReturnGeometry = False
        DataSets("ZL").shpFragments = Await DataSets("ZL").shpShapeFileTable.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)
        For Each park In DataSets("ZL").shpFragments
            Conservati = park.GetAttributeValue("Conservati").ToString
            If Existing.Contains(Conservati) Then
                skipped += 1
            Else
                count += 1
                If count = 1 Then
                    line = Join(park.Attributes.Keys.ToArray, ",")
                    line &= ",X_COORD,Y_COORD,Region,Council,District"
                    logWriter.WriteLine(line)
                End If

                ' Empty envelope
                name = park.GetAttributeValue("Name").ToString
                env = Nothing
                myQueryFilter.WhereClause = "Conservati='" & Conservati & "'"    ' query parameters
                myQueryFilter.ReturnGeometry = True
                myQueryFilter.OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
                Dim fragments = Await DataSets("ZL").shpShapeFileTable.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)           ' run query
                For Each fragment In fragments
                    env = EnvelopeUnion(env, fragment.Geometry.Extent)
                Next
                centroid = env.GetCenter    ' get center of envelope
                ' Find which geographical area this is
                myQueryFilter.Geometry = GeometryEngine.Project(New MapPoint(centroid.X, centroid.Y, SpatialReferences.Wgs84), Coastline.SpatialReference) ' point to find
                myQueryFilter.SpatialRelationship = Esri.ArcGISRuntime.Data.SpatialRelationship.Within
                Dim area = Await Coastline.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)  ' find shape that contains point
                If Not area.Any Then
                    region = "? ? ? ? ?"
                Else
                    region = area(0).GetAttributeValue("name").ToString
                    If region.StartsWith("North Island", StringComparison.CurrentCulture) Then
                        region = "North Island"
                    ElseIf region.StartsWith("South Island", StringComparison.CurrentCulture) Then
                        region = "South Island"
                    End If
                End If
                ' Find regional council
                myQueryFilter.Geometry = GeometryEngine.Project(New MapPoint(centroid.X, centroid.Y, SpatialReferences.Wgs84), RegionalCouncils.SpatialReference)
                Dim council = Await RegionalCouncils.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)  ' find shape that contains point
                If Not council.Any Then
                    RegionalCouncil = "? ? ? ? ?"
                Else
                    RegionalCouncil = Replace(council(0).GetAttributeValue("REGC2017_N").ToString, " Region", "")
                End If
                ' Find district
                myQueryFilter.Geometry = GeometryEngine.Project(New MapPoint(centroid.X, centroid.Y, SpatialReferences.Wgs84), Districts.SpatialReference)
                Dim district = Await Districts.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)  ' find shape that contains point
                If Not district.Any Then
                    DistrictName = "? ? ? ? ?"
                Else
                    DistrictName = Replace(district(0).GetAttributeValue("TA2017_NAM").ToString, " District", "")
                End If
                SetText(TextBox1, String.Format("{0}(Skipped {1})/{2}: {3}: Centroid at {4:f5},{5:f5}", count, skipped, DataSets("ZL").shpFragments.Count, name, centroid.X, centroid.Y))
                values = park.Attributes.Values.Select(Function(s) """" & If(s, String.Empty).ToString & """").ToArray()
                line = Join(values, ",") & String.Format(",""{0:f5}"",""{1:f5}"",""{2}"",""{3}"",""{4}""", centroid.X, centroid.Y, region, RegionalCouncil, DistrictName)
                logWriter.WriteLine(line)
                logWriter.Flush()
            End If
        Next
        SetText(TextBox1, "Done")
        logWriter.Close()
    End Sub

    Public Shared Function GetTruePolygonCenter(ByVal pgnPolygon As Polygon) As MapPoint
        ' Find the centroid of a polygon, i.e. the average of all the points
        Contract.Requires(pgnPolygon IsNot Nothing AndAlso Not pgnPolygon.IsEmpty)
        Contract.Requires(pgnPolygon.Parts.Any)
        Try
            Dim iNumberOfMapPoints As Integer = 0
            Dim centroidX As Double = 0
            Dim centroidY As Double = 0
            For Each part In pgnPolygon.Parts
                For Each point In part.Points
                    centroidX += point.X
                    centroidY += point.Y
                    iNumberOfMapPoints += 1
                Next
            Next
            Return New MapPoint(centroidX / iNumberOfMapPoints, centroidY / iNumberOfMapPoints, pgnPolygon.SpatialReference)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Sub GenerateRegionsToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles GenerateRegionsToolStripMenuItem.Click
        ' Generate list of Region & District for all VKFF parks where these are missing
        Dim sql As SQLiteCommand
        Dim SQLdr As SQLiteDataReader

        Dim sqlWriter As New System.IO.StreamWriter(String.Format("regions-{0}.sql", DateTime.Now.ToString("yyMMdd-HHmm")), False)
        Dim Updates As New List(Of String)
        Dim region As String, district As String, count As Integer = 0, skipped As Integer = 0, name As String, state As String
        Dim WWFFID As String, X_COORD As Double, Y_COORD As Double, ParkID As Integer

        ' Output SQL data
        sqlWriter.WriteLine("/* Parks with Region/District missing. File created {0} */", DateTime.Now.ToString(YMDHMS))
        Using connect As New SQLiteConnection(PARKSdb)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            ' select all parks in this state with a GISID
            sql.CommandText = "SELECT * FROM parks WHERE WWFFID LIKE 'VKFF-%' AND WWFFID<>'VKFF-0000' AND Longitude>0 AND (Region is null OR Region='' OR District is null OR District='') ORDER BY WWFFID"
            SQLdr = sql.ExecuteReader()
            While SQLdr.Read()
                Updates.Clear()
                ParkID = SQLdr.Item("ParkID")
                WWFFID = SQLdr.Item("WWFFID").ToString
                X_COORD = CDbl(SQLdr.Item("Longitude").ToString)
                Y_COORD = CDbl(SQLdr.Item("Latitude").ToString)
                name = SQLdr.Item("Name").ToString
                state = SQLdr.Item("State").ToString
                region = SQLdr.Item("Region").ToString
                district = SQLdr.Item("District").ToString
                count += 1
                ' Find data using online service
                GetRegionDistrict(X_COORD, Y_COORD, region, district)
                If region <> "" And district <> "" Then
                    SetText(TextBox1, $"{count } ({skipped }) {WWFFID } {region } -> {district }")
                    Updates.Add($"`Region`='{SQLEscape(region) }'")
                    Updates.Add($"`District`='{SQLEscape(district) }'")
                Else
                    skipped += 1
                End If
                If Updates.Any Then
                    sqlWriter.WriteLine($"/* WWFFID {WWFFID } */")  ' comment
                    sqlWriter.WriteLine($"UPDATE `PARKS` SET {Join(Updates.ToArray, ",") } WHERE `ParkID`={ParkID };")
                    sqlWriter.Flush()
                End If
            End While
            SQLdr.Close()
            SetText(TextBox1, "Done - results in " & CType(sqlWriter.BaseStream, FileStream).Name)
            sqlWriter.Close()
        End Using
    End Sub
    Private Shared Sub GetRegionDistrict(longitude As Double, latitude As Double, ByRef Region As String, ByRef District As String)
        ' get the region and district for a lat/lon
        Dim request As System.Net.HttpWebRequest, response As System.Net.HttpWebResponse
        Dim params As String, url As String, sourcecode As String
        Dim remove As String() = {"New South Wales -", "Victoria -", "Queensland -", "South Australia -", "Western Australia -", "Tasmania -", "Northern Territory -", "Other Territories -"}

        params = "?f=json"          ' return json
        params &= "&geometryType=esriGeometryPoint"     ' look for a point (centroid)
        params &= $"&geometry={longitude:f5},{latitude:f5}"   ' point to look for
        params &= "&inSR={'wkid' : 4326}"                                      ' WGS84 datum
        params &= "&spatialRel=esriSpatialRelWithin"
        params &= "&returnGeometry=false"                   ' don't need the geometry
        params &= "&outFields=*"                            ' return all fields, even though we only use the default

        ' Get region
        url = GEOSERVER_SA4 & params
        request = HttpWebRequest.Create(url)
        response = request.GetResponse()
        Using sr As New System.IO.StreamReader(response.GetResponseStream())
            sourcecode = sr.ReadToEnd()
        End Using
        Dim Jo As JObject = JObject.Parse(sourcecode)
        If Jo.HasValues And Jo.Item("features").Any Then
            Region = Jo.Item("features")(0)("attributes")("SA4_NAME_2016").ToString
            ' remove redundant state name
            For Each st In remove
                Region = Replace(Region, st, "")
            Next
            Region = Trim(Region)
        End If
        If Region <> "" Then
            ' Get District
            url = GEOSERVER_SA3 & params
            request = HttpWebRequest.Create(url)
            response = request.GetResponse()
            Using sr As New System.IO.StreamReader(response.GetResponseStream())
                sourcecode = sr.ReadToEnd()
            End Using
            Jo = JObject.Parse(sourcecode)
            If Jo.HasValues And Jo.Item("features").Any Then
                District = Jo.Item("features")(0)("attributes")("SA3_NAME_2016").ToString
            End If
        End If
    End Sub

    Private Sub ImportIOTARefsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportIOTARefsToolStripMenuItem.Click
        ' Get the list of IOTA definitions and add to database
        ' Dim url As String = "https://www.iota-world.org/rest/get/iota/fulllist?api_key=BOM6PQXRU8O427TFRDYH"
        Dim refno As String, name As String, latitude_max As Double, latitude_min As Double, longitude_max As Double, longitude_min As Double
        Dim sql As SQLiteCommand, cmd As String, count As Integer = 0

        Using connect As New SQLiteConnection("Data Source=IOTA.db")
            connect.Open()  ' open database
            sql = connect.CreateCommand
            sql.CommandText = "begin"
            sql.ExecuteNonQuery()       ' start transaction for speed
            sql.CommandText = "DELETE FROM iota"   ' remove existing data
            sql.ExecuteNonQuery()

            ' There must be some code missing here - get url

            Dim IOTAjson As String
            Dim IOTAreader As New System.IO.StreamReader("response.json")
            IOTAjson = IOTAreader.ReadToEnd
            IOTAreader.Close()
            Dim IOTA = JsonConvert.DeserializeXmlNode(IOTAjson, "contents")      ' convert JSON to XML
            Dim refs As XmlNodeList = IOTA.SelectNodes("contents/content")
            For Each node As XmlElement In refs
                refno = node("refno").InnerText
                name = node("name").InnerText
                name = name.Replace("'", "''")   ' escape single quotes
                latitude_max = node("latitude_max").InnerText
                latitude_min = node("latitude_min").InnerText
                If latitude_max < latitude_min Then Swap(latitude_max, latitude_min)   ' Inexplicably, latitudes are back to front in southern hemisphere
                longitude_max = node("longitude_max").InnerText
                longitude_min = node("longitude_min").InnerText
                If longitude_max < longitude_min Then Swap(longitude_max, longitude_min)   ' Inexplicably, longitudes are back to front in western hemisphere
                cmd = String.Format("INSERT INTO IOTA ({0},{1},{2},{3},{4},{5}) VALUES ('{6}','{7}',{8},{9},{10},{11})", "refno", "name", "latitude_max", "latitude_min", "longitude_max", "longitude_min", refno, name, latitude_max, latitude_min, longitude_max, longitude_min)
                sql.CommandText = cmd
                sql.ExecuteNonQuery()
                count += 1
            Next
            sql.CommandText = "end"
            sql.ExecuteNonQuery()       ' end transaction
            connect.Close()
            SetText(TextBox1, $"Done - {count } IOTA refs added")
        End Using
    End Sub

    Private Sub FindIOTAForParksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindIOTAForParksToolStripMenuItem.Click
        Dim SERVER = "https://services.ga.gov.au/gis/rest/services/NM_Labelling_and_Boundaries/MapServer/identify"    ' Server with lots of outlines
        Dim sqlParks As SQLiteCommand, sqlIOTA As SQLiteCommand
        Dim SQLdr As SQLiteDataReader, SQLdrIOTA As SQLiteDataReader
        Dim count As Integer = 0, skipped As Integer = 0, ParkName As String, IslandName As String, IOTAref As String, IOTAgroup As String
        Dim WWFFID As String, TYPE As String, X_COORD As Double, Y_COORD As Double, State As String, Island As Boolean
        Dim point As MapPoint
        Dim GETfields As New NameValueCollection()

        Using myWebClient As New WebClient(),
            logWriter As New System.IO.StreamWriter("IOTA.html", False),
            connectIOTA As New SQLiteConnection("Data Source=IOTA.db"),
            connectParks As New SQLiteConnection(PARKSdb)

            Dim htmlFile As String = CType(logWriter.BaseStream, FileStream).Name
            connectIOTA.Open()  ' open database
            connectParks.Open()
            sqlParks = connectParks.CreateCommand
            sqlIOTA = connectIOTA.CreateCommand

            logWriter.WriteLine("<table border=1>")
            logWriter.WriteLine("<tr><th>WWFFID</th><th>Park name</th><th>State</th><th>IOTA ref</th><th>Island name</th><th>IOTA Group</th></tr>")

            ' select all parks
            sqlParks.CommandText = "SELECT * FROM parks WHERE WWFFID LIKE 'VKFF-%' AND WWFFID<>'VKFF-0000' ORDER BY WWFFID"
            SQLdr = sqlParks.ExecuteReader()
            While SQLdr.Read()
                count += 1                           ' count number we have scanned
                WWFFID = SQLdr.Item("WWFFID").ToString
                ParkName = SQLdr.Item("Name").ToString
                TYPE = SQLdr.Item("Type").ToString
                X_COORD = CDbl(SQLdr.Item("Longitude").ToString)
                Y_COORD = CDbl(SQLdr.Item("Latitude").ToString)
                State = SQLdr.Item("State").ToString
                Island = False
                IslandName = ""
                Select Case WWFFID
                ' Some islands are not on the map, so hardwire them
                    Case "VKFF-0098" : Island = True : IslandName = "Christmas Island"
                    Case "VKFF-0295", "VKFF-1409", "VKFF-0392" : Island = True : IslandName = "Lord Howe Island"
                    Case "VKFF-0392" : Island = True : IslandName = "Norfolk Island"
                    Case "VKFF-0423" : Island = True : IslandName = "Cocos (Keeling) Island"
                    Case "VKFF-0565" : Island = True : IslandName = "Heard and MacDonald Islands"
                    Case "VKFF-0566" : Island = True : IslandName = "Macquarie Island"
                    Case "VKFF-0571" : Island = True : IslandName = "Antarctica"
                    Case "VKFF-0573" : Island = True : IslandName = "Mellish Reef"
                    Case "VKFF-0574" : Island = True : IslandName = "Willis Island"
                    Case Else
                        ' Find island if any
                        point = New MapPoint(X_COORD, Y_COORD, SpatialReferences.Wgs84)
                        point = GeometryEngine.Project(point, SpatialReference.Create(4283))  ' convert point to datum of map
                        ParkName = SQLdr.Item("Name").ToString
                        SetText(TextBox1, $"Scanning {WWFFID } {ParkName } ({count }) Found: {found }")
                        ' Find data using online service
                        Dim buffer = GeometryEngine.BufferGeodetic(point, 100, LinearUnits.Meters)          ' create 100m rectangle around point
                        GETfields.Clear()
                        GETfields.Add("f", "json")
                        GETfields.Add("geometryType", "esriGeometryPoint")
                        GETfields.Add("geometry", $"{point.X:f5},{point.Y:f5}")
                        GETfields.Add("mapExtent", $"{buffer.Extent.XMin:f5},{buffer.Extent.YMin:f5},{buffer.Extent.XMax:f5},{buffer.Extent.YMax:f5}")
                        GETfields.Add("layers", "all:31")       ' Islands
                        GETfields.Add("tolerance", "1")
                        GETfields.Add("imageDisplay", "600,550,96")
                        GETfields.Add("returnGeometry", "false")

                        ' Identify the point
                        Dim responseArray As Byte() = myWebClient.UploadValues(SERVER, GETfields)
                        Dim sourcecode = System.Text.Encoding.UTF8.GetString(responseArray)     ' convert to string
                        System.Threading.Thread.Sleep(100)      ' don't hammer the server - delay 100mS
                        'Dim fs = FeatureSet.FromJson(sourcecode)        ' convert JSON to featureSet
                        ' For reasons unknown "FeatureSet.FromJson(sourcecode)" does not work
                        ' Convert JSON to XML instead
                        Dim xml As XmlDocument = JsonConvert.DeserializeXmlNode(sourcecode)
                        If xml.HasChildNodes Then
                            Island = True
                            IslandName = xml.DocumentElement("value").InnerText
                            IslandName = Regex.Replace(IslandName, "\b(\w|['-])+\b", evaluator)   ' convert case
                        End If
                End Select
                If Island Then
                    ' Get the IOTA ref
                    sqlIOTA.CommandText = $"SELECT * FROM iota WHERE ({Y_COORD } BETWEEN latitude_min AND latitude_max) AND ({X_COORD } BETWEEN longitude_min AND longitude_max)"
                    SQLdrIOTA = sqlIOTA.ExecuteReader()
                    IOTAref = "??????"
                    IOTAgroup = "??????"
                    While SQLdrIOTA.Read()
                        IOTAref = SQLdrIOTA.Item("refno")
                        IOTAgroup = SQLdrIOTA.Item("name")
                    End While
                    SQLdrIOTA.Close()
                    logWriter.WriteLine($"<tr><th>{WWFFID }</th><th>{ParkName } {TYPE }</th><th>{State }</th><th>{IOTAref }</th><th>{IslandName }</th><th>{IOTAgroup }</th></tr>")
                    logWriter.Flush()
                    found += 1
                End If
            End While
            logWriter.WriteLine("</table>")
            logWriter.Close()
            SetText(TextBox1, $"Scaned {count } Found: {found }")
            Process.Start(htmlFile)
        End Using
    End Sub

    Private Sub CrossCheckAreaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CrossCheckAreaToolStripMenuItem.Click
        ' Check area in parks database with that in shapefile
        ' Each shapefile dataset comes with an adjacent .dbf (dBase IV) file that contains the metadata.
        ' It is quicker to search the metadata than to retrieve and sum polygons.

        Const tolerance = 0.05  ' 5 percent
        Dim sql As SQLiteCommand, sqldr As SQLiteDataReader, WWFFID As String, DataSet As String, GISID As String, Area As Integer
        Dim query As String, GIS_AREA As Double, count As Integer = 0, mismatch As Integer = 0
        Dim dBaseDataReader As System.Data.OleDb.OleDbDataReader
        Dim ParkData As NameValueCollection

        Using connect As New SQLiteConnection(PARKSdb),
            logWriter As New System.IO.StreamWriter("AreaCrosscheck.html", False),
            sqlWriter As New System.IO.StreamWriter("AreaCrosscheck.sql", False)
            Dim htmlFile As String = CType(logWriter.BaseStream, FileStream).Name
            logWriter.WriteLine($"Parks where area difference exceeds {tolerance * 100 }%<br><br>")
            sqlWriter.WriteLine($"/* Parks where area difference exceeds {tolerance * 100 }% */")
            logWriter.WriteLine("<table border=1>")
            logWriter.WriteLine(String.Format("<tr><th>{0}</th><th>{1}</th><th>{2}</th><th>{3}</th><th>{4}</th></tr>", "WWFFID", "Name", "State", "Parks Area", "GIS_AREA"))
            connect.Open()  ' open database
            sql = connect.CreateCommand
            sql.CommandText = "SELECT DISTINCT WWFFID FROM GISmapping ORDER BY WWFFID"
            sqldr = sql.ExecuteReader()
            While sqldr.Read()
                count += 1
                WWFFID = sqldr.Item("WWFFID").ToString
                ParkData = GetParkData(WWFFID)
                DataSet = ParkData("DataSet")
                GISID = ParkData("GISIDListQuoted")
                Area = ParkData("Area")
                TextBox1.Text = $"Checking {WWFFID }: Count: {count }, Mismatch: {mismatch }"
                Application.DoEvents()        ' update text box

                query = DataSets(DataSet).BuildWhere(GISID)
                Using dBaseCommand As New System.Data.OleDb.OleDbCommand(String.Format("SELECT * FROM {0} WHERE {1}", DataSets(DataSet).dbfTableName, query), DataSets(DataSet).dbfConnection)
                    dBaseDataReader = dBaseCommand.ExecuteReader()
                    ' Sum all areas of park
                    GIS_AREA = 0
                    While dBaseDataReader.Read
                        GIS_AREA += dBaseDataReader(DataSets(DataSet).AreaField).ToString / DataSets(DataSet).dbfAreaScale
                    End While
                    dBaseDataReader.Close()
                End Using

                Dim ratio As Single = 1 - (Min(Area, CInt(GIS_AREA)) / Max(Area, CInt(GIS_AREA)))       ' measure of difference. If 0 no difference
                If ratio > tolerance Then
                    mismatch += 1
                    logWriter.WriteLine($"<tr><td>{WWFFID }</td><td>{ParkData("Name") } {ParkData("Type") }</td><td>{ParkData("State") }</td><td>{Area }</td><td>{GIS_AREA:f2}</td></tr>")
                    logWriter.Flush()
                    sqlWriter.WriteLine($"/* Updating {WWFFID }: P&P area={Area }, GIS area={GIS_AREA } */")
                    sqlWriter.WriteLine($"UPDATE `PARKS` SET `Area`={GIS_AREA:f2} WHERE `ParkID`={ParkData("ParkID") };")
                    sqlWriter.Flush()
                End If
            End While
            logWriter.WriteLine("</table>")
            sqldr.Close()
            connect.Close()
            TextBox1.Text = "Done"
            Process.Start(htmlFile)
        End Using
    End Sub

    ReadOnly evaluator As MatchEvaluator = AddressOf TitleCase

    Public Shared Function TitleCase(ByVal m As Match) As String
        Return m.Value(0).ToString().ToUpper() & m.Value.Substring(1).ToLower()
    End Function

    Shared Sub Swap(Of T)(ByRef a As T, ByRef b As T)
        ' Swap variables a and b
        Dim temp As T
        temp = a
        a = b
        b = temp
    End Sub

    Private Sub GenerateRegionsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles GenerateRegionsToolStripMenuItem1.Click
        ' generate a list of SOTA regions, and their bounding box
        Dim connect As SQLiteConnection ' declare the connections
        Dim SOTAsql As SQLiteCommand, Regionsql As SQLiteCommand
        Dim SOTAsqldr As SQLiteDataReader
        Dim count As Integer = 0

        ' Connect to SOTA database
        connect = New SQLiteConnection(SOTAdb)
        connect.Open()  ' open database
        Regionsql = connect.CreateCommand
        SOTAsql = connect.CreateCommand

        Regionsql.CommandText = "begin"
        Regionsql.ExecuteNonQuery()
        Regionsql.CommandText = "DELETE FROM Region"       ' clear existing data
        Regionsql.ExecuteNonQuery()
        SOTAsql.CommandText = "select substr(SummitCode,1,6) as RegionCode, max(Latitude) as LatMax, min(Latitude) as LatMin, max(Longitude) as LonMax, min(Longitude) as LonMin FROM SOTA GROUP BY substr(SummitCode,1,6)"    ' Find all unique region codes
        SOTAsqldr = SOTAsql.ExecuteReader()
        While SOTAsqldr.Read()
            count += 1
            Dim RegionCode As String = SOTAsqldr.Item("RegionCode").ToString
            Dim LatMax As Double = CDbl(SOTAsqldr.Item("LatMax").ToString)
            Dim LatMin As Double = CDbl(SOTAsqldr.Item("LatMin").ToString)
            If LatMax < LatMin Then Swap(LatMax, LatMin)
            Dim LonMax As Double = CDbl(SOTAsqldr.Item("LonMax").ToString)
            Dim LonMin As Double = CDbl(SOTAsqldr.Item("LonMin").ToString)
            Dim rg As New Esri.ArcGISRuntime.Geometry.Envelope(LonMin, LatMin, LonMax, LatMax, SpatialReferences.Wgs84)
            Dim buffer = GeometryEngine.BufferGeodetic(rg, 500, LinearUnits.Meters)          ' expand region by 500m
            Regionsql.CommandText = $"INSERT INTO Region (RegionCode,LatMax,LatMin,LonMax,LonMin) VALUES ('{RegionCode }',{buffer.Extent.YMax },{buffer.Extent.YMin },{buffer.Extent.XMax },{buffer.Extent.XMin })"
            Regionsql.ExecuteNonQuery()
        End While
        Regionsql.CommandText = "end"
        Regionsql.ExecuteNonQuery()
        connect.Close()
        SetText(TextBox1, $"Done: added {count } regions")
    End Sub

    Private Sub ImportPromenenceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportPromenenceToolStripMenuItem.Click
        ' Import the parks data file into SQLite
        Dim count As Integer = 0    ' count of lines read
        Dim sql As SQLiteCommand

        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("prominence.csv"), connect As New SQLiteConnection(SOTAdb)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            MyReader.HasFieldsEnclosedInQuotes = True
            Dim currentRow As String()
            Dim cmd As String
            Dim fields As String = "", values As String

            sql.CommandText = "begin"
            sql.ExecuteNonQuery()       ' start transaction for speed
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    count += 1
                    If count = 1 Then
                        ' remove existing data
                        sql.CommandText = "DELETE FROM prominence"   ' remove existing data
                        sql.ExecuteNonQuery()
                        fields = Join(currentRow, ",")
                    Else
                        values = Join(currentRow, ",")
                        cmd = "INSERT INTO prominence (" & fields & ") VALUES (" & values & ")"
                        sql.CommandText = cmd
                        sql.ExecuteNonQuery()
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped." & vbCrLf & "Import aborted")
                    sql.CommandText = "rollback"
                    sql.ExecuteNonQuery()
                    GoTo done
                End Try
            End While

            sql.CommandText = "end"
            sql.ExecuteNonQuery()
done:
            connect.Close()
        End Using
        MsgBox(count - 1 & " prominence imported", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Import complete")
    End Sub

    Private Sub MatchSummitsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MatchSummitsToolStripMenuItem.Click
        ' Const SERVER = "https://services.ga.gov.au/gis/rest/services/NM_Labelling_and_Boundaries/MapServer/identify"    ' Server with lots of outlines

        Dim GETfields As New NameValueCollection()
        Dim count As Integer = 0    ' count of peaks tested
        Dim found As Integer = 0    ' count of summits found
        Dim sql As SQLiteCommand, sqldr As SQLiteDataReader
        Dim SOTAsql As SQLiteCommand, SOTAsqldr As SQLiteDataReader
        Dim Promsql As SQLiteCommand, Promsqldr As SQLiteDataReader
        Dim delta As Double             ' distance between peak and SOTA summit

        Using logWriter As New System.IO.StreamWriter("ProminentPeaks.html", False), myWebClient As New WebClient(), connect As New SQLiteConnection(SOTAdb)
            Dim htmlFile As String = CType(logWriter.BaseStream, FileStream).Name

            connect.Open()  ' open database
            sql = connect.CreateCommand
            SOTAsql = connect.CreateCommand
            Promsql = connect.CreateCommand
            sql.CommandText = "select * FROM Region ORDER BY RegionCode"    ' do all region in turn
            sqldr = sql.ExecuteReader()
            While sqldr.Read()
                ' Get the details of the region
                Dim RegionCode As String = sqldr.Item("RegionCode").ToString
                Dim LatMax As Single = CSng(sqldr.Item("LatMax").ToString)
                Dim LatMin As Single = CSng(sqldr.Item("LatMin").ToString)
                Dim LonMax As Single = CSng(sqldr.Item("LonMax").ToString)
                Dim LonMin As Single = CSng(sqldr.Item("LonMin").ToString)

                logWriter.WriteLine("<h1>Summits in region " & RegionCode & "</h1>")
                logWriter.WriteLine($"Region bounds: Top left {LatMax:f5},{LonMin:f5}: Bottom right {LatMin:f5},{LonMax:f5}</br>")
                logWriter.WriteLine("<table border=1>")
                logWriter.WriteLine("<tr><th colspan=4>'Kirmse' prominence data</th><th colspan=2>SOTA data</th></tr>")
                logWriter.WriteLine("<tr><th>Latitude</th><th>Longitude</th><th>Elevation</br>(m)</th><th>Prominence</br>(m)</th><th>SOTA ref</th><th>Name</th><th>Delta</br>(m)</th></tr>")
                ' Extract the peaks in this region
                Promsql.CommandText = String.Format("SELECT * FROM prominence WHERE Latitude BETWEEN {0:f5} AND {1:f5} AND Longitude BETWEEN {2:f5} AND {3:f5} AND Prominence>={4}", LatMin, LatMax, LonMin, LonMax, 150 / MetersperFoot)
                Promsqldr = Promsql.ExecuteReader()
                While Promsqldr.Read()
                    ' find matching SOTA peak
                    Dim PeakLat As Single = CSng(Promsqldr.Item("Latitude").ToString)
                    Dim PeakLon As Single = CSng(Promsqldr.Item("Longitude").ToString)
                    Dim elevation As Integer = Promsqldr.Item("Elevation").ToString
                    Dim prominence As Integer = Promsqldr.Item("Prominence").ToString
                    Dim peak As New MapPoint(PeakLon, PeakLat, SpatialReferences.Wgs84)       ' location of peak
                    Dim buffer = GeometryEngine.BufferGeodetic(peak, 300, LinearUnits.Meters)          ' create 100m circle around point
                    ' Find a matching SOTA summit
                    Dim SummitCode As String = ""
                    Dim SummitName As String = ""
                    count += 1
                    SOTAsql.CommandText = $"SELECT * FROM SOTA WHERE Latitude BETWEEN {buffer.Extent.YMin:f4} AND {buffer.Extent.YMax:f4} AND Longitude BETWEEN {buffer.Extent.XMin:f4} AND {buffer.Extent.XMax:f4}"
                    SOTAsqldr = SOTAsql.ExecuteReader()
                    While SOTAsqldr.Read()
                        SummitCode = SOTAsqldr.Item("SummitCode").ToString
                        SummitName = SOTAsqldr.Item("SummitName").ToString
                        Dim SummitLatitude As Double = SOTAsqldr.Item("Latitude").ToString
                        Dim SummitLongitude As Double = SOTAsqldr.Item("Longitude").ToString
                        Dim Summit As New MapPoint(SummitLongitude, SummitLatitude, SpatialReferences.Wgs84)
                        delta = GeometryEngine.DistanceGeodetic(peak, Summit, LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic).Distance
                        logWriter.WriteLine($"<tr><td>{PeakLat:f4}</td><td>{PeakLon:f4}</td><td align='right'>{elevation * MetersperFoot:f1}</td><td align='right'>{prominence * MetersperFoot:f1}</td><td>{SummitCode }</td><td>{SummitName }</td><td align='right'>{delta:f0}</td>")
                        found += 1
                    End While
                    If String.IsNullOrEmpty(SummitCode) Then
                        'try to find a mountain name
                        'GETfields.Clear()
                        'GETfields.Add("f", "json")
                        'GETfields.Add("geometryType", "esriGeometryPoint")
                        'GETfields.Add("geometry", String.Format("{0:f5},{1:f5}", peak.X, peak.Y))
                        'GETfields.Add("mapExtent", String.Format("{0:f5},{1:f5},{2:f5},{3:f5}", buffer.Extent.XMin, buffer.Extent.YMin, buffer.Extent.XMax, buffer.Extent.YMax))
                        'GETfields.Add("layers", "all:13")       ' releif features
                        'GETfields.Add("tolerance", "1")
                        'GETfields.Add("imageDisplay", "600,550,96")
                        'GETfields.Add("returnGeometry", "false")
                        'Dim responseArray As Byte() = myWebClient.UploadValues(SERVER, GETfields)
                        'Dim sourcecode = System.Text.Encoding.UTF8.GetString(responseArray)     ' convert to string
                        'System.Threading.Thread.Sleep(100)      ' don't hammer the server - delay 100mS
                        ''Dim fs = FeatureSet.FromJson(sourcecode)        ' convert JSON to featureSet
                        '' For reasons unknown "FeatureSet.FromJson(sourcecode)" does not work
                        '' Convert JSON to XML instead
                        'Dim xml As XmlDocument = JsonConvert.DeserializeXmlNode(sourcecode)
                        'If xml.HasChildNodes Then
                        '    SummitName = xml.DocumentElement("value").InnerText
                        '    SummitName = "<font color='red'>" & Regex.Replace(SummitName, "\b(\w|['-])+\b", evaluator) & "</font>"  ' convert case
                        'End If

                        logWriter.WriteLine(String.Format("<tr><td>{0:f4}</td><td>{1:f4}</td><td align='right'>{2:f1}</td><td align='right'>{3:f1}</td><td>{4}</td><td>{5}</td><td></td>", PeakLat, PeakLon, elevation * MetersperFoot, prominence * MetersperFoot, SummitCode, SummitName))
                    End If
                    SetText(TextBox1, $"Tested {count } Found {found }")
                    SOTAsql.Reset()
                End While
                Promsql.Reset()
                logWriter.WriteLine("</table>")
            End While
            logWriter.Close()
            connect.Close()
            SetText(TextBox1, $"Done: Tested {count } Found {found }")
            Process.Start(htmlFile)
        End Using
    End Sub

    Private Async Sub ImportWWFFDirectoryCSVToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportWwffdirectorycsvToolStripMenuItem.Click
        ' Import  the WWFF directory data file into SQLite
        Dim count As Integer = 0, inserted As Integer = 0 ' count of lines read
        Dim sql As SQLiteCommand, sourcecode As String

        Dim url As New Uri("https://wwff.co/wwff-data/wwff_directory.csv")  ' the wwff directory at wwff.co
        TextBox2.Text = $"requesting WWFF directory from {url }"
        Dim webReq = CType(WebRequest.Create(url), HttpWebRequest)
        webReq.Timeout = 5 * 60 * 1000        ' 5 min timeout
        webReq.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)" ' trick server into thinking it's browser access
        Using response As WebResponse = Await webReq.GetResponseAsync().ConfigureAwait(True)
            Dim content As New MemoryStream()
            Using responseStream As Stream = response.GetResponseStream()
                Await responseStream.CopyToAsync(content).ConfigureAwait(True)
            End Using

            ' convert the MemoryStream into a string
            Using sr = New StreamReader(content)
                content.Position = 0    ' rewind to start
                sourcecode = sr.ReadToEnd()
                ' Write to file
                Dim file As System.IO.StreamWriter
                file = My.Computer.FileSystem.OpenTextFileWriter("wwff_directory.csv", False)
                file.Write(sourcecode)
                file.Close()
            End Using
        End Using
        TextBox2.AppendText($"{vbCrLf }received {sourcecode.Length } bytes of data")

        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("wwff_directory.csv"), connect As New SQLiteConnection(PARKSdb)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            MyReader.HasFieldsEnclosedInQuotes = True
            Dim currentRow As String()
            Dim fields As New List(Of String), fieldlist As String, values As New List(Of String), valuelist As String

            sql.CommandText = "begin"
            sql.ExecuteNonQuery()       ' start transaction for speed
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    count += 1

                    If count = 1 Then
                        ' remove existing data
                        sql.CommandText = "DELETE FROM WWFF"   ' remove existing data
                        sql.ExecuteNonQuery()
                        ' build the sql command
                        fields.Clear()
                        values.Clear()
                        fields = currentRow.ToList  ' list of field names
                        fieldlist = String.Join(",", fields)   ' csv list of fields
                        For Each field In fields
                            values.Add($"@{field }")   ' prepend field symbol
                        Next
                        valuelist = String.Join(",", values)  ' csv list of value placeholders
                        sql.CommandText = $"INSERT INTO WWFF ({fieldlist }) VALUES ({valuelist })" ' sql command with placeholders
                        sql.Prepare()   ' prepare command for speed
                    Else
                        If currentRow(0).StartsWith("VKFF", StringComparison.CurrentCulture) Or currentRow(0).StartsWith("ZLFF", StringComparison.CurrentCulture) Then      ' only add VK and ZL to table
                            For i = 0 To currentRow.Length - 1
                                currentRow(i) = currentRow(i).Replace("\", "")   ' remove rogue backslashes
                                currentRow(i) = currentRow(i).Replace("'", "''")   ' escape single quotes
                            Next
                            ' add parameters
                            sql.Parameters.Clear()
                            For i = 0 To values.Count - 1
                                sql.Parameters.AddWithValue(values(i), currentRow(i))
                            Next
                            sql.ExecuteNonQuery()
                            inserted += 1
                        End If
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.", vbCritical + vbOKOnly, "Input error")
                End Try
            End While

            sql.CommandText = "end"
            sql.ExecuteNonQuery()
done:
            connect.Close()
            MsgBox($"{count - 1 } parks scanned, {inserted } inserted", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "Import complete")
        End Using
    End Sub

    Private Sub CrossCheckWWFFPnPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CrossCheckWWFFPnPToolStripMenuItem.Click
        ' Cross check WWFF directory with Parks & Peaks data
        Const MinDistance = 5       ' minimum accptable distance netween centroids (km)
        Dim connect As SQLiteConnection ' declare the connection
        Dim sqlWWFF As SQLiteCommand, sqlPnP As SQLiteCommand
        Dim sqlWWFFdr As SQLiteDataReader, sqlPnPdr As SQLiteDataReader
        Dim errors As Integer = 0, count As Integer = 0, errStr As String
        Dim errList As New List(Of String)
        Dim WWFFID As String            ' park reference
        Dim NameWWFF As String, NamePnP As String
        Dim type As String
        Dim LatWWFF As Single, LonWWFF As Single, LatPnP As Single, LonPNP As Single
        Dim continentWWFF As String
        Dim dxccWWFF As String, dxccPnP As String
        Dim stateWWFF As String, statePnP As String
        Dim statusWWFF As String, statusPnP As String
        Dim found As Boolean
        Dim pastels() As String = {"#d0fffe", "#fffddb", "#e4ffde", "#ffd3fd", "#ffe7d3", "#e3e3f2", "#ddc7c1", "#cbc8c9", "#d7d8bf", "#006E6D"}
        Dim CheckList As New List(Of String), FirstPass As Boolean = True
        ' Key is WWFF state field. Value is PnP state field
        Dim StateNames As New Dictionary(Of String, String) From
            {
                {"AN-VK0", "VK0"},
                {"VK-ACT", "VK1"},
                {"VK-NSW", "VK2"},
                {"VK-VIC", "VK3"},
                {"VK-QLD", "VK4"},
                {"VK-SA", "VK5"},
                {"VK-WA", "VK6"},
                {"VK-TAS", "VK7"},
                {"VK-NT", "VK8"},
                {"VK9L", "VK9"},
                {"VK9M", "VK9"},
                {"VK9N", "VK9"},
                {"VK9W", "VK9"},
                {"VK9C", "VK9"},
                {"VK9X", "VK9"},
                {"VK0H", "VK0"},
                {"VK0M", "VK0"},
                {"ZL", "ZL"},
                {"ZL7", "ZL"},
                {"ZL8", "ZL"},
                {"ZL9", "ZL"},
                {"AN-ZL5", "ZL"},
                {"E5-N", "ZL"},
                {"ZK3", "ZL"},
                {"E6", "ZL"}
            }

        If WWFFCheck.ShowDialog = DialogResult.OK Then
            Application.UseWaitCursor = True
            Dim logWriter As New System.IO.StreamWriter("WWFFcrosscheck.html", False)
            Dim htmlFile As String = CType(logWriter.BaseStream, FileStream).Name

            logWriter.WriteLine(String.Format("Cross-check of WWFF directory against Parks&Peaks data {0}<br><br>", Now))
            logWriter.WriteLine("<table border=1><tr><th>reference</th><th>WWFF Name</th><th>errors</th></tr>")

            connect = New SQLiteConnection(PARKSdb)
            connect.Open()  ' open database
            sqlWWFF = connect.CreateCommand
            sqlPnP = connect.CreateCommand

            sqlWWFF.CommandText = "SELECT * FROM WWFF WHERE status IN ('active','Active') AND reference LIKE 'VKFF-%' ORDER BY reference ASC"      ' scan directory
            sqlWWFFdr = sqlWWFF.ExecuteReader()
            While sqlWWFFdr.Read()
                count += 1
                errList.Clear()
                WWFFID = sqlWWFFdr.Item("reference")
                SetText(TextBox1, $"Checking {WWFFID }/{count }")
                LatWWFF = CSng(sqlWWFFdr.Item("latitude").ToString)
                LonWWFF = CSng(sqlWWFFdr.Item("longitude").ToString)
                NameWWFF = sqlWWFFdr.Item("name")
                continentWWFF = sqlWWFFdr.Item("continent") ' always OC
                dxccWWFF = sqlWWFFdr.Item("dxcc")           ' only VK or ZL
                stateWWFF = sqlWWFFdr.Item("state")         ' VK-NSW, ....
                statusWWFF = sqlWWFFdr.Item("status")
                Dim websiteWWFF As String = sqlWWFFdr.Item("website").ToString
                Dim notesWWFF As String = sqlWWFFdr.Item("notes").ToString

                sqlPnP.CommandText = $"SELECT * FROM parks WHERE WWFFID='{WWFFID }'"      ' find the matching PnP entry
                sqlPnPdr = sqlPnP.ExecuteReader()
                found = False
                While sqlPnPdr.Read()
                    found = True
                    LatPnP = CSng(sqlPnPdr.Item("Latitude").ToString)
                    LonPNP = CSng(sqlPnPdr.Item("Longitude").ToString)
                    type = sqlPnPdr.Item("type").ToString
                    NamePnP = sqlPnPdr.Item("Name").ToString & " " & LongNames(type)
                    dxccPnP = sqlPnPdr.Item("DXCC")
                    statePnP = sqlPnPdr.Item("State")
                    statusPnP = sqlPnPdr.Item("Status")
                    Dim websitePnP As String = sqlPnPdr.Item("HTTPLink").ToString
                    Dim notesPnP As String = sqlPnPdr.Item("Notes").ToString

                    If WWFFCheck.CheckBox1.Checked Or WWFFCheck.CheckBoxAll.Checked Then
                        If FirstPass Then CheckList.Add(WWFFCheck.CheckBox1.Text)
                        If NameWWFF <> NamePnP Then
                            errList.Add(String.Format("<span style='background-color:{0}'>Name mismatch ({1})</span>", pastels(6), NamePnP))
                        End If
                    End If

                    If WWFFCheck.CheckBox2.Checked Or WWFFCheck.CheckBoxAll.Checked Then
                        If FirstPass Then CheckList.Add(WWFFCheck.CheckBox2.Text)
                        ' Remove any scheme so we can compare paths
                        If (websiteWWFF IsNot Nothing) Then
                            websiteWWFF = Replace(websiteWWFF, "http://", "")
                            websiteWWFF = Replace(websiteWWFF, "https://", "")
                            websitePnP = Replace(websitePnP, "http://", "")
                            websitePnP = Replace(websitePnP, "https://", "")
                            If String.IsNullOrEmpty(websiteWWFF) Then websiteWWFF = "<blank>"
                            If String.IsNullOrEmpty(websitePnP) Then websitePnP = "<blank>"
                            If websiteWWFF.Length > 0 And websiteWWFF <> "-" And websiteWWFF <> websitePnP Then
                                errList.Add(String.Format("<span style='background-color:{0}'>Websites differ (WWFF) {1} - (PnP) {2}</span>", pastels(4), HttpUtility.HtmlEncode(websiteWWFF), HttpUtility.HtmlEncode(websitePnP)))
                            End If
                        End If
                        ' Test if site exists
                        'If Len(websiteWWFF) > 4 Then
                        '    If Not websiteWWFF.StartsWith("http") Then websiteWWFF = "http://" & websiteWWFF     ' add default scheme
                        '    Dim req As System.Net.WebRequest
                        '    Dim res As System.Net.WebResponse
                        '    req = System.Net.WebRequest.Create(websiteWWFF)
                        '    req.Timeout = 30 * 1000
                        '    req.Method = "HEAD"        ' only retrieve headers, not body
                        '    Try
                        '        res = req.GetResponse()
                        '    Catch ex As WebException
                        '        Dim stat As WebExceptionStatus = ex.Status
                        '        If stat = WebExceptionStatus.ProtocolError Then
                        '            Dim httpResponse As HttpWebResponse = CType(ex.Response, HttpWebResponse)
                        '            If httpResponse.StatusCode = 404 Then
                        '                errList.Add(String.Format("<span style='background-color:{0}'>WWFF website not found {1}", pastels(8), HttpUtility.HtmlEncode(websiteWWFF)))
                        '            Else
                        '                Dim i As Integer = 1
                        '            End If
                        '        End If
                        '    End Try
                        'End If
                    End If

                    If WWFFCheck.CheckBox3.Checked Or WWFFCheck.CheckBoxAll.Checked Then
                        If FirstPass Then CheckList.Add(WWFFCheck.CheckBox3.Text)
                        If String.IsNullOrEmpty(notesPnP) Then notesPnP = "<blank>"
                        If notesWWFF.Length > 0 And notesWWFF <> "-" And notesWWFF <> notesPnP Then
                            errList.Add(String.Format("<span style='background-color:{0}'>Notes differ (WWFF) {1} - (PnP) {2}</span>", pastels(7), HttpUtility.HtmlEncode(notesWWFF), HttpUtility.HtmlEncode(notesPnP)))
                        End If
                    End If

                    If WWFFCheck.CheckBox4.Checked Or WWFFCheck.CheckBoxAll.Checked Then
                        If FirstPass Then CheckList.Add(WWFFCheck.CheckBox4.Text)
                        If Not StateNames.ContainsKey(stateWWFF) Then
                            errList.Add(String.Format("<span style='background-color:{0}'>WWFF state {1} not found in StateNames</span>", pastels(2), stateWWFF))
                        Else
                            If StateNames(stateWWFF) <> statePnP Then errList.Add($"<span style='background-color:{pastels(2) }'>State mismatch (WWFF={stateWWFF } vs PnP={statePnP })</span>")
                        End If
                    End If

                    If WWFFCheck.CheckBox5.Checked Or WWFFCheck.CheckBoxAll.Checked Then
                        If FirstPass Then CheckList.Add(WWFFCheck.CheckBox5.Text)
                        Dim DXCCcontinent = continentWWFF & "/" & stateWWFF
                        If DXCCcontinent <> Replace(dxccPnP, " ", "") Then
                            errList.Add(String.Format("<span style='background-color:{0}'>dxcc mismatch (WWFF={1} vs PnP={2})</span>", pastels(1), DXCCcontinent, dxccPnP))
                        End If
                    End If

                    If WWFFCheck.CheckBox6.Checked Or WWFFCheck.CheckBoxAll.Checked Then
                        If FirstPass Then CheckList.Add(WWFFCheck.CheckBox6.Text)
                        ' Check for single missing lat/lon
                        If LatWWFF = 0 And LonWWFF <> 0 Then errList.Add($"<p style='background-color:{pastels(0) }'>Latitude missing from WWFF</p>") Else
                        If LatWWFF <> 0 And LonWWFF = 0 Then errList.Add($"<p style='background-color:{pastels(0) }'>Longitude missing from WWFF</p>") Else
                        If LatPnP = 0 And LonPNP <> 0 Then errList.Add($"<p style='background-color:{pastels(0) }'>Latitude missing from PnP</p>") Else
                        If LatPnP <> 0 And LonPNP = 0 Then errList.Add($"<p style='background-color:{pastels(0) }'>Longitude missing from PnP</p>") Else
                        ' Calculate the distance between centroids
                        Dim WWFFpoint As New MapPoint(LatWWFF, LonWWFF, SpatialReferences.Wgs84)
                        Dim PnPpoint As New MapPoint(LatPnP, LonPNP, SpatialReferences.Wgs84)
                        WWFFpoint = GeometryEngine.Project(WWFFpoint, SpatialReferences.WebMercator)
                        PnPpoint = GeometryEngine.Project(PnPpoint, SpatialReferences.WebMercator)
                        Dim distance = GeometryEngine.DistanceGeodetic(WWFFpoint, PnPpoint, LinearUnits.Kilometers, AngularUnits.Degrees, GeodeticCurveType.Geodesic).Distance
                        If distance > MinDistance Then
                            errList.Add($"<span style='background-color:{pastels(3) }'>Distance between centroids ({Int(distance) }) exceeds {MinDistance } km</span>")
                        End If
                    End If

                    If WWFFCheck.CheckBox7.Checked Then
                        If FirstPass Then CheckList.Add(WWFFCheck.CheckBox7.Text)
                        If (statusPnP.ToLower <> statusWWFF) Then errList.Add(String.Format("<span style='background-color:{0}'>Status differs - WWFF={1} vs PnP={2}</span>", pastels(8), statusWWFF, statusPnP))
                    End If

                    FirstPass = False
                End While
                If Not found Then
                    errList.Add(String.Format("<span style='background-color:{0}'>Park not found in Parks&amp;Peaks</span>", pastels(5)))
                End If

                If errList.Any Then
                    errors += errList.Count
                    errStr = String.Join("<br>", errList.ToArray())
                    logWriter.WriteLine(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>", WWFFID, NameWWFF, errStr))
                    logWriter.Flush()
                End If
                sqlPnPdr.Close()
            End While

            sqlWWFFdr.Close()
            connect.Close()
            logWriter.WriteLine($"</table><br>Total of {errors } errors in {count } parks")
            logWriter.WriteLine($"<br>Checks performed {String.Join(",", CheckList.ToArray) }")
            logWriter.Close()
            SetText(TextBox1, $"Done: Total of {errors } errors in {count } parks")
            Application.UseWaitCursor = False
            Process.Start(htmlFile)         ' display file
        End If
    End Sub

    Private Async Sub GenerateKMLForParkToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles GenerateKMLForParkToolStripMenuItem.Click
        ' Create KML for single selected park
        Dim sql As SQLiteCommand
        Dim SQLdr As SQLiteDataReader
        Dim SOTAsql As SQLiteCommand, parks() As String, last_park As String

        Try
            Dim park As String = InputBox("Enter park reference, or range separated by semicolon")
            park = UCase(park)
            parks = park.Split(";")     ' split into potentially range of parks
            Dim first() As String = parks(0).Split("-")  ' split into prefix and number

            Dim this_park As String = parks(0)    ' start at first park
            Dim this_number As Integer = first(1)            ' current park number
            If parks.Length = 1 Then last_park = parks(0) Else last_park = parks(1)  ' stop at this park

            If park.Length > 0 Then
                While this_park <= last_park
                    Using connect As New SQLiteConnection(PARKSdb), SOTAconnect As New SQLiteConnection(SOTAdb)
                        connect.Open()  ' open database
                        sql = connect.CreateCommand

                        SOTAconnect.Open()  ' open database
                        SOTAsql = SOTAconnect.CreateCommand
                        ' select all parks in this state with a GISID
                        sql.CommandText = $"SELECT * FROM parks WHERE WWFFID='{this_park }'"
                        SQLdr = sql.ExecuteReader()
                        While SQLdr.Read()
                            TextBox1.Text = $"Generating {this_park }"
                            Application.DoEvents()
                            Await ParkToKML(this_park, SOTAsql, False).ConfigureAwait(False)
                        End While
                        connect.Close()
                        SOTAconnect.Close()
                    End Using
                    this_number += 1
                    this_park = $"{first(0) }-{this_number:0000}"
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
        End Try
        SetText(TextBox1, "Done")
    End Sub

    Private Sub GenerateSOTAKMLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateKMLToolStripMenuItem1.Click
        ' Generate SOTA position KML files for selected states
        ' SOTA summit source data only has 4 dec. places, so replicate that precision
        Dim count As Integer = 0    ' count of summits
        Dim SOTAconnect As SQLiteConnection ' declare the connection
        Dim SOTAsql As SQLiteCommand
        Dim SOTAsqldr As SQLiteDataReader
        Dim SOTA1sql As SQLiteCommand
        Dim SOTA1sqldr As SQLiteDataReader
        Dim i As Integer
        Dim state As String, extent As Envelope, files As New List(Of StreamWriter)
        Dim ParkList As New List(Of String)

        With Form3
            If .ShowDialog() = DialogResult.OK Then
                Dim logWriter As New System.IO.StreamWriter("log.txt", True)
                Dim htmlFile As String = CType(logWriter.BaseStream, FileStream).Name
                logWriter.WriteLine(String.Format("{0} - KML generation started", DateTime.Now.ToString(YMDHMS)))
                Application.UseWaitCursor = True
                SOTAconnect = New SQLiteConnection(SOTAdb)
                SOTAconnect.Open()  ' open database
                SOTAsql = SOTAconnect.CreateCommand
                SOTA1sql = SOTAconnect.CreateCommand
                For i = 0 To .ctrls.Count - 2
                    Dim box As KeyValuePair(Of String, CheckBox) = .ctrls.Item(i)    ' extract the checkbox
                    If box.Value.Checked Or .ctrls.Item(.ctrls.Count - 1).Value.Checked Then
                        ' Create kml file for this state
                        state = box.Key
                        count = 0
                        Dim BaseFilename As String = String.Format("{0}\files\SOTA-{1}", Application.StartupPath, state)
                        TextBox2.AppendText($"creating SOTA KML for {state }{vbCrLf }")
                        Dim RemoteWriter As New System.IO.StreamWriter(BaseFilename & ".kml", False)
                        Dim LocalWriter As New System.IO.StreamWriter(BaseFilename & "_local.kml", False)
                        files.Clear()
                        files.Add(RemoteWriter)
                        files.Add(LocalWriter)

                        Dim pins As String          ' folder for pin icons
                        For Each f In files
                            If CType(f.BaseStream, FileStream).Name.Contains("_local") Then pins = "pins/" Else pins = PnPurl & "pins/" ' link to pins
                            f.WriteLine(KMLheader)
                            ' Draw all the regions
                            f.WriteLine("<Folder>")
                            f.WriteLine("<name>Regions</name>")
                            f.WriteLine("<visibility>0</visibility>")
                            f.WriteLine("<Style id='region'>")
                            f.WriteLine("<LineStyle><width>2</width><colorMode>random</colorMode></LineStyle>")
                            f.WriteLine("<PolyStyle><color>7f0000ff</color><fill>0</fill></PolyStyle>")
                            f.WriteLine("<LabelStyle><color>8078FA3C</color><scale>3</scale></LabelStyle>")
                            f.WriteLine("<IconStyle><scale>0</scale></IconStyle>")
                            f.WriteLine("</Style>")
                            extent = Nothing        ' Extent of all summits
                            SOTAsql.CommandText = "SELECT * FROM Region WHERE Like('" & state & "%',RegionCode)=1 ORDER BY RegionCode"
                            SOTAsqldr = SOTAsql.ExecuteReader()
                            While SOTAsqldr.Read()
                                Dim RegionCode As String = SOTAsqldr.Item("RegionCode")
                                Dim LatMax As Double = SOTAsqldr.Item("LatMax")
                                Dim LatMin As Double = SOTAsqldr.Item("LatMin")
                                Dim LonMax As Double = SOTAsqldr.Item("LonMax")
                                Dim LonMin As Double = SOTAsqldr.Item("LonMin")
                                Dim env As New Envelope(New MapPoint(LonMin, LatMin, SpatialReferences.Wgs84), New MapPoint(LonMax, LatMax, SpatialReferences.Wgs84))
                                extent = EnvelopeUnion(extent, env)        ' add region boundary to extent
                                Dim polygon As Polygon = GeometryEngine.Buffer(env, 0)      ' convert to polygon
                                f.WriteLine("<Placemark>")
                                f.WriteLine("<styleUrl>#region</styleUrl>")
                                f.WriteLine("<ExtendedData>")
                                f.WriteLine("<Data name='Association'><value>{0}</value></Data>", Replace(XMLencode(SOTAsqldr.Item("AssociationName").ToString), " ", "&amp;nbsp;"))
                                f.WriteLine("<Data name='Region'><value>{0}</value></Data>", Replace(XMLencode(SOTAsqldr.Item("RegionName").ToString), " ", "&amp;nbsp;"))
                                f.WriteLine("<Data name='North'><value>{0:f4}</value></Data>", LatMax)
                                f.WriteLine("<Data name='South'><value>{0:f4}</value></Data>", LatMin)
                                f.WriteLine("<Data name='West'><value>{0:f4}</value></Data>", LonMin)
                                f.WriteLine("<Data name='East'><value>{0:f4}</value></Data>", LonMax)
                                Dim area As Long = GeometryEngine.AreaGeodetic(extent) / 10000
                                f.WriteLine("<Data name='Area (ha)'><value>{0}</value></Data>", area)
                                f.WriteLine("</ExtendedData>")
                                f.WriteLine("<name>{0}</name><visibility>0</visibility>", RegionCode)   ' regions initially invisible
                                f.WriteLine("<MultiGeometry>")
                                Dim labelPoint As MapPoint = env.GetCenter
                                f.WriteLine("<Point><coordinates>{0:f4},{1:f4}</coordinates></Point>", labelPoint.X, labelPoint.Y)
                                f.WriteLine("<Polygon>")
                                f.WriteLine("<outerBoundaryIs>")
                                f.WriteLine("<LinearRing><coordinates>")
                                For Each segment In polygon.Parts(0)
                                    f.WriteLine("{0:f4},{1:f4}", segment.StartPoint.X, segment.StartPoint.Y)
                                Next
                                With polygon.Parts(0)
                                    f.WriteLine("{0:f4},{1:f4}", .StartPoint.X, .StartPoint.Y)    ' close polygon
                                End With
                                f.WriteLine("</coordinates></LinearRing>")
                                f.WriteLine("</outerBoundaryIs>")
                                f.WriteLine("</Polygon>")
                                f.WriteLine("</MultiGeometry>")
                                If extent IsNot Nothing Then
                                    ' Add LookAt for this region
                                    Dim Range As Double = CalcRange(extent)
                                    Dim LookAtPoint As MapPoint = extent.GetCenter      ' center of extent
                                    If state = "ZL" Then
                                        ' the ZL envelop crosses date line, so LookAtPoint is wrong side of earth
                                        LookAtPoint = New MapPoint(LookAtPoint.X + 180, LookAtPoint.Y)
                                    End If
                                    f.WriteLine("<LookAt><longitude>{0:f4}</longitude><latitude>{1:f4}</latitude><range>{2:f0}</range><heading>0</heading><tilt>0</tilt></LookAt>", LookAtPoint.X, LookAtPoint.Y, Range)
                                End If
                                f.WriteLine("</Placemark>")
                            End While
                            f.WriteLine("</Folder>")
                            SOTAsqldr.Close()

                            ' Create styles for pins
                            For p As Integer = 1 To 10
                                Dim styles() As String = {"", "_caution", "_park"}      ' 3 style variants
                                For Each s In styles
                                    f.WriteLine("<Style id='{1}{2}_normal'><IconStyle><Icon><href>{0}{1}{2}.png</href></Icon></IconStyle><LabelStyle><scale>0</scale></LabelStyle></Style>", pins, p, s)
                                    f.WriteLine("<Style id='{1}{2}_highlight'><IconStyle><Icon><href>{0}{1}{2}.png</href></Icon></IconStyle><LabelStyle><scale>1</scale></LabelStyle></Style>", pins, p, s)
                                    f.WriteLine("<StyleMap id='{0}{1}'>", p, s)
                                    f.WriteLine("<Pair><key>normal</key><styleUrl>#{0}{1}_normal</styleUrl></Pair>", p, s)
                                    f.WriteLine("<Pair><key>highlight</key><styleUrl>#{0}{1}_highlight</styleUrl></Pair>", p, s)
                                    f.WriteLine("</StyleMap>")
                                Next
                            Next
                            ' Create placemarks for all summits
                            extent = Nothing
                            SOTAsql.CommandText = "SELECT * FROM `SOTA` LEFT JOIN `GEheight` ON `SOTA`.`SummitCode`=`GEheight`.`SummitCode` WHERE Like('" & state & "%',`SOTA`.`SummitCode`)=1 ORDER BY `SOTA`.`SummitCode`"
                            SOTAsqldr = SOTAsql.ExecuteReader()
                            While SOTAsqldr.Read()
                                f.WriteLine("<Placemark>")
                                Dim SummitCode As String = SOTAsqldr.Item("SummitCode")
                                Dim points As Integer = SOTAsqldr.Item("Points")
                                Dim longitude As Double = SOTAsqldr.Item("GridRef1")
                                Dim latitude As Double = SOTAsqldr.Item("GridRef2")
                                Dim height As Integer = SOTAsqldr.Item("AltM")
                                Dim GEheight As Integer = SOTAsqldr.Item("elevation")
                                Dim GEresolution As Integer = SOTAsqldr.Item("resolution")
                                extent = EnvelopeUnion(extent, New MapPoint(longitude, latitude, SpatialReferences.Wgs84))
                                f.WriteLine($"<name>{SummitCode } - {SOTAsqldr.Item("SummitName") }</name>")
                                f.WriteLine("<description><![CDATA[<style>table, th, td {white-space:nowrap; }</style>")
                                f.WriteLine("<table border='1'><tr><th>Item</th><th>Value</th></tr>")
                                f.WriteLine($"<tr><td>Summit Code</td><td>{SummitCode }</td></tr>")
                                f.WriteLine($"<tr><td>Summit Name</td><td>{SOTAsqldr.Item("SummitName") }</td></tr>")
                                f.WriteLine($"<tr><td>Association Name</td><td>{SOTAsqldr.Item("AssociationName") }</td></tr>")
                                f.WriteLine($"<tr><td>Region Name</td><td>{SOTAsqldr.Item("RegionName") }</td></tr>")
                                f.WriteLine($"<tr><td>Altitude (m)</td><td>{height }</td></tr>")
                                f.WriteLine($"<tr><td>Longitude</td><td>{longitude:f4}</td></tr>")
                                f.WriteLine($"<tr><td>Latitude</td><td>{latitude:f4}</td></tr>")
                                f.WriteLine($"<tr><td>Points</td><td>{points }</td></tr>")
                                f.WriteLine($"<tr><td>Bonus Points</td><td>{SOTAsqldr.Item("BonusPoints") }</td></tr>")
                                f.WriteLine("<tr><td>{0}</td><td>{1}</td></tr>", "Valid from", SOTAsqldr.Item("ValidFrom"))
                                f.WriteLine("<tr><td>{0}</td><td>{1}</td></tr>", "Valid to", SOTAsqldr.Item("ValidTo"))
                                f.WriteLine("<tr><td>{0}</td><td>{1}m/{2}m</td></tr>", "GE height/resolution", GEheight, GEresolution)
                                ' Make list of parks with a relationship with this summit
                                SOTA1sql.CommandText = $"SELECT * FROM SummitsInParks WHERE SummitCode='{SummitCode }' ORDER BY WWFFID"
                                ParkList.Clear()
                                SOTA1sqldr = SOTA1sql.ExecuteReader()
                                Dim style As String = $"#{points }"        ' assume normal style
                                While SOTA1sqldr.Read()
                                    Dim s As String = ""
                                    If SOTA1sqldr.Item("WithIn").ToString = "Y" Then
                                        s = "Within "
                                        style = $"#{points }_park"
                                    Else
                                        s = "Edge "
                                        style = $"#{points }_caution"
                                    End If
                                    s += SOTA1sqldr.Item("WWFFID").ToString
                                    ParkList.Add(s)
                                End While
                                If ParkList.Any Then
                                    f.WriteLine("<tr><td>{0}</td><td>{1}</td></tr>", "Parks", Join(ParkList.ToArray, "<br>"))
                                End If
                                SOTA1sqldr.Close()
                                f.WriteLine("</table>]]></description>")
                                f.WriteLine($"<Point><coordinates>{longitude:f4},{latitude:f4},{height }</coordinates></Point>")
                                f.WriteLine("<styleUrl>{0}</styleUrl>", style)
                                f.WriteLine("</Placemark>")
                                count += 1
                            End While
                            SOTAsqldr.Close()

                            If extent IsNot Nothing Then
                                ' Add LookAt for this state
                                Dim Range As Double = CalcRange(extent)
                                Dim LookAtPoint As MapPoint = extent.GetCenter      ' center of extent
                                If state = "ZL" Then
                                    ' the ZL envelop crosses date line, so LookAtPoint is wrong side of earth
                                    LookAtPoint = New MapPoint(LookAtPoint.X + 180, LookAtPoint.Y)
                                End If
                                f.WriteLine($"<LookAt><longitude>{LookAtPoint.X:f4}</longitude><latitude>{LookAtPoint.Y:f4}</latitude><range>{Range:f0}</range><heading>0</heading><tilt>0</tilt></LookAt>")
                            End If

                            ' now create Activation Zone (AZ) folder
                            f.WriteLine("<Folder>")
                            f.WriteLine("<name>Activation Zones</name>")
                            f.WriteLine("<open>1</open>")
                            f.WriteLine("<description><![CDATA[<b>These elements will flood fill an area 25m below a SOTA summit indicating the approximate Activation Zone.</b><br><br>It comes with many disclaimers, caveats and cautions.
    <ol>
<li>It uses elevation data intrinsic to Google Earth. Its accuracy varies depending on location.</li>
<li>It is based on SOTA data. If that data is incorrect, i.e. the summit height is incorrect, then the flooded area will be incorrect. If you notice any errors, please approach your region SOTA rep.</li>
<li>The outline of the flooded area is only approximate, and as good as the elevation data. You must check the true extent on the ground with a GPS.</li>
</ol><br>
Double click the summit to fly there.<br>
Select the checkbox (make visible) to see the activation zone.<br><br>
Note that the AZ is only that area which is contiguous with the summit, i.e. the area within a 25m contour line below the summit. Any nearby ""islands"" of land are not part of the AZ.<br><br>
If there is no activation zone shown, and the summit appears below the plane, then the most likely explanation is that the summit height is incorrect.
If needed, you can move the AZ plane up and down. Go to Right mouse | Properties | Altitude and you can enter a new value for the height.
This allows you to view a relative AZ which is quite useful.
]]></description>")
                            ' Create style for polygons
                            f.WriteLine("<Style id=""AZ"">")
                            f.WriteLine("<PolyStyle>")
                            f.WriteLine("<color>c0ff0000</color>")  ' blue
                            f.WriteLine("</PolyStyle>")
                            f.WriteLine("</Style>")
                            SOTAsql.CommandText = "SELECT * FROM SOTA WHERE Like('" & state & "%',SummitCode)=1 ORDER BY SummitCode"
                            SOTAsqldr = SOTAsql.ExecuteReader()
                            While SOTAsqldr.Read()
                                Dim Height As Integer = CInt(SOTAsqldr.Item("AltM")) - 25
                                Dim Code As String = SOTAsqldr.Item("SummitCode")
                                Dim Name As String = SOTAsqldr.Item("SummitName")
                                Dim longitude As Double = SOTAsqldr.Item("GridRef1")
                                Dim latitude As Double = SOTAsqldr.Item("GridRef2")
                                Dim summit As New MapPoint(longitude, latitude, SpatialReferences.Wgs84)
                                Dim buffer As Geometry = GeometryEngine.BufferGeodetic(summit, 2500, LinearUnits.Meters)    ' create circular buffer around summit
                                f.WriteLine("<Placemark>")
                                f.WriteLine($"<name>{Code } - {Name }</name>")
                                f.WriteLine("<visibility>0</visibility>")
                                f.WriteLine("<styleUrl>#AZ</styleUrl>")
                                f.WriteLine("<LookAt>")
                                f.WriteLine($"<longitude>{longitude:f4}</longitude>")
                                f.WriteLine($"<latitude>{latitude:f4}</latitude>")
                                f.WriteLine("<altitude>0</altitude>")
                                f.WriteLine("<range>5000</range>")
                                f.WriteLine("<heading>0</heading>")
                                f.WriteLine("<tilt>0</tilt>")
                                f.WriteLine("</LookAt>")
                                f.WriteLine("<Polygon>")
                                f.WriteLine("<extrude>0</extrude>")
                                f.WriteLine("<altitudeMode>absolute</altitudeMode>")
                                f.WriteLine("<tessellate>0</tessellate>")
                                f.WriteLine("<outerBoundaryIs><LinearRing>")
                                f.WriteLine("<coordinates>")
                                f.WriteLine($"{buffer.Extent.XMin:f4},{buffer.Extent.YMin:f4},{Height } {buffer.Extent.XMin:f4},{buffer.Extent.YMax:f4},{Height } {buffer.Extent.XMax:f4},{buffer.Extent.YMax:f4},{Height } {buffer.Extent.XMax:f4},{buffer.Extent.YMin:f4},{Height } {buffer.Extent.XMin:f4},{buffer.Extent.YMin:f4},{Height }")
                                f.WriteLine("</coordinates>")
                                f.WriteLine("</LinearRing></outerBoundaryIs>")
                                f.WriteLine("</Polygon>")
                                f.WriteLine("</Placemark>")
                            End While
                            SOTAsqldr.Close()
                            f.WriteLine("</Folder>")
                            f.WriteLine(KMLfooter)
                            f.Close()
                        Next

                        ' compress to zip file
                        System.IO.File.Delete(BaseFilename & ".kmz")
                        Dim zip As ZipArchive = ZipFile.Open(BaseFilename & ".kmz", ZipArchiveMode.Create)    ' create new archive file
                        zip.CreateEntryFromFile(BaseFilename & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
                        zip.Dispose()
                        Dim msg As String = $"{count / 2 } SOTA summits generated"
                        TextBox2.AppendText($"{msg }{vbCrLf }")
                        logWriter.WriteLine(msg)
                        TextBox2.AppendText($"Generated data in {BaseFilename }.kmz{vbCrLf }")
                    End If
                Next
                Application.UseWaitCursor = False
                logWriter.Close()
                SOTAconnect.Close()
            End If
        End With
    End Sub

    Shared Function EnvelopeUnion(extent As Envelope, env As Envelope) As Envelope
        ' Perform a union of envelopes
        ' If currently nothing, then inialise
        ' If not empty, then union env
        Dim result As Envelope

        If extent Is Nothing OrElse extent.IsEmpty Then
            result = env.Extent
        ElseIf env.IsEmpty Then
            result = extent.Extent
        Else
            Contract.Requires(extent.SpatialReference.Equals(env.SpatialReference), "Spatial reference must be same")
            result = GeometryEngine.CombineExtents(extent, env)
        End If
        Return result
    End Function

    Shared Function EnvelopeUnion(extent As Envelope, env As MapPoint) As Envelope
        ' Perform a union of envelope and mappoint

        Dim EnvPnt As Envelope = env.Extent     '  Convert to envelope
        Return EnvelopeUnion(extent, EnvPnt)
    End Function

    Private Sub MakeIconsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeIconsToolStripMenuItem.Click
        ' Make a series of icons for SOTA summits
        ' Uses the same colours as used in SOTA mapping
        Dim i As Integer
        Dim colors As New List(Of Color) From {                             ' color of pins
                Color.FromArgb(255, 76, 76, 255),       ' 1
                Color.FromArgb(255, 79, 159, 255),      ' 2
                Color.FromArgb(255, 76, 233, 255),      ' 3
                Color.FromArgb(255, 76, 255, 188),      ' 4
                Color.FromArgb(255, 82, 255, 117),      ' 5
                Color.FromArgb(255, 117, 255, 76),      ' 6
                Color.FromArgb(255, 195, 255, 76),      ' 7
                Color.FromArgb(255, 255, 228, 78),      ' 8
                Color.FromArgb(255, 255, 159, 79),      ' 9
                Color.FromArgb(255, 255, 79, 79)}       ' 10
        Dim triangle As PointF() = {New PointF(0F, 63.0F), New PointF(63.0F, 63.0F), New PointF(32.0F, 0F), New PointF(0F, 63.0F)}  ' a triangle
        Dim drawRect As New Rectangle(0, 25.0F, 64.0F, 40.0F)
        Dim BaseFilename As String = $"{ Application.StartupPath }\files\pins\"

        Dim count As Integer = 0
        Dim TrianglePen As Pen, SaveFile As String

        Using drawFont As New Font("Arial", 24, FontStyle.Bold),
             drawBrush As New SolidBrush(Color.Black),
             drawFormat As New StringFormat,
             RedBrush As New SolidBrush(Color.Red),
             RedPen As New Pen(RedBrush, 5),
             GreenBrush As New SolidBrush(Color.Green),
             GreenPen As New Pen(GreenBrush, 5)

            drawFormat.Alignment = StringAlignment.Center
            drawFormat.LineAlignment = StringAlignment.Center
            ' Make pin for each summit value
            For i = 1 To colors.Count
                For style = 1 To 3
                    ' Make 3 styles of pin
                    ' 1 = normal, 2 = caution, 3 = in park
                    Dim bm As New Bitmap(64, 64, Imaging.PixelFormat.Format32bppArgb)      ' create a transparent bitmap
                    Dim gr As Graphics = Graphics.FromImage(bm)
                    Select Case style
                        Case 2
                            SaveFile = $"{BaseFilename }{i }_caution.png"
                            TrianglePen = RedPen
                        Case 3
                            SaveFile = $"{BaseFilename }{i }_park.png"
                            TrianglePen = GreenPen
                        Case Else
                            SaveFile = $"{BaseFilename }{i }.png"
                            TrianglePen = Nothing
                    End Select
                    ' Draw the coloured triangle
                    Dim br As Brush = New SolidBrush(colors(i - 1))
                    gr.FillPolygon(br, triangle)
                    ' draw triangle boundary
                    If TrianglePen IsNot Nothing Then gr.DrawPolygon(TrianglePen, triangle)
                    ' Draw the point value label
                    gr.DrawString(CStr(i), drawFont, drawBrush, drawRect, drawFormat)
                    bm.Save(SaveFile, Imaging.ImageFormat.Png)
                    count += 1
                    br.Dispose()
                    bm.Dispose()
                Next
            Next
            Dim msg As String = $"{count } icons generated"
            SetText(TextBox1, msg)
        End Using
    End Sub

    Private Sub FixRegionTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FixRegionTableToolStripMenuItem.Click
        ' Add AssociationName and RegionName fields to Region table
        Dim SOTAsql As SQLiteCommand, SOTA1sql As SQLiteCommand
        Dim SOTAsqldr As SQLiteDataReader

        Using SOTAconnect As New SQLiteConnection(SOTAdb)
            SOTAconnect.Open()  ' open database
            SOTAsql = SOTAconnect.CreateCommand
            SOTA1sql = SOTAconnect.CreateCommand
            SOTAsql.CommandText = "select distinct RegionCode,sota.AssociationName,sota.RegionName from Region join SOTA where sota.SummitCode like region.RegionCode || '%'"
            SOTAsqldr = SOTAsql.ExecuteReader()
            While SOTAsqldr.Read()
                SOTA1sql.CommandText = $"UPDATE Region SET AssociationName=""{SOTAsqldr.Item("AssociationName")}"",RegionName=""{SOTAsqldr.Item("RegionName")}"" WHERE RegionCode=""{SOTAsqldr.Item("RegionCode") }"""
                SOTA1sql.ExecuteNonQuery()
            End While
            SOTAconnect.Close()
        End Using
    End Sub

    Shared Function PolygonArea(polygon As ReadOnlyPart) As Double
        ' Calculate the area of a polygon using the 'Shoelace' or Gauss's formula
        ' https://en.wikipedia.org/wiki/Shoelace_formula
        ' if result <0 then CW winding, else CCW
        Contract.Requires(polygon IsNot Nothing AndAlso polygon.Any, "Illegal polygon")
        Dim result As Double = 0
        If polygon.Count > 2 Then           ' ignore degenerate polygon
            For Each s In polygon
                result += s.StartPoint.X * s.EndPoint.Y - s.EndPoint.X * s.StartPoint.Y
            Next
        End If
        Return result / 2
    End Function

    Shared Function KMLColor(alpha As Integer, blue As Integer, green As Integer, red As Integer) As String
        ' Construct a KML color in hex form aabbggrr
        Contract.Requires(alpha >= 0 And alpha <= 255, "Alpha must be between 0 and 255")
        Contract.Requires(blue >= 0 And blue <= 255, "Blue must be between 0 and 255")
        Contract.Requires(green >= 0 And green <= 255, "Green must be between 0 and 255")
        Contract.Requires(red >= 0 And red <= 255, "Red must be between 0 and 255")
        Return $"{alpha:x2}{blue:x2}{green:x2}{red:x2}"
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Initialize all the data sets
        'DataSets.Add("CAPAD_T", New DataSet("CAPAD_T", "F:/GIS Data/CAPAD/2018/CAPAD2018_terrestrial.shp"))
        'DataSets.Add("CAPAD_M", New DataSet("CAPAD_M", "F:/GIS Data/CAPAD/2018/CAPAD2018_marine.shp"))
        DataSets.Add("CAPAD_T", New DataSet("CAPAD_T", "F:/GIS Data/CAPAD/2020/CAPAD2020_terrestrial.shp"))
        DataSets.Add("CAPAD_M", New DataSet("CAPAD_M", "F:/GIS Data/CAPAD/2020/CAPAD2020_marine.shp"))
        DataSets.Add("VIC_PARKS", New DataSet("VIC_PARKS", "F:/GIS Data/PARKRES/parkres.shp"))
        DataSets.Add("ZL", New DataSet("ZL", "F:/GIS Data/NZ/DOC_PublicConservationAreas_2017_06_01.shp"))
    End Sub

    Private Sub FindPARKRESInCAPADToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindPARKRESInCAPADToolStripMenuItem.Click
        ' Find any parks from set VIC_PARKS that exist in CAPAD_T
        Dim count As Integer = 0, total As Integer = 0, found As Integer = 0   ' count of parks
        Dim Name As String, TypeAbbr As String, WWFFID As String, PRIMS_ID As String
        Dim sql As SQLiteCommand
        Dim SQLdr As SQLiteDataReader
        Dim dBaseDataReader As System.Data.OleDb.OleDbDataReader
        Dim where As String

        Using connect As New SQLiteConnection(PARKSdb), logWriter_name As New System.IO.StreamWriter("found_parkres.html", False)
            Dim htmlFileName As String = CType(logWriter_name.BaseStream, FileStream).Name
            connect.Open()  ' open database
            sql = connect.CreateCommand

            ' Get count of parks to process
            sql.CommandText = "SELECT COUNT(DISTINCT WWFFID) AS COUNT FROM GISmapping WHERE DataSet='VIC_PARKS'"
            SQLdr = sql.ExecuteReader()
            While SQLdr.Read()
                total = SQLdr.Item("COUNT")
            End While
            SQLdr.Close()

            logWriter_name.WriteLine("<table border=1>")
            logWriter_name.WriteLine("<tr><th>WWFFID</th><th>Name</th><th>VIC_PARKS id</th><th>CAPAD_T id</td></tr>")

            ' Select all the parks that use PARKRES
            sql.CommandText = "SELECT * FROM parks JOIN GISmapping USING(WWFFID) WHERE GISmapping.DataSet='VIC_PARKS' GROUP BY parks.WWFFID ORDER BY Name"
            SQLdr = sql.ExecuteReader()
            While SQLdr.Read()
                count += 1
                ' name to look for in CAPAD_T
                WWFFID = SQLdr.Item("WWFFID")
                Name = SQLdr.Item("NAME")
                TypeAbbr = SQLdr.Item("Type")
                PRIMS_ID = SQLdr.Item("GISID")
                SetText(TextBox1, $"Searching: {Name } {TypeAbbr } {count }/{total } Found: {found }")
                Application.DoEvents()
                where = $"NAME='{Name }' AND TYPE_ABBR='{TypeAbbr }'"
                Using dBaseCommand As New System.Data.OleDb.OleDbCommand($"SELECT * FROM {DataSets("CAPAD_T").dbfTableName } WHERE {where }", DataSets("CAPAD_T").dbfConnection)
                    dBaseDataReader = dBaseCommand.ExecuteReader()
                    While dBaseDataReader.Read
                        logWriter_name.WriteLine($"<tr><td>{WWFFID }</td><td>{Name } {TypeAbbr }</td><td>{PRIMS_ID }</td><td>{dBaseDataReader("PA_ID") }</td></tr>")
                        logWriter_name.Flush()
                        found += 1
                    End While
                    dBaseDataReader.Close()
                End Using
            End While
            SQLdr.Close()
            logWriter_name.WriteLine("</table>")
            logWriter_name.WriteLine($"<br>{found } found out of total {total }")
            logWriter_name.Close()
            Process.Start(htmlFileName)
        End Using
    End Sub

    '=======================================================================================================
    ' Items on WWFF menu
    '=======================================================================================================
    Private Sub InputCookiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InputCookiesToolStripMenuItem.Click
        Cookies = InputBox("For Firefox use F12 | Network | Headers | Request Headers | Raw Headers", "Enter cookies", "")
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        ' Upload list of new parks to WWFF
        Const uploaded = "uploaded.txt"        ' contains list of previously uploaded parks
        Dim uploadedlist As New List(Of String)       ' list of park names already uploaded
        Dim mandatory As New List(Of String)(New String() {"NAME", "TYPE ABBR", "LATITUDE", "LONGITUDE", "IUCN", "STATE", "HTTPLINK"})   ' mandatory fields
        Dim continent As String = "OC"    ' Oceania
        Dim DXCC As String = "VK"       ' Australia
        ' Allow flexiblity in input to have VK1 or ACT as state
        Dim StateName As New Dictionary(Of String, String) From {
            {"VK1", "VK-ACT"}, {"ACT", "VK-ACT"},
            {"VK2", "VK-NSW"}, {"NSW", "VK-NSW"},
            {"VK3", "VK-VIC"}, {"VIC", "VK-VIC"},
            {"VK4", "VK-QLD"}, {"QLD", "VK-QLD"},
            {"VK5", "VK-SA"}, {"SA", "VK-SA"},
            {"VK6", "VK-WA"}, {"WA", "VK-WA"},
            {"VK7", "VK-TAS"}, {"TAS", "VK-TAS"},
            {"VK8", "VK-NT"}, {"NT", "VK-NT"}
        }
        Dim states As New List(Of String)(New String() {"VK-ACT", "VK-NSW", "VK-NT", "VK-QLD", "VK-SA", "VK-TAS", "VK-VIC", "VK-WA"})
        Dim IUCN As New List(Of String)(New String() {"Cat Ia", "Cat Ib", "Cat II", "Cat III", "Cat IV", "Cat V", "Cat VI", "WHS", "Biosphere", "Natura2000", "Ramsar", "None"})    ' list of IUCN
        Dim columns As New List(Of String), form_fields As New List(Of String)
        Dim header As New List(Of String), value As String
        Dim errors As String
        Dim ref As String           ' park reference number
        Dim name As String          ' park name
        Dim POSTfields As New NameValueCollection()
        Dim wwffco As String = "http://wwff.co/wp-admin/admin.php"      ' WWFF admin page
        Dim uriString As String = "https://wwff.co/wp-admin/admin.php?page=logsearch-manage-directory"      ' WWFF upload page

        If String.IsNullOrEmpty(Cookies) Then
            MsgBox("You must use ""Input cookies"" first", MsgBoxStyle.Critical + vbOKOnly, "Cookies")
        Else
            Using openFileDialog1 As New OpenFileDialog(), myWebClient As New WebClient()
                openFileDialog1.Filter = "csv files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*"
                If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    TextBox2.Clear()
                    Try
                        If Not File.Exists(uploaded) Then
                            ' Create a file to write to.
                            Using sw As StreamWriter = File.CreateText(uploaded)
                            End Using
                        Else
                            ' Read list of previously uploaded parks
                            Dim myloadedlist = System.IO.File.ReadAllLines(uploaded)        ' read file contents (a list)
                            For Each ref In myloadedlist
                                uploadedlist.Add(ref)
                            Next
                        End If
                        Dim afile As New FileIO.TextFieldParser(openFileDialog1.FileName)
                        Dim CurrentRecord As String() ' this array will hold each line of data
                        afile.TextFieldType = FileIO.FieldType.Delimited
                        afile.Delimiters = New String() {","}
                        afile.HasFieldsEnclosedInQuotes = True

                        ' Initialise the WebClient
                        Dim myCache = New CredentialCache From {
                            {New Uri(wwffco), "Basic", New NetworkCredential("vk3ohm", "rubbish")}
                        }
                        myWebClient.Credentials = myCache
                        Dim credentials As String = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes("vk3ohm" & ":" & "rubbish"))
                        myWebClient.Headers.Add("Authorization", $"Basic {credentials }")
                        myWebClient.Headers.Add("Content-type", "application/x-www-form-urlencoded")
                        myWebClient.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
                        myWebClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:60.0) Gecko/20100101 Firefox/60.0")
                        myWebClient.Headers.Add("Accept-Language", "en-US, en;q=0.5")
                        myWebClient.Headers.Add("Referer", WebUtility.UrlEncode(uriString))
                        'myWebClient.Headers.Add("Connection", "keep-alive")
                        myWebClient.Headers.Add("Upgrade-Insecure-Requests", "1")
                        myWebClient.Headers.Add("DNT", "1")
                        myWebClient.Headers.Add("Host", "wwff.co")

                        myWebClient.Headers.Add("Cookie", Cookies)
                        myWebClient.Proxy = Nothing

                        ' parse the actual file
                        Dim count As Integer = 0
                        Do While Not afile.EndOfData
                            Try
                                CurrentRecord = afile.ReadFields
                                errors = ""
                                count += 1
                                If count = 1 Then
                                    ' Check the header
                                    header = CurrentRecord.ToList
                                    ' Upper case header
                                    For i = 0 To header.Count - 1
                                        header(i) = Trim(header(i).ToUpper())
                                    Next
                                    For Each st In mandatory
                                        If Not header.Contains(st) Then
                                            MsgBox("File does not contain an " & st & " column", vbCritical + vbOKOnly, "Mandatory column missing")
                                            Exit Sub
                                        End If
                                    Next
                                Else
                                    Try
                                        ' data line - process it
                                        Dim indx As Integer
                                        errors = ""
                                        With POSTfields
                                            .Clear()
                                            .Add("command", "Create")
                                            .Add("progName", "VKFF")
                                            .Add("dxccName", "VKFF")
                                            .Add("newStatus", "proposed")
                                            .Add("newLocator", "-")
                                            .Add("newRegion", "-")
                                            .Add("newContinent", continent)
                                            .Add("newDxcc", DXCC)
                                        End With
                                        name = ""
                                        For Each col As String In header
                                            If col <> "" Then
                                                indx = header.IndexOf(col)
                                                If indx >= 0 Then
                                                    ' found the column in the header. Validate it and create form POST variables
                                                    value = CurrentRecord(indx)      ' value in column
                                                    Dim svalue As Single
                                                    If IsNumeric(value) Then svalue = CSng(value) Else svalue = 0
                                                    Select Case col
                                                        Case "NAME"
                                                            Dim TypeAbbr As String = CurrentRecord(header.IndexOf("TYPE ABBR"))
                                                            Dim Type = LongNames(TypeAbbr)
                                                            name = Trim(CurrentRecord(header.IndexOf("NAME"))) & " " & Type     ' park reference
                                                            If uploadedlist.Contains(name) Then
                                                                ' skip - already uploaded
                                                                TextBox2.AppendText(value & " previously uploaded - skipped" & vbCrLf)
                                                                Application.DoEvents()
                                                                GoTo Skip
                                                            End If
                                                            If Len(name) > 96 Then errors &= "The Name exceeds 96 characters"
                                                            POSTfields.Add("newName", name)
                                                        Case "STATE"
                                                            For Each kvp In StateName
                                                                If value = kvp.Key Then value = kvp.Value
                                                            Next
                                                            If Not states.IndexOf(value) >= 0 Then errors &= "Invalid State"
                                                            POSTfields.Add("newState", value)     ' State
                                                        Case "LATITUDE"
                                                            If svalue = 0 Or svalue < -90 Or svalue > 90 Then errors &= "Latitude Is illegal"
                                                            POSTfields.Add("newLatitude", svalue)
                                                        Case "LONGITUDE"
                                                            If svalue = 0 Or svalue < -180 Or svalue > 180 Then errors &= "Longitude Is illegal"
                                                            POSTfields.Add("newLongitude", svalue)
                                                        Case "HTTPLINK"
                                                            If Len(value) > 128 Then errors &= "The website exceeds 128 characters"
                                                            If Not ValidUrl(value) Then errors &= "Website malformed"
                                                            value = Replace(value, "http://", "")           ' remove redundant schemes
                                                            value = Replace(value, "https://", "")
                                                            POSTfields.Add("newWebsite", value)
                                                        Case "NOTES"
                                                            If Len(value) > 96 Then errors &= "The notes exceed 96 characters"
                                                            POSTfields.Add("newNotes", value)
                                                        Case "IOTA"
                                                            Dim rgx = New Regex("^(AF|AN|AS|EU|NA|OC|SA)-\d\d\d$")
                                                            If Not rgx.IsMatch(value) Then errors &= "Invalid IOTA"
                                                        Case "IUCN"
                                                            Select Case value
                                                                Case "IA"
                                                                    value = "Cat Ia"
                                                                Case "IB"
                                                                    value = "Cat Ib"
                                                                Case "II"
                                                                    value = "Cat II"
                                                                Case "III"
                                                                    value = "Cat III"
                                                                Case "IV"
                                                                    value = "Cat IV"
                                                                Case "V"
                                                                    value = "Cat V"
                                                                Case "VI"
                                                                    value = "Cat VI"
                                                                Case Else
                                                                    value = "None"
                                                            End Select
                                                            If Not IUCN.IndexOf(value) >= 0 Then errors &= "Invalid IUCN"
                                                            POSTfields.Add("newIUCNcat", value)
                                                    End Select
                                                End If
                                            End If
                                        Next

                                        If String.IsNullOrEmpty(errors) Then
                                            Try
                                                TextBox2.AppendText("Uploading " & name & vbCrLf)
                                                Application.DoEvents()
                                                ' Add headers (some are cleared after each call)
                                                With myWebClient.Headers
                                                    .Clear()
                                                    .Add("Authorization", $"Basic {credentials }")
                                                    .Add("Content-type", "application/x-www-form-urlencoded")
                                                    .Add("Accept", "text/html, Application / xhtml + Xml, Application / Xml;q=0.9,*/*;q=0.8")
                                                    .Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:60.0) Gecko/20100101 Firefox/60.0")
                                                    .Add("Accept-Language", "en-US, en;q=0.5")
                                                    .Add("Referer", WebUtility.UrlEncode(uriString))
                                                    '.Add("Connection", "keep-alive")
                                                    .Add("Upgrade-Insecure-Requests", "1")
                                                    .Add("DNT", "1")
                                                    .Add("Host", "wwff.co")
                                                    .Add("Cookie", Cookies)
                                                End With
                                                ' Send the POST
                                                Dim responseArray As Byte() = myWebClient.UploadValues(uriString, "POST", POSTfields)
                                                Dim responseStr = System.Text.Encoding.UTF8.GetString(responseArray)     ' convert to string
                                                ' Check response indicates success
                                                Dim rgx = New Regex("New reference '(VKFF-\d\d\d\d)' created")
                                                Dim m As Match = rgx.Match(responseStr)
                                                If m.Success Then
                                                    'Save reference as uploaded
                                                    ref = m.Groups(1).Value
                                                    Using sw As StreamWriter = File.AppendText(uploaded)
                                                        sw.WriteLine(name)
                                                    End Using
                                                    uploadedlist.Add(name)
                                                    If MsgBox(ref & " - " & name & " created" & vbCrLf & "Continue?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Success") <> MsgBoxResult.Yes Then Stop
                                                Else
                                                    MsgBox("Upload failed", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Data error in line " & count)
                                                    Stop
                                                End If
                                            Catch ex As WebException
                                                MessageBox.Show("Server error: " & ex.Message)
                                                Exit Do
                                            End Try
                                        Else
                                            MsgBox(errors, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Data error in line " & count)
                                            Exit Do
                                        End If
                                    Catch ex As Exception
                                        Dim trace = New Diagnostics.StackTrace(ex, True)
                                        Dim line As String = trace.ToString.Substring(trace.ToString.Length - 6)
                                        MessageBox.Show("Line " & count & " - " & "Source line " & line & ex.Message)
                                    End Try
                                End If
                            Catch ex As FileIO.MalformedLineException
                                MessageBox.Show("Cannot read file from disk. Original error: " & ex.Message)
                                Stop
                            End Try
Skip:
                        Loop
                        afile.Close()
                    Catch Ex As Exception
                        MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
                    End Try
                End If
            End Using
        End If
    End Sub
    Private Sub AddParksToParksdbToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddParksToParksdbToolStripMenuItem.Click
        ' Add parks from upload file to parks table so we can generate Region/District and boundary
        ' Also update GISmapping table
        ' Generate incremental update files
        Dim fields As New NameValueCollection
        Dim mandatory As New List(Of String)(New String() {"WWFFID", "NAME", "TYPE ABBR", "LATITUDE", "LONGITUDE", "STATE", "DECLARED", "DATASET", "GISID", "IUCN", "IBRA", "AUTHORITY"})   ' mandatory fields
        Dim columns As New List(Of String)
        Dim header As New List(Of String)
        Dim errors As String, Region As String, District As String, latitude As Double, longitude As Double
        Dim connect As SQLiteConnection ' declare the connection
        Dim sql As SQLiteCommand, ret As Integer
        Dim HTTPlink As String
        Dim StateName As New Dictionary(Of String, String) From {
            {"ACT", "VK1"},
            {"NSW", "VK2"},
            {"VIC", "VK3"},
            {"QLD", "VK4"},
            {"SA", "VK5"},
            {"WA", "VK6"},
            {"TAS", "VK7"},
            {"NT", "VK8"}
        }

        Using parkssql As New System.IO.StreamWriter("parksUpdate.sql", False),   ' SQL update commands for parks.sql
            GISmappingsql As New System.IO.StreamWriter("GISmappingUpdate.sql", False),   ' SQL update commands for GISmapping.sql
            openFileDialog1 As New OpenFileDialog With {
            .Filter = "csv files (*.csv)|*.txt;*.csv|All files (*.*)|*.*"
        }
            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                TextBox1.Clear()
                Try
                    parkssql.WriteLine($"/* updates to parks.sql created {Now.ToString(YMDHMS) } */")
                    GISmappingsql.WriteLine($"/* updates to GISmapping.sql created {Now.ToString(YMDHMS) } */")
                    Using afile As New FileIO.TextFieldParser(openFileDialog1.FileName)
                        Dim CurrentRecord As String() ' this array will hold each line of data
                        afile.TextFieldType = FileIO.FieldType.Delimited
                        afile.Delimiters = New String() {","}
                        afile.HasFieldsEnclosedInQuotes = True

                        connect = New SQLiteConnection(PARKSdb)             ' SQLite database containing PARKS data
                        connect.Open()  ' open database
                        sql = connect.CreateCommand
                        Dim count As Integer = 0
                        Do While Not afile.EndOfData
                            Try
                                CurrentRecord = afile.ReadFields
                                errors = ""
                                count += 1
                                If count = 1 Then
                                    ' Check the header
                                    header = CurrentRecord.ToList
                                    ' Upper case header
                                    For i = 0 To header.Count - 1
                                        header(i) = Trim(header(i).ToUpper())
                                    Next
                                    For Each st In mandatory
                                        If Not header.Contains(st) Then
                                            MsgBox("File does not contain an " & st & " column", vbCritical + vbOKOnly, "Mandatory column missing")
                                            Exit Sub
                                        End If
                                    Next
                                Else
                                    fields.Clear()
                                    ' extract park data
                                    Dim wwffid As String = Trim(CurrentRecord(header.IndexOf("WWFFID")))
                                    fields.Add("`WWFFID`", wwffid)
                                    fields.Add("`Name`", SQLEscape(Trim(CurrentRecord(header.IndexOf("NAME")))))
                                    fields.Add("`Type`", Trim(CurrentRecord(header.IndexOf("TYPE ABBR"))))
                                    latitude = CurrentRecord(header.IndexOf("LATITUDE"))
                                    fields.Add("`Latitude`", latitude)
                                    longitude = CurrentRecord(header.IndexOf("LONGITUDE"))
                                    fields.Add("`Longitude`", longitude)
                                    fields.Add("`IUCN`", Trim(CurrentRecord(header.IndexOf("IUCN"))))
                                    fields.Add("`IBRA`", Trim(CurrentRecord(header.IndexOf("IBRA"))))
                                    fields.Add("`Management`", Trim(CurrentRecord(header.IndexOf("AUTHORITY"))))
                                    fields.Add("`createDate`", Trim(CurrentRecord(header.IndexOf("DECLARED"))))
                                    Dim DataSet As String = Trim(CurrentRecord(header.IndexOf("DATASET")))
                                    fields.Add("`DataSet`", DataSet)
                                    Dim GISID As String = Trim(CurrentRecord(header.IndexOf("GISID")))
                                    fields.Add("`GISID`", GISID)
                                    Dim State As String = Trim(CurrentRecord(header.IndexOf("STATE")))
                                    ' Convert to VKn style
                                    For Each kvp In StateName
                                        If State = kvp.Key Then State = kvp.Value
                                    Next
                                    fields.Add("`State`", State)
                                    fields.Add("`IOTAID`", "OC-001")    ' hard-wired
                                    ' Do reverse lookup on Statename
                                    Dim st = ""
                                    For Each kvp In StateName
                                        If kvp.Value = State Then st = kvp.Key
                                    Next
                                    If st = "" Then Throw New System.Exception($"State {State } not recognised")
                                    fields.Add("`DXCC`", $"OC / VK-{st }")
                                    fields.Add("`HASC`", $"AU.{st.Substring(0, 2) }")
                                    HTTPlink = ""
                                    If header.Contains("HTTPLINK") Then
                                        HTTPlink = Trim(CurrentRecord(header.IndexOf("HTTPLINK")))
                                        If HTTPlink <> "" Then fields.Add("`HTTPLink`", HTTPlink)
                                    End If
                                    If header.IndexOf("AREA") >= 0 Then
                                        Dim area As String = CInt(CurrentRecord(header.IndexOf("AREA")))
                                        fields.Add("`Area`", $"{area}")     ' escape single quote
                                    End If
                                    If header.IndexOf("NOTES") >= 0 Then
                                        Dim notes As String = Trim(CurrentRecord(header.IndexOf("NOTES")))
                                        fields.Add("`Notes`", SQLEscape(notes))     ' escape single quote
                                    End If
                                    fields.Add("`Status`", "pending")
                                    Region = ""
                                    District = ""
                                    GetRegionDistrict(longitude, latitude, Region, District)
                                    If Region <> "" Then fields.Add("`Region`", SQLEscape(Region))
                                    If District <> "" Then fields.Add("`District`", SQLEscape(District))

                                    Try
                                        ' insert into parks table
                                        Dim field_names As String = String.Join(",", fields.AllKeys)    ' CSV list of fields
                                        Dim value_list As New List(Of String)
                                        For Each value In fields.AllKeys
                                            value_list.Add($"'{fields(value)}'")
                                        Next
                                        Dim field_values As String = String.Join(",", value_list)       ' quoted list of CSV values
                                        sql.CommandText = $"INSERT OR REPLACE INTO `parks` ({field_names}) VALUES ({field_values});"
                                        ret = sql.ExecuteNonQuery()
                                        Me.TextBox1.Text = $"adding {wwffid }"
                                        Application.DoEvents()
                                        count += 1
                                        If ret = 0 Then MsgBox("SQL Error", vbCritical + vbOKOnly, sql.CommandText)
                                        parkssql.WriteLine(sql.CommandText)
                                        ' insert into GISmapping table
                                        Dim GISIDlist() As String = Split(GISID, ",")   ' there might be more than 1 GISID
                                        For i = 0 To GISIDlist.Length - 1
                                            sql.CommandText = $"INSERT OR REPLACE INTO `GISmapping` (`WWFFID`,`DataSet`,`GISID`) VALUES ('{wwffid }','{DataSet }','{GISIDlist(i) }');"
                                            ret = sql.ExecuteNonQuery()
                                            If ret = 0 Then MsgBox("SQL Error", vbCritical + vbOKOnly, sql.CommandText)
                                            GISmappingsql.WriteLine(sql.CommandText)
                                        Next
                                    Catch ex As Exception
                                        MessageBox.Show(wwffid & " " & sql.CommandText & vbCrLf & ex.Message, "Insert error")
                                    End Try
                                End If
                            Catch ex As Exception
                                Dim trace = New Diagnostics.StackTrace(ex, True)
                                Dim line As String = Strings.Right(trace.ToString, 6)
                                MessageBox.Show($"{ex.Message }{vbCrLf }Line: {line }", "Exception")
                            End Try
                        Loop
                        connect.Close()
                    End Using

                    MsgBox($"{count } parks added", vbInformation + vbOKOnly, "Done")
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End Using
    End Sub

    Shared Function ValidUrl(url As String) As Boolean
        ' Validate a URL
        Dim validatedUri As Uri = Nothing

        If (Uri.TryCreate(url, UriKind.Absolute, validatedUri)) Then
            ' If true: validatedUri Contains a valid Uri. Check For the scheme In addition.
            Return (True)
        End If
        Return False
    End Function

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Stop    ' exit program
    End Sub

    Private Sub ProcessPark(ByRef sql As SQLiteCommand, PA_IDList As List(Of String), ByRef Maxmapping As Integer)
        ' Process the current list of PA_ID for a park
        Dim temp As List(Of String), sqldr As SQLiteDataReader, WWFFID As String, DataSet As String = ""
        If PA_IDList.Count > 1 Then
            Maxmapping = Max(Maxmapping, PA_IDList.Count)   ' record largest number of mappings
            ' Look for this WWFFID
            temp = PA_IDList.ToList
            For i = 0 To temp.Count - 1
                temp(i) = "'" & temp(i) & "'" ' surround value with quotes
            Next
            ' find which park this is
            WWFFID = ""
            sql.CommandText = $"SELECT * FROM parks where GISID in ({String.Join(",", temp.ToArray) })"
            sqldr = sql.ExecuteReader()
            While sqldr.Read()
                WWFFID = sqldr.Item("WWFFID")
                DataSet = sqldr.Item("DataSet")
            End While
            sqldr.Close()
            ' Add mapping
            If Not String.IsNullOrEmpty(WWFFID) Then
                TextBox1.Text = "Adding mapping for " & WWFFID
                Application.DoEvents()
                For Each id In PA_IDList
                    sql.CommandText = $"REPLACE INTO GISmapping (WWFFID,DataSet,GISID) VALUES ('{WWFFID }','{DataSet }','{id }')"
                    sql.ExecuteNonQuery()
                Next
            End If
        End If
    End Sub

    Private Function GetParkData(WWFFID As String) As NameValueCollection
        ' Retrieve all park data as a name value collection
        ' Input parameter is park id
        Contract.Requires(Not String.IsNullOrEmpty(WWFFID), "Illegal WWFFID")
        Dim result As NameValueCollection = Nothing
        Dim sql As SQLiteCommand, SQLdr As SQLiteDataReader

        Using connect As New SQLiteConnection(PARKSdb)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            ' Add all fields from parks.csv to park data
            sql.CommandText = $"SELECT * FROM parks LEFT JOIN SHIRES ON parks.ShireID=SHIRES.ShireId WHERE WWFFID='{WWFFID }'"
            SQLdr = sql.ExecuteReader()
            SQLdr.Read()
            If SQLdr.HasRows Then
                result = GetParkData(SQLdr)
                connect.Close()
            Else
                MsgBox("GetParkData: no data for " & WWFFID, vbCritical + vbOKOnly, "Error")
            End If
            Return result
        End Using
    End Function

    Private Function GetParkData(SQLDataReader As SQLiteDataReader) As NameValueCollection
        ' Retrieve all park data as a name value collection
        ' Input parameter is SQLiteDataReader
        Dim result As New NameValueCollection, IDs As New List(Of String), DataSet As String = Nothing, LongName As String, WWFFID As String
        Dim sql As SQLiteCommand, SQLdr As SQLiteDataReader

        result.Clear()
        If Not SQLDataReader.HasRows Then
            Return result
        Else
            Using connect As New SQLiteConnection(PARKSdb)
                WWFFID = SQLDataReader.Item("WWFFID").ToString
                connect.Open()  ' open database
                sql = connect.CreateCommand
                ' Create list of GISID's
                IDs.Clear()
                sql.CommandText = $"SELECT * FROM GISmapping WHERE WWFFID='{WWFFID }' ORDER BY GISID"
                SQLdr = sql.ExecuteReader()
                While SQLdr.Read()
                    DataSet = SQLdr("DataSet").ToString
                    IDs.Add(SQLdr("GISID").ToString)
                End While
                SQLdr.Close()
                If DataSet IsNot Nothing Then
                    result.Add("DataSet", DataSet)
                    result.Add("GISIDList", String.Join(",", IDs.ToArray))  ' list of GISID without quotes
                    If IDs.Any Then
                        For i = 0 To IDs.Count - 1
                            IDs(i) = "'" & IDs(i) & "'"       ' surround with quotes
                        Next
                        result.Add("GISIDListQuoted", String.Join(",", IDs.ToArray))  ' produce comma separated quoted list
                    End If
                End If
                ' Add all fields from parks.db to park data
                Dim val As Double
                Dim fields As NameValueCollection = SQLDataReader.GetValues
                For Each key In fields.AllKeys
                    If key <> "GISID" And key <> "DataSet" Then
                        Dim value As String = fields(key)
                        Select Case key
                            Case "Area", "Latitude", "Longitude"
                                ' Sometimes these can be blank, empty, null. Convert them to 0
                                If Not Double.TryParse(value, val) Then value = "0"
                            Case "ShireID"
                                value = value.Split(",")(0)     ' joined field returns field twice
                        End Select
                        result.Add(key, value)    ' GISID & DataSet are obsolete
                    End If
                Next
                connect.Close()
                ' Construct long name for park
                If LongNames.ContainsKey(result("Type")) Then
                    LongName = result("Name") & " " & LongNames(result("Type"))
                Else
                    LongName = result("Name") & " " & result("Type")
                    MsgBox($"Park {WWFFID } has no long name for abbreviation {result("Type") }", vbExclamation + vbOKOnly, "Program error")
                End If
                result.Add("LongName", LongName)
                Return result
            End Using
        End If
    End Function

    Private Sub ValidateHyperlinkToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ValidateHyperlinkToolStripMenuItem.Click
        ' Check the hyperlink of all parks
        Dim sql As SQLiteCommand
        Dim SQLdr As SQLiteDataReader
        Dim status As String, url As String, resp As HttpWebResponse
        Dim total As Integer = 0, OK As Integer = 0, IllFormed As Integer = 0, NotFound As Integer = 0, Missing As Integer = 0
        Dim ParkData As NameValueCollection, StatusCode As String

        Using connect As New SQLiteConnection(PARKSdb), logWriter As New System.IO.StreamWriter("hyperlink.html", False)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            logWriter.WriteLine("<table border=1>")
            logWriter.WriteLine("<tr><th>WWFFID</th><th>Name</th><th>State</th><th>HTTPLink</th><th>Status</th><th>HTTP error</th></tr>")
            ' Get basic statistics
            sql.CommandText = "SELECT * FROM parks LEFT JOIN SHIRES ON parks.ShireID=SHIRES.ShireId WHERE lower(Status) IN ('active', 'pending') ORDER BY WWFFID"
            SQLdr = sql.ExecuteReader()
            While SQLdr.Read()
                StatusCode = ""
                total += 1
                ' Validate URL
                ParkData = GetParkData(SQLdr)       ' list of all data for this park
                If IsDBNull(SQLdr.Item("HTTPLink")) Then
                    url = ""
                Else
                    url = SQLdr.Item("HTTPLink")
                End If
                If String.IsNullOrEmpty(url) Then
                    status = "<font color='orange'>Missing</font>"
                    Missing += 1
                Else
                    If Uri.IsWellFormedUriString(url, UriKind.Absolute) Then
                        Dim request As HttpWebRequest = WebRequest.Create(url)
                        request.AllowAutoRedirect = True
                        request.Proxy = Nothing
                        request.Timeout = 30 * 1000     ' 30s
                        request.KeepAlive = False
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
                        resp = Nothing
                        Try
                            resp = request.GetResponse
                            status = "<font color='green'>OK</font>"
                            OK += 1
                        Catch ex As WebException
                            StatusCode = ex.Message
                            status = "<font color='red'>Not found</font>"
                            NotFound += 1
                        Finally
                            If resp IsNot Nothing Then
                                StatusCode = resp.StatusCode.ToString
                                resp.Close()
                            End If
                        End Try
                    Else
                        status = "Ill formed"
                        IllFormed += 1
                    End If
                End If
                TextBox1.Text = $"Checking {ParkData("WWFFID") } {ParkData("Longname") }: {total }"
                Application.DoEvents()
                logWriter.WriteLine($"<tr><td>{ParkData("WWFFID") }</td><td>{ParkData("Longname") }</td><td>{ParkData("State") }</td><td><a href='{url }'>{url }</a></td><td>{status }</td><td>{StatusCode }</td></tr>")
                logWriter.Flush()
            End While
            connect.Close()
            logWriter.WriteLine("</table>")
            logWriter.WriteLine("<br>{0} total links validated. {1} OK, {2} Missing, {3} Ill formed, {4} Not Found", total, OK, Missing, IllFormed, NotFound)
            logWriter.Close()
            TextBox1.Text = "Done"
        End Using
    End Sub

    Private Sub GuessHyperlinkToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuessHyperlinkToolStripMenuItem.Click
        ' Guess hyperlinks for parks without one. Test if they are OK, and write to sql update file if OK.

        Dim sql As SQLiteCommand
        Dim SQLdr As SQLiteDataReader
        Dim url As String, urls As New List(Of String), u As String = ""
        Dim total As Integer = 0, OK As Integer = 0, found As Integer = 0, Errors As Integer = 0
        Dim ParkData As NameValueCollection, result As String = "", ParkName As String
        Dim SA_Parks As New List(Of String) From {"Adelaide", "Adelaide_Hills", "Barossa", "Limestone_Coast", "Limestone-Coast", "Eyre_Peninsula", "Kangaroo-Island", "Flinders-Ranges-Outback", "yorke_peninsula", "Murray_River", "Murray-River", "Fleurieu_Peninsula", "Clare_Valley", "south_east", "upper_spencer_gulf", "far_west_coast"}
        Dim SA_MarineParks As New List(Of String) From {"limestone-coast", "eyre-peninsula", "kangaroo-island", "yorke-peninsula", "fleurieu-peninsula", "south-east", "upper-spencer-gulf", "far-west-coast"}
        Dim ZL_Regions As New List(Of String) From {"canterbury", "taranaki", "nelson-tasman", "west-coast", "stewart-island-rakiura", "northland", "auckland", "coromandel", "waikato", "bay-of-plenty", "east-coast", "central-north-island", "hawkes-bay", "manawatu-whanganui", "wairarapa", "wellington-kapiti", "marlborough", "otago", "southland", "fiordland", "chatham-islands"}

        ' Tasmanian park names, and their base value
        Dim TAS_base As New Dictionary(Of String, String) From {
            {"Arthur-Pieman Conservation Area", "768"},
            {"Bay of Fires Conservation Area", "3999"},
            {"Granite Point Conservation Area", "5849"},
            {"Peter Murrell Reserves", "4036"},
            {"Scamander Conservation Area", "4061"},
            {"Southport Lagoon Conservation Area", "4080"},
            {"St Helens Point Conservation Area", "4099"},
            {"Tamar River Conservation Area", "4118"},
            {"Waterhouse Conservation Area", "4140"},
            {"Tom Gibson Nature Reserve", "4158"},
            {"Humbug Point Nature Recreation Area", "4176"},
            {"Kate Reed Nature Recreation Area", "19454"},
            {"Trevallyn Nature Recreation Area", "10399"},
            {"Coningham Nature Recreation Area", "25588"},
            {"Gunns Plains Cave State Reserve", "5854"},
            {"Hastings Caves State Reserve", "4212"},
            {"Liffey Falls State Reserve", "4235"},
            {"Notley Fern Gorge State Reserve", "4253"},
            {"St Columba Falls State Reserve", "4271"},
            {"St Marys Pass State Reserve", "4293"},
            {"Moulting Lagoon Game Reserve", "4319"},
            {"Opossum Bay Marine Conservation Area", "42194"},
            {"Monk Bay Marine Conservation Area", "42198"},
            {"Cloudy Bay Marine Conservation Area", "42202"},
            {"Central Channel Marine Conservation Area", "42206"},
            {"Simpsons Point Marine Conservation Area", "42214"},
            {"Roberts Point Marine Conservation Area", "42218"},
            {"Huon Estuary Marine Conservation Area", "42210"},
            {"Hippolyte Rocks Marine Conservation Area", "42222"},
            {"Sloping Island Marine Conservation Area", "42226"},
            {"Waterfall?Fortescue Marine Conservation Area", "42230"},
            {"Blackman Rivulet Marine Conservation Area", "42234"},
            {"South Arm Marine Conservation Area", "42238"},
            {"Port Cygnet Marine Conservation Area", "42242"},
            {"River Derwent Marine Conservation Area", "42256"},
            {"Cradle Mountain National Park", "3297"},
            {"Douglas-Apsley National Park", "3330"},
            {"Freycinet National Park", "3363"},
            {"Hartz Mountains National Park", "3396"},
            {"Kent Group National Park", "3429"},
            {"Lake St Clair National Park", "3462"},
            {"Maria Island National Park", "3495"},
            {"Mole Creek Karst National Park", "3530"},
            {"Mt Field National Park", "3589"},
            {"Mt William National Park", "3622"},
            {"Narawntapu National Park", "3665"},
            {"Rocky Cape National Park", "3698"},
            {"Savage River National Park", "3732"},
            {"South Bruny National Park", "3765"},
            {"Southwest National Park", "3801"},
            {"Strzelecki National Park", "3834"},
            {"Tasman National Park", "3868"},
            {"Walls of Jerusalem National Park", "3904"},
            {"Wild Rivers National Park", "3937"},
            {"Tinderbox Marine Reserve", "5501"},
            {"Maria Island Marine Reserve", "2910"},
            {"Ninepin Point Marine Reserve", "2926"},
            {"Governor Island Marine Reserve", "3094"},
            {"Kent Group Marine Reserve", "3110"},
            {"Port Davey Marine Reserve", "3126"},
            {"Macquarie Island Marine Reserve", "3142"},
            {"Cradle Mountain-Lake St Clair National Park", "3297"}
        }

        If Form3.ShowDialog() = DialogResult.OK Then
            Using connect As New SQLiteConnection(PARKSdb),
             logWriter As New System.IO.StreamWriter(String.Format("guess_url-{0}.html", DateTime.Now.ToString("yyMMdd-HHmm")), False),
             sqlWriter As New System.IO.StreamWriter(String.Format("guess_url-{0}.sql", DateTime.Now.ToString("yyMMdd-HHmm")), False)
                Dim selected As List(Of String) = Form3.Selected
                logWriter.WriteLine("<table border=1>")
                logWriter.WriteLine("<tr><th>WWFFID</th><th>Name</th><th>State</th><th>URL guess</th><th>Result</th></tr>")
                connect.Open()  ' open database
                sql = connect.CreateCommand
                ' Select parks that need a proper URL.
                ' Wikipedia and protectedplanet.net URL's are allowed, but opportunities are sought for a better one.
                sql.CommandText = "SELECT * FROM parks LEFT JOIN SHIRES ON parks.ShireID=SHIRES.ShireId WHERE (`HTTPLink` ISNULL Or `HTTPLink`='' OR `HTTPLink` LIKE '%www2.dec.wa.gov.au%' OR `HTTPLink` LIKE '%www.tams.act.gov.au%' OR `HTTPLink` LIKE '%protectedplanet%' OR `HTTPLink` LIKE '%wikipedia%' OR `HTTPLink` LIKE '%www.environment.nsw.gov.au%' OR (`HTTPLink` LIKE '%nationalparks.nsw.gov.au%' AND `HTTPLink` NOT LIKE '%visit%')) AND lower(Status) IN ('active', 'pending') ORDER BY State, Name"
                SQLdr = sql.ExecuteReader()
                While SQLdr.Read()
                    ParkData = GetParkData(SQLdr)       ' list of all data for this park
                    ParkData("Name") = ParkData("Name").Trim()     ' in case of extraneous spaces
                    ParkData("LongName") = ParkData("LongName").Replace("  ", " ")  ' in case of extraneous spaces
                    ParkData("LongName") = ParkData("LongName").Replace(" - ", "-")
                    ParkData("Name") = ParkData("Name").Replace(" - ", "-")
                    Dim OriginalURL As String = ParkData("HTTPLink")
                    If selected.IndexOf(ParkData("State")) >= 0 Then
                        total += 1
                        ' Guess a URL based on state
                        urls.Clear()
                        Select Case ParkData("State")
                            Case "VK1"
                                url = "https://www.environment.act.gov.au/parks-conservation/parks-and-reserves/find-a-park/canberra-nature-park/" & LCase(ParkData("Longname").Replace(" ", "-"))
                                url = Replace(url, "mt-", "mount-")     ' they like the long version of mt in the ACT
                                urls.Add(url)
                            Case "VK2"
                                urls.Add("https://www.nationalparks.nsw.gov.au/visit-a-park/parks/" & Replace(LCase(ParkData("Longname")), " ", "-"))
                            Case "VK3"
                                Dim ln As String = LCase(LongNames(ParkData("Type")))
                                Dim split = ln.Split(" ")
                                Dim FirstLetter As New List(Of String)
                                For Each st In split
                                    FirstLetter.Add(st.Substring(0, 1))
                                Next
                                Dim suffix = Join(FirstLetter.ToArray, ".") & "."
                                urls.Add("https://parkweb.vic.gov.au/explore/parks/" & Replace(LCase(ParkData("Name")), " ", "-") & "-" & suffix)
                            Case "VK4"
                                urls.Add("https://parks.des.qld.gov.au/parks/" & Replace(LCase(ParkData("Name")), " ", "-") & "/")
                                ParkName = LCase(String.Format("{0} {1}", LongNames(ParkData("Type")), ParkData("Name"))).Replace(" ", "-")
                                urls.Add("https://wetlandinfo.des.qld.gov.au/wetlands/facts-maps/" & ParkName)
                            Case "VK5"
                                ' Need to test all regions for match
                                ParkName = LCase(ParkData("LongName").Replace(" ", "-"))
                                For Each u In SA_Parks
                                    urls.Add($"https://www.parks.sa.gov.au/find-a-park/Browse_by_region/{u }/{ParkName }")
                                Next
                                If ParkName.Contains("marine") Or ParkName.Contains("island") Then
                                    ParkName = LCase(ParkData("Name").Replace(" ", "-"))
                                    For Each u In SA_MarineParks
                                        urls.Add($"https://www.environment.sa.gov.au/marineparks/find-a-park/{u }/{ParkName }")
                                    Next
                                Else
                                    For Each u In SA_Parks
                                        urls.Add($"https://www.environment.sa.gov.au/parks/Find_a_Park/Browse_by_region/{u }/{ParkName }")
                                    Next
                                End If
                            Case "VK6"
                                urls.Add("https://parks.dpaw.wa.gov.au/park/" & Replace(LCase(ParkData("Name")), " ", "-"))
                            Case "VK7"
                                If TAS_base.ContainsKey(ParkData("Longname")) Then
                                    urls.Add("https://www.parks.tas.gov.au/index.aspx?base=" & TAS_base(ParkData("Longname")))
                                End If
                            Case "VK8"
                                urls.Add("https://nt.gov.au/leisure/parks-reserves/find-a-park/find-a-park-to-visit/" & LCase(ParkData("Longname").Replace(" ", "-")))
                            Case "ZL"
                                For Each u In ZL_Regions
                                    urls.Add($"https://www.doc.govt.nz/parks-and-recreation/places-to-go/{u }/places/" & LCase(ParkData("Longname").Replace("  ", " ").Replace(" ", "-")))
                                Next
                        End Select
                        urls.Add($"https://en.wikipedia.org/wiki/{ParkData("LongName").Replace(" ", "_") }")    ' last resort is wikipedia
                        If urls.Any Then
                            For Each u In urls
                                u = Replace(u, "'", "")     ' remove any quotes
                                result = TestURL(u)         ' test the url
                                If result = "OK" Then
                                    Exit For
                                End If
                            Next
                            If result = "OK" Then
                                If u <> OriginalURL Then
                                    ' Produce SQL to UPDATE parks
                                    sqlWriter.WriteLine($"/* {ParkData("WWFFID") } - {ParkData("State") } */")
                                    sqlWriter.WriteLine($"UPDATE `PARKS` SET `HTTPLink`='{SQLEscape(u) }' WHERE `WWFFID`='{ParkData("WWFFID") }';")
                                    sqlWriter.Flush()
                                    u = $"<a href='{u }'>{u }</a>"
                                    found += 1
                                Else
                                    u = "None better"
                                End If
                                OK += 1
                            Else
                                u = "Not found"
                                Errors += 1
                            End If
                        End If
                        TextBox1.Text = $"Checking {ParkData("State") } {ParkData("WWFFID") } {ParkData("Longname") }: {found }/{total }"
                        Application.DoEvents()
                        logWriter.WriteLine($"<tr><td>{ParkData("WWFFID") }</td><td>{ParkData("Longname") }</td><td>{ParkData("State") }</td><td>{u }</td><td>{result }</td></tr>")
                        logWriter.Flush()
                    End If
                End While
                connect.Close()
                logWriter.WriteLine("</table>")
                logWriter.WriteLine("<br>{0} total links guessed. {1} OK, {2} errors", total, OK, Errors)
                logWriter.Close()
                sqlWriter.Close()
                TextBox1.Text = "Done"
            End Using
        End If
    End Sub

    Shared Function TestURL(url As String) As String
        ' Test a URL
        Dim result As String
        Dim resp As HttpWebResponse = Nothing
        Dim request As HttpWebRequest = WebRequest.Create(url)
        request.AllowAutoRedirect = True
        request.Proxy = Nothing
        request.Timeout = 30 * 1000     ' 30s
        request.KeepAlive = False
        ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls

        Try
            resp = request.GetResponse
            result = "OK"
        Catch ex As WebException
            result = ex.Message
        Finally
            If resp IsNot Nothing Then resp.Close()
        End Try
        Return result
    End Function

    Private Sub GetLGADataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetLGADataToolStripMenuItem.Click
        ' Get all the 2021 LGA data
        Const GEO_SERVER = "https://censusdata.abs.gov.au/arcgis/rest/services/ASGS2021/LGA/MapServer/1/query"
        Dim POSTfields As NameValueCollection, responseStr As String, resp As Byte() = {}
        Dim sql As SQLiteCommand, count As Integer = 0

        POSTfields = New NameValueCollection From {
            {"f", "json"},          ' return json
            {"where", "1=1"},
            {"returnCountOnly", "false"},
            {"returnIdsOnly", "false"},
            {"returnGeometry", "false"},              ' don't need the geometry
            {"outFields", "*"}
        }
        Using myWebClient As New WebClient
            Try
                myWebClient.Headers.Add("accept", "text/html,Application/xhtml+Xml,Application/Xml;q=0.9,Image/avif,Image/webp,Image/apng,*/*;q=0.8,Application/signed-exchange;v=b3;q=0.9")
                myWebClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36")
                resp = myWebClient.UploadValues(GEO_SERVER, "POST", POSTfields)    ' query map server
                TextBox1.Text = "retrieved " & resp.Length & " bytes of data"
                Application.DoEvents()
                responseStr = System.Text.Encoding.UTF8.GetString(resp)
                Dim Jo As JObject = JObject.Parse(responseStr)
                If Jo.HasValues And Jo.Item("features").Any Then
                    Using connect As New SQLiteConnection(PARKSdb)
                        connect.Open()  ' open database
                        sql = connect.CreateCommand
                        sql.CommandText = "BEGIN TRANSACTION"   ' start transaction
                        sql.ExecuteNonQuery()
                        sql.CommandText = "DELETE FROM LGA_2021"   ' remove existing data
                        sql.ExecuteNonQuery()
                        ' Iterate through each feature and save
                        For Each f In Jo.Item("features")
                            Dim LGA_NAME As String = f("attributes")("LGA_NAME_2021").ToString.Replace("'", "''")
                            sql.CommandText = $"INSERT INTO LGA_2021 (OBJECTID,LGA_CODE_2021,LGA_NAME_2021,STATE_CODE_2021,STATE_NAME_2021,AREA_ALBERS_SQKM) VALUES ('{f("attributes")("OBJECTID") }','{f("attributes")("LGA_CODE_2021") }','{LGA_NAME }','{f("attributes")("STATE_CODE_2021") }','{f("attributes")("STATE_NAME_2021") }','{f("attributes")("AREA_ALBERS_SQKM") }')"
                            sql.ExecuteNonQuery()
                            count += 1
                            TextBox1.Text = $"Added {LGA_NAME } : {count }"
                            Application.DoEvents()
                        Next
                        sql.CommandText = "COMMIT"   ' end transaction
                        sql.ExecuteNonQuery()
                        connect.Close()
                    End Using
                    MsgBox(count & " LGA added", vbInformation + vbOKOnly, "LGA added")
                End If
            Catch ex As Exception
                MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOK, "Web request failed")
            End Try
        End Using
    End Sub

    Private Sub FindChangesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindChangesToolStripMenuItem.Click
        ' Find the changes between 2020 data and 2021 data
        Dim sql As SQLiteCommand, sqldr As SQLiteDataReader, deleted As Integer = 0, added As Integer = 0
        Dim old_LGA As String = "2020", new_LGA As String = "2021"
        Using connect As New SQLiteConnection(PARKSdb), logWriter As New System.IO.StreamWriter("LGAchanges.html", False)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            ' firstly find deletions, i.e. LGA in {old_LGA} that are not in {new_LGA}
            sql.CommandText = $"select * from LGA_{old_LGA} left join LGA_{new_LGA} ON LGA_CODE_{old_LGA}=LGA_CODE_{new_LGA} WHERE LGA_{new_LGA}.OBJECTID isnull ORDER BY LGA_CODE_{old_LGA}"
            sqldr = sql.ExecuteReader
            logWriter.WriteLine($"LGA deleted in {new_LGA}")
            logWriter.WriteLine("<table border=1>")
            logWriter.WriteLine("<tr><th>LGA code</th><th>Name</th><th>State code</th></tr>")
            While sqldr.Read
                logWriter.WriteLine($"<tr><td>{sqldr($"LGA_CODE_{old_LGA}") }</td><td>{sqldr($"LGA_NAME_{old_LGA}") }</td><td>{sqldr($"STATE_CODE_{old_LGA}") }</td></tr>")
                deleted += 1
            End While
            logWriter.WriteLine("</table><br>")
            sqldr.Close()
            ' secondly find additions, i.e. LGA in {new_LGA} that are not in {old_LGA}
            sql.CommandText = $"select * from LGA_{new_LGA} left join LGA_{old_LGA} ON LGA_CODE_{old_LGA}=LGA_CODE_{new_LGA} WHERE LGA_{old_LGA}.OBJECTID isnull AND NOT LGA_NAME_{new_LGA} LIKE ""Migratory%"" AND NOT LGA_NAME_{new_LGA} LIKE ""No usual%"" ORDER BY LGA_CODE_{new_LGA}"
            sqldr = sql.ExecuteReader
            logWriter.WriteLine($"LGA added in {new_LGA}")
            logWriter.WriteLine("<table border=1>")
            logWriter.WriteLine("<tr><th>LGA code</th><th>Name</th><th>State code</th></tr>")
            While sqldr.Read
                logWriter.WriteLine($"<tr><td>{sqldr($"LGA_CODE_{new_LGA}") }</td><td>{sqldr($"LGA_NAME_{new_LGA}") }</td><td>{sqldr($"STATE_CODE_{new_LGA}") }</td></tr>")
                added += 1
            End While
            logWriter.WriteLine("</table>")
            sqldr.Close()
            TextBox1.Text = $"{deleted } LGA deleted and {added } LGA added"
        End Using
    End Sub

    ' State numbers as they appear in LGA listing
    ReadOnly LGA_States As New Dictionary(Of Integer, String) From {
            {8, "VK1"},
            {1, "VK2"},
            {2, "VK3"},
            {3, "VK4"},
            {4, "VK5"},
            {5, "VK6"},
            {6, "VK7"},
            {7, "VK8"},
            {9, "Other"}
            }

    Private Sub CreateLGAListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateLGAListToolStripMenuItem.Click
        ' Create an Excel workbook with LGA list
        Dim xls_Path As String = Application.StartupPath & "\LGA.xls"
        Dim csvWriter As New System.IO.StreamWriter("VKSHIRES.csv", False)  ' CSV file for VKCL
        Dim sql As SQLiteCommand, sqldr As SQLiteDataReader, row As Integer
        Dim objApp As Excel.Application
        Dim objBook As Excel._Workbook

        ' Remove the XLS file if it already exists.
        File.Delete(xls_Path)

        Dim objBooks As Excel.Workbooks
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet
        Dim range As Excel.Range
        Dim misValue As Object = Reflection.Missing.Value

        ' Before we export anything, check there are no duplicate shire codes
        Using connect As New SQLiteConnection(PARKSdb)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            sql.CommandText = "SELECT *,COUNT(*) AS C FROM `SHIRES` GROUP BY `ShireID` HAVING C>1"
            sqldr = sql.ExecuteReader()
            While sqldr.Read()
                MsgBox("There is a duplicate shire ID " & sqldr![ShireID], vbAbort + vbOK, "Duplicate ShireID")
            End While
            sqldr.Close()

            ' Create a new instance of Excel and start a new workbook.
            objApp = New Excel.Application()
            objBooks = objApp.Workbooks
            objBook = objBooks.Add(misValue)
            objSheets = objBook.Worksheets
            ' Make an info sheet
            objSheet = objSheets(1)
            objSheet.Name = "Info"
            range = objSheet.Range("A1") : range.Value = "List of local government area codes"
            range = objSheet.Range("A2") : range.Value = "LGA codes are as at 2020"
            range = objSheet.Range("A3") : range.Value = "Listing produced " & Now()
            range = objSheet.Range("A4") : range.Value = "Errors/omissions to vk3ohm@wia.org.au"
            objSheet.Columns.EntireColumn.AutoFit()

            csvWriter.WriteLine(String.Format("# LGA 2020 data produced {0}", Now()))
            For Each state In LGA_States.Reverse
                TextBox1.Text = "Creating state " & state.Value
                objSheets.Add()     ' insert sheet at front of book
                objSheet = objSheets(1)
                objSheet.Name = state.Value
                row = 1
                sql.CommandText = "SELECT * FROM `SHIRES` WHERE `ShireState`='" & state.Value & "' ORDER BY `ShireName`"
                sqldr = sql.ExecuteReader()
                While sqldr.Read()
                    range = objSheet.Range("A" & row) : range.Value = sqldr![ShireName]
                    range = objSheet.Range("B" & row) : range.Value = sqldr![ShireID]
                    row += 1
                    csvWriter.WriteLine(String.Format("{0},""{1}"",""{2}""", sqldr![ShireID], sqldr![LGALookup], LongShire(sqldr![LGALookup])))
                End While
                sqldr.Close()
                objSheet.Columns.EntireColumn.AutoFit()
            Next
            ' Save and close the new populated spreadsheet
            objBook.SaveAs(xls_Path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
             Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            objBook.Close(True, misValue, misValue)
            objApp.Quit()
            range = Nothing
            objSheet = Nothing
            objSheets = Nothing
            objBooks = Nothing
            connect.Close()
            csvWriter.Close()
            TextBox1.Text = "Done"
        End Using
    End Sub

    Private Sub FixSHIRESTableToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles FixSHIRESTableToolStripMenuItem.Click
        ' Using the latest LGA data, fix the PnP SHIRES table to match
        ' 1. Look for entries in ABS LGA data that do not exist in SHIRES, and create INSERT
        ' 2. Look for entries in SHIRES data that do not exist in ABS LGA data, and create DELETE
        ' 3. Check spelling of name

        Const YEAR = "2020"     ' year of LGA data
        Dim sql As SQLiteCommand, sqldr As SQLiteDataReader, count As Integer

        Using connect As New SQLiteConnection(PARKSdb), sqlWriter As New System.IO.StreamWriter("SHIRES_fix.sql", False)

            sqlWriter.WriteLine($"/* Validation of SHIRES table against ABS LGA data ({YEAR }) at {Now } */")
            connect.Open()  ' open database
            sql = connect.CreateCommand

            ' Do some sanity checking on SHIRES table. Can't fix if inconsistant
            ' Look for duplicate ShireID
            sql.CommandText = "SELECT *,COUNT(*) AS C FROM `SHIRES` GROUP BY `ShireID` HAVING C>1"
            sqldr = sql.ExecuteReader()
            While sqldr.Read()
                MsgBox("There is a duplicate shire ID " & sqldr("ShireID"), vbAbort + vbOK, "Duplicate ShireID")
            End While
            sqldr.Close()
            ' Look for duplicate LGA
            sql.CommandText = "SELECT *,COUNT(*) AS C FROM `SHIRES` GROUP BY `LGA` HAVING C>1"
            sqldr = sql.ExecuteReader()
            While sqldr.Read()
                MsgBox("There is a duplicate LGA " & sqldr("LGA"), vbAbort + vbOK, "Duplicate LGA")
            End While
            sqldr.Close()

            ' 2. Look for entries in SHIRES data that do not exist in ABS LGA data, and create DELETE
            sql.CommandText = $"Select * From SHIRES As a left Join LGA_{YEAR } As b On a.LGA=b.LGA_CODE_{YEAR } Where b.objectid isnull"
            sqldr = sql.ExecuteReader()
            count = 0
            While sqldr.Read()
                sqlWriter.WriteLine($"DELETE FROM `SHIRES` WHERE `id`={sqldr![id] }; /* Deleting {sqldr![LGA] }, {sqldr![ShireID] }, {sqldr![ShireName] } */")
                count += 1
            End While
            sqlWriter.WriteLine($"/* {count } shires exist in SHIRES data that do not exist in ABS LGA data */")
            sqlWriter.WriteLine()
            sqldr.Close()

            ' 3. Look for entries in ABS LGA data that do not exist in SHIRES, and create INSERT
            Dim NewCodes As New Dictionary(Of String, String) From {
                {"10500", "BX2"},    ' Bayside
                {"12160", "CT2"},    ' Cootamundra
                {"12390", "DU2"},    ' Dubbo
                {"40150", "AL5"}     ' Adelaide Plains
                }
            sql.CommandText = $"Select * From LGA_{YEAR } As a left Join SHIRES As b On a.LGA_CODE_{YEAR }=b.LGA Where b.id isnull"
            sqldr = sql.ExecuteReader()
            count = 0
            While sqldr.Read()
                Dim ShireId As String = "XX" & count
                Dim LGAcode As String = sqldr($"LGA_CODE_{YEAR }").ToString
                Dim LGAname As String = sqldr($"LGA_NAME_{YEAR }").ToString
                Dim StateCode As String = LGA_States(sqldr($"STATE_CODE_{YEAR }"))
                If NewCodes.ContainsKey(LGAcode) Then
                    ShireId = NewCodes(LGAcode)
                End If
                sqlWriter.WriteLine($"INSERT INTO `SHIRES` (`LGA`,`ShireState`,`LGALookup`,`ShireName`,`ShireID`,`SD`,`SSD`,`SLA`) VALUES ('{LGAcode }','{StateCode }','{LGAname }','{LongShire(LGAname) }','{ShireId }','0','0','0');")
                count += 1
            End While
            sqlWriter.WriteLine($"/* {count } shires exist in ABS LGA data that do not exist in SHIRES */")
            sqlWriter.WriteLine()
            sqldr.Close()

            ' 4. Check spelling of shire name
            sql.CommandText = $"Select * From SHIRES As a Join LGA_{YEAR } As b On a.LGA=b.LGA_CODE_{YEAR } WHERE a.LGALookup<>b.LGA_NAME_{YEAR }"
            sqldr = sql.ExecuteReader()
            count = 0
            While sqldr.Read()
                Dim LGAcode As String = $"LGA_CODE_{YEAR }"
                Dim LGAname As String = $"LGA_NAME_{YEAR }"
                If sqldr("LGALookup").ToString <> sqldr(LGAcode).ToString Then
                    sqlWriter.WriteLine(String.Format("UPDATE `SHIRES` SET `LGALookup`='{1}',`ShireName`='{2}' WHERE `id`={0}; /* Updating '{3}' to '{1}' */", sqldr![id], sqldr(LGAname).replace("'", "''"), LongShire(LGAcode).Replace("'", "''"), sqldr("LGALookup").replace("'", "''")))
                    count += 1
                End If
            End While
            sqlWriter.WriteLine($"/* {count } shires had a missing or changed name */")
            sqlWriter.WriteLine()
            sqldr.Close()

            '5. Check spelling of ShireName matches LGALookup
            sql.CommandText = "Select * From SHIRES ORDER BY ShireName"
            sqldr = sql.ExecuteReader()
            count = 0
            While sqldr.Read()
                Dim LongName As String
                LongName = LongShire(sqldr("LGALookup").ToString)
                If LongName <> sqldr![ShireName].ToString Then
                    sqlWriter.WriteLine(String.Format("UPDATE `SHIRES` SET `ShireName`='{1}' WHERE `id`={0}; /* Updating '{2}' to '{1}' */", sqldr![id], LongName.Replace("'", "''"), sqldr![ShireName].ToString))
                    count += 1
                End If
            End While
            sqlWriter.WriteLine($"/* {count } shires had long name changed */")
            sqlWriter.WriteLine()
            sqldr.Close()
            sqlWriter.Close()
            TextBox1.Text = "Done"
        End Using
    End Sub

    Shared Function LongShire(shire As String) As String
        ' Convert a shire name with a shire code in it to long form
        Dim result As String = shire
        Dim LongNames As New Dictionary(Of String, String) From {
            {"(A)", "Shire Council"},
            {"(DC)", "District Council"},
            {"(S)", "Shire Council"},
            {"(C)", "City Council"},
            {"(RC)", "Rural City Council"},
            {"(R)", "Regional Council"},
            {"(AC)", "Aboriginal Council"},
            {"(M)", "Municipality"},
            {"(T)", "Town Council"},
            {"(CGC)", "Community Government Council"},
            {"(RegC)", "Regional Council"},
            {"(IC)", "Island Council"},
            {"(B)", "Borough"},
            {"(Tas.)", "(Tas.)"},
            {"(Vic.)", "(Vic.)"},
            {"(NSW.)", "(NSW.)"},
            {"(Qld.)", "(Qld.)"}
            }
        Contract.Requires(Not String.IsNullOrEmpty(shire), "Illegal shire")
        If shire.Contains("(") Then
            For Each kvp As KeyValuePair(Of String, String) In LongNames
                Dim key As String = kvp.Key.Replace("(", "\(").Replace(")", "\)").Replace(".", "\.")   ' escape brackets and dots
                result = Regex.Replace(result, key, kvp.Value)
            Next
            Contract.Requires(shire <> result, $"Could not convert '{shire }' to long version")  ' check short name was recognised
        End If
        Return result
    End Function

    Private Async Sub FindOverlappingParksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindOverlappingParksToolStripMenuItem.Click
        ' Make list of parks that overlap
        Dim myQueryFilter1 As New QueryParameters, myQueryFilter2 As New QueryParameters
        Dim overlaps As FeatureQueryResult
        Dim SqlParks As SQLiteCommand, SQLParksdr As SQLiteDataReader
        Dim SqlMapping As SQLiteCommand, SQLMappingdr As SQLiteDataReader
        Dim WWFFID As String, OverlappedWWFFID As String, ThisParkData As NameValueCollection, OverlappedParkData As NameValueCollection

        ' Get all parks with overlaps
        myQueryFilter1.WhereClause = "OVERLAP='2'"    ' query parameters
        myQueryFilter1.ReturnGeometry = True
        myQueryFilter1.OutSpatialReference = SpatialReferences.Wgs84
        overlaps = Await DataSets("CAPAD_T").shpShapeFileTable.QueryFeaturesAsync(myQueryFilter1).ConfigureAwait(False)           ' run query
        'TextBox2.Text = "Found " & overlaps.Count & " overlapping parks" & vbCrLf

        Using connect As New SQLiteConnection(PARKSdb)
            connect.Open()  ' open database
            SqlParks = connect.CreateCommand
            SqlMapping = connect.CreateCommand

            ' Search for the overlapping park(s)
            For Each overlap In overlaps
                ' Find WWFFID for this park
                SqlParks.CommandText = String.Format("SELECT * FROM parks AS A JOIN GISmapping AS B ON A.WWFFID=B.WWFFID WHERE B.GISID='{0}' AND Status IN ('active','Active', 'pending')", overlap.GetAttributeValue("PA_ID").ToString)
                SQLParksdr = SqlParks.ExecuteReader()
                If SQLParksdr.Read() Then
                    WWFFID = SQLParksdr.Item("WWFFID").ToString
                    TextBox1.Text = "Checking " & WWFFID
                    ThisParkData = GetParkData(WWFFID)
                    ' Now look for where parks intersect
                    myQueryFilter2.Geometry = overlap.Geometry
                    myQueryFilter2.WhereClause = String.Format("PA_ID <> '{0}'", overlap.GetAttributeValue("PA_ID").ToString)    ' ignore self
                    myQueryFilter2.SpatialRelationship = SpatialRelationship.Intersects
                    myQueryFilter2.ReturnGeometry = True
                    myQueryFilter2.OutSpatialReference = SpatialReferences.Wgs84
                    Dim parks = Await DataSets("CAPAD_T").shpShapeFileTable.QueryFeaturesAsync(myQueryFilter2).ConfigureAwait(False)
                    For Each park In parks
                        ' Check if overlapped park is WWFF
                        SqlMapping.CommandText = String.Format("SELECT * FROM GISmapping WHERE GISID='{0}'", park.GetAttributeValue("PA_ID").ToString)
                        SQLMappingdr = SqlMapping.ExecuteReader()
                        If SQLMappingdr.Read() Then
                            OverlappedWWFFID = SQLMappingdr.Item("WWFFID").ToString
                            OverlappedParkData = GetParkData(OverlappedWWFFID)
                            Dim over As Geometry = GeometryEngine.Intersection(overlap.Geometry, park.Geometry)   ' geometry of overlapping area
                            If over IsNot Nothing AndAlso over.Extent IsNot Nothing Then
                                Dim LabelPoint As MapPoint = GeometryEngine.LabelPoint(over)
                                If Not LabelPoint.IsEmpty Then
                                    TextBox2.AppendText(String.Format("{0} - {1} in {2} has an overlap with {3} - {4} in {5}" & vbCrLf, WWFFID, ThisParkData("Name"), ThisParkData("State"), OverlappedWWFFID, OverlappedParkData("Name"), OverlappedParkData("State")))
                                    Dim Area As Double = GeometryEngine.AreaGeodetic(over, AreaUnits.Hectares)
                                    TextBox2.AppendText(String.Format("Overlap at {0:f5},{1:f5} Area {2:f1} ha" & vbCrLf, LabelPoint.Y, LabelPoint.X, Area))
                                End If
                            End If
                        End If
                        SQLMappingdr.Close()
                    Next
                End If
                SQLParksdr.Close()
            Next
            TextBox2.AppendText("Done")
            connect.Close()
        End Using
    End Sub

    Private Sub MakeSKAOverlayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeSKAOverlayToolStripMenuItem.Click
        Dim SKA As New MapPoint(116.658889, -26.704167, New SpatialReference(4283))    ' location of Square Kilometer Array in GDA94 datum
        Dim kml As New XmlDocument, rootNode As XmlElement, docNode As XmlElement, desNode As XmlElement, CDATA As XmlCDataSection, pm As XmlElement, st As XmlElement, st1 As XmlElement
        Dim at As XmlElement, co As XmlElement, pt As XmlElement, ls As XmlElement, nm As XmlElement
        Dim BaseFileName As String = Application.StartupPath & "\files\VK6\SKA"

        TextBox2.Text = "Make overlay files for Square Kilomenter Array"
        Dim SKA84 As MapPoint = GeometryEngine.Project(SKA, SpatialReferences.Wgs84)     ' convert to WGS84
        Dim decNode = kml.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
        kml.AppendChild(decNode)
        rootNode = kml.CreateElement("kml")
        rootNode.SetAttribute("xmlns", "http://www.opengis.net/kml/2.2")
        docNode = kml.CreateElement("Document")

        desNode = kml.CreateElement("description")
        CDATA = kml.CreateCDataSection("Overlay to show extent of Square Kilometer Array exclusion zones. Created by VK3OHM")
        desNode.AppendChild(CDATA)
        docNode.AppendChild(desNode)

        st = kml.CreateElement("Style")
        st.SetAttribute("id", "Inner")
        st1 = kml.CreateElement("PolyStyle")
        st1.InnerXml = "<color>7f0000ff</color><fill>1</fill>"
        st.AppendChild(st1)
        docNode.AppendChild(st)

        st = kml.CreateElement("Style")
        st.SetAttribute("id", "Outer")
        st1 = kml.CreateElement("PolyStyle")
        st1.InnerXml = "<color>7f00ffff</color><fill>1</fill>"
        st.AppendChild(st1)
        docNode.AppendChild(st)

        st = kml.CreateElement("Style")
        st.SetAttribute("id", "Coord")
        st1 = kml.CreateElement("PolyStyle")
        st1.InnerXml = "<color>7fffff00</color><fill>1</fill>"
        st.AppendChild(st1)
        docNode.AppendChild(st)

        pm = kml.CreateElement("Placemark")
        pm.SetAttribute("id", "SKA")
        at = kml.CreateElement("name")
        at.InnerText = "SKA"
        pm.AppendChild(at)

        desNode = kml.CreateElement("description")
        CDATA = kml.CreateCDataSection("The Square Kilometer Array (SKA)")
        desNode.AppendChild(CDATA)
        pm.AppendChild(desNode)

        pt = kml.CreateElement("Point")
        co = kml.CreateElement("coordinates")
        co.InnerText = $"{SKA84.X:f5},{SKA84.Y:f5}"
        pt.AppendChild(co)
        pm.AppendChild(pt)
        docNode.AppendChild(pm)

        ' Create range circles
        pm = kml.CreateElement("Placemark")
        nm = kml.CreateElement("name")
        nm.InnerText = "Inner zone"
        pm.AppendChild(nm)
        desNode = kml.CreateElement("description")
        CDATA = kml.CreateCDataSection("The Inner zone 0 - 70km. Primary use astronomy. There are limitations on emissions in the range 70MHz to 25.25GHz.")
        desNode.AppendChild(CDATA)
        pm.AppendChild(desNode)
        nm = kml.CreateElement("styleUrl")
        nm.InnerText = "#Inner"
        pm.AppendChild(nm)
        ls = Annulus(kml, SKA84, 0, 70 * 1000, "Inner")     ' Inner Zone
        pm.AppendChild(ls)
        docNode.AppendChild(pm)

        pm = kml.CreateElement("Placemark")
        nm = kml.CreateElement("name")
        nm.InnerText = "Outer zone"
        pm.AppendChild(nm)
        desNode = kml.CreateElement("description")
        CDATA = kml.CreateCDataSection("The Outer zone 70 - 150km. Primary use astronomy. There are limitations on emissions in the range 70MHz to 25.25GHz.")
        desNode.AppendChild(CDATA)
        pm.AppendChild(desNode)
        nm = kml.CreateElement("styleUrl")
        nm.InnerText = "#Outer"
        pm.AppendChild(nm)
        ls = Annulus(kml, SKA84, 70 * 1000, 150 * 1000, "Outer")     ' Outer Zone
        pm.AppendChild(ls)
        docNode.AppendChild(pm)

        pm = kml.CreateElement("Placemark")
        nm = kml.CreateElement("name")
        nm.InnerText = "Coordination zone"
        pm.AppendChild(nm)
        desNode = kml.CreateElement("description")
        CDATA = kml.CreateCDataSection("The Coordination zone 150 - 260km. There are limitations on emissions in the range 70MHz to 25.25GHz.")
        desNode.AppendChild(CDATA)
        pm.AppendChild(desNode)
        nm = kml.CreateElement("styleUrl")
        nm.InnerText = "#Coord"
        pm.AppendChild(nm)
        ls = Annulus(kml, SKA84, 150 * 1000, 260 * 1000, "Coord")     ' Coordination Zone
        pm.AppendChild(ls)
        docNode.AppendChild(pm)
        '         kmlWriter.WriteLine("<ExtendedData><Data name=""north""><value>{0:f5}</value></Data><Data name=""south""><value>{1:f5}</value></Data><Data name=""east""><value>{2:f5}</value></Data><Data name=""west""><value>{3:f5}</value></Data></ExtendedData>", extent.YMax, extent.YMin, extent.XMax, extent.XMin)
        ' Write out the file
        rootNode.AppendChild(docNode)
        kml.AppendChild(rootNode)
        Dim writer As New XmlTextWriter(BaseFileName & ".kml", Nothing) With {
            .Formatting = Xml.Formatting.Indented
        }
        kml.Save(writer)
        writer.Close()
        ' compress to zip file
        System.IO.File.Delete(BaseFileName & ".kmz")
        Dim zip As ZipArchive = ZipFile.Open(BaseFileName & ".kmz", ZipArchiveMode.Create)    ' create new archive file
        zip.CreateEntryFromFile(BaseFileName & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
        zip.Dispose()
        TextBox2.AppendText($"{vbCrLf }SKA overlay files created in {BaseFileName }.kmz")
    End Sub

    Shared Function Annulus(Xml As XmlDocument, p As MapPoint, InnerRadius As Single, OuterRadius As Single, id As String) As XmlElement
        ' Generate an Annulus with inner and outer radius. Return Polygon
        Dim pg As XmlElement, co As XmlElement, el As XmlElement, coords As New List(Of String)
        Dim ob As XmlElement, ib As XmlElement, lr As XmlElement, buffer As Polygon

        Contract.Requires(Xml IsNot Nothing, "Bad XML document")
        Contract.Requires(p IsNot Nothing And Not p.IsEmpty, "Bad point")
        Contract.Requires(InnerRadius > 0, "Bad inner radius")
        Contract.Requires(OuterRadius > InnerRadius, "Bad outer radius")
        Contract.Requires(Not String.IsNullOrEmpty(id), "Bad id")

        coords.Clear()
        pg = Xml.CreateElement("Polygon")        ' return a polygon
        pg.SetAttribute("id", id)
        el = Xml.CreateElement("tessellate")
        el.InnerText = "1"
        pg.AppendChild(el)
        ' generate the outer boundary
        ob = Xml.CreateElement("outerBoundaryIs")
        lr = Xml.CreateElement("LinearRing")
        co = Xml.CreateElement("coordinates")
        buffer = GeometryEngine.BufferGeodetic(p, OuterRadius, LinearUnits.Meters)        ' generate circle
        coords.Clear()
        For Each pnt As MapPoint In buffer.Parts(0).Points
            coords.Add($"{pnt.X:f5},{pnt.Y:f5}")
        Next
        co.InnerText = String.Join(" ", coords)
        lr.AppendChild(co)
        ob.AppendChild(lr)
        pg.AppendChild(ob)

        ' generate the inner boundary
        If InnerRadius > 0 Then
            ib = Xml.CreateElement("innerBoundaryIs")
            lr = Xml.CreateElement("LinearRing")
            co = Xml.CreateElement("coordinates")
            buffer = GeometryEngine.BufferGeodetic(p, InnerRadius, LinearUnits.Meters)        ' generate circle
            coords.Clear()
            For Each pnt As MapPoint In buffer.Parts(0).Points
                coords.Add($"{pnt.X:f5},{pnt.Y:f5}")
            Next
            co.InnerText = String.Join(" ", coords)
            lr.AppendChild(co)
            ib.AppendChild(lr)
            pg.AppendChild(ib)
        End If
        Return pg
    End Function

    Private Sub MakeRQZOverlayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeRQZOverlayToolStripMenuItem.Click
        ' Make overlay for other RQZ
        Dim AGD66 As New SpatialReference(4202)     ' AGD66
        Dim kml As New XmlDocument(), docNode As XmlElement, rootNode As XmlElement, desNode As XmlElement, CDATA As XmlCDataSection
        Dim st As XmlElement, st1 As XmlElement, pm As XmlElement, nm As XmlElement, ls As XmlElement, pt As XmlElement, co As XmlElement, at As XmlElement
        Dim BaseFileName As String = Application.StartupPath & "\files\RQZ"

        Dim RQZ As New List(Of RadioQuietZone) From {
            New RadioQuietZone("Parkes", "Parkes Observatory", CoordinateFormatter.FromLatitudeLongitude("-32° 59' 59.8657, 148° 15' 44.3591", AGD66), 100 * 1000),
            New RadioQuietZone("Narrabri", "Paul Wild Observatory", CoordinateFormatter.FromLatitudeLongitude("-30° 59' 52.084, 149° 32' 56.327", AGD66), 100 * 1000),
            New RadioQuietZone("Coonabarabran", "Mopra Observatory", CoordinateFormatter.FromLatitudeLongitude("-31° 16' 4.451, 149° 5' 58.732", AGD66), 100 * 1000),
            New RadioQuietZone("Hobart", "Mount Pleasant Observatory", CoordinateFormatter.FromLatitudeLongitude("-42° 48' 12.9207, 147° 26' 25.854", AGD66), 100 * 1000),
            New RadioQuietZone("Ceduna", "Ceduna Observatory", CoordinateFormatter.FromLatitudeLongitude("-31° 52' 8.8269, 133° 48' 35.3748", AGD66), 100 * 1000),
            New RadioQuietZone("Canberra", "Deep Space Communication Complex", CoordinateFormatter.FromLatitudeLongitude("-35° 23' 54, 148° 58' 40", AGD66), 100 * 1000)
        }

        TextBox2.Text = "Make overlay files for Radio Quiet zones"
        Dim decNode = kml.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
        kml.AppendChild(decNode)
        rootNode = kml.CreateElement("kml")
        rootNode.SetAttribute("xmlns", "http://www.opengis.net/kml/2.2")
        docNode = kml.CreateElement("Document")

        desNode = kml.CreateElement("description")
        CDATA = kml.CreateCDataSection("Overlay to show extent of Radio Quiet Zones. Created by VK3OHM")
        desNode.AppendChild(CDATA)
        docNode.AppendChild(desNode)

        st = kml.CreateElement("Style")
        st.SetAttribute("id", "Circle")
        st1 = kml.CreateElement("PolyStyle")
        st1.InnerXml = "<color>7f0000ff</color><fill>1</fill>"
        st.AppendChild(st1)
        docNode.AppendChild(st)

        For Each QuietZone As RadioQuietZone In RQZ
            pm = kml.CreateElement("Placemark")
            pm.SetAttribute("id", QuietZone.Name)
            at = kml.CreateElement("name")
            at.InnerText = QuietZone.Name
            pm.AppendChild(at)

            desNode = kml.CreateElement("description")
            CDATA = kml.CreateCDataSection("Radio Quiet Zone - " & QuietZone.Name & " at " & QuietZone.Location)
            desNode.AppendChild(CDATA)
            pm.AppendChild(desNode)

            pt = kml.CreateElement("Point")
            co = kml.CreateElement("coordinates")
            co.InnerText = $"{QuietZone.Position.X:f5},{QuietZone.Position.Y:f5}"
            pt.AppendChild(co)
            pm.AppendChild(pt)
            docNode.AppendChild(pm)

            pm = kml.CreateElement("Placemark")
            nm = kml.CreateElement("name")
            nm.InnerText = QuietZone.Name
            pm.AppendChild(nm)
            docNode.AppendChild(pm)

            nm = kml.CreateElement("styleUrl")
            nm.InnerText = "#Circle"
            pm.AppendChild(nm)
            ls = Annulus(kml, QuietZone.Position, 0, QuietZone.Radius, "Circle")     ' Outer Zone
            pm.AppendChild(ls)
            docNode.AppendChild(pm)
        Next

        ' Write out the file
        rootNode.AppendChild(docNode)
        kml.AppendChild(rootNode)
        Using writer As New XmlTextWriter(BaseFileName & ".kml", Nothing) With {
            .Formatting = Xml.Formatting.Indented
        }
            kml.Save(writer)
            writer.Close()
        End Using
        ' compress to zip file
        System.IO.File.Delete(BaseFileName & ".kmz")
        Dim zip As ZipArchive = ZipFile.Open(BaseFileName & ".kmz", ZipArchiveMode.Create)    ' create new archive file
        zip.CreateEntryFromFile(BaseFileName & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
        zip.Dispose()
        Me.TextBox1.Text = "Done"
        TextBox2.AppendText($"{vbCrLf }RQZ overlay files created in {BaseFileName }.kmz")
    End Sub

    Private Sub Extract2018CommentsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Extract2018CommentsToolStripMenuItem.Click
        ' Extract comments from 2018 CAPAD
        Dim sql As SQLiteCommand, sqldr As SQLiteDataReader, sqldrdbf As System.Data.OleDb.OleDbDataReader, LastState As String = ""

        Using connect As New SQLiteConnection(PARKSdb), logWriter As New System.IO.StreamWriter("CAPAD 2018 comments.html", False)
            logWriter.WriteLine("<table border=1>")
            logWriter.WriteLine("<tr><th>WWFFID</th><th>PA_ID</th><th>PA_PID</th><th>Name</th><th>Type</th><th>Comment</th></tr>")
            connect.Open()  ' open database
            sql = connect.CreateCommand

            ' Search all parks
            sql.CommandText = "SELECT * FROM parks AS A JOIN GISmapping as B ON A.WWFFID=B.WWFFID and B.Dataset='CAPAD_T' ORDER BY State,Name,Type"
            sqldr = sql.ExecuteReader()
            While sqldr.Read()
                ' Get comments
                Dim query = $"SELECT * FROM CAPAD2018_terrestrial WHERE PA_ID='{sqldr.Item("GISID")}'"
                Using sqldbf As New System.Data.OleDb.OleDbCommand(query, DataSets("CAPAD_T").dbfConnection)
                    sqldrdbf = sqldbf.ExecuteReader()
                    While sqldrdbf.Read()
                        If sqldrdbf.Item("COMMENTS").ToString <> "" Then
                            If sqldr.Item("State").ToString <> LastState Then
                                logWriter.WriteLine($"<tr><td colspan=6 align='center'><b><big>{sqldr.Item("State")}</big></b></td></tr>")     ' State header
                                LastState = sqldr.Item("State").ToString
                            End If
                            logWriter.WriteLine($"<tr><td>{sqldr.Item("WWFFID")}</td><td>{sqldr.Item("GISID")}</td><td>{sqldrdbf.Item("PA_PID")}</td><td>{sqldr.Item("Name")}</td><td>{sqldr.Item("Type")}</td><td>{sqldrdbf.Item("COMMENTS")}</td></tr>")
                            TextBox1.Text = sqldr.Item("WWFFID").ToString
                            Application.DoEvents()
                            logWriter.Flush()
                        End If
                    End While
                    sqldrdbf.Close()
                End Using
            End While
            sqldr.Close()
            logWriter.WriteLine("</table>")
            logWriter.Close()
            TextBox1.Text = "Done"
        End Using
    End Sub
    Private Sub Extract2020CommentsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Extract2020CommentsToolStripMenuItem.Click
        ' Extract comments from 2020 CAPAD
        Dim sql As SQLiteCommand, sqldr As SQLiteDataReader, sqldrdbf As System.Data.OleDb.OleDbDataReader, LastState As String = ""

        Using connect As New SQLiteConnection(PARKSdb), logWriter As New System.IO.StreamWriter("CAPAD 2020 comments.html", False)
            logWriter.WriteLine("<table border=1>")
            logWriter.WriteLine("<tr><th>WWFFID</th><th>PA_ID</th><th>PA_PID</th><th>Name</th><th>Type</th><th>Comment</th></tr>")
            connect.Open()  ' open database
            sql = connect.CreateCommand

            ' Search all parks
            sql.CommandText = "SELECT * FROM parks AS A JOIN GISmapping as B ON A.WWFFID=B.WWFFID and B.Dataset='CAPAD_T' ORDER BY State,Name,Type"
            sqldr = sql.ExecuteReader()
            While sqldr.Read()
                ' Get comments
                Dim query = $"SELECT * FROM {DataSets("CAPAD_T").dbfTableName } WHERE PA_ID='{sqldr.Item("GISID")}'"
                Using sqldbf As New System.Data.OleDb.OleDbCommand(query, DataSets("CAPAD_T").dbfConnection)
                    sqldrdbf = sqldbf.ExecuteReader()
                    While sqldrdbf.Read()
                        If sqldrdbf.Item("COMMENTS").ToString <> "" Then
                            If sqldr.Item("State").ToString <> LastState Then
                                logWriter.WriteLine($"<tr><td colspan=6 align='center'><b><big>{sqldr.Item("State")}</big></b></td></tr>")     ' State header
                                LastState = sqldr.Item("State").ToString
                            End If
                            logWriter.WriteLine($"<tr><td>{sqldr.Item("WWFFID")}</td><td>{sqldr.Item("GISID")}</td><td>{sqldrdbf.Item("PA_PID")}</td><td>{sqldr.Item("Name")}</td><td>{sqldr.Item("Type")}</td><td>{sqldrdbf.Item("COMMENTS")}</td></tr>")
                            TextBox1.Text = sqldr.Item("WWFFID").ToString
                            Application.DoEvents()
                            logWriter.Flush()
                        End If
                    End While
                    sqldrdbf.Close()
                End Using
            End While
            sqldr.Close()
            logWriter.WriteLine("</table>")
            logWriter.Close()
            TextBox1.Text = "Done"
        End Using
    End Sub
    Private Sub TestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestToolStripMenuItem.Click
        Dim latitude As Double = 5.5, longitude As Double = 6.6
        Dim params = System.Web.HttpUtility.ParseQueryString(String.Empty)  ' tricky way to create an HttpValueCollection, which you can't do directly because it's internal
        params.Add("f", "json")          ' return json
        params.Add("geometryType", "esriGeometryPoint")   ' look for a point (centroid)
        params.Add("geometry", $"{longitude:f5},{latitude:f5}")   ' point to look for
        params.Add("inSR", "{'wkid' : 4326}")                                  ' WGS84 datum
        params.Add("spatialRel", "esriSpatialRelWithin")
        params.Add("returnGeometry", "false")                  ' don't need the geometry
        params.Add("outFields", "*")                          ' return all fields, even though we only use the default

        Dim url = params.ToString
    End Sub

    Private Sub GetSpotElevationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetSpotElevationToolStripMenuItem.Click
        Dim elevation As Single, resolution As Single
        Dim point As String = InputBox("Enter position as long,lat", "Enter position")
        Dim position As String() = point.Split(",")
        GetSpotElevation(position(0), position(1), elevation, resolution)
    End Sub

    Private Shared Sub GetSpotElevation(longitude As Double, latitude As Double, ByRef elevation As Single, ByRef resolution As Single)
        ' Get a spot elevation from google maps API
        Dim request As HttpWebRequest, response As System.Net.HttpWebResponse, sourcecode As String
        Dim url As New Uri($"https://maps.googleapis.com/maps/api/elevation/json?locations={latitude:f5},{longitude:f5}&key={GOOGLE_API_KEY }")
        request = WebRequest.Create(url)
        response = request.GetResponse()
        Using sr As New StreamReader(response.GetResponseStream())
            sourcecode = sr.ReadToEnd()
        End Using
        Dim jsonObject As JObject = JObject.Parse(sourcecode)
        Dim results As JArray = jsonObject("results")
        elevation = results.First("elevation")
        resolution = results.First("resolution")
    End Sub

    Private Async Sub ImportSOTADatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportSOTADatabaseToolStripMenuItem.Click
        ' Import the current SOTA database
        Dim sourcecode As String, lines As List(Of String)
        Dim header As String, header_list As New List(Of String), fields As New List(Of String), deleted As Integer, TEXTfields As New List(Of Integer)
        Dim exclude As New List(Of String) From {"ActivationCount", "ActivationDate", "ActivationCall"}     ' exclude volitile fields
        Dim url As New Uri("https://www.sotadata.org.uk/summitslist.csv")
        Dim dt As String = $"{DateTime.Now:yyyyMMddHHmmss}"

        TextBox2.Text = $"requesting SOTA database from {url }"
        ' read the data asynchronously. This avoids a "ConextSwitchDeadlock" error because it takes longer than 60 secs
        Dim webReq = CType(WebRequest.Create(url), HttpWebRequest)
        webReq.Timeout = 5 * 60 * 1000        ' 5 min timeout
        Using response As WebResponse = Await webReq.GetResponseAsync().ConfigureAwait(True)
            Dim content As New MemoryStream()
            Using responseStream As Stream = response.GetResponseStream()
                Await responseStream.CopyToAsync(content).ConfigureAwait(True)
            End Using
            ' convert the MemoryStream into a string
            Using sr = New StreamReader(content)
                content.Position = 0    ' rewind to start
                sourcecode = sr.ReadToEnd()
            End Using
        End Using
        TextBox2.AppendText($"{vbCrLf }received {sourcecode.Length } bytes of data")
        lines = sourcecode.Split(vbLf).ToList
        header_list = lines(1).Split(",").ToList

        TextBox2.AppendText($"{vbCrLf }extracted {lines.Count } lines of data")
        ' process each line
        Dim summits As Integer = 0       ' count of summits
        lines.RemoveAt(lines.Count - 1)      ' remove last blank line
        Using connect As New SQLiteConnection(SOTAdb)
            Dim sqlcmd As SQLiteCommand
            connect.Open()  ' open database
            sqlcmd = connect.CreateCommand
            sqlcmd.CommandText = "BEGIN TRANSACTION"   ' start transaction
            deleted = sqlcmd.ExecuteNonQuery()
            sqlcmd.CommandText = "DELETE FROM `SOTA`"   ' remove existing data
            deleted = sqlcmd.ExecuteNonQuery()
            TextBox2.AppendText($"{vbCrLf }deleted {deleted } SOTA summits")
            header = String.Join(",", header_list)      ' csv list of field names
            Dim values_list As New List(Of String)
            For Each f As String In header_list
                values_list.Add($"@{f }")
            Next
            Dim values As String = String.Join(",", values_list)        ' csv list of parameter placeholders
            sqlcmd.CommandText = $"INSERT INTO `SOTA` ({header }) VALUES ({values })"   ' add this summit
            sqlcmd.Prepare()        ' compile SQL for repeated use
            For line = 2 To lines.Count - 1
                If lines(line).StartsWith("VK") Or lines(line).StartsWith("ZL") Then
                    summits += 1
                    fields = lines(line).Split(",").ToList
                    sqlcmd.Parameters.Clear()
                    Dim indx As Integer
                    For indx = 0 To header_list.Count - 1
                        fields(indx) = fields(indx).TrimStart("""")        ' remove double quotes
                        fields(indx) = fields(indx).TrimEnd("""")
                        sqlcmd.Parameters.AddWithValue(values_list(indx), fields(indx))
                    Next
                    sqlcmd.ExecuteNonQuery() ' add this summit
                    Application.DoEvents()
                    If (summits Mod 1000) = 0 Then TextBox2.AppendText($"{vbCrLf }Added {summits } summits")
                End If
            Next
            sqlcmd.CommandText = "COMMIT"   ' end transaction
            deleted = sqlcmd.ExecuteNonQuery()
            ExportCSV(connect, "SOTA", $"SOTA_{dt }.csv", exclude)    ' export after database
        End Using
        TextBox2.AppendText($"{vbCrLf }data for {summits } summits extracted")
        TextBox2.AppendText($"{vbCrLf }CSV export in SOTA_{dt }.csv")
    End Sub

    Private Shared Sub ExportCSV(connect As SQLiteConnection, table As String, filename As String, Optional exclude As List(Of String) = Nothing)
        ' export table to csv file
        ' connect = existing SQLite connection
        ' table = name of table
        ' filename = name of csv file
        ' exclude = optional list of columns to exclude

        Dim sql As SQLiteCommand, sqldr As SQLiteDataReader, count As Integer = 0
        Dim names As New List(Of String), types As New List(Of String), valuelist As New List(Of String)
        Dim csv As New StreamWriter(filename)
        sql = connect.CreateCommand
        sql.CommandText = $"SELECT * FROM {table }"
        sqldr = sql.ExecuteReader()
        Dim values(sqldr.FieldCount - 1) As Object
        While sqldr.Read()
            If count = 0 Then
                ' get data for header
                Dim header As New List(Of String)
                For i = 0 To sqldr.FieldCount - 1
                    Dim name As String = sqldr.GetName(i)
                    names.Add(name)
                    types.Add(sqldr.GetDataTypeName(i))
                    If exclude Is Nothing Or Not exclude.Contains(name) Then header.Add(name)
                Next
                csv.WriteLine(String.Join(",", header))  ' header line
            End If

            sqldr.GetValues(values)     ' read of values at once
            valuelist.Clear()
            For i = LBound(values) To UBound(values)
                If exclude Is Nothing Or Not exclude.Contains(names(i)) Then
                    If types(i) = "TEXT" Then
                        valuelist.Add($"""{values(i) }""")      ' enclose in double quotes
                    Else
                        valuelist.Add(values(i))
                    End If
                End If
            Next
            csv.WriteLine(String.Join(",", valuelist))
            count += 1
        End While
        sqldr.Close()
        csv.Close()
    End Sub
    Private Sub GetGEElevationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetGEElevationToolStripMenuItem.Click
        ' Get missing GE elevation data for SOTA summits
        Dim count As Integer    ' count of summits
        Using SOTAconnect As New SQLiteConnection(SOTAdb) ' declare the connection        
            Dim SOTAsql As SQLiteCommand
            Dim SOTAsqldr As SQLiteDataReader
            Dim SOTAsqlw As SQLiteCommand

            Dim SummitCode As String, latitude As Single, longitude As Single, elevation As Single, resolution As Single

            Application.UseWaitCursor = True
            SOTAconnect.Open()  ' open database
            SOTAsql = SOTAconnect.CreateCommand
            SOTAsqlw = SOTAconnect.CreateCommand
            SOTAsqlw.CommandText = $"INSERT INTO `GEheight` (`SummitCode`,`elevation`,`resolution`) VALUES (@SummitCode,@elevation,@resolution)"     ' update height/resolution
            SOTAsqlw.Prepare()

            SOTAsql.CommandText = "SELECT * FROM `SOTA` LEFT JOIN `GEheight` ON `SOTA`.`SummitCode`=`GEheight`.`SummitCode` WHERE `GEheight`.`SummitCode` ISNULL"     ' select all summits with missing elevation
            SOTAsqldr = SOTAsql.ExecuteReader()
            count = 0
            While SOTAsqldr.Read()
                SummitCode = SOTAsqldr.Item("SummitCode")
                longitude = CSng(SOTAsqldr.Item("GridRef1"))
                latitude = CSng(SOTAsqldr.Item("GridRef2"))
                GetSpotElevation(longitude, latitude, elevation, resolution)    ' get GE height
                SOTAsqlw.Parameters.Clear()
                SOTAsqlw.Parameters.AddWithValue("@SummitCode", SummitCode)
                SOTAsqlw.Parameters.AddWithValue("@elevation", CInt(elevation))
                SOTAsqlw.Parameters.AddWithValue("@resolution", CInt(resolution))
                SOTAsqlw.ExecuteNonQuery()
                count += 1
                Application.DoEvents()
            End While
            SOTAsqldr.Close()
        End Using
        TextBox2.AppendText($"Retrieved heights for {count } summits{vbCrLf }")
        Application.UseWaitCursor = False
    End Sub

    Private Sub NAVMANPOIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NAVMANPOIToolStripMenuItem.Click
        ' make data file of SOTA summits for NAVMAN
        Dim count As Integer    ' count of summits
        Using SOTAconnect As New SQLiteConnection(SOTAdb) ' declare the connection       
            Dim SOTAsql As SQLiteCommand
            Dim SOTAsqldr As SQLiteDataReader
            Dim SummitCode As String, latitude As Single, longitude As Single, Name As String
            Dim poi As New StreamWriter("SOTA_NAVMAN.csv", False)

            Application.UseWaitCursor = True
            SOTAconnect.Open()  ' open database
            SOTAsql = SOTAconnect.CreateCommand

            SOTAsql.CommandText = "SELECT * FROM `SOTA` WHERE SummitCode LIKE 'VK%' ORDER BY SummitCode"
            SOTAsqldr = SOTAsql.ExecuteReader()
            count = 0
            While SOTAsqldr.Read()
                SummitCode = SOTAsqldr.Item("SummitCode")
                Name = SOTAsqldr.Item("SummitName")
                longitude = CSng(SOTAsqldr.Item("GridRef1"))
                latitude = CSng(SOTAsqldr.Item("GridRef2"))
                poi.WriteLine($"{longitude:f9},{latitude:f8},{SummitCode } {Name }")
                count += 1
                Application.DoEvents()
            End While
            SOTAsqldr.Close()
            TextBox2.AppendText($"Created {count } SOTA POI{vbCrLf }")
            poi.Close()
        End Using
        Application.UseWaitCursor = False
    End Sub

    '=======================================================================================================
    ' Items on Silos menu
    '=======================================================================================================
    Private Sub RailwayStationsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RailwayStationsToolStripMenuItem.Click
        Const MAPSERVER = "https://services.ga.gov.au/gis/rest/services/NM_Transport_Infrastructure/MapServer/4/"
        Const BASEFILENAME = "Railway_Stations"
        Dim POSTfields As NameValueCollection, count As Integer = 0, offset As Integer = 0
        Dim resp As Byte() = {}, responseStr As String
        Dim maxRecordCount As Integer   ' max number of records retrievable with 1 query
        Dim Jo As JObject
        Dim ExtendedData As New Dictionary(Of String, String)
        Dim KMLurl As New Dictionary(Of String, String)

        ' Get the capabilities of this map
        POSTfields = New NameValueCollection From {{"f", "json"}}
        Using myWebClient As New WebClient
            Try
                resp = myWebClient.UploadValues(MAPSERVER, "POST", POSTfields)    ' query map server
            Catch ex As WebException
                MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
            End Try
            responseStr = System.Text.Encoding.UTF8.GetString(resp)
            Jo = JObject.Parse(responseStr)
            maxRecordCount = Jo("maxRecordCount")
            ' Extract icons
            For Each uniqueValue In Jo("drawingInfo")("renderer")("uniqueValueInfos")
                Dim value As String = uniqueValue("value")
                If value <> "<Null>" Then
                    KMLurl.Add(value, uniqueValue("symbol")("url"))
                    ' create image file
                    Dim fileBytes As Byte(), streamImage As Bitmap
                    fileBytes = Convert.FromBase64String(uniqueValue("symbol")("imageData"))    ' convert Base64 to byte array
                    Using ms As New MemoryStream(fileBytes)
                        streamImage = Image.FromStream(ms)
                        streamImage.Save($"{uniqueValue("symbol")("url") }.png", System.Drawing.Imaging.ImageFormat.Png)    ' save image in png format
                    End Using
                End If
            Next

            ' Prepare some POST fields for a request to the map server. We use POST because the requests are too large for a GET
            POSTfields = New NameValueCollection From {
                {"f", "geojson"},
                {"where", "1=1"},
                {"GeometryType", "esriGeometryEnvelope"},
                {"spatialRel", "esriSpatialRelIntersects"},
                {"outFields", "*"},
                {"outSR", "WGS84"},
                {"returnGeometry", "true"},          '  need the geometry
                {"orderByFields", "name"},
                {"returnTrueCurves", "false"},
                {"returnIdsOnly", "false"},
                {"returnCountOnly", "false"},
                {"resultRecordCount", maxRecordCount}, ' maximum number of records we can receive
                {"returnZ", "false"},
                {"returnM", "false"},
                {"returnDistinctValues", "false"},
                {"returnExtentOnly", "false"},
                {"resultOffset", CStr(offset)}
            }

            ' Create kml header
            Using kmlWriter As New System.IO.StreamWriter($"{BASEFILENAME }.kml")
                kmlWriter.WriteLine(KMLheader)
                kmlWriter.WriteLine("<description><![CDATA[<style>table, th, td {white-space:nowrap; }</style>")
                kmlWriter.WriteLine("<table>")
                kmlWriter.WriteLine("<tr><td>Data produced by Marc Hillman - VK3OHM/VK3IP</td></tr>")
                Dim utc As String = String.Format("{0:dd MMM yyyy hh:mm:ss UTC}", DateTime.UtcNow)     ' time now in UTC
                kmlWriter.WriteLine($"<tr><td>Data extracted from {MAPSERVER } on {utc }.</td></tr>")
                kmlWriter.WriteLine("</table>")
                kmlWriter.WriteLine("]]>")
                kmlWriter.WriteLine("</description>")
                ' Create some icon styles
                For Each url In KMLurl
                    kmlWriter.WriteLine($"<Style id=""{url.Key }"">")
                    kmlWriter.WriteLine("<IconStyle>")
                    kmlWriter.WriteLine("<scale>0.75</scale>")
                    kmlWriter.WriteLine($"<Icon><href>{url.Value }.png</href></Icon>")
                    kmlWriter.WriteLine("</IconStyle>
                <LabelStyle>
                  <color>00000000</color>
                  <scale>0.000000</scale>
                </LabelStyle>
                <PolyStyle>
                  <color>ff000000</color>
                  <outline>0</outline>
                </PolyStyle>
            </Style>")
                Next

                Try
                    Dim done = False
                    While Not done
                        resp = Array.Empty(Of Byte)()   ' clear  array
                        Try
                            resp = myWebClient.UploadValues($"{MAPSERVER }query/", "POST", POSTfields)    ' query map server
                            TextBox1.Text = "retrieved " & resp.Length & " bytes of data at offset " & offset
                            responseStr = System.Text.Encoding.UTF8.GetString(resp)
                            My.Computer.FileSystem.WriteAllText("GeoJson.json", responseStr, True)
                            ' The GeoJSON output produced by the mapserver seems to be an old standard. Need to do some fixups for current version
                            responseStr = responseStr.Replace("Feature Layer", "FeatureCollection")
                            Dim pattern As String = "\""crs\"":.*?},"
                            responseStr = Regex.Replace(responseStr, pattern, "")
                            pattern = ",\""exceededTransferLimit\"":.*?}"
                            responseStr = Regex.Replace(responseStr, pattern, "}")
                            Jo = JObject.Parse(responseStr)
                            If Jo.HasValues And Jo("features").Any Then
                                For Each feature In Jo("features")
                                    count += 1   ' count number of segments
                                    ' collect Extended Data
                                    ExtendedData.Clear()
                                    ExtendedData.Add("Name", feature("properties")("name").ToString)
                                    ExtendedData.Add("Status", feature("properties")("railstationstatus").ToString)
                                    ExtendedData.Add("Type", feature("properties")("featuretype").ToString)
                                    If Not IsNothing(feature("properties")("textnote")) Then ExtendedData.Add("Note", feature("properties")("textnote").ToString)
                                    ' Create Placemark
                                    kmlWriter.WriteLine($"<Placemark id='{HttpUtility.HtmlEncode(ExtendedData("Name")) }'>")
                                    kmlWriter.WriteLine($"<styleUrl>#{ExtendedData("Status") }</styleUrl>")
                                    kmlWriter.WriteLine($"<name>{ExtendedData("Name") }</name>")

                                    ' write out ExtendedData
                                    kmlWriter.WriteLine("<ExtendedData>")
                                    For Each item In ExtendedData
                                        kmlWriter.WriteLine($"<Data name=""{item.Key }""><value>{item.Value }</value></Data>")
                                    Next
                                    kmlWriter.WriteLine("</ExtendedData>")
                                    ' plot the point
                                    kmlWriter.WriteLine("<Point>")
                                    kmlWriter.WriteLine("<altitudeMode>clampToGround</altitudeMode>")
                                    kmlWriter.WriteLine($"<coordinates>{feature("geometry")("coordinates")(0):f5},{feature("geometry")("coordinates")(1):f5}</coordinates>")
                                    kmlWriter.WriteLine("</Point>")
                                    kmlWriter.WriteLine("</Placemark>")
                                Next
                            Else done = True
                            End If
                            Application.DoEvents()
                            offset += maxRecordCount      ' next block
                            POSTfields("resultOffset") = CStr(offset)
                        Catch ex As Exception
                            done = True
                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOK, "Web request failed")
                        End Try
                    End While
                Catch ex As Exception
                    MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                End Try
                TextBox1.Text = $"Done. {count } stations extracted"
                POSTfields = Nothing
                kmlWriter.WriteLine(KMLfooter)
                kmlWriter.Close()
                ' compress to zip file
                System.IO.File.Delete(BASEFILENAME & ".kmz")
                Dim zip As ZipArchive = ZipFile.Open(BASEFILENAME & ".kmz", ZipArchiveMode.Create)    ' create new archive file
                zip.CreateEntryFromFile(BASEFILENAME & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
                ' add icon images
                For Each url In KMLurl
                    zip.CreateEntryFromFile($"{url.Value }.png", $"{url.Value }.png", CompressionLevel.Optimal)
                Next
                zip.Dispose()
            End Using
        End Using
    End Sub

    Private Sub RailwayTracksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RailwayTracksToolStripMenuItem.Click
        Const MAPSERVER = "https://services.ga.gov.au/gis/rest/services/NM_Transport_Infrastructure/MapServer/7/"
        Const BASEFILENAME = "Railway"
        Dim POSTfields As NameValueCollection, count As Integer = 0
        Dim resp As Byte() = {}, responseStr As String
        Dim maxRecordCount As Integer   ' max number of records retrievable with 1 query
        Dim Jo As JObject
        Dim ExtendedData As New Dictionary(Of String, String)
        Dim KMLColor As New Dictionary(Of String, String)     ' color by status
        Dim KMLWidth As New Dictionary(Of String, Integer)     ' width by status

        ' Get the capabilities of this map
        POSTfields = New NameValueCollection From {{"f", "json"}}
        Using myWebClient As New WebClient
            Try
                resp = myWebClient.UploadValues(MAPSERVER, "POST", POSTfields)    ' query map server
            Catch ex As WebException
                MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
            End Try
            responseStr = System.Text.Encoding.UTF8.GetString(resp)
            Jo = JObject.Parse(responseStr)
            maxRecordCount = Jo("maxRecordCount")
            ' Extract drawing info so it can be converted to KML style
            For Each uniqueValue In Jo("drawingInfo")("renderer")("uniqueValueInfos")
                KMLColor.Add(uniqueValue("value"), $"{uniqueValue("symbol")("color")(3):x2}{uniqueValue("symbol")("color")(2):x2}{uniqueValue("symbol")("color")(1):x2}{uniqueValue("symbol")("color")(0):x2}")    ' KML hex color
                KMLWidth.Add(uniqueValue("value"), uniqueValue("symbol")("width"))                                                                  ' track width
            Next

            ' Prepare some POST fields for a request to the map server. We use POST because the requests are too large for a GET
            POSTfields = New NameValueCollection From {
                {"f", "geojson"},
                {"where", "featuretype='Railway'"},
                {"GeometryType", "esriGeometryEnvelope"},
                {"spatialRel", "esriSpatialRelIntersects"},
                {"outFields", "routename,status,sectionname,gauge,tracks,textnote"},
                {"outSR", "WGS84"},
                {"returnGeometry", "true"},          '  need the geometry
                {"orderByFields", "routename"},
                {"returnTrueCurves", "true"},
                {"returnIdsOnly", "false"},
                {"returnCountOnly", "false"},
                {"resultRecordCount", maxRecordCount}, ' maximum number of records we can receive
                {"returnZ", "false"},
                {"returnM", "false"},
                {"returnDistinctValues", "false"},
                {"returnExtentOnly", "false"}
            }

            ' Create kml header
            Using kmlWriter As New System.IO.StreamWriter($"{BASEFILENAME }.kml")
                kmlWriter.WriteLine(KMLheader)
                kmlWriter.WriteLine("<description><![CDATA[<style>table, th, td {white-space:nowrap; }</style>")
                kmlWriter.WriteLine("<table>")
                kmlWriter.WriteLine("<tr><td>Data produced by Marc Hillman - VK3OHM/VK3IP</td></tr>")
                Dim utc As String = $"{DateTime.UtcNow:dd MMM yyyy hh:mm:ss} UTC"     ' time now in UTC
                kmlWriter.WriteLine($"<tr><td>Data extracted from {MAPSERVER } on {utc }.</td></tr>")
                kmlWriter.WriteLine("</table>")
                kmlWriter.WriteLine("]]>")
                kmlWriter.WriteLine("</description>")
                ' Create some line styles
                For Each item In KMLColor
                    Dim key As String = item.Key
                    kmlWriter.WriteLine($"<Style id=""{key }""><LineStyle><color>{KMLColor(key) }</color><width>{KMLWidth(key) }</width></LineStyle></Style>")
                Next
                Try
                    Dim offset As Integer = 0
                    POSTfields.Add("resultOffset", CStr(offset))
                    Dim done = False
                    While Not done
                        Dim off As String = POSTfields("resultOffset")
                        resp = Array.Empty(Of Byte)()   ' clear  array
                        Try
                            resp = myWebClient.UploadValues($"{MAPSERVER }query/", "POST", POSTfields)    ' query map server
                            TextBox1.Text = "retrieved " & resp.Length & " bytes of data at offset " & offset
                            responseStr = System.Text.Encoding.UTF8.GetString(resp)
                            Jo = JObject.Parse(responseStr)
                            If Jo.HasValues And Jo("features").Any Then
                                For Each feature In Jo("features")
                                    count += 1   ' count number of segments
                                    ' convert each rail segment to a line string
                                    ExtendedData.Clear()
                                    ExtendedData.Add("Route Name", feature("properties")("routename").ToString)
                                    If Not IsNothing(feature("properties")("sectionname")) Then ExtendedData.Add("Section Name", feature("properties")("sectionname").ToString)
                                    ExtendedData.Add("Status", feature("properties")("status").ToString)
                                    If Not IsNothing(feature("properties")("gauge")) Then ExtendedData.Add("Gauge", feature("properties")("gauge").ToString)
                                    If Not IsNothing(feature("properties")("tracks")) Then ExtendedData.Add("Tracks", feature("properties")("tracks").ToString)
                                    If Not IsNothing(feature("properties")("textnote")) Then ExtendedData.Add("Note", feature("properties")("textnote").ToString)

                                    kmlWriter.WriteLine($"<Placemark id='{ExtendedData("Route Name") }'>")
                                    kmlWriter.WriteLine($"<styleUrl>#{ExtendedData("Status") }</styleUrl>")
                                    kmlWriter.WriteLine($"<name>{ExtendedData("Route Name") }</name>")

                                    ' write out ExtendedData
                                    kmlWriter.WriteLine("<ExtendedData>")
                                    For Each item In ExtendedData
                                        kmlWriter.WriteLine($"<Data name=""{item.Key }""><value>{SecurityElement.Escape(item.Value) }</value></Data>")
                                    Next
                                    kmlWriter.WriteLine("</ExtendedData>")

                                    ' Extract all the points in the linestring
                                    Dim coords As New List(Of String)
                                    coords.Clear()
                                    For Each pnt In feature("geometry")("coordinates")
                                        coords.Add($"{pnt(0):f5},{pnt(1):f5}")
                                    Next
                                    kmlWriter.WriteLine("<LineString>")
                                    kmlWriter.WriteLine("<altitudeMode>clampToGround</altitudeMode>")
                                    Dim st As String = Join(coords.ToArray, " ")
                                    kmlWriter.WriteLine($"<coordinates>{st }</coordinates>")
                                    kmlWriter.WriteLine("</LineString>")
                                    kmlWriter.WriteLine("</Placemark>")
                                Next
                            Else done = True
                            End If
                            Application.DoEvents()
                            offset += maxRecordCount      ' next block
                            POSTfields("resultOffset") = CStr(offset)
                        Catch ex As WebException
                            done = True
                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOK, "Web request failed")
                        End Try
                    End While
                Catch ex As WebException
                    MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                End Try
                TextBox1.Text = $"Done. {count } segments extracted"
                POSTfields = Nothing
                kmlWriter.WriteLine(KMLfooter)
                kmlWriter.Close()
                ' compress to zip file
                System.IO.File.Delete(BASEFILENAME & ".kmz")
                Dim zip As ZipArchive = ZipFile.Open(BASEFILENAME & ".kmz", ZipArchiveMode.Create)    ' create new archive file
                zip.CreateEntryFromFile(BASEFILENAME & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
                zip.Dispose()
            End Using
        End Using
    End Sub

    Private Sub TestToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TestToolStripMenuItem1.Click
        Dim fc As New FeatureCollection
        Dim json As String = "{
  ""type"": ""FeatureCollection"",
  ""extra_fc_member"": ""foo"",
  ""features"":
  [
    {
      ""type"": ""Feature"",
      ""extra_feat_member"": ""bar"",
      ""geometry"": {
        ""type"": ""Point"",
        ""extra_geom_member"": ""baz"",
        ""coordinates"": [ 2, 49, 3, 100, 101 ]
      },
      ""properties"": {
        ""a_property"": ""foo""
      }
    }
  ]
}"
        fc = FeatureCollection.FromJson(json)

    End Sub

    Private Sub DownloadSiloscsvToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DownloadSiloscsvToolStripMenuItem.Click
        ' Download silos.csv file
        Dim url As String, millis As Long = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(), count As Integer = 0
        Dim header As String, heading As String, values As String, SILOS_sql As SQLiteCommand, currentRow() As String, valueList As New List(Of String)
        Dim ratings As New Dictionary(Of String, Integer)
        Dim columns As New Dictionary(Of String, Integer), column As Integer, active As Integer = 0, deleted As Integer = 0

        ' download the file
        Using webClient As New WebClient
            Try
                url = $"https://www.silosontheair.com/data/silos.csv?a={millis }"   ' random a= parameter to defeat caching
                webClient.DownloadFile(url, "silos.csv")
                SetText(TextBox1, "silos.csv downloaded")
            Catch ex As WebException
                MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
            End Try
        End Using

        Using connect As New SQLiteConnection(SILOSdb), csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("silos.csv")
            csvReader.TextFieldType = FileIO.FieldType.Delimited
            csvReader.SetDelimiters(",")
            connect.Open()  ' open database
            SILOS_sql = connect.CreateCommand
            SILOS_sql.CommandText = "BEGIN TRANSACTION"   ' start transaction
            SILOS_sql.ExecuteNonQuery()
            SILOS_sql.CommandText = "DELETE FROM SILOS"   ' delete existing data
            SILOS_sql.ExecuteNonQuery()
            While Not csvReader.EndOfData
                Try
                    count += 1
                    If count = 1 Then
                        header = csvReader.ReadLine.ToLower
                        valueList.Clear()
                        column = 0
                        For Each heading In header.Split(",").ToList
                            columns.Add(heading, column)        ' list of column offsets
                            valueList.Add($"@{heading }")      '  add @ symbol
                            column += 1
                        Next
                        values = String.Join(",", valueList.ToArray)
                        SILOS_sql.CommandText = $"INSERT INTO silos ({header }) VALUES ({values })"
                        Try
                            SILOS_sql.Prepare()
                        Catch ex As SQLiteException
                            MsgBox($"Line {count } " & ex.Message & vbCrLf & ex.StackTrace & "Prepare error.")
                        End Try
                    Else
                        currentRow = csvReader.ReadFields       ' read the fields
                        SILOS_sql.Parameters.Clear()            ' add parameter list
                        For i = 0 To currentRow.Length - 1
                            SILOS_sql.Parameters.AddWithValue(valueList(i), currentRow(i))
                        Next
                        If String.IsNullOrEmpty(currentRow(columns("not_after"))) Then
                            active += 1
                            ' total ratings field
                            Dim rating As String = currentRow(columns("rating"))
                            If ratings.ContainsKey(rating) Then ratings(rating) += 1 Else ratings.Add(rating, 1)    ' accumulate ratings
                        Else
                            deleted += 1
                        End If
                        Try
                            SILOS_sql.ExecuteNonQuery()     ' do the insert
                        Catch ex As SQLiteException
                            MsgBox($"Line {count } " & ex.Message & vbCrLf & ex.StackTrace & "INSERT error.")
                        End Try
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox($"Line {count } " & ex.Message & vbCrLf & ex.StackTrace & "is not valid and will be skipped.")
                End Try
            End While
            SILOS_sql.CommandText = "COMMIT"   ' start transaction
            SILOS_sql.ExecuteNonQuery()
        End Using
        TextBox1.Text = $"{count } lines read"
        AppendText(TextBox2, $"Active: {active } Deleted: {deleted }{vbCrLf }")
        AppendText(TextBox2, $"Ratings count{vbCrLf }")
        For Each entry In ratings
            AppendText(TextBox2, $"{ entry.Key } {entry.Value }{vbCrLf }")
        Next
    End Sub
    Private Sub GetsilosFromOpenStreetMapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetsilosFromOpenStreetMapToolStripMenuItem.Click
        ' Get all item in OpenStreetMap tagged as "silo"
        Const SERVER = "https://lz4.overpass-api.de/api/interpreter"
        Const MAPSERVER = "https://services.ga.gov.au/gis/rest/services/NM_Transport_Infrastructure/MapServer/7/"    ' railway tracks
        Const DISTANCE = "1000"  ' max distance of silo from railway
        Dim POSTfields As New NameValueCollection, resp As Byte() = {}, responseStr As String, ql As String, count As Integer = 0, processed As Integer = 0, Jo As JObject, style As String
        Dim Yes As Integer = 0, No As Integer = 0, Maybe As Integer = 0, GIS As Integer = 0
        Dim TagsofInterest As New List(Of String) From {"name", "material", "operator", "product", "content", "crop",
            "building:colour", "building:height", "building:levels", "building:material",
            "roof:colour", "roof:height", "building:shape", "roof:material",
            "damaged", "height", "storage"}     ' tags in OSM data to extract
        Dim ExtendedData As New Dictionary(Of String, String), lat As Double, lon As Double, name As String, id As String

        ql = "[bbox:-43.6345972634,113.338953078,-10.6681857235,153.569469029];(node [""man_made""=""silo""]; way [""man_made""=""silo""];); out body geom; >;"    ' retrieve silos with bounding box of Australia
        POSTfields.Clear()
        POSTfields.Add("data", ql)
        Using myWebClient As New WebClient
            Try
                resp = Array.Empty(Of Byte)()  ' clear  array
                resp = myWebClient.UploadValues(SERVER, "POST", POSTfields)    ' query map server
            Catch ex As WebException
                MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
            End Try
            responseStr = System.Text.Encoding.UTF8.GetString(resp)

            Dim doc As New XmlDocument
            Using kmlWriter As New System.IO.StreamWriter("silos.kml"), csvWriter As New System.IO.StreamWriter("silos.csv")
                ' Create kml header
                kmlWriter.WriteLine(KMLheader)
                kmlWriter.WriteLine("<Style id=""Yes""><IconStyle><Icon><href>https://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href></Icon></IconStyle></Style>")
                kmlWriter.WriteLine("<Style id=""Maybe""><IconStyle><Icon><href>https://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon></IconStyle></Style>")
                kmlWriter.WriteLine("<Style id=""No""><IconStyle><Icon><href>https://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href></Icon></IconStyle></Style>")
                ' Create csv header
                csvWriter.Write("id,lon,lat,style,")
                csvWriter.WriteLine(Join(TagsofInterest.ToArray, ","))
                doc.LoadXml(responseStr)    ' read the XML

                Dim nodelist As XmlNodeList = doc.SelectNodes("//node | //way")    ' select all data nodes
                count = nodelist.Count
                TextBox1.Text = $"Retrieved {count } silos"
                Application.DoEvents()
                For Each node As XmlElement In nodelist
                    id = node.GetAttribute("id")    ' unique OpenStreetMap id
                    ExtendedData.Clear()
                    ' find the location
                    Select Case node.Name
                        Case "node"
                            ExtendedData.Add("Geometry", "Point")
                            lat = CDbl(node.GetAttribute("lat"))
                            lon = CDbl(node.GetAttribute("lon"))
                        Case "way"
                            ExtendedData.Add("Geometry", "Polygon")
                            ' point is centroid of polygon
                            Dim bounds As XmlElement = node.SelectSingleNode("bounds")
                            Dim p1 As New MapPoint(CDbl(bounds.GetAttribute("minlon")), CDbl(bounds.GetAttribute("minlat")), SpatialReferences.Wgs84)
                            Dim p2 As New MapPoint(CDbl(bounds.GetAttribute("maxlon")), CDbl(bounds.GetAttribute("maxlat")), SpatialReferences.Wgs84)
                            Dim extent As New Envelope(p1, p2)
                            lon = extent.GetCenter().X
                            lat = extent.GetCenter().Y
                    End Select
                    ' find interesting tags and save in ExtendedData
                    Dim tags As XmlNodeList = node.SelectNodes("tag")     ' select all tag
                    For Each tag As XmlElement In tags
                        Dim k As String = tag.GetAttribute("k")
                        If TagsofInterest.Contains(k) Then  ' save interesting tag
                            Dim v As String = tag.GetAttribute("v")
                            ExtendedData.Add(k, v)
                        End If
                    Next
                    ' try to find a name
                    If ExtendedData.ContainsKey("name") Then
                        name = ExtendedData("name")
                    Else
                        name = ""
                    End If
                    kmlWriter.WriteLine($"<Placemark id=""{id }"">")
                    ' write out ExtendedData
                    kmlWriter.WriteLine("<ExtendedData>")
                    For Each item In ExtendedData
                        kmlWriter.WriteLine($"<Data name=""{item.Key }""><value>{SecurityElement.Escape(item.Value) }</value></Data>")
                    Next
                    kmlWriter.WriteLine("</ExtendedData>")

                    style = "Maybe"     ' default style
                    If ExtendedData.ContainsKey("name") Then
                        With ExtendedData("name")
                            If .ToLower.Contains("wheat") _
                               Or .ToLower.Contains("grain") _
                               Or .ToLower.Contains("silo") Then style = "Yes"
                            If .ToLower.Contains("cfa") _
                                    Or .ToLower.Contains("water") _
                                    Or .ToLower.Contains("tank") _
                                    Or .ToLower.Contains("coal") Then style = "No"
                        End With
                    End If
                    If ExtendedData.ContainsKey("operator") Then
                        With ExtendedData("operator")
                            If .ToLower.Contains("ausbulk") _
                               Or .ToLower.Contains("grain") _
                               Or .ToLower.Contains("vittera") Then style = "Yes"
                            If .ToLower.Contains("penfold") _
                               Or .ToLower.Contains("winery") _
                               Or .ToLower.Contains("plantagenet") _
                               Or .ToLower.Contains("ridley") _
                               Or .ToLower.Contains("mentelle") _
                               Or .ToLower.Contains("wolf blass") Then style = "No"
                        End With
                    End If
                    If ExtendedData.ContainsKey("product") Then
                        With ExtendedData("product")
                            If .ToLower.Contains("wheat") _
                               Or .ToLower.Contains("grain") Then style = "Yes"
                            If .ToLower.Contains("water") _
                               Or .ToLower.Contains("wine") Then style = "No"
                        End With
                    End If
                    If ExtendedData.ContainsKey("content") Then
                        With ExtendedData("content")
                            If .ToLower.Contains("wheat") Or .ToLower.Contains("grain") Then style = "Yes"
                            If .ToLower.Contains("water") Or .ToLower.Contains("concrete") Or .ToLower.Contains("coal") Or .ToLower.Contains("silage") Then style = "No"
                        End With
                    End If
                    If ExtendedData.ContainsKey("crop") Then
                        With ExtendedData("crop")
                            If .ToLower.Contains("cement") Then style = "No" Else style = "Yes"
                        End With
                    End If
                    If ExtendedData.ContainsKey("storage") Then
                        With ExtendedData("storage")
                            style = "Yes"
                        End With
                    End If
                    If ExtendedData.ContainsKey("material") Then
                        With ExtendedData("material")
                            If .ToLower.Contains("water") Then style = "No"
                        End With
                    End If
                    If style = "Maybe" Then
                        ' See if placemarker near railway track
                        GIS += 1
                        With POSTfields
                            .Clear()
                            .Add("f", "geojson")
                            .Add("where", "featuretype='Railway'")
                            .Add("GeometryType", "esriGeometryPoint")
                            .Add("Geometry", $"{lon:f6},{lat:f6}")
                            .Add("distance", DISTANCE)
                            .Add("units", "esriSRUnit_Meter")
                            .Add("spatialRel", "esriSpatialRelIntersects")
                            .Add("inSR", "WGS84")
                            .Add("returnGeometry", "false")          '  don't need the geometry
                            .Add("returnCountOnly", "true")
                        End With
                        Try
                            resp = Array.Empty(Of Byte)()  ' clear  array
                            resp = myWebClient.UploadValues($"{MAPSERVER }query/", "POST", POSTfields)    ' query map server
                            responseStr = System.Text.Encoding.UTF8.GetString(resp)
                            Jo = JObject.Parse(responseStr)
                            If Jo("count") <> "0" Then style = "Yes"
                        Catch ex As Exception
                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                        End Try
                    End If
                    Select Case style
                        Case "No" : No += 1
                        Case "Yes" : Yes += 1
                        Case "Maybe" : Maybe += 1
                    End Select

                    kmlWriter.WriteLine($"<styleUrl>#{style }</styleUrl>")
                    If name <> "" Then kmlWriter.WriteLine($"<name>{SecurityElement.Escape(name) }</name>")
                    kmlWriter.WriteLine("<Point>")
                    kmlWriter.WriteLine($"<coordinates>{lon:f6},{lat:f6}</coordinates>")
                    kmlWriter.WriteLine("</Point>")
                    kmlWriter.WriteLine("</Placemark>")
                    processed += 1
                    ' write csv data
                    csvWriter.Write($"{id },{lon:f6},{lat:f6},{style }")
                    For Each item In TagsofInterest
                        csvWriter.Write(",")
                        If ExtendedData.ContainsKey(item) Then csvWriter.Write(ExtendedData(item))
                    Next
                    csvWriter.WriteLine()
                    TextBox1.Text = $"Processed {processed } of {count }"
                    Application.DoEvents()
                Next
                kmlWriter.WriteLine(KMLfooter)
                kmlWriter.Close()
                csvWriter.Close()
                TextBox1.Text = $"Extracted {count } silos.  Yes={Yes },  Maybe={Maybe },  No={No },  GIS={GIS }"
            End Using
        End Using
    End Sub

    Private Sub ConsolidateSilosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsolidateSilosToolStripMenuItem.Click
        ' Each individual silo in a cluster of silos has a separate pin.
        ' Find groups of silos that are close and consolidate them.
        ' Use most recently produced silos.csv for data
        Const CLOSENESS = 500     ' distance below which two silos are considered one
        Dim silos As New Dictionary(Of String, SiloData)
        Dim currentRow As String(), count As Integer = 0, header As String() = {}, lon As Double, lat As Double

        Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("silos.csv")
            csvReader.TextFieldType = FileIO.FieldType.Delimited
            csvReader.SetDelimiters(",")
            While Not csvReader.EndOfData
                Try
                    count += 1
                    currentRow = csvReader.ReadFields()
                    If count = 1 Then
                        header = currentRow     ' save for later
                    Else
                        silos.Add(currentRow(0), New SiloData(currentRow.ToArray))
                        Dim currentField As String
                        For Each currentField In currentRow
                            ' MsgBox(currentField)
                        Next
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
            End While
        End Using
        TextBox1.Text = $"{count } lines read"
        ' Now calculate distances to all other silos
        Dim total = silos.Count   ' total number of silos
        count = 0
        For Each siloOuter In silos
            If Not siloOuter.Value.Consolidated Then
                ' Calculate distances for all silos
                For Each siloInner In silos
                    If Not siloInner.Value.Consolidated Then siloInner.Value.GeoDistance(siloOuter.Value)
                Next
                ' find all silos within 50 meters of this one
                Dim close As New List(Of String)
                close.Clear()
                For Each silo In silos
                    If silo.Key <> siloOuter.Key And Not silo.Value.Consolidated And silo.Value.Distance <= CLOSENESS Then close.Add(silo.Key)
                Next
                If close.Any Then
                    ' Consolidate silos.
                    ' Calculate centroid
                    lon = CDbl(siloOuter.Value.Data(1))
                    lat = CDbl(siloOuter.Value.Data(2))
                    For Each silo In close
                        silos(silo).Consolidated = True
                        lon += CDbl(silos(silo).Data(1))
                        lat += CDbl(silos(silo).Data(2))
                        ' Consolidate the metadata
                        Select Case siloOuter.Value.Data(3) ' style
                            Case "Yes" ' do nothing
                            Case "Maybe" : If (silos(silo).Data(3) = "Yes") Then siloOuter.Value.Data(3) = "Yes"
                            Case "No" : siloOuter.Value.Data(3) = silos(silo).Data(3)
                        End Select
                        For i = 4 To 20
                            If String.IsNullOrEmpty(siloOuter.Value.Data(i)) Then siloOuter.Value.Data(i) = silos(silo).Data(i)
                        Next
                    Next
                    ' Now calculate centroid
                    lon /= (close.Count + 1)
                    lat /= (close.Count + 1)
                    ' Consolidate all points into first in group
                    siloOuter.Value.Data(1) = $"{lon:f6}"
                    siloOuter.Value.Data(2) = $"{lat:f6}"
                End If
            End If
            count += 1
            TextBox1.Text = $"{count }/{total } processed"
            Application.DoEvents()
        Next
        ' Write out the consolidated data
        Using csvWriter As New System.IO.StreamWriter("silosConsolidated.csv")
            count = 0
            csvWriter.WriteLine(String.Join(",", header))
            For Each silo In silos
                If Not silo.Value.Consolidated Then
                    count += 1
                    Dim csv As String = String.Join(",", silo.Value.Data)
                    csvWriter.WriteLine(csv)
                End If
            Next
            csvWriter.Close()
        End Using
        TextBox1.Text = $"{total } silos consolidated to {count }"
    End Sub

    Private Sub FindIsolatedSilosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindIsolatedSilosToolStripMenuItem.Click
        Const CLOSENESS = 60     ' distance above which silo is considered isolated (km)
        Dim silos As New Dictionary(Of String, SiloDataIsol), silo As SiloDataIsol, remote As New List(Of SiloDataIsol)
        Dim count As Integer = 0
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader

        Using connect As New SQLiteConnection(SILOSdb)
            connect.Open()  ' open database
            sqlcmd = connect.CreateCommand()
            sqlcmd.CommandText = "SELECT * FROM silos"     ' select all silos
            SQLdr = sqlcmd.ExecuteReader()
            While SQLdr.Read()
                silo = New SiloDataIsol(SQLdr("locality"), CDbl(SQLdr("lng")), CDbl(SQLdr("lat")), SQLdr("state"))
                silos.Add(SQLdr("id"), silo)
            End While

            TextBox1.Text = $"{count } lines read"
            ' Now calculate distances to all other silos
            Dim total = silos.Count   ' total number of silos
            count = 0
            remote.Clear()

            For Each siloOuter In silos
                ' Calculate distances for all silos
                For Each siloInner In silos
                    siloInner.Value.GeoDistance(siloOuter.Value)    ' calculate distance to all other silos
                Next
                ' find closest silo
                siloOuter.Value.Distance = 5000     ' nearest neighbour
                For Each s In silos
                    If s.Key <> siloOuter.Key And s.Value.Distance < siloOuter.Value.Distance Then
                        siloOuter.Value.Distance = s.Value.Distance     ' save smallest distance
                    End If
                Next
                If siloOuter.Value.Distance > CLOSENESS Then
                    ' silo is isolated
                    remote.Add(siloOuter.Value.Clone)
                End If
                count += 1
                TextBox1.Text = $"{count }/{total } processed"
                Application.DoEvents()
            Next
        End Using
        ' print sorted list
        remote.Sort(Function(x, y) y.Distance.CompareTo(x.Distance))    ' sort by distance
        Using htmlWriter As New System.IO.StreamWriter("silosIsolated.html")
            htmlWriter.WriteLine("<table border=1>")
            htmlWriter.WriteLine("<tr><th>Locality</th><th>lon</th><th>lat</th><th>state</th><th>distance (km)</th></tr>")
            For Each r In remote
                htmlWriter.WriteLine($"<tr><td>{r.Locality}</td><td>{r.Lon:f5}</td><td>{r.Lat:f5}</td><td>{r.State}</td><td>{r.Distance:f0}</td></tr>")
            Next
            htmlWriter.WriteLine("</table>")
        End Using
        TextBox1.Text = $"{remote.Count } remote silos found"
    End Sub

    Private Sub MakeTestADIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeTestADIToolStripMenuItem.Click
        ' Make a test ADI file
        Const ACTIVATIONS = 100   ' number of activations
        Dim dt As DateTime = #2/3/2020 05:01#
        ' List of all valid silo codes
        Dim silocodes As New List(Of String) From {"ABN2", "ALA5", "ALR2", "ALT2", "ALI2", "ALA2", "ANO3", "ANP3", "APA5", "ARS3", "ARN2", "KMK2", "ARK2", "ARE2", "ARL2", "ATA2", "AVA3", "BNA2", "BCH3", "BLA5", "BLG6", "BLL3", "BLE2", "BLU6", "BLR2", "BNN3", "BRE2", "BRN2", "GRN2", "YDH2", "BRA2", "WRN2", "BRT4", "BRB2", "BRE3", "BRT3", "BTT2", "PRF7", "BCN6", "BLA3", "BCM2", "MRL2", "BLN2", "BLL4", "BLA2", "BNL2", "BRG2", "BRK3", "BRN3", "BLH3", "GLL3", "BLI2", "BLA4", "BNI6", "BNY2", "BNI2", "BNW2", "BNN2", "BRP3", "BRW2", "BDN6", "BGE2", "BGA2", "BGI2", "BGH2", "BLN3", "BNN4", "BRT2", "BRR3", "BRK2", "BRG3", "BWS5", "CMN2", "BRZ2", "BRR2", "BRA3", "BRW4", "BRM3", "TLR2", "BRY2", "BRD4", "BRN6", "BRE6", "SHS2", "BRC3", "BCO5", "BCE3", "BNY6", "BNA4", "BNL6", "BNO2", "BRH2", "NRM6", "BRA5", "BRI2", "BRO2", "BRC2", "BTE5", "CLL2", "CLI6", "CMA4", "CMI5", "CNE3", "CNY2", "CNA2", "CPA4", "CRL2", "CRA3", "CRH6", "CRN2", "CRP3", "CCS4", "YLA3", "CHN3", "CHH3", "CHA4", "CHK3", "CLT4", "CLN4", "CBM3", "CCA3", "CLN3", "CLD3", "CMG2", "CNN2", "CNG2", "CLH2", "CLN2", "CNN5", "CNE2", "CTA2", "CPE3", "SWR3", "CRA2", "CRW2", "CSE3", "CWE3", "CWL5", "GLN2", "NNA2", "HLD2", "WGL6", "CRK6", "NYX3", "CRE2", "CRK2", "CRO2", "CRK5", "CLR2", "CLA3", "CMS5", "CMK2", "CNA5", "CNR2", "CRB2", "CRS2", "CRO3", "DLY4", "GCT4", "NRE4", "NEA4", "MCE6", "DNH3", "DRK5", "DLA2", "DRG2", "DVH3", "DVT7", "DGA3", "DMA3", "DNN3", "DMH3", "RCG3", "DNO4", "NNO4", "ARA3", "DND3", "DNE3", "DKE3", "DRA4", "DBO2", "KPN4", "DLA4", "DLL2", "DMS3", "DNO2", "DNY3", "DRI2", "KWA6", "ECA3", "EDI2", "EDE5", "ELE3", "EMD4", "EML2", "ERA2", "EBT2", "MDK2", "EGA2", "EME2", "FRT5", "FNY2", "RDD2", "FRS2", "FRL4", "FRS5", "GBN6", "ARO2", "GNN2", "GRA2", "GLG3", "GRM5", "GRE2", "GDY2", "GNE4", "GRL2", "GLE5", "GLH3", "GLY3", "GLO2", "GLL2", "BRD2", "SGI2", "GLI2", "STA2", "GLG2", "GNA2", "GRT3", "GRG3", "GVN4", "GRE3", "GWD3", "GRH6", "GRD2", "GRS2", "GRN3", "GRP2", "GRF2", "BRU2", "WRA2", "GRG2", "GLE2", "GLU2", "WSI2", "GNH2", "NWH2", "GND2", "GRY2", "PNP2", "GYN3", "HME5", "HRN2", "HRD2", "HRN4", "HYY2", "HNY2", "HLN2", "HDN4", "HLK2", "HPD2", "HPN3", "HRM3", "HNR3", "HYN6", "IDY4", "ILO2", "INL2", "PPL7", "JMN5", "JPT3", "GYA3", "ELM3", "JRE2", "BGN4", "JMT4", "JNN4", "JNE2", "JNG3", "KRI4", "KLE6", "KMH2", "KNA3", "KPE5", "KPA5", "KRA5", "KTA3", "KTH5", "KLN6", "KRG3", "KCO2", "SLY3", "KKA2", "KMA5", "KNH4", "DTA4", "KNY4", "KNE2", "KNL3", "KLG3", "KLI6", "KLN3", "KNL2", "KRN3", "KYA5", "LDH2", "LHH3", "LKA3", "LKO2", "LKE6", "LLT3", "LMO5", "LSS3", "LTM6", "LCT3", "LLR3", "LNA3", "LTD3", "LLY3", "LCK5", "LCT2", "GLA4", "LNS5", "LRN3", "LXN5", "LBK3", "MCR4", "MLA5", "MLU4", "MNG3", "MNH2", "MNA2", "MNL2", "MRR2", "MTA2", "MTG2", "MNA4", "MTN3", "MCG6", "MNN2", "YRA3", "MRR3", "GMA3", "WSN6", "MRE3", "PRA2", "KRA3", "MRA2", "MRG2", "MRO2", "MLG2", "MLS4", "MLG6", "MLN4", "MLE2", "MNW6", "MNA5", "EJG6", "MNO2", "NLN3", "MNP3", "MRI3", "MRM3", "MTO3", "MTE3", "MTK3", "CNT2", "MLN2", "MNH5", "MLT3", "MML2", "MRA6", "MRE2", "MRS3", "MRH2", "MLI2", "MNE2", "MNN4", "MRA4", "MKN6", "MLA6", "SLE2", "MNB2", "MRT3", "MRA5", "MRI2", "MRE5", "MRL3", "CRY3", "MRA3", "MYK3", "NGN3", "NNY3", "NNE4", "NRE5", "NRN2", "NRI2", "NRT2", "NRA2", "MCE2", "NRS2", "WYA2", "NRE2", "NTK3", "NTA3", "NTY3", "NVE2", "NWA2", "NWE6", "NWN2", "NHL3", "DPR3", "NNA3", "NRG3", "NRR2", "NRD4", "YVL2", "NRN4", "NLL3", "NMH3", "NNN6", "NNA5", "NYG6", "NYT3", "NYN3", "GRW2", "TRE2", "NYN2", "NYK2", "OKY4", "OKE3", "OLE2", "OTA2", "MRD7", "ORO5", "WNG3", "GLA3", "NNG3", "KML3", "OKS3", "OYN3", "OWN5", "PNA3", "PRA5", "PRG5", "PRS2", "EGD2", "PTK3", "PKE5", "MCI2", "PKL2", "PNG5", "PRE3", "PNL3", "PWG6", "PCA3", "PRN3", "PMO3", "PNY6", "PNP6", "PNA4", "PNO5", "PRA3", "PTA6", "PLS2", "PCA5", "PRY6", "PRD2", "PRN5", "PRE4", "PRE5", "PRI3", "PRR2", "PRE2", "PCN2", "PYD3", "QMK3", "QNY2", "BRX2", "QNA2", "DHG2", "QRI2", "QRN5", "RNW3", "PLT3", "ALA3", "RND2", "RNS2", "RVE6", "RYD3", "RZK2", "RDL5", "RFN2", "RNE2", "RBN5", "RBE3", "RCR3", "RMA4", "RSY3", "RWA2", "RZE2", "RDA5", "RPP3", "RTN3", "SDH5", "STD3", "SNO2", "SNR2", "SLE3", "SRN3", "SHS3", "SHE3", "SKN3", "SNN5", "STS6", "SPD3", "SPT3", "SPE2", "SPE4", "STL3", "STS3", "STL2", "LWE2", "MRN2", "STH2", "STY5", "SML2", "SNE3", "STN3", "SWL3", "TBA2", "WRN5", "TLA5", "TLN2", "TLA3", "TMA2", "TNA3", "TRA4", "TRN6", "TRE5", "TRK3", "TTN3", "TDY3", "TLD3", "TMR2", "TMY3", "THN4", "THG2", "THS4", "THN2", "THD5", "THH2", "THO4", "THS6", "TCE2", "TNN6", "TNA5", "TCL2", "TCH2", "TMT2", "TBH4", "TLE5", "TTL2", "TRA3", "TTM2", "TRI2", "BRL2", "MNR2", "BDR2", "TRG6", "KDE2", "TRL2", "GBY2", "TLE2", "TLL2", "THS2", "BRV2", "TMY5", "TNH3", "TRF3", "TTE3", "ULI2", "ULA3", "UNL3", "UNE2", "UNA5", "CLE2", "URA2", "URT2", "URY2", "VCS3", "WLR3", "WDE5", "WGN6", "WKE5", "WLL3", "WTE3", "WKL2", "WLO5", "WLA2", "WLN2", "WLA4", "WLR2", "WLP3", "WNI5", "WNN4", "MRM2", "WRD2", "WRE3", "WRA4", "WRL3", "WBL3", "BTA3", "WRL2", "WRO5", "WRK4", "WTM3", "WTA3", "WTA2", "WDN3", "WDN2", "WTE2", "WWA2", "WJA2", "WLO2", "WRU3", "WRK2", "WSH6", "WSE2", "WSD6", "WSE3", "WSR3", "NRG2", "CLA2", "WSG2", "WHA5", "WHN2", "WCN6", "WDI2", "WLA3", "WLE2", "WLS6", "WLT2", "WNA3", "WRA5", "WRY2", "WRL5", "WLY5", "WMA2", "WDK2", "WMG3", "WRN3", "WRH3", "WDA5", "WNU3", "WYM6", "WYG2", "LNE3", "WYF3", "YPT3", "YNC3", "YNO2", "YNE5", "YRI2", "YRA4", "YRG3", "YLG6", "YLA5", "YLN4", "YLT3", "YNA2", "YRN6", "YRK2", "YNA5", "YTG6", "YNG2", "YNN2"}

        Using adiWriter As New System.IO.StreamWriter("SiOTA_test.adi")
            For silo As Integer = 0 To ACTIVATIONS - 1
                QSO(adiWriter, dt.ToString("yyyyMMdd"), dt.ToString("HHmm"), "20m", "SSB", silocodes(index:=silo))
                dt = dt.AddMinutes(1)
                QSO(adiWriter, dt.ToString("yyyyMMdd"), dt.ToString("HHmm"), "40m", "CW", silocodes(index:=silo))
                dt = dt.AddMinutes(1)
                QSO(adiWriter, dt.ToString("yyyyMMdd"), dt.ToString("HHmm"), "2m", "SSB", silocodes(index:=silo))
                dt = dt.AddMinutes(1)
            Next
        End Using
        TextBox1.Text = $"{ACTIVATIONS } activations created"
    End Sub

    Private Shared Sub QSO(writer As StreamWriter, QSO_date As String, time As String, band As String, mode As String, silocode As String)
        ' write a complete QSO
        writer.Write(ADIF("STATION_CALLSIGN", "VK3OHM"))
        writer.Write(ADIF("QSO_DATE", QSO_date))
        writer.Write(ADIF("TIME_ON", time))
        writer.Write(ADIF("BAND", band))
        writer.Write(ADIF("MODE", mode))
        writer.Write(ADIF("MY_SIG", "SIOTA"))
        writer.Write(ADIF("MY_SIG_INFO", silocode))
        writer.WriteLine(ADIF("EOR"))
    End Sub
    Private Shared Function ADIF(field As String, Optional value As String = "")
        ' Construct an ADIF field
        If field = "EOR" Then Return "<EOR>" Else Return $"<{field }:{value.Length }>{value } "
    End Function


    Private Sub MakeSQLTestDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeSQLTestDataToolStripMenuItem.Click
        ' Make some test data for SiOTA. Create SQL data to add to log table
        Const FIRSTTESTUSER = 1000      ' id of start of test data
        Const FIRSTSILO = 699      ' id of first silo
        Dim values As New List(Of String)

        Dim user_id As Integer, i As Integer, siloIndx As Integer, siloIndxSave As Integer
        ' Large list of real callsigns
        Dim callsigns As String() = {"VK2KAL", "VK2KY", "VK4DR", "VK2WEL", "VK2ZLA", "VK2YH", "VK2RRT", "VK2JWA", "VK2YVA", "VK2KH", "VK2ZRJ", "VK2KMY", "VK2BVA", "VK3ZDG", "VK2YGA", "VK2DWV", "VK1ZAT", "VK2XTG", "VK1IA", "VK9JA", "VK2BHH", "VK2ADN", "VK2TC", "VK2AWZ", "VK2ANS", "VK2BRJ", "VK2KPA", "VK2BDF", "VK2YTG", "VK2SEM", "VK1RA", "VK1ARA", "VK2OG", "VK2ZTN", "VK2KIV", "VK4DGU", "VK2NYL", "VK2VK", "VK2EFC", "VK2SG", "VK2BCQ", "VK2IBA", "VK2JMA", "VK2KIB", "VK2WQ", "VK2DKZ", "VK2ZIA", "VK2DBO", "VK2RBO", "VK4HDF", "VK2YUS", "VK2NYA", "VK2BAI", "VK2RJ", "VK2BMB", "VK2GQ", "VK2KHB", "VK2BNP", "VK2KMT", "VK2AIG", "VK2CBL", "VK2NPG", "VK2FWB", "VK2CZX", "VK4DHY", "VK2BBL", "VK2FOE", "VK2BTV", "VK2DVZ", "VK1SB", "VK2HI", "VK2BIF", "VK2YVB", "VK2DAT", "VK2JAP", "VK2BB", "VK2DNT", "VK2ABY", "VK2KYP", "VK3FGE", "VK2DLT", "VK2KMB", "VK2VJR", "VK2CRB", "VK2VKY", "VK2AIB", "VK2ABB", "VK2RT", "VK4GAV", "VK2AXB", "VK2XRB", "VK2XBZ", "VK2QA", "VK2DQ", "VK2BMA", "VK2BM", "VK2YGR", "VK2JE", "VK2ZCE", "VK2DB"}

        Using sqlWriter As New System.IO.StreamWriter("testData.sql")

            ' Create user records
            sqlWriter.WriteLine("-- Create user records")
            sqlWriter.WriteLine($"DELETE FROM users WHERE id>={FIRSTTESTUSER };")
            user_id = FIRSTTESTUSER
            sqlWriter.WriteLine("INSERT INTO users (`id`,`email`,`not_before`)")
            values.Clear()
            For Each callsign In callsigns
                values.Add($"  ({user_id },'{callsign }@wia.org.au','2020-01-01 00:00:00')")
                user_id += 1
            Next
            sqlWriter.WriteLine($"VALUES ({vbCrLf }{String.Join("," & vbCrLf, values.ToArray) }{vbCrLf });")

            ' Create callsign records
            sqlWriter.WriteLine("-- Create callsign records")
            sqlWriter.WriteLine($"DELETE FROM callsigns WHERE id>={FIRSTTESTUSER };")
            user_id = FIRSTTESTUSER
            sqlWriter.WriteLine("INSERT INTO callsigns (`user_id`,`callsign`,`not_before`)")
            values.Clear()
            For Each callsign In callsigns
                values.Add($"  ({user_id },'{callsign }','2020-01-01 00:00:00')")
                user_id += 1
            Next
            sqlWriter.WriteLine($"VALUES ({vbCrLf }{String.Join("," & vbCrLf, values.ToArray) }{vbCrLf });")

            ' Create QSO records
            user_id = FIRSTTESTUSER
            siloIndx = FIRSTSILO
            sqlWriter.WriteLine("-- Create QSO records")
            sqlWriter.WriteLine($"DELETE FROM logs WHERE `user_id`>={FIRSTTESTUSER };")
            TestCase(sqlWriter, user_id, siloIndx, 1, 1)
            TestCase(sqlWriter, user_id, siloIndx, 1, 2)
            TestCase(sqlWriter, user_id, siloIndx, 2, 3)
            siloIndxSave = siloIndx
            For i = 1 To 10
                user_id += 1
                TestCase(sqlWriter, user_id, siloIndx, 1, 1)
                TestCase(sqlWriter, user_id, siloIndx, 1, 2)
                TestCase(sqlWriter, user_id, siloIndx, 1, 3)
                TestCase(sqlWriter, user_id, siloIndx, 1, 4)
                TestCase(sqlWriter, user_id, siloIndx, 2, 2)
                TestCase(sqlWriter, user_id, siloIndx, 3, 2)
            Next
            siloIndx = siloIndxSave
            For i = 1 To 10
                user_id += 1
                TestCase(sqlWriter, user_id, siloIndx, 1, 1)
                TestCase(sqlWriter, user_id, siloIndx, 1, 2)
                TestCase(sqlWriter, user_id, siloIndx, 1, 3)
                TestCase(sqlWriter, user_id, siloIndx, 1, 4)
                TestCase(sqlWriter, user_id, siloIndx, 2, 2)
            Next
            siloIndx = siloIndxSave
            For i = 1 To 10
                user_id += 1
                TestCase(sqlWriter, user_id, siloIndx, 1, 1)
                TestCase(sqlWriter, user_id, siloIndx, 1, 2)
                TestCase(sqlWriter, user_id, siloIndx, 1, 3)
                TestCase(sqlWriter, user_id, siloIndx, 1, 4)
            Next
            siloIndx = siloIndxSave
            For i = 1 To 10
                user_id += 1
                TestCase(sqlWriter, user_id, siloIndx, 1, 1)
                TestCase(sqlWriter, user_id, siloIndx, 1, 2)
                TestCase(sqlWriter, user_id, siloIndx, 1, 3)
            Next
        End Using
        TextBox1.Text = "Done"
    End Sub

    Private Shared Sub TestCase(sqlWriter As System.IO.StreamWriter, user_id As Integer, ByRef siloIndx As Integer, hf As Integer, nonhf As Integer)
        Static dt As DateTime = #2021-02-03 01:00#    ' time of QSO
        ' very large list of silos to choose from - way more than we'll need
        Static silos As String() = {"VK-ABN2", "VK-ALA5", "VK-ALA3", "VK-ALR2", "VK-ALT2", "VK-ALI2", "VK-ALA2", "VK-ANO3", "VK-ANP3", "VK-APA5", "VK-ARO2", "VK-ARS3", "VK-ARN2", "VK-ARN5", "VK-ARK2", "VK-ARA3", "VK-ARE2", "VK-ARL2", "VK-ATA2", "VK-AVA3", "VK-BNA2", "VK-BCH3", "VK-BGN4", "VK-BLA5", "VK-BLG6", "VK-BLL3", "VK-BLE2", "VK-BLU6", "VK-BLR2", "VK-BNN3", "VK-BRE2", "VK-BRN2", "VK-BRA2", "VK-BRT4", "VK-BRB2", "VK-BRE3", "VK-BRT3", "VK-BTA3", "VK-BTT2", "VK-BCN6", "VK-BLA3", "VK-BCM2", "VK-BLN2", "VK-BLL4", "VK-BLA2", "VK-BNL2", "VK-BNS2", "VK-BRX2", "VK-BRG2", "VK-BRK3", "VK-BRN3", "VK-BLH3", "VK-BLI2", "VK-BLA4", "VK-BNY2", "VK-BNI2", "VK-BNW2", "VK-BNN2", "VK-BRP3", "VK-BRW2", "VK-BDN6", "VK-BGE2", "VK-BGA2", "VK-BGI2", "VK-BGH2", "VK-BGT3", "VK-BLT6", "VK-BLN3", "VK-BNN4", "VK-BRT2", "VK-BRR3", "VK-BRD6", "VK-BRK2", "VK-BRG3", "VK-BWS5", "VK-BYK6", "VK-BRZ2", "VK-BRR2", "VK-BRA3", "VK-BRW4", "VK-BRM3", "VK-BRH5", "VK-BRY2", "VK-BRD4", "VK-BRN6", "VK-BRE6", "VK-BRU2", "VK-BRC3", "VK-BRD2", "VK-BCO5", "VK-BCE3", "VK-BDR2", "VK-BNY6", "VK-BNA4", "VK-BNO2", "VK-BRL2", "VK-BRH2", "VK-BRV2", "VK-BRA5", "VK-BRI2", "VK-BRI6", "VK-BRO2", "VK-BRC2", "VK-BRU3", "VK-BTE5", "VK-CLL2", "VK-CLI6", "VK-CLW3", "VK-CLA2", "VK-CMA4", "VK-CMI5", "VK-CMN2", "VK-CNE3", "VK-CNY2", "VK-CNA2", "VK-CPA4", "VK-CRL2", "VK-CRA3", "VK-CRH6", "VK-CRN6", "VK-CRN2", "VK-CRP3", "VK-CCS4", "VK-CNT2", "VK-CHN3", "VK-CHH3", "VK-CHA4", "VK-CHK3", "VK-CLT4", "VK-CLN4", "VK-CBM3", "VK-CCA3", "VK-CLN3", "VK-CLD3", "VK-CMG2", "VK-CNN2", "VK-CNG2", "VK-CLH2", "VK-CLN2", "VK-CNN5", "VK-CNE2", "VK-CTA2", "VK-CPE3", "VK-CRA2", "VK-CRY3", "VK-CRW2", "VK-CSE3", "VK-CWE3", "VK-CWL5", "VK-CWA2", "VK-CRK6", "VK-CRE2", "VK-CRK2", "VK-CRO2", "VK-CRK5", "VK-CLR2", "VK-CLA3", "VK-CLE2", "VK-CMS5", "VK-CMK2", "VK-CNN6", "VK-CNA5", "VK-CNR2", "VK-CRB2", "VK-CRS2", "VK-CRO3", "VK-DHG2", "VK-DLY4", "VK-DNH3", "VK-DRK5", "VK-DLA2", "VK-DNN3", "VK-DRG2", "VK-DTA4", "VK-DVH3", "VK-DVT7", "VK-DPR3", "VK-DGA3", "VK-DMA3", "VK-DMH3", "VK-DNO4", "VK-DND3", "VK-DDE6", "VK-DNE3", "VK-DKE3", "VK-DRA4", "VK-DBO2", "VK-DLA4", "VK-DLL2", "VK-DMS3", "VK-DNO2", "VK-DNY3", "VK-DRI2", "VK-ECA3", "VK-EDI2", "VK-EDE5", "VK-EJG6", "VK-ELM3", "VK-ELE3", "VK-EMD4", "VK-EML2", "VK-ERA2", "VK-EBT2", "VK-EDA5", "VK-EGA2", "VK-EME2", "VK-FRT5", "VK-FNY2", "VK-FRS2", "VK-FRL4", "VK-FRS5", "VK-GBN6", "VK-GRR6", "VK-GLA3", "VK-GLL3", "VK-GMA3", "VK-GNN2", "VK-GRA2", "VK-GRN2", "VK-GLG3", "VK-GRM5", "VK-GRE2", "VK-GDY2", "VK-GLL2", "VK-GLO2", "VK-GNE4", "VK-GRL2", "VK-GLE5", "VK-GLN2", "VK-GLH3", "VK-GLY3", "VK-GBY2", "VK-GCT4", "VK-GDS6", "VK-GLA4", "VK-GLI2", "VK-GLG2", "VK-GMG6", "VK-GNS4", "VK-GNT4", "VK-GNA2", "VK-GRT3", "VK-GRG3", "VK-GVN4", "VK-GRE3", "VK-NRE5", "VK-GWD3", "VK-GYA3", "VK-GRW2", "VK-GRH6", "VK-GRD2", "VK-GRS2", "VK-GRN3", "VK-GRP2", "VK-GRF2", "VK-GRG2", "VK-GLE2", "VK-GLU2", "VK-GNH2", "VK-GND2", "VK-GRY2", "VK-GYN3", "VK-HME5", "VK-HRN2", "VK-HRD2", "VK-HRN4", "VK-HYY2", "VK-HNY2", "VK-HLN2", "VK-HDN4", "VK-HLK2", "VK-HLD2", "VK-HPD2", "VK-HPN3", "VK-HRM3", "VK-HNR3", "VK-HYN6", "VK-IDY4", "VK-ILO2", "VK-INL2", "VK-JCN3", "VK-JCP6", "VK-JMN5", "VK-JNE4", "VK-JPT3", "VK-JRE2", "VK-JMT4", "VK-JNN4", "VK-JNE2", "VK-JNG3", "VK-KDE2", "VK-KRI4", "VK-KLE6", "VK-KMH2", "VK-KMK2", "VK-KNA3", "VK-KPE5", "VK-KPA5", "VK-KRA3", "VK-KRA5", "VK-KTE3", "VK-KTA3", "VK-KTH5", "VK-KLN6", "VK-KRG3", "VK-KCO2", "VK-KML3", "VK-KLA5", "VK-KKA2", "VK-KLN4", "VK-KMA5", "VK-KNY4", "VK-KNH4", "VK-KNE2", "VK-KNL3", "VK-KJP6", "VK-KKY6", "VK-KLG3", "VK-KRA6", "VK-KLI6", "VK-KLN3", "VK-KNL2", "VK-KPN4", "VK-KRN3", "VK-KWA6", "VK-KYA5", "VK-LDH2", "VK-LHH3", "VK-LKA3", "VK-LKO2", "VK-LKE6", "VK-LKG6", "VK-LLT3", "VK-LMO5", "VK-LSS3", "VK-LCT3", "VK-LNS3", "VK-LWE2", "VK-LLR3", "VK-LNA3", "VK-LTD3", "VK-LLY3", "VK-LCK5", "VK-LCT2", "VK-LNE3", "VK-LNS5", "VK-LRN3", "VK-LXN5", "VK-LBK3", "VK-MCR4", "VK-MCU4", "VK-YNG2", "VK-MLA5", "VK-MLU4", "VK-MNG3", "VK-MNH2", "VK-MNA2", "VK-MNL2", "VK-MRD7", "VK-MRM2", "VK-MRO3", "VK-MRR2", "VK-MSY3", "VK-MTA2", "VK-MTG2", "VK-MCE2", "VK-MCE6", "VK-MNA4", "VK-MTN3", "VK-MCG6", "VK-MNN2", "VK-MRR3", "VK-MRE3", "VK-MRA2", "VK-MRG2", "VK-MRO2", "VK-MRN2", "VK-MCI2", "VK-MLG2", "VK-MLS4", "VK-MLG6", "VK-MLT5", "VK-MLN4", "VK-MLE2", "VK-MNW6", "VK-MNA5", "VK-MNO2", "VK-MNP3", "VK-MRI3", "VK-MRM3", "VK-MRL2", "VK-MTO3", "VK-MTE3", "VK-MTK3", "VK-MGR6", "VK-MLN2", "VK-MNH5", "VK-MLT3", "VK-MML2", "VK-MRA6", "VK-MRE2", "VK-MRS2", "VK-MRS3", "VK-MRH2", "VK-MLI2", "VK-MNN6", "VK-MNE2", "VK-MNN4", "VK-MRA4", "VK-EGD2", "VK-MKN6", "VK-MLA6", "VK-MNR2", "VK-MNP6", "VK-MNB2", "VK-MRT3", "VK-MRA5", "VK-MRI2", "VK-MRE5", "VK-MRL3", "VK-MRA3", "VK-MYK3", "VK-NGN3", "VK-NNY3", "VK-NNE4", "VK-NRN2", "VK-NRI2", "VK-NRT2", "VK-NRA2", "VK-NRE2", "VK-NRS2", "VK-NTI3", "VK-NTK3", "VK-NTA3", "VK-NEA4", "VK-NTY3", "VK-NVE2", "VK-NWA2", "VK-NWE6", "VK-NWN2", "VK-NHL3", "VK-NNA3", "VK-NNA2", "VK-NNO4", "VK-NRM6", "VK-NRN6", "VK-NRG3", "VK-NRE4", "VK-NRR2", "VK-NRD4", "VK-NRG2", "VK-NRN4", "VK-NLN3", "VK-NLL3", "VK-NMH3", "VK-NNG3", "VK-NNN6", "VK-NNA5", "VK-NWH2", "VK-NYG6", "VK-NYT3", "VK-NYN3", "VK-NYX3", "VK-NYN2", "VK-NYK2", "VK-OKY4", "VK-OKS3", "VK-OKE3", "VK-OLE2", "VK-OTA2", "VK-ORO5", "VK-OYN3", "VK-OWN5", "VK-PNA3", "VK-PRA5", "VK-PRG5", "VK-PRS2", "VK-PSE5", "VK-PTK3", "VK-PKE5", "VK-PKL2", "VK-PNP2", "VK-PNG5", "VK-PPL7", "VK-PRE3", "VK-PNL3", "VK-PCA3", "VK-PRN3", "VK-PMO3", "VK-PNE3", "VK-PNY6", "VK-PNP6", "VK-PNA4", "VK-PNO5", "VK-PRA3", "VK-PRA2", "VK-PLS2", "VK-PCA5", "VK-PRD5", "VK-PRY6", "VK-PRS5", "VK-PRL2", "VK-PRD2", "VK-PRD3", "VK-PRN5", "VK-PRE4", "VK-PRE5", "VK-PRI3", "VK-PRR2", "VK-PRE2", "VK-PRF7", "VK-PCN2", "VK-PLT3", "VK-PYD3", "VK-QRG6", "VK-QMK3", "VK-QNY2", "VK-QNA2", "VK-QRI2", "VK-QRN5", "VK-RNW3", "VK-RND2", "VK-RNS2", "VK-RVE6", "VK-RYD3", "VK-RZK2", "VK-RDD2", "VK-RDL5", "VK-RFN2", "VK-RNE2", "VK-RCG3", "VK-RBN5", "VK-RBE3", "VK-RCR3", "VK-RMA4", "VK-RSY3", "VK-RSY5", "VK-RWA2", "VK-RZE2", "VK-RDA5", "VK-RPP3", "VK-RSH3", "VK-RTN3", "VK-SDH5", "VK-SLY3", "VK-SNO2", "VK-SNR2", "VK-SLE3", "VK-SGI2", "VK-SRN3", "VK-SHS3", "VK-SHE3", "VK-SHS2", "VK-SKN3", "VK-SLE2", "VK-SNN5", "VK-STS6", "VK-STA2", "VK-SPD3", "VK-SPT3", "VK-SPE2", "VK-SPE4", "VK-STD3", "VK-STL3", "VK-STS3", "VK-STL2", "VK-STL5", "VK-STN5", "VK-STH2", "VK-STY5", "VK-SML2", "VK-SNE3", "VK-STN3", "VK-SWL3", "VK-SWR3", "VK-TBA2", "VK-TLR2", "VK-TLA5", "VK-TLN2", "VK-TLA3", "VK-TMA2", "VK-TMN6", "VK-TNA3", "VK-TRA4", "VK-TRN6", "VK-TRE5", "VK-TRE2", "VK-TRK3", "VK-TTN3", "VK-TDY3", "VK-TLD3", "VK-TMR2", "VK-TMY3", "VK-THN4", "VK-THG2", "VK-THS4", "VK-THN2", "VK-THS2", "VK-THD5", "VK-THH2", "VK-THO4", "VK-THS6", "VK-TCE2", "VK-TNN6", "VK-TNA5", "VK-TCL2", "VK-TCH2", "VK-TMT2", "VK-TBH4", "VK-TLE5", "VK-TTL2", "VK-TWY4", "VK-TRA3", "VK-TTM2", "VK-MDK2", "VK-TRI2", "VK-TRG6", "VK-TRL2", "VK-TLE2", "VK-TLL2", "VK-TMY5", "VK-TNH3", "VK-TRF3", "VK-TTE3", "VK-ULI2", "VK-ULA3", "VK-UNL3", "VK-UNE2", "VK-UNA5", "VK-URA2", "VK-URT2", "VK-URY2", "VK-VCS3", "VK-WAA3", "VK-WDE5", "VK-WGN6", "VK-WKE5", "VK-WLL3", "VK-WTE3", "VK-WKL2", "VK-WLG2", "VK-WLO3", "VK-WLO5", "VK-WLA2", "VK-WLN2", "VK-WLA4", "VK-WLR3", "VK-WLR2", "VK-WLP3", "VK-WNI5", "VK-WNN4", "VK-WNG3", "VK-WRN2", "VK-WRD2", "VK-WRE3", "VK-WRA4", "VK-WRL3", "VK-WRO5", "VK-WRK4", "VK-WTM3", "VK-WTA3", "VK-WRL2", "VK-WTA2", "VK-WDN3", "VK-WDN2", "VK-WTE2", "VK-WWA2", "VK-WJA2", "VK-WLO2", "VK-WLD6", "VK-WRU3", "VK-WRK2", "VK-WSH6", "VK-WSE2", "VK-WSE6", "VK-WSD6", "VK-WSE3", "VK-WSR3", "VK-WSN6", "VK-WSG2", "VK-WHA5", "VK-WHN2", "VK-WCN6", "VK-WDI2", "VK-WLA3", "VK-WLE2", "VK-WLS6", "VK-WLT2", "VK-WNA3", "VK-WRA2", "VK-WRA5", "VK-WRY2", "VK-WRL5", "VK-WGL6", "VK-WLY5", "VK-WMA2", "VK-WDK2", "VK-WMG3", "VK-WRN3", "VK-WRH3", "VK-WBL3", "VK-WDA5", "VK-WNU3", "VK-WNR5", "VK-WRN5", "VK-WYM6", "VK-WYG2", "VK-WYA2", "VK-WYF3", "VK-WYA3", "VK-YPT3", "VK-YBH3", "VK-YNC3", "VK-YNO2", "VK-YNE5", "VK-YRI2", "VK-YRA4", "VK-YRA3", "VK-YRG3", "VK-YLG6", "VK-YLA5", "VK-YLN4", "VK-YLT3", "VK-YNA2", "VK-YVL2", "VK-YRN6", "VK-YRK2", "VK-YDH2", "VK-YNA5", "VK-YNE3", "VK-YNN2", "VK-YLA3"}
        Static Callsigns As String() = {"VK3OHM", "VK3IP", "VK3TIN", "VK3ABC", "VK3DEF", "VK3GHI", "VK3JKL", "VK3MNO", "VK3PQR", "VK3STU", "VK3VWX"}

        Dim silo_id As Integer = siloIndx
        Dim i As Integer, values As New List(Of String)

        sqlWriter.WriteLine($"-- Test Case silo_id= {silo_id} hf QSO={hf } non HF QSO={nonhf }")
        values.Clear()
        sqlWriter.WriteLine("INSERT INTO users (`station`,`user_id`,`silo_id`,`time_on`,`mode`,`band`,`call`)")
        For i = 1 To hf
            values.Add($"  ('VK3OHM',{user_id },{silo_id },'{dt:yyyy-MM-dd HH:mm}','CW','20m','{Callsigns(i) }')")
            dt = dt.AddMinutes(1)    ' auto increment to next minute
        Next
        For i = 1 To nonhf
            values.Add($"  ('VK3OHM',{user_id },{silo_id },'{dt:yyyy-MM-dd HH:mm}','SSB','2m','{Callsigns(i) }')")
            dt = dt.AddMinutes(1)    ' auto increment to next minute
        Next
        sqlWriter.WriteLine($"VALUES ({vbCrLf }{String.Join("," & vbCrLf, values.ToArray) }{vbCrLf });")
        siloIndx += 1     ' next silo
    End Sub

    Private Sub SiloMapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SiloMapToolStripMenuItem.Click
        ' Make a map of silos for PnP site
        Dim count As Integer = 0
        Dim now As String = DateTime.UtcNow.ToString("O")   ' UTC date/time in ISO 8601 format
        Dim coords As New List(Of String), buffer As Polygon
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader
        Dim BaseFilename As String = "SiOTA"
        Dim pins As New Dictionary(Of String, String) From {
            {"Silo Art", "https://maps.google.com/mapfiles/kml/paddle/ylw-stars.png"},
            {"Very Interesting", "https://maps.google.com/mapfiles/kml/paddle/ylw-square.png"},
            {"Interesting", "https://maps.google.com/mapfiles/kml/paddle/grn-diamond.png"},
            {"Just a Silo", "https://maps.google.com/mapfiles/kml/paddle/blu-circle.png"},
            {"Unrated", "https://maps.google.com/mapfiles/kml/paddle/red-blank.png"}
            }
        Dim states As New Dictionary(Of String, String) From {
            {"New South Wales", "VK2"},
            {"Victoria", "VK3"},
            {"Queensland", "VK4"},
            {"South Australia", "VK5"},
            {"Western Australia", "VK6"},
            {"Tasmania", "VK7"},
            {"Northern Territory", "VK8"}
            }

        Using connect As New SQLiteConnection(SILOSdb)
            connect.Open()  ' open database
            For Each state In states
                Dim BaseFilenameState As String = $"{BaseFilename }-{state.Value }"
                Using kmlWriter As New System.IO.StreamWriter($"{BaseFilenameState }.kml")
                    kmlWriter.WriteLine(KMLheader)
                    kmlWriter.WriteLine("<description><![CDATA[<style>table, th, td {white-space:nowrap; }</style>")
                    kmlWriter.WriteLine("<table>")
                    kmlWriter.WriteLine("<tr><td>Data produced by Marc Hillman - VK3OHM/VK3IP</td></tr>")
                    Dim utc As String = String.Format("{0:dd MMM yyyy hh:mm:ss UTC}", DateTime.UtcNow)     ' time now in UTC
                    kmlWriter.WriteLine($"<tr><td>Data extracted from SiOTA on {utc }.</td></tr>")
                    kmlWriter.WriteLine($"<tr><td>Silo activation zones for {state.Value }</tr></td>")
                    kmlWriter.WriteLine("</table>")
                    kmlWriter.WriteLine("]]>")
                    kmlWriter.WriteLine("</description>")
                    ' Create styles for each type of pin
                    For Each pin In pins
                        kmlWriter.WriteLine($"<Style id='{pin.Key }'>")
                        kmlWriter.WriteLine("<LineStyle><width>2</width><color>ffffffff</color></LineStyle>")   ' white outline
                        kmlWriter.WriteLine($"<IconStyle><Icon><href>{pin.Value }</href></Icon><scale>1</scale></IconStyle>")   ' pin
                        kmlWriter.WriteLine($"<PolyStyle><color>{KMLColor(PolyAlpha, 0, 0, 255) }</color><fill>1</fill><outline>1</outline></PolyStyle>")
                        kmlWriter.WriteLine("</Style>")
                    Next

                    sqlcmd = connect.CreateCommand()
                    sqlcmd.CommandText = $"SELECT * FROM silos WHERE not_before<='{now }' AND not_after='' and state='{state.Key }' ORDER BY silo_code"     ' select active silos
                    SQLdr = sqlcmd.ExecuteReader()
                    While SQLdr.Read()
                        Try
                            count += 1
                            ' check that entry is valid date
                            ' silo is valid. add silo to map
                            kmlWriter.WriteLine("<Placemark>")
                            kmlWriter.WriteLine($"<name>{SQLdr("silo_code") } - {SQLdr("name") }</name>")
                            Dim style As String
                            If SQLdr("arty").ToLower = "true" Then style = "#Silo Art" Else style = $"#{SQLdr("rating") }"
                            kmlWriter.WriteLine($"<styleUrl>{style }</styleUrl>")
                            kmlWriter.WriteLine("<ExtendedData>")
                            kmlWriter.WriteLine($"<Data name='Silo code'><value>{SQLdr("silo_code") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='Position'><value>{SQLdr("lat") },{SQLdr("lng") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='Grid'><value>{SQLdr("locator") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='Name'><value>{SQLdr("name") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='Locality'><value>{SQLdr("locality") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='State'><value>{SQLdr("state") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='LGA code'><value>{SQLdr("lga_code") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='Railway'><value>{SQLdr("railway") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='Rail status'><value>{SQLdr("rail_status") }</value></Data>")
                            kmlWriter.WriteLine($"<Data name='Rating'><value>{SQLdr("rating") }</value></Data>")
                            If Not String.IsNullOrEmpty(SQLdr("park").ToString) Then kmlWriter.WriteLine($"<Data name='Nearby park'><value>{SQLdr("park") }</value></Data>")
                            Dim link As String = HtmlEncode($"<a href='{SQLdr("street_view") }'>Link</a>")
                            kmlWriter.WriteLine($"<Data name='Street View'><value>{link }</value></Data>")
                            kmlWriter.WriteLine("</ExtendedData>")
                            Dim p As New MapPoint(CDbl(SQLdr("lng")), CDbl(SQLdr("lat")), SpatialReferences.Wgs84)  ' coordinate of silo
                            kmlWriter.WriteLine("<MultiGeometry>")
                            kmlWriter.WriteLine("<Point>")
                            kmlWriter.WriteLine($"<coordinates>{p.X:f5},{p.Y:f5}</coordinates>")
                            kmlWriter.WriteLine("</Point>")
                            kmlWriter.WriteLine("<Polygon>")
                            kmlWriter.WriteLine("<tessellate>1</tessellate>")
                            kmlWriter.WriteLine("<outerBoundaryIs>")
                            kmlWriter.WriteLine("<LinearRing>")
                            kmlWriter.WriteLine("<coordinates>")
                            buffer = GeometryEngine.BufferGeodetic(p, SILO_ACTIVATION_ZONE, LinearUnits.Meters)        ' generate 1km circle
                            coords.Clear()
                            For Each pnt As MapPoint In buffer.Parts(0).Points
                                coords.Add($"{pnt.X:f5},{pnt.Y:f5}")
                            Next
                            coords.Add(coords(0))    ' close the circle
                            kmlWriter.WriteLine(String.Join(" ", coords))
                            kmlWriter.WriteLine("</coordinates>")
                            kmlWriter.WriteLine("</LinearRing>")
                            kmlWriter.WriteLine("</outerBoundaryIs>")
                            kmlWriter.WriteLine("</Polygon>")
                            kmlWriter.WriteLine("</MultiGeometry>")
                            kmlWriter.WriteLine("</Placemark>")
                        Catch ex As Exception
                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                        End Try
                    End While
                    kmlWriter.WriteLine(KMLfooter)
                End Using
                ' compress to zip file
                System.IO.File.Delete($"{BaseFilenameState }.kmz")
                Dim zip As ZipArchive = ZipFile.Open($"{BaseFilenameState }.kmz", ZipArchiveMode.Create)    ' create new archive file
                zip.CreateEntryFromFile($"{BaseFilenameState }.kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
                zip.Dispose()
            Next
        End Using
        SetText(TextBox1, $"{count } placemarks created")
    End Sub
    Private Async Sub SilosNearParksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SilosInParksToolStripMenuItem.Click
        ' Find silos that overlap parks
        Dim datenow As String = DateTime.UtcNow.ToString("O")   ' UTC date/time in ISO 8601 format
        Dim sqlcmd_park As SQLiteCommand, SQLdr_park As SQLiteDataReader
        Dim sqlcmd_silo As SQLiteCommand, SQLdr_silo As SQLiteDataReader
        Dim myQueryFilter As QueryParameters, extent As Envelope
        Dim silo_overlaps As New NameValueCollection()  ' list of silo and overlapping parks
        Dim park_overlaps As New NameValueCollection()  ' list of parks and overlapping silos
        Dim count As Integer = 0, found As Integer = 0, total As Integer = 0, started As DateTime = Now()
        Dim ParkData = New NameValueCollection, GISIDListQuoted As String, where As String

        Using connect_silo As New SQLiteConnection(SILOSdb),
            connect_park As New SQLiteConnection(PARKSdb),
            htmlWriter As New StreamWriter("SilosNearParks.html")

            htmlWriter.AutoFlush = True
            connect_silo.Open()  ' open database
            connect_park.Open()
            htmlWriter.WriteLine("<table border=1>")
            htmlWriter.WriteLine("<tr><th>WWFFID</th><th>Name</th><th>Silo</th><th>Name</th></tr>")
            sqlcmd_park = connect_park.CreateCommand()
            sqlcmd_park.CommandText = $"SELECT count(*) as total FROM parks WHERE Status IN ('active','Active') AND State IN ('VK2','VK3','VK4','VK5','VK6','VK7','VK8')"     ' select active parks
            SQLdr_park = sqlcmd_park.ExecuteReader()
            SQLdr_park.Read()
            total = SQLdr_park("total")
            SQLdr_park.Close()
            sqlcmd_park.CommandText = $"SELECT * FROM parks WHERE Status IN ('active','Active') AND State IN ('VK2','VK3','VK4','VK5','VK6','VK7','VK8')"     ' select active parks
            SQLdr_park = sqlcmd_park.ExecuteReader()
            While SQLdr_park.Read() ' Search each park
                count += 1
                ParkData = GetParkData(SQLdr_park("WWFFID").ToString)
                Dim WWFFID As String = SQLdr_park("WWFFID")
                Dim Name As String = $"{SQLdr_park("Name") } {SQLdr_park("Type") }"
                Dim ds As String = ParkData("Dataset")
                ' Get park shape and find extent
                ' Add border of 1km to extent
                ' Only check silos within expanded extent
                GISIDListQuoted = ParkData("GISIDListQuoted")
                If Not GISIDListQuoted = "" Then
                    where = DataSets(ds).BuildWhere(GISIDListQuoted)           ' build query statement
                    myQueryFilter = New QueryParameters With {
                        .WhereClause = where,    ' query parameters
                        .ReturnGeometry = True,
                        .OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
                        }
                    DataSets(ds).shpFragments = Await DataSets(ds).shpShapeFileTable.QueryFeaturesAsync(myQueryFilter)           ' run query
                    ' find the extent of the park
                    If DataSets(ds).shpFragments.Count > 0 Then
                        extent = Nothing
                        For Each fragment In DataSets(ds).shpFragments
                            extent = EnvelopeUnion(extent, fragment.Geometry.Extent)
                        Next
                        ' Increase extent by silo activation zone
                        Dim buffer As Geometry = GeometryEngine.BufferGeodetic(extent, SILO_ACTIVATION_ZONE, LinearUnits.Meters)              ' silo location with buffer
                        extent = buffer.Extent
                        ' Now search for silos within extent
                        sqlcmd_silo = connect_silo.CreateCommand()
                        sqlcmd_silo.CommandText = $"SELECT * FROM silos WHERE not_before<='{datenow }' AND not_after='' AND lat BETWEEN {extent.Extent.YMin } AND {extent.Extent.YMax } AND lng BETWEEN {extent.Extent.XMin } AND {extent.Extent.XMax }"     ' select silos in this state
                        SQLdr_silo = sqlcmd_silo.ExecuteReader()
                        While SQLdr_silo.Read()
                            ' test each candidate silo - there may be more than one
                            SetText(TextBox1, $"testing {SQLdr_silo("silo_code") }")
                            Dim silo As New MapPoint(SQLdr_silo("lng"), SQLdr_silo("lat"), SpatialReferences.Wgs84)      ' silo coordinate
                            Dim AZ As Geometry = GeometryEngine.BufferGeodetic(silo, SILO_ACTIVATION_ZONE, LinearUnits.Meters)              ' silo location with buffer
                            ' test for intersection of park and silo, i.e. buffer and shpFragments intersect
                            For Each fragment In DataSets(ds).shpFragments
                                If GeometryEngine.Intersects(AZ, fragment.Geometry) Then
                                    htmlWriter.WriteLine($"<tr><td>{WWFFID }</td><td>{Name }</td><td>{SQLdr_silo("silo_code") }</td><td>{SQLdr_silo("name") }</td></tr>")
                                    silo_overlaps.Add(SQLdr_silo("silo_code"), WWFFID)
                                    park_overlaps.Add(WWFFID, SQLdr_silo("silo_code"))
                                    found += 1
                                    Exit For
                                End If
                            Next
                        End While
                        SQLdr_silo.Close()
                    Else
                        ' MsgBox($"{ParkData("WWFFID") }, PA_ID={GISIDListQuoted } has no fragments", vbCritical + vbOKOnly, "Data error")
                    End If
                End If
                SetText(TextBox1, $"{count }/{total } parks searched, {found } silos found. Finish {TogoFormat(started, count, total) }")
            End While
            SQLdr_park.Close()
            htmlWriter.WriteLine("</table>")
        End Using

        ' Create SQL files
        Using siloSQLWriter As New StreamWriter("SilosNearParks.sql"),
            parkSQLWriter As New StreamWriter("ParksNearSilos.sql")
            ' Create SQL from overlap list
            siloSQLWriter.WriteLine("START TRANSACTION;")
            For Each silo In silo_overlaps.AllKeys
                Dim values() As String = silo_overlaps.GetValues(silo)   ' get all values for this park
                Dim parklist As String = Join(values, ",")          ' make csv list
                siloSQLWriter.WriteLine($"UPDATE silos SET park='{parklist }' WHERE silo_code='{silo }';") ' construct SQL
            Next
            siloSQLWriter.WriteLine("COMMIT;")
            ' Create SQL from overlap list
            siloSQLWriter.WriteLine("START TRANSACTION;")
            For Each park In park_overlaps.AllKeys
                Dim values() As String = park_overlaps.GetValues(park)   ' get all values for this park
                Dim silolist As String = Join(values, ",")          ' make csv list
                parkSQLWriter.WriteLine($"UPDATE parks SET silos='{silolist }' WHERE WWFFID='{park }';") ' construct SQL
            Next
            siloSQLWriter.WriteLine("COMMIT;")
        End Using
        SetText(TextBox1, $"Done: {count }/{total } parks searched, {found } silos found")
    End Sub


    Private Async Sub ExperimentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExperimentToolStripMenuItem.Click

        ' Find silos that overlap parks
        Dim datenow As String = DateTime.UtcNow.ToString("O")   ' UTC date/time in ISO 8601 format
        Dim sqlcmd_park As SQLiteCommand, SQLdr_park As SQLiteDataReader
        Dim sqlcmd_silo As SQLiteCommand, SQLdr_silo As SQLiteDataReader
        Dim myQueryFilter As QueryParameters, extent As Envelope = Nothing
        Dim silo_overlaps As New NameValueCollection()  ' list of silo and overlapping parks
        Dim park_overlaps As New NameValueCollection()  ' list of parks and overlapping silos
        Dim count As Integer = 0, found As Integer = 0, total As Integer = 0, started As DateTime = Now()
        Dim ParkData = New NameValueCollection, GISIDListQuoted As String, where As String
        Dim watchTotal As New Stopwatch(), watchExtent As New Stopwatch(), watchFetch As New Stopwatch(), watchTest As New Stopwatch()
        Dim fragments As FeatureQueryResult = Nothing


        ' Open all datasets
        For Each ds In DatasetDict.Keys
            Dim fi As New IO.FileInfo(DatasetDict(ds)("Shapefile"))
            Dim ext As String = fi.Extension
            Select Case ext
                Case ".shp"
                    DatasetDict(ds)("shpShapeFileTable") = Await ShapefileFeatureTable.OpenAsync(DatasetDict(ds)("Shapefile")).ConfigureAwait(True)
                Case ".gdb"
                    ' gdb not supported yet
                    DatasetDict(ds)("shpShapeFileTable") = Await Geodatabase.OpenAsync(DatasetDict(ds)("Shapefile")).ConfigureAwait(True)
                Case Else
                    Throw New InvalidOperationException($"{DatasetDict(ds)("shpShapeFileTable")} does not have a recognised extension")
            End Select
        Next

        Using connect_silo As New SQLiteConnection(SILOSdb),
            connect_park As New SQLiteConnection(PARKSdb),
            htmlWriter As New StreamWriter("SilosNearParks.html")

            htmlWriter.AutoFlush = True
            connect_silo.Open()  ' open database
            connect_park.Open()
            htmlWriter.WriteLine("<table border=1>")
            htmlWriter.WriteLine("<tr><th>WWFFID</th><th>Name</th><th>Silo</th><th>Name</th></tr>")
            sqlcmd_park = connect_park.CreateCommand()
            sqlcmd_park.CommandText = $"SELECT count(*) as total FROM parks WHERE Status IN ('active','Active') AND State IN ('VK2','VK3','VK4','VK5','VK6','VK7','VK8')"     ' select active parks
            SQLdr_park = sqlcmd_park.ExecuteReader()
            SQLdr_park.Read()
            total = SQLdr_park("total")
            SQLdr_park.Close()
            sqlcmd_park.CommandText = $"SELECT * FROM parks WHERE Status IN ('active','Active') AND State IN ('VK2','VK3','VK4','VK5','VK6','VK7','VK8')"     ' select active parks
            SQLdr_park = sqlcmd_park.ExecuteReader()
            While SQLdr_park.Read() ' Search each park
                watchTotal.Restart()
                count += 1
                ParkData = GetParkData(SQLdr_park("WWFFID").ToString)
                Dim WWFFID As String = SQLdr_park("WWFFID")
                Dim Name As String = $"{SQLdr_park("Name") } {SQLdr_park("Type") }"
                Dim ds As String = ParkData("Dataset")
                ' Get park shape and find extent
                ' Add border of 1km to extent
                ' Only check silos within expanded extent
                GISIDListQuoted = ParkData("GISIDListQuoted")
                If Not GISIDListQuoted = "" Then
                    where = Dataset_BuildWhere(ds, GISIDListQuoted)           ' build query statement
                    myQueryFilter = New QueryParameters With {
                        .WhereClause = where,    ' query parameters
                        .OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
                        }
                    watchExtent.Restart()
                    extent = Await DatasetDict(ds)("shpShapeFileTable").QueryExtentAsync(myQueryFilter)           ' run query
                    watchExtent.Stop()
                    'Dim fragments As FeatureQueryResult = DatasetDict(ds)("shpFragments")
                    ' find the extent of the park
                    If Not extent.IsEmpty Then
                        'watchExtent = Stopwatch.StartNew()
                        ''extent = Aggregate fragment In fragments Into AggrExtent(fragment)
                        'extent = fragments.First.Geometry.Extent
                        'For i = 1 To fragments.Count - 1
                        '    extent = GeometryEngine.CombineExtents(extent, fragments(i).Geometry)
                        'Next
                        'watchExtent.Stop()
                        ' Increase extent by silo activation zone
                        Dim buffer As Geometry = GeometryEngine.BufferGeodetic(extent, SILO_ACTIVATION_ZONE, LinearUnits.Meters)              ' silo location with buffer
                        extent = buffer.Extent
                        ' Now search for silos within extent
                        sqlcmd_silo = connect_silo.CreateCommand()
                        sqlcmd_silo.CommandText = $"SELECT * FROM silos WHERE not_before<='{datenow }' AND not_after='' AND lat BETWEEN {extent.Extent.YMin } AND {extent.Extent.YMax } AND lng BETWEEN {extent.Extent.XMin } AND {extent.Extent.XMax }"     ' select silos in this state
                        SQLdr_silo = sqlcmd_silo.ExecuteReader()
                        watchFetch.Restart()
                        If SQLdr_silo.HasRows Then fragments = Await DatasetDict(ds)("shpShapeFileTable").QueryFeaturesAsync(myQueryFilter)           ' run query
                        watchFetch.Stop()
                        watchTest.Restart()
                        While SQLdr_silo.Read()
                            SetText(TextBox1, $"testing {SQLdr_silo("silo_code") }")
                            Dim silo As New MapPoint(SQLdr_silo("lng"), SQLdr_silo("lat"), SpatialReferences.Wgs84)      ' silo coordinate
                            Dim AZ As Geometry = GeometryEngine.BufferGeodetic(silo, SILO_ACTIVATION_ZONE, LinearUnits.Meters)              ' silo location with buffer
                            ' test for intersection of park and silo, i.e. buffer and shpFragments intersect
                            'Dim fragmentsEnumerable = fragments.AsEnumerable
                            For Each fragment In fragments
                                If GeometryEngine.Intersects(AZ, fragment.Geometry) Then
                                    htmlWriter.WriteLine($"<tr><td>{WWFFID }</td><td>{Name }</td><td>{SQLdr_silo("silo_code") }</td><td>{SQLdr_silo("name") }</td></tr>")
                                    silo_overlaps.Add(SQLdr_silo("silo_code"), WWFFID)
                                    park_overlaps.Add(WWFFID, SQLdr_silo("silo_code"))
                                    found += 1
                                    Exit For
                                End If
                            Next
                        End While
                        watchTest.Stop()
                        SQLdr_silo.Close()
                    Else
                        ' MsgBox($"{ParkData("WWFFID") }, PA_ID={GISIDListQuoted } has no fragments", vbCritical + vbOKOnly, "Data error")
                    End If
                End If
                watchTotal.Stop()
                AppendText(TextBox2, $"{count,4 } Extent:{watchExtent.ElapsedMilliseconds,5} Fetch:{watchFetch.ElapsedMilliseconds,5} Test:{watchTest.ElapsedMilliseconds,5} Total:{watchTotal.ElapsedMilliseconds,5}{vbCrLf}")
                SetText(TextBox1, $"{count }/{total } parks searched, {found } silos found. Finish {TogoFormat(started, count, total) }")
            End While
            SQLdr_park.Close()
            htmlWriter.WriteLine("</table>")
        End Using

        ' Create SQL files
        Using siloSQLWriter As New StreamWriter("SilosNearParks.sql"),
            parkSQLWriter As New StreamWriter("ParksNearSilos.sql")
            ' Create SQL from overlap list
            siloSQLWriter.WriteLine("START TRANSACTION;")
            For Each silo In silo_overlaps.AllKeys
                Dim values() As String = silo_overlaps.GetValues(silo)   ' get all values for this park
                Dim parklist As String = Join(values, ",")          ' make csv list
                siloSQLWriter.WriteLine($"UPDATE silos SET park='{parklist }' WHERE silo_code='{silo }';") ' construct SQL
            Next
            siloSQLWriter.WriteLine("COMMIT;")
            ' Create SQL from overlap list
            siloSQLWriter.WriteLine("START TRANSACTION;")
            For Each park In park_overlaps.AllKeys
                Dim values() As String = park_overlaps.GetValues(park)   ' get all values for this park
                Dim silolist As String = Join(values, ",")          ' make csv list
                parkSQLWriter.WriteLine($"UPDATE parks SET silos='{silolist }' WHERE WWFFID='{park }';") ' construct SQL
            Next
            siloSQLWriter.WriteLine("COMMIT;")
        End Using
        SetText(TextBox1, $"Done: {count }/{total } parks searched, {found } silos found")
    End Sub

    Private Async Sub Experiment2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Experiment2ToolStripMenuItem.Click
        ' Find silos that overlap parks
        Dim datenow As String = DateTime.UtcNow.ToString("O")   ' UTC date/time in ISO 8601 format
        Dim sqlcmd_park As SQLiteCommand, SQLdr_park As SQLiteDataReader
        Dim sqlcmd_silo As SQLiteCommand, SQLdr_silo As SQLiteDataReader
        Dim myQueryFilter As QueryParameters
        Dim silo_overlaps As New NameValueCollection()  ' list of silo and overlapping parks
        Dim park_overlaps As New NameValueCollection()  ' list of parks and overlapping silos
        Dim count As Integer = 0, found As Integer = 0, total As Integer = 0, started As DateTime = Now()
        Dim watchTotal As New Stopwatch(), watchQuery As New Stopwatch(), watchTest As New Stopwatch()
        Dim overlaps As FeatureQueryResult, buffer As Geometry, dataset As String

        Using connect_silo As New SQLiteConnection(SILOSdb),
            connect_park As New SQLiteConnection(PARKSdb),
            htmlWriter As New StreamWriter("SilosNearParks.html")

            htmlWriter.AutoFlush = True
            connect_silo.Open()  ' open database
            sqlcmd_silo = connect_silo.CreateCommand()
            connect_park.Open()
            sqlcmd_park = connect_park.CreateCommand()

            ' Count total number of silos
            sqlcmd_silo.CommandText = $"SELECT COUNT(*) as total FROM silos WHERE not_before<='{datenow }' AND not_after=''"     ' all active silos
            SQLdr_silo = sqlcmd_silo.ExecuteReader()
            SQLdr_silo.Read()
            total = SQLdr_silo("total")
            SQLdr_silo.Close()

            htmlWriter.WriteLine("<table border=1>")
            htmlWriter.WriteLine("<tr><th>WWFFID</th><th>Name</th><th>Silo</th><th>Name</th></tr>")

            ' test each silo in turn
            sqlcmd_silo.CommandText = $"SELECT * FROM silos WHERE not_before<='{datenow }' AND not_after=''"     ' all active silos
            SQLdr_silo = sqlcmd_silo.ExecuteReader()
            While SQLdr_silo.Read() ' Search each silo
                watchTotal.Restart()
                count += 1
                buffer = GeometryEngine.BufferGeodetic(New MapPoint(SQLdr_silo("lng"), SQLdr_silo("lat"), SpatialReferences.Wgs84), SILO_ACTIVATION_ZONE, LinearUnits.Meters)
                myQueryFilter = New QueryParameters With {
                        .WhereClause = "1=1",    ' query parameters
                        .Geometry = buffer,
                        .SpatialRelationship = SpatialRelationship.Overlaps,
                        .OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
                        }
                watchQuery.Restart()
                dataset = "CAPAD_T"
                overlaps = Await DataSets(dataset).shpShapeFileTable.QueryFeaturesAsync(myQueryFilter)           ' run query
                watchQuery.Stop()
                ' find the extent of the park
                watchTest.Restart()
                If overlaps.Any Then
                    For Each park In overlaps
                        ' We have overlaps. Determine if they are WWFF parks
                        sqlcmd_park.CommandText = $"SELECT * FROM GISmapping WHERE DataSet='{dataset}' AND GISID='{park.GetAttributeValue("PA_ID")}'"
                        SQLdr_park = sqlcmd_park.ExecuteReader()
                        While SQLdr_park.Read()
                            Dim WWFFID As String = SQLdr_park("WWFFID")
                            htmlWriter.WriteLine($"<tr><td>{WWFFID }</td><td>{Name }</td><td>{SQLdr_silo("silo_code") }</td><td>{SQLdr_silo("name") }</td></tr>")
                            silo_overlaps.Add(SQLdr_silo("silo_code"), WWFFID)
                            park_overlaps.Add(WWFFID, SQLdr_silo("silo_code"))
                            found += 1
                        End While
                        SQLdr_park.Close()
                    Next
                End If
                watchTest.Stop()
                watchTotal.Stop()
                AppendText(TextBox2, $"{count,4 } Query:{watchQuery.ElapsedMilliseconds,5} Test:{watchTest.ElapsedMilliseconds,5} Total:{watchTotal.ElapsedMilliseconds,5}{vbCrLf}")
                SetText(TextBox1, $"{count }/{total } parks searched, {found } silos found. Finish {TogoFormat(started, count, total) }")
            End While
            SQLdr_silo.Close()
            htmlWriter.WriteLine("</table>")
        End Using

        ' Create SQL files
        Using siloSQLWriter As New StreamWriter("SilosNearParks.sql"),
            parkSQLWriter As New StreamWriter("ParksNearSilos.sql")
            ' Create SQL from overlap list
            siloSQLWriter.WriteLine("START TRANSACTION;")
            For Each silo In silo_overlaps.AllKeys
                Dim values() As String = silo_overlaps.GetValues(silo)   ' get all values for this park
                Dim parklist As String = Join(values, ",")          ' make csv list
                siloSQLWriter.WriteLine($"UPDATE silos SET park='{parklist }' WHERE silo_code='{silo }';") ' construct SQL
            Next
            siloSQLWriter.WriteLine("COMMIT;")
            ' Create SQL from overlap list
            siloSQLWriter.WriteLine("START TRANSACTION;")
            For Each park In park_overlaps.AllKeys
                Dim values() As String = park_overlaps.GetValues(park)   ' get all values for this park
                Dim silolist As String = Join(values, ",")          ' make csv list
                parkSQLWriter.WriteLine($"UPDATE parks SET silos='{silolist }' WHERE WWFFID='{park }';") ' construct SQL
            Next
            siloSQLWriter.WriteLine("COMMIT;")
        End Using
        SetText(TextBox1, $"Done: {count }/{total } parks searched, {found } silos found")
    End Sub
    Public Function Dataset_BuildWhere(ds As String, ByVal GISID As String) As String
        ' Build a database WHERE clause for query
        ' GISID is a comma separated list of quoted ID's

        Contract.Requires(Not String.IsNullOrEmpty(GISID), "Bad GISID")
        Dim where As String = ""
        Select Case ds
            Case "CAPAD_T", "CAPAD_M", "ZL"
                where = $"{DatasetDict(ds)("shpIDField") } IN ({GISID })"
            Case "VIC_PARKS"
                where = $"{DatasetDict(ds)("shpIDField") } IN ({GISID.Replace("'", "") })"    ' remove quotes if any
        End Select
        Contract.Ensures(Not String.IsNullOrEmpty(where))
        Return where
    End Function
    Private Function TogoFormat(started As DateTime, count As Integer, total As Integer) As String
        ' Do a time to go time calculation, and return time in d h m s
        Dim rate As Long = DateDiff(DateInterval.Second, started, Now()) / count      ' seconds per item
        Dim togo As Long = (total - count) * rate                                      ' seconds to go
        Dim d As Integer = togo / (24 * 3600)   ' days
        togo = togo Mod (24 * 3600)
        Dim h As Integer = togo / 3600          ' hours
        togo = togo Mod 3600
        Dim m As Integer = togo / 60            ' minutes
        togo = togo Mod 60
        Dim s As Integer = togo                 ' seconds
        Dim finishIn As String
        If d > 0 Then
            finishIn = $"{d}d {h:00}h {m:00}m {s:00}s"
        ElseIf h > 0 Then
            finishIn = $"{h}h {m:00}m {s:00}s"
        ElseIf m > 0 Then
            finishIn = $"{m}m {s:00}s"
        Else
            finishIn = $"{s}s"
        End If
        Return finishIn
    End Function
    Private Sub ExtractALLMissingDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtractALLMissingDataToolStripMenuItem.Click
        ' Update all missing data
        TextBox2.Clear()
        Using sqlWriter As New System.IO.StreamWriter("siloUpdates.sql")
            UpdateLocality(sqlWriter)
            UpdateRailway(sqlWriter)
            UpdateMural(sqlWriter)
            UpdateLGA(sqlWriter)
            UpdateStreetView(sqlWriter)
        End Using
        TextBox2.AppendText($"Updates complete")
    End Sub

    Private Sub GuessLinkToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuessLinkToolStripMenuItem.Click
        ' Guess and test wiki link for town
        Using sqlWriter As New System.IO.StreamWriter("siloLocality.sql")
            UpdateLocality(sqlWriter)
        End Using
    End Sub
    Private Sub UpdateLocality(sqlWriter As StreamWriter)
        Dim count As Integer = 0
        Dim found As Integer = 0, notFound As Integer = 0
        Dim silo_code As String, columns As New Dictionary(Of String, Integer)
        Dim url As Uri, req As System.Net.HttpWebRequest, resp As System.Net.HttpWebResponse
        Dim now As String = DateTime.UtcNow.ToString("O")   ' UTC date/time in ISO 8601 format
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader, transaction As Boolean = False

        TextBox2.AppendText($"Creating Locality updates{vbCrLf }")
        Using connect As New SQLiteConnection(SILOSdb)

            sqlWriter.WriteLine("-- Creating Locality updates")
            connect.Open()  ' open database
            sqlcmd = connect.CreateCommand()
            sqlcmd.CommandText = $"SELECT * FROM silos WHERE not_before<='{now }' AND not_after='' AND link=''"     ' select active silos with no link
            SQLdr = sqlcmd.ExecuteReader()
            While SQLdr.Read()
                Try
                    count += 1
                    ' find locality in wikipedia
                    silo_code = SQLdr("silo_code")
                    Dim town As String = $"{SQLdr("locality") }, {SQLdr("state") }".Replace(" ", "_")
                    url = New Uri($"https://en.wikipedia.org/wiki/{town }")
                    ' test the url
                    req = System.Net.HttpWebRequest.Create(url)
                    req.Method = "HEAD"     ' don't need body
                    Try
                        resp = req.GetResponse()
                        resp.Close()
                        found += 1
                        If Not transaction Then
                            sqlWriter.WriteLine("START TRANSACTION;")
                            transaction = True
                        End If
                        sqlWriter.WriteLine($"UPDATE silos SET link='{url }' WHERE silo_code='{silo_code }';")
                    Catch ex As WebException
                        If ex.Message.Contains("404") Then
                            ' Page not found
                            ' sqlWriter.WriteLine($"-- no link found for {town }")
                            notFound += 1
                        Else
                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                        End If
                    End Try
                    req = Nothing
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
                TextBox1.Text = $"{found } found, {notFound } not found, {count } total"
                Application.DoEvents()
            End While
            If transaction Then sqlWriter.WriteLine("COMMIT;")
            sqlWriter.WriteLine($"-- {found } found, {count } total")
        End Using
        TextBox2.AppendText($"Done {found } found, {notFound } not found, {count } total{vbCrLf }")
    End Sub
    Private Sub ExtractRailwayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtractRailwayToolStripMenuItem.Click
        Using sqlWriter As New System.IO.StreamWriter("silosRailway.sql")
            UpdateRailway(sqlWriter)
        End Using
    End Sub
    Private Sub UpdateRailway(sqlWriter As StreamWriter)
        ' Find closest railway line for each silo
        Const CLOSENESS = 500     ' maximum distance from railway
        Const MAPSERVER = "https://services.ga.gov.au/gis/rest/services/NM_Transport_Infrastructure/MapServer/7/"    ' mapserver for railways
        Dim count As Integer = 0, found As Integer = 0, notFound As Integer = 0
        Dim route As String, status As String
        Dim resp As Byte(), responseStr As String, transaction As Boolean = False
        Dim now As String = DateTime.UtcNow.ToString("O")   ' UTC date/time in ISO 8601 format
        Dim silo_code As String, POSTfields As NameValueCollection
        Dim Jo As JObject, railways As New List(Of Railway)
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader

        TextBox2.AppendText($"Creating Railway updates{vbCrLf }")
        sqlWriter.WriteLine("-- Creating Railway updates")
        Using connect As New SQLiteConnection(SILOSdb)
            connect.Open()  ' open database
            sqlcmd = connect.CreateCommand()
            sqlcmd.CommandText = "SELECT * FROM silos WHERE railway='' AND not_before<='{now }' AND not_after='' "     ' select missing railways
            SQLdr = sqlcmd.ExecuteReader()
            While SQLdr.Read()
                count += 1
                ' find nearest railway
                If String.IsNullOrEmpty(SQLdr("railway")) Or String.IsNullOrEmpty(SQLdr("rail_status")) Then
                    silo_code = SQLdr("silo_code")
                    ' Prepare some POST fields for a request to the map server. We use POST because the requests are too large for a GET
                    POSTfields = Nothing        ' destroy existing values
                    POSTfields = New NameValueCollection From
                    {
                        {"f", "geojson"},
                        {"where", "featuretype='Railway'"},
                        {"GeometryType", "esriGeometryPoint"},
                        {"inSR", "WGS84"},
                        {"geometry", $"{SQLdr("lng") },{SQLdr("lat") }"},
                        {"spatialRel", "esriSpatialRelIntersects"},
                        {"outFields", "routename, status"},
                        {"returnGeometry", "false"},          '  need the geometry
                        {"returnIdsOnly", "false"},
                        {"returnCountOnly", "false"},
                        {"returnZ", "false"},
                        {"returnM", "false"},
                        {"returnDistinctValues", "false"},
                        {"returnExtentOnly", "false"},
                        {"distance", CLOSENESS},
                        {"units", "esriSRUnit_Meter"}
                    }
                    Using myWebClient As New WebClient
                        Try
                            resp = Array.Empty(Of Byte)()   ' clear  array
                            resp = myWebClient.UploadValues($"{MAPSERVER }query/", "POST", POSTfields)    ' query map server
                            responseStr = System.Text.Encoding.UTF8.GetString(resp)
                            Jo = JObject.Parse(responseStr)
                            If Jo.HasValues And Jo("features").Any Then
                                ' extract (possibly multiple) near railways
                                railways.Clear()
                                For Each feature In Jo("features")
                                    route = feature("properties")("routename").ToString
                                    status = feature("properties")("status").ToString
                                    If Not String.IsNullOrEmpty(route) Then railways.Add(New Railway(route, status))    ' save railway details
                                Next
                                If railways.Any Then
                                    found += 1
                                    railways.Sort() ' sort into importance order
                                    route = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(railways(0).Route.ToLower())   ' convert to Camel case
                                    'route = route.Replace(" Railway", "")
                                    'route = route.Replace(" Line", "")
                                    'route = route.Replace(" Branch", "")
                                    status = railways(0).Status
                                    If Not transaction Then
                                        sqlWriter.WriteLine("START TRANSACTION;")
                                        transaction = True
                                    End If
                                    sqlWriter.WriteLine($"UPDATE silos SET railway='{route }',rail_status='{status }' WHERE silo_code='{silo_code }';")
                                Else
                                    notFound += 1
                                    ' sqlWriter.WriteLine($"-- Railway name is blank for silo {silo_code }")
                                End If
                            Else
                                ' sqlWriter.WriteLine($"-- No railway found for silo {silo_code }")
                                notFound += 1
                            End If
                        Catch ex As WebException
                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOK, "Web request failed")
                        End Try
                    End Using
                End If
                TextBox1.Text = $"{found } found, {notFound } not found, {count } total"
                Application.DoEvents()
            End While
            If transaction Then sqlWriter.WriteLine("COMMIT;")
            sqlWriter.WriteLine($"-- {found } found, {notFound } not found, {count } total")
        End Using
        TextBox2.AppendText($"Done {found } found, {notFound } not found, {count } total{vbCrLf }")
    End Sub
    Private Sub GenerateMuralLinkToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateMuralLinkToolStripMenuItem.Click
        ' Generate link to mural site if one does not already exist
        Using sqlWriter As New System.IO.StreamWriter("siloMurals.sql")
            UpdateMural(sqlWriter)
        End Using
    End Sub
    Private Sub UpdateMural(sqlWriter As StreamWriter)
        Dim count As Integer = 0
        Dim found As Integer = 0
        Dim silo_code As String, columns As New Dictionary(Of String, Integer)
        Dim url As Uri, req As System.Net.HttpWebRequest, resp As System.Net.HttpWebResponse
        Dim now As String = DateTime.UtcNow.ToString("O")   ' UTC date/time in ISO 8601 format
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader, transaction As Boolean = False

        TextBox2.AppendText($"Creating Mural updates{vbCrLf }")
        Using connect As New SQLiteConnection(SILOSdb)
            sqlWriter.WriteLine("-- Creating Mural updates")
            connect.Open()  ' open database
            sqlcmd = connect.CreateCommand()
            sqlcmd.CommandText = $"SELECT * FROM silos WHERE not_before<='{now }' AND not_after='' AND arty='true' AND NOT comment LIKE '%**Mural:**%'"     ' select active silos with no mural
            SQLdr = sqlcmd.ExecuteReader()
            While SQLdr.Read()
                Try
                    count += 1
                    ' guess mural link
                    silo_code = SQLdr("silo_code")
                    Dim siloName As String = SQLdr("locality")
                    Dim locality As String = siloName.ToLower.Replace(" ", "-")
                    url = New Uri($"https://www.australiansiloarttrail.com/{locality }")
                    ' test the url
                    req = System.Net.HttpWebRequest.Create(url)
                    req.Method = "HEAD"     ' don't need body
                    req.UserAgent = "Mozilla/5.0"       ' required else 403 error
                    Try
                        resp = req.GetResponse()
                        resp.Close()
                        found += 1
                        If Not transaction Then
                            sqlWriter.WriteLine("START TRANSACTION;")
                            transaction = True
                        End If
                        sqlWriter.WriteLine($"UPDATE silos SET comment='**Mural:** [{siloName } silos]({url })' WHERE silo_code='{silo_code }';")
                    Catch ex As WebException
                        If ex.Message.Contains("403") Or ex.Message.Contains("404") Then
                            ' Page not found
                            ' sqlWriter.WriteLine($"-- no link found for {siloName }")
                        Else
                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                        End If
                    End Try
                    req = Nothing
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
                TextBox1.Text = $"{found } found, {count } total"
                Application.DoEvents()
            End While
            If transaction Then sqlWriter.WriteLine("COMMIT;")
            sqlWriter.WriteLine($"-- {found } found, {count } total")
        End Using
        TextBox2.AppendText($"Done {found } found, {count } total{vbCrLf }")
    End Sub
    Private Sub AddLGACodesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddLGACodesToolStripMenuItem.Click
        ' Add LGA code for each silo
        Using sqlWriter As New System.IO.StreamWriter("silosLGA.sql")
            UpdateLGA(sqlWriter)
        End Using
    End Sub

    Private Sub UpdateLGA(sqlWriter As StreamWriter)
        ' Produce updates for LGA
        Const GEO_SERVER = "https://geo.abs.gov.au/arcgis/rest/services/ASGS2021/LGA/MapServer"
        Dim count As Integer = 0, found As Integer = 0, notFound As Integer = 0
        Dim POSTfields As NameValueCollection
        Dim resp As Byte(), responseStr As String
        Dim silo_code As String, lga_code As Integer
        Dim now As String = DateTime.UtcNow.ToString("O")   ' UTC date/time in ISO 8601 format
        Dim Jo As JObject
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader, ShireSqlcmd As SQLiteCommand, ShireSqldr As SQLiteDataReader
        Dim ShireID As String, transaction As Boolean = False

        TextBox2.AppendText($"Creating LGA updates{vbCrLf }")
        Using connect As New SQLiteConnection(SILOSdb),
                parksConnect As New SQLiteConnection(PARKSdb),
                myWebClient As New WebClient

            ' get service metadata to find extent of data
            POSTfields = Nothing        ' destroy existing values
            POSTfields = New NameValueCollection From {
                    {"f", "json"}}         ' return json
            resp = Array.Empty(Of Byte)()   ' clear  array
            resp = myWebClient.UploadValues(GEO_SERVER, "POST", POSTfields)    ' query map server
            responseStr = System.Text.Encoding.UTF8.GetString(resp)
            Jo = JObject.Parse(responseStr)
            Dim sr = New SpatialReference(CInt(Jo.Item("spatialReference")("latestWkid"))) ' datum used for the extent
            Dim p1 As New MapPoint(Jo.Item("fullExtent")("xmin"), Jo.Item("fullExtent")("ymin"), sr)    ' one corner of extent
            Dim p2 As New MapPoint(Jo.Item("fullExtent")("xmax"), Jo.Item("fullExtent")("ymax"), sr)    ' one corner of extent
            Dim extent As New Envelope(p1, p2)          ' envelope for Australia
            extent = GeometryEngine.Project(extent, SpatialReferences.Wgs84)    ' the extent of Australia

            myWebClient.Headers.Add("accept", "text/html, Application / xhtml + Xml, Application / Xml;q=0.9, Image / avif, Image / webp, Image / apng,*/*;q=0.8, Application / signed - exchange;v=b3;q=0.9")
            sqlWriter.WriteLine("-- Creating LGA updates")
            connect.Open()
            sqlcmd = connect.CreateCommand()
            parksConnect.Open()
            ShireSqlcmd = parksConnect.CreateCommand()
            sqlcmd.CommandText = "Select * FROM silos where trim(lga_code) ='' AND not_before<='{now }' AND not_after=''"      ' select all silos with missing lga code
            SQLdr = sqlcmd.ExecuteReader()
            While SQLdr.Read()
                count += 1
                silo_code = SQLdr("silo_code")
                ' fetch LGA data
                POSTfields = Nothing        ' destroy existing values
                POSTfields = New NameValueCollection From {
                    {"f", "json"},          ' return json
                    {"geometryType", "esriGeometryPoint"},
                    {"sr", "4326"}, ' WGS84
                    {"geometry", $"{SQLdr("lng") }, {SQLdr("lat") }"},
                    {"spatialRel", "esriSpatialRelWithin"},
                    {"layers", "LGA"},
                    {"mapExtent", $"{extent.XMin},{extent.YMin},{extent.XMax},{extent.YMax}"},
                    {"imageDisplay", $"1000,1000,96"},
                    {"tolerance", 1},
                    {"returnGeometry", "false"},              ' don't need the geometry
                    {"outFields", "*"}
                }
                Try
                    resp = Array.Empty(Of Byte)()   ' clear  array
                    resp = myWebClient.UploadValues(GEO_SERVER & "/identify", "POST", POSTfields)    ' query map server
                    responseStr = System.Text.Encoding.UTF8.GetString(resp)
                    Jo = JObject.Parse(responseStr)
                    If Jo.HasValues And Jo.Item("results").Any Then
                        lga_code = Jo.Item("results")(0)("attributes")("LGA_CODE_2021")
                        ' Now convert lga_code into ShireID
                        ShireSqlcmd.CommandText = $"SELECT * FROM SHIRES where LGA={lga_code }"
                        ShireSqldr = ShireSqlcmd.ExecuteReader()
                        If ShireSqldr.Read() Then
                            ShireID = ShireSqldr("ShireID")
                            If Not transaction Then
                                sqlWriter.WriteLine("START TRANSACTION;")
                                transaction = True
                            End If
                            sqlWriter.WriteLine($"UPDATE silos SET lga_code='{ShireID }' WHERE silo_code='{silo_code }';")
                            found += 1
                        Else
                            sqlWriter.WriteLine($"-- No ShireID found for LGA_CODE {lga_code } on silo {silo_code }")
                            notFound += 1
                        End If
                        ShireSqldr.Close()
                    Else
                        sqlWriter.WriteLine($"-- No LGA code found for {silo_code }")
                        notFound += 1
                    End If
                Catch ex As WebException
                    MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOK, "Web request failed")
                End Try
                TextBox1.Text = $"Processed {notFound }/{found }, {count }"
                Application.DoEvents()
            End While
            If transaction Then sqlWriter.WriteLine("COMMIT;")
            sqlWriter.WriteLine($"-- {found } found, {count } total")
        End Using
        TextBox2.AppendText($"Done {found } found, {notFound } not found, {count } total{vbCrLf }")
    End Sub

    Private Sub CrossCheckSiloCodeWithLGAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CrossCheckSiloCodeWithLGAToolStripMenuItem.Click
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader, count As Integer = 0, RealState As String
        ' State codes used in SA4
        Dim states As New NameValueCollection From {
            {"1", "New South Wales"},
            {"2", "Victoria"},
            {"3", "Queensland"},
            {"4", "South Australia"},
            {"5", "Western Australia"},
            {"6", "Tasmania"},
            {"7", "Northern Territory"},
            {"8", "Australian Capital Territory"},
            {"9", "Other Territories"}
            }
        Using connect As New SQLiteConnection(SILOSdb), htmlWriter As New System.IO.StreamWriter("silosCodeLGA.html")
            connect.Open()
            ' Cross check last character of LGA code (known correct) with last character of silo_code (suspect)
            sqlcmd = connect.CreateCommand()
            sqlcmd.CommandText = "Select * From silos Where substr(lga_code, 3, 1) <> substr(silo_code, 7, 1) And not_before<='{now }' AND not_after='' order by silo_code"      ' select all silos with missing lga code
            SQLdr = sqlcmd.ExecuteReader()
            htmlWriter.WriteLine("Cross check of silo_code and lga_code<br>")
            htmlWriter.WriteLine("<table border=1>")
            htmlWriter.WriteLine($"<tr><th>Silo code</th><th>Name</th><th>State</th><th>LGA code</th></tr>")
            While SQLdr.Read()
                count += 1
                htmlWriter.WriteLine($"<tr><td>{SQLdr("silo_code") }</td><td>{SQLdr("name") }</td><td>{SQLdr("state") }</td><td>{SQLdr("lga_code") }</td></tr>")
                htmlWriter.Flush()
            End While
            SQLdr.Close()
            htmlWriter.WriteLine("</table>")

            ' Now check locality
            htmlWriter.WriteLine("<br>Cross check of locality and link<br>")
            htmlWriter.WriteLine("<table border=1>")
            htmlWriter.WriteLine($"<tr><th>Silo code</th><th>Name</th><th>Town</th><th>Link</th></tr>")
            sqlcmd.CommandText = $"SELECT * FROM silos WHERE not_before<='{Now }' AND not_after='' AND link!='' order by silo_code"     ' select active silos with link
            SQLdr = sqlcmd.ExecuteReader()
            While SQLdr.Read()
                ' find locality in wikipedia
                Dim town As String = $"{SQLdr("locality") }, {SQLdr("state") }".Replace(" ", "_")
                If Not SQLdr("link").EndsWith(town) Then
                    htmlWriter.WriteLine($"<tr><td>{SQLdr("silo_code") }</td><td>{SQLdr("name") }</td><td>{town }</td><td>{SQLdr("link") }</td></tr>")
                    htmlWriter.Flush()
                    count += 1
                End If
            End While
            SQLdr.Close()
            htmlWriter.WriteLine("</table>")

            ' Check state
            htmlWriter.WriteLine("<br>Cross check of lat/lon and state<br>")
            htmlWriter.WriteLine("<table border=1>")
            htmlWriter.WriteLine($"<tr><th>Silo code</th><th>Name</th><th>State</th><th>Actual state</th></tr>")
            sqlcmd.CommandText = $"SELECT * FROM silos WHERE not_before<='{Now }' AND not_after='' order by silo_code"     ' select active silos with link
            SQLdr = sqlcmd.ExecuteReader()
            Using MyWebClient As New WebClient
                While SQLdr.Read()
                    ' Prepare some POST fields for a request to the map server. We use POST because the requests are too large for a GET
                    Dim POSTfields = New NameValueCollection From {
                                                        {"f", "json"},
                                                        {"geometryType", "esriGeometryPoint"},
                                                        {"geometry", $"{SQLdr("lng"):f5},{SQLdr("lat"):f5}"},
                                                        {"inSR", "4326"},
                                                        {"spatialRel", "esriSpatialRelIntersects"},
                                                        {"returnGeometry", "false"},
                                                        {"outFields", "*"}
                                                    }
                    ' return all fields, even though we only use the default
                    Dim resp As Byte() = MyWebClient.UploadValues(GEOSERVER_SA4, "POST", POSTfields)
                    Dim responseStr = System.Text.Encoding.UTF8.GetString(resp)
                    Dim Jo As JObject = JObject.Parse(responseStr)
                    If Jo.HasValues And Jo("features").Any Then
                        Dim StateCode As String = Jo("features")(0)("attributes")("STATE_CODE_2016")
                        RealState = states(StateCode)
                        If RealState <> SQLdr("state") Then
                            count += 1
                            htmlWriter.WriteLine($"<tr><td>{SQLdr("silo_code") }</td><td>{SQLdr("name") }</td><td>{SQLdr("state") }</td><td>{RealState }</td></tr>")
                            htmlWriter.Flush()
                        End If
                    Else
                        MsgBox($"Could not find state for silo {SQLdr("silo_code") } at {SQLdr("lat"):f5},{SQLdr("lng"):f5}", vbCritical + vbOKOnly, "Can't resolve state")
                    End If
                    POSTfields = Nothing
                End While
            End Using
            htmlWriter.WriteLine("</table>")
            TextBox1.Text = $"Done. {count } discrepancies found"
        End Using
    End Sub
    Private Sub NearestRailwayStationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NearestRailwayStationToolStripMenuItem.Click
        ' Compare silo name with Railway Station name within 1km
        Const MAPSERVER = "https://services.ga.gov.au/gis/rest/services/NM_Transport_Infrastructure/MapServer/4/query"
        Dim count As Integer = 0, found As Integer = 0
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader
        Dim myTI As TextInfo = New CultureInfo("en-US", False).TextInfo

        Using myWebClient As New WebClient,
        connect As New SQLiteConnection(SILOSdb),
        htmlWriter As New System.IO.StreamWriter("silosClosestStation.html")
            htmlWriter.WriteLine("<br>Stations within 1km of silo<br>")
            htmlWriter.WriteLine("<table border=1>")
            htmlWriter.WriteLine($"<tr><th>Silo code</th><th>Name</th><th>Station</th></tr>")
            connect.Open()
            sqlcmd = connect.CreateCommand()
            sqlcmd.CommandText = "Select * From silos Where not_before<='{now }' AND not_after='' order by silo_code"      ' select all silos with missing lga code
            SQLdr = sqlcmd.ExecuteReader()
            While SQLdr.Read()
                Try
                    count += 1
                    Dim silo As New MapPoint(SQLdr("lng"), SQLdr("lat"), SpatialReferences.Wgs84)
                    Dim buffer As Geometry = GeometryEngine.BufferGeodetic(silo, SILO_ACTIVATION_ZONE, LinearUnits.Meters)              ' silo location with buffer
                    Dim POSTfields = New NameValueCollection From {
                                                            {"f", "geojson"},
                                                            {"geometryType", "esriGeometryPolygon"},
                                                            {"geometry", buffer.ToJson},
                                                            {"inSR", "4326"},
                                                            {"spatialRel", "esriSpatialRelIntersects"},
                                                            {"returnGeometry", "false"},
                                                            {"outFields", "*"}
                                                        }
                    Dim resp As Byte() = myWebClient.UploadValues(MAPSERVER, "POST", POSTfields)
                    Dim responseStr = System.Text.Encoding.UTF8.GetString(resp)
                    Dim Jo As JObject = JObject.Parse(responseStr)
                    If Jo.HasValues And Jo("features").Any Then
                        Dim Station As String = Trim(Jo("features")(0)("properties")("name"))
                        Station = myTI.ToTitleCase(LCase(Station))  ' ToTitleCase requires lower case input !!!
                        If Station <> SQLdr("name") Then
                            found += 1
                            htmlWriter.WriteLine($"<tr><td>{SQLdr("silo_code") }</td><td>{SQLdr("name") }</td><td>{Station }</td></tr>")
                            htmlWriter.Flush()
                        End If
                    End If
                    TextBox1.Text = $"Checked {count }: found {found }"
                    Application.DoEvents()
                Catch ex As WebException
                    MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                End Try
            End While
            htmlWriter.WriteLine("</table>")
            TextBox1.Text = $"Done. {found } stations found"
        End Using
    End Sub
    Private Sub UpdateStreetViewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateStreetViewToolStripMenuItem.Click
        Using sqlWriter As New System.IO.StreamWriter("StreetView.sql")
            UpdateStreetView(sqlWriter)
        End Using
    End Sub

    Sub UpdateStreetView(sqlwriter As StreamWriter)
        ' generate Street View url
        ' metadata doco  https://developers.google.com/maps/documentation/streetview/overview#url_parameters
        Const MAPSERVER = "https://maps.googleapis.com/maps/api/streetview/metadata"
        Const SEARCH_RADIUS = 400       ' Street View search radius
        Const SILO_WIDTH = 100        ' Assumed real width of silo
        Const SILO_HEIGHT = 40        ' Assumed real height of silo
        Dim count As Integer = 0, found As Integer = 0, NotFound As Integer = 0
        Dim sqlcmd As SQLiteCommand, SQLdr As SQLiteDataReader
        Dim ViewPoint As MapPoint, SiloPoint As MapPoint, bearing As GeodeticDistanceResult, dist As Integer, fov As Integer, pitch As Integer, StreetViewURL As String
        Dim transaction As Boolean = False

        TextBox2.AppendText($"Creating Street View updates{vbCrLf }")
        Using myWebClient As New WebClient,
        connect As New SQLiteConnection(SILOSdb),
                htmlWriter As New System.IO.StreamWriter("StreetView.html")
            sqlwriter.WriteLine("-- Creating Street View updates")
            htmlWriter.WriteLine("<table border=1><tr><th>Code</th><th>Name/Link</th><th>Distance (m)</th></tr>")
            connect.Open()
            sqlcmd = connect.CreateCommand()
            'sqlcmd.CommandText = "Select * from silos where not_before<='{now }' AND not_after='' AND street_view = '' order by silo_code"      ' select all silos with missing Street View
            sqlcmd.CommandText = "Select * from silos where not_before<='{now }' AND not_after=''  AND street_view = '' order by silo_code"      ' select all silos
            SQLdr = sqlcmd.ExecuteReader()
            While SQLdr.Read()
                count += 1
                ' manually inserted URLs contain fov/calc parameter. Skip them
                If Not (SQLdr("street_view").ToString.Contains("fov=") And Not SQLdr("street_view").ToString.Contains("&calc")) Then
                    Try
                        StreetViewURL = ""
                        dist = 0
                        Dim GETfields As New Dictionary(Of String, String) From {
                            {"location", $"{SQLdr("lat"):f6},{SQLdr("lng"):f6}"},
                            {"radius", SEARCH_RADIUS},
                            {"key", GOOGLE_STREET_VIEW_KEY}
                        }
                        If (SQLdr("arty") = "false") Then GETfields.Add("source", "outdoor")
                        Dim url As String = QueryHelpers.AddQueryString(MAPSERVER, GETfields)
                        'Dim resp As Byte() = myWebClient.UploadValues(MAPSERVER, "GET", GETfields)     ' this should work, but doesn't
                        Dim resp As Byte() = myWebClient.DownloadData(url)
                        Dim responseStr = System.Text.Encoding.UTF8.GetString(resp)
                        Dim Jo As JObject = JObject.Parse(responseStr)
                        If Jo.HasValues And Jo("status") IsNot Nothing Then
                            If Jo("status") = "OK" Then
                                found += 1
                                SiloPoint = New MapPoint(SQLdr("lng"), SQLdr("lat"), SpatialReferences.Wgs84)     ' location of silo
                                ViewPoint = New MapPoint(Jo("location")("lng"), Jo("location")("lat"), SpatialReferences.Wgs84)    ' point of street view
                                bearing = GeometryEngine.DistanceGeodetic(ViewPoint, SiloPoint, LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic) ' calculate bearing to silo
                                dist = bearing.Distance ' extract distance to silo
                                pitch = CInt(RadtoDeg(Atan(SILO_HEIGHT / (2 * dist))))               ' aim at point half way up silo
                                fov = Min(120, CInt(RadtoDeg(2 * Atan(SILO_WIDTH / (2 * dist)))))    ' field of view calculation
                                StreetViewURL = $"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={ViewPoint.Y:f6},{ViewPoint.X:f6}&heading={bearing.Azimuth1:f0}&pitch={pitch }&fov={fov }&calc"   ' construct a URL
                                If Not transaction Then
                                    sqlwriter.WriteLine("START TRANSACTION;")
                                    transaction = True
                                End If
                                sqlwriter.WriteLine($"UPDATE silos SET Street_View='{StreetViewURL }' WHERE silo_code='{SQLdr("silo_code") }';")
                            Else
                                ' No Street view here
                                NotFound += 1
                            End If
                            htmlWriter.Write($"<tr><td>{SQLdr("silo_code") }</td>")
                            If StreetViewURL = "" Then
                                htmlWriter.Write($"<td>{SQLdr("name") }</td>")
                            Else
                                htmlWriter.Write($"<td><a href='{StreetViewURL }'>{SQLdr("name") }</a></td>")
                            End If
                            htmlWriter.WriteLine($"<td>{dist }</td></tr>")
                        End If
                    Catch ex As WebException
                        MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                    End Try
                End If
                TextBox1.Text = $"processed {count}"
                htmlWriter.Flush()
                sqlwriter.Flush()
                Application.DoEvents()
            End While
            htmlWriter.WriteLine("</table>")
            If transaction Then sqlwriter.WriteLine("COMMIT;")
            sqlwriter.WriteLine($"-- {found } found, {count } total")
        End Using
        TextBox2.AppendText($"Done {found } found, {NotFound } not found, {count } total{vbCrLf }")
    End Sub
    Function DegtoRad(deg As Single) As Single
        ' Convert degrees to radians
        Return deg * PI / 180
    End Function
    Function RadtoDeg(rad As Single) As Single
        ' Convert radians to degrees
        Return rad * 180 / PI
    End Function
    Private Sub DecodeStreetViewURLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DecodeStreetViewURLToolStripMenuItem.Click
        ' decode, and recode, a street view URL
        StreetView.ShowDialog()
    End Sub
    '======================================================================================
    ' BIG menu
    '======================================================================================
    Private Sub GetLatlonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetLatlonToolStripMenuItem.Click
        ' calculate any missing lat/lon for big items
        Dim currentRow As String(), count As Integer = 0, found As Integer = 0
        Dim columns As New Dictionary(Of String, Integer)
        Dim url As Uri
        Dim req As System.Net.HttpWebRequest
        Dim resp As System.Net.HttpWebResponse
        Dim Jo As JObject
        Dim lat As Double, lng As Double

        Using csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("big.csv", System.Text.Encoding.ASCII),
    csvWriter As New System.IO.StreamWriter("bigUpdated.csv")
            csvReader.TextFieldType = FileIO.FieldType.Delimited
            csvReader.SetDelimiters(",")
            csvReader.TrimWhiteSpace = True
            While Not csvReader.EndOfData
                Try
                    count += 1
                    currentRow = csvReader.ReadFields()
                    If count = 1 Then
                        ' make list of column headers and their offset
                        columns.Clear()
                        For i = LBound(currentRow) To UBound(currentRow)
                            columns.Add(currentRow(i).ToLower, i)
                        Next
                        csvWriter.WriteLine(Join(currentRow, ","))      ' write the header line
                    Else
                        ' find item in google Maps
                        Dim address As String = $"{currentRow(columns("name")) }, {currentRow(columns("state")) }"
                        Dim bounds As String = "-43.6345972634,113.338953078|-10.6681857235,153.569469029"  ' bounding box for australia
                        url = New Uri($"https://maps.googleapis.com/maps/api/geocode/json?address={address }&bounds={bounds }&key={GOOGLE_API_KEY }")
                        ' test the url
                        req = System.Net.HttpWebRequest.Create(url)
                        Try
                            resp = req.GetResponse()
                            Dim responseString As String = New StreamReader(resp.GetResponseStream()).ReadToEnd()
                            resp.Dispose()
                            Jo = JObject.Parse(responseString)
                            Select Case (Jo("status"))
                                Case "OK"
                                    With Jo("results")(0)("geometry")("location")
                                        lat = CDbl(Jo("results")(0)("geometry")("location")("lat"))
                                        currentRow(columns("lat")) = $"{lat:f6}"
                                        lng = CDbl(Jo("results")(0)("geometry")("location")("lng"))
                                        currentRow(columns("lon")) = $"{lng:f6}"
                                    End With
                                    found += 1
                                Case "ZERO_RESULTS"
                                Case Else
                                    MsgBox(Jo("status"), vbAbort + vbOKOnly, "Google Maps API error")
                                    Stop
                            End Select
                        Catch ex As WebException
                            MsgBox(ex.Message & vbCrLf & ex.StackTrace, vbCritical + vbOKOnly, "Exception")
                        End Try
                        req = Nothing
                        ' ready for csv output
                        For i = 0 To UBound(currentRow)
                            currentRow(i) = Csv(currentRow(i))
                        Next
                        csvWriter.WriteLine(Join(currentRow, ","))      ' write the data line
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
                SetText(TextBox1, $"found {found }, processed {count }")
            End While
        End Using
        SetText(TextBox1, $"Done. found {found }, processed {count }")
    End Sub

    Private Sub ExtractFromPositieMapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtractFromPositieMapToolStripMenuItem.Click
        ' extract big things from postie's map of big things
        Dim doc As New XmlDocument, placemarks As XmlNodeList
        Dim name As String, state As String, town As String, count As Integer = 0, notes As String
        Dim coords As String, lat As Double, lon As Double, type As String

        doc.Load("Big Things of Australia.kml")  ' open the data
        Dim nsmgr As New XmlNamespaceManager(doc.NameTable)     ' create namespace manager
        nsmgr.AddNamespace("kml", "http://www.opengis.net/kml/2.2")
        placemarks = doc.SelectNodes("//kml:Placemark", nsmgr)   ' select all placemarks
        Using csvWriter As New System.IO.StreamWriter("BigThings.csv")
            csvWriter.WriteLine("Name,Location,State,lat,lon,Built,Size,Notes,Image")   ' header
            For Each place As XmlNode In placemarks
                count += 1
                type = Csv(place.SelectSingleNode("kml:ExtendedData/kml:Data[@name='Type']/kml:value", nsmgr).InnerText.Trim)
                If type <> "Real Thing" And type <> "Painted Silo" Then
                    ' Ignore real things and painted silos
                    name = Csv(place.SelectSingleNode("kml:name", nsmgr).InnerText.Trim)
                    If Not name.StartsWith("Big") Then name = $"Big {name }"   ' force to start with "Big"
                    town = Csv(place.SelectSingleNode("kml:ExtendedData/kml:Data[@name='Town']/kml:value", nsmgr).InnerText.Trim)
                    state = place.SelectSingleNode("kml:ExtendedData/kml:Data[@name='State']/kml:value", nsmgr).InnerText.Trim
                    coords = place.SelectSingleNode("kml:Point/kml:coordinates", nsmgr).InnerText.Trim
                    notes = Csv(place.SelectSingleNode("kml:ExtendedData/kml:Data[@name='description']/kml:value", nsmgr).InnerText.Trim)
                    Dim s As String() = coords.Split(",")
                    lon = CDbl(s(0))
                    lat = CDbl(s(1))
                    SetText(TextBox1, $"{name }, {town }, {state } count={count }")
                    csvWriter.WriteLine($"{name },{town },{state },{lat:f5},{lon:f5},,,{notes },")
                End If
            Next
        End Using
    End Sub

    Shared Function Csv(st As String) As String
        ' escape, and enclose in double quotes, a string suitable for csv files
        Dim result As String
        Contract.Requires(st IsNot Nothing, "String is null")
        result = st
        If st.Contains(",") Or st.Contains("""") Or st.Contains(vbCr) Or st.Contains(vbLf) Then
            result = result.Replace("""", """""")   ' escape double quotes  
            result = $"""{result }"""      ' enclose in double quotes
        End If
        Return result
    End Function

    Private Async Sub ExtractSuburbsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtractSuburbsToolStripMenuItem.Click
        ' Extract list of suburbs and calculate centroid
        ' Open the Shape File
        Dim ShapeFile As ShapefileFeatureTable, myQueryFilter As New QueryParameters, suburbs As FeatureQueryResult, count As Integer, name As String, state_code As Integer
        Dim states() As String = {"Illegal", "NSW", "VIC", "QLD", "SA", "WA", "TAS", "NT", "ACT", "Other"}
        ShapeFile = Await ShapefileFeatureTable.OpenAsync("F:\GIS Data\State Suburbs\SSC_2016_AUST.shp").ConfigureAwait(False)
        With myQueryFilter
            .WhereClause = "1=1"    ' query parameters
            .OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
            .ReturnGeometry = True
        End With
        suburbs = Await ShapeFile.QueryFeaturesAsync(myQueryFilter)          ' run query
        Using logWriter As New System.IO.StreamWriter("suburbs.csv", False)
            logWriter.WriteLine("name,state,longitude,latitude")
            count = 0
            ' There seems to be a bug iterating a collection after an Await Async. Can't use "for each", so have to fetch each element explicitly
            ' Very much slower
            For Each suburb As Feature In suburbs
                count += 1
                If Not suburb.Geometry.IsEmpty Then
                    name = suburb.Attributes("SSC_NAME16").ToString
                    name = Regex.Replace(name, " \(.*\)$", "")      ' remove anything in brackets at end of name
                    state_code = CInt(suburb.Attributes("STE_CODE16").ToString)
                    Dim center As MapPoint = suburb.Geometry.Extent.GetCenter       ' get center of geometry
                    logWriter.WriteLine($"{Csv(name)},{states(state_code)},{center.X:0.####},{center.Y:0.####}")
                End If
                SetText(TextBox1, $"processed {count }")
            Next
            SetText(TextBox1, $"Done. processed {count }")
        End Using
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub
End Class

Class Railway : Implements IComparable(Of Railway)
    ' Capture route details
    Public Property Route As String
    Public Property Status As String
    Public Sub New(route As String, status As String)
        _Route = route
        _Status = status
    End Sub
    Public Function CompareTo(other As Railway) As Integer Implements IComparable(Of Railway).CompareTo
        Dim RailwayTypes As New List(Of String) From {"RAILWAY", "LINE", "BRANCH"}, words As String()
        Dim thisIndex As Integer, otherIndex As Integer

        If Me.Status = other.Status Then
            ' status the same - result depends on route name
            words = Me.Route.Split(" ")    ' split this route into words
            Dim RailwayType = words(words.Length - 1)
            thisIndex = RailwayTypes.IndexOf(RailwayType)
            If thisIndex = -1 Then thisIndex = RailwayTypes.Count     ' last in list
            words = other.Route.Split(" ")    ' split this route into words
            RailwayType = words(words.Length - 1)
            otherIndex = RailwayTypes.IndexOf(RailwayType)
            If otherIndex = -1 Then otherIndex = RailwayTypes.Count     ' last in list
            Return thisIndex.CompareTo(otherIndex)
        Else
            ' Operational beats Abandoned
            Return -1 * Status.CompareTo(other.Status)   ' exploit fact that Abandoned/Operational are in reverse alpha order
        End If
    End Function
End Class
Class SiloData
    Public Property Data As String()   ' all data fields from csv
    Public Property Distance As Double   ' calculated distance to all other points
    Public Property Consolidated As Boolean
    Sub New(d As String())
        Data = d
        Consolidated = False
    End Sub
    Sub GeoDistance(other As SiloData)
        ' calculate distance to other silo, but other if I'm not already consolidated
        Dim gdr As GeodeticDistanceResult
        gdr = GeometryEngine.DistanceGeodetic(New MapPoint(CDbl(Me.Data(1)), CDbl(Me.Data(2)), SpatialReferences.Wgs84), New MapPoint(CDbl(other.Data(1)), CDbl(other.Data(2)), SpatialReferences.Wgs84), LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic)
        Me.Distance = gdr.Distance
    End Sub
End Class

Class SiloDataIsol
    Public Property Distance As Double   ' calculated distance to all other points
    Public Property Lon As Double   ' longitude
    Public Property Lat As Double   ' latitude
    Public Property Locality As String   ' name
    Public Property State As String
    Sub New(locality As String, lon As Double, lat As Double, state As String)
        _Locality = locality
        _Lon = lon
        _Lat = lat
        _State = state
    End Sub
    Sub GeoDistance(other As SiloDataIsol)
        ' calculate distance to other silo, but other if I'm not already consolidated
        Dim gdr As GeodeticDistanceResult
        gdr = GeometryEngine.DistanceGeodetic(New MapPoint(CDbl(Me.Lon), CDbl(Me.Lat), SpatialReferences.Wgs84), New MapPoint(CDbl(other.Lon), CDbl(other.Lat), SpatialReferences.Wgs84), LinearUnits.Kilometers, AngularUnits.Degrees, GeodeticCurveType.Geodesic)
        Me._Distance = gdr.Distance
    End Sub
    Function Clone() As SiloDataIsol
        ' Create a clone of this object
        Dim s As New SiloDataIsol(Locality, Lon, Lat, State) With {
            .Distance = Distance
        }
        Return s
    End Function
End Class

Public Class RadioQuietZone
    ' class representing a radio quiet zone
    Public ReadOnly Property Location As String
    Public ReadOnly Property Name As String      ' Friendly name of facility
    Public ReadOnly Property Position As MapPoint    ' location
    Public ReadOnly Property Radius As Single        ' radius of quiet zone in m 
    Sub New(location As String, name As String, position As MapPoint, radius As Single)
        If String.IsNullOrEmpty(location) Then
            Throw New ArgumentException($"'{NameOf(location)}' cannot be null or empty", NameOf(location))
        End If

        If String.IsNullOrEmpty(name) Then
            Throw New ArgumentException($"'{NameOf(name)}' cannot be null or empty", NameOf(name))
        End If

        _Location = location
        _Name = name
        _Position = GeometryEngine.Project(position, SpatialReferences.Wgs84)
        _Radius = radius
    End Sub
End Class

' User Defined Function "REGEXP" for SQLite
<SQLiteFunction(Name:="REGEXP", Arguments:=2, FuncType:=FunctionType.Scalar)>
Public Class Regexp : Inherits SQLiteFunction
    Public Overrides Function Invoke(ByVal args() As Object) As Object
        If (args Is Nothing) OrElse (args.Length <> 2) Then Return Nothing  ' something wrong with call parameters
        Dim rg As New Regex(args(0).ToString)   ' pattern
        Return rg.IsMatch(args(1).ToString)     ' string to match
    End Function

End Class
Public Module Extensions

    <Extension()>
    Public Function AggrExtent(ByVal fragments As IEnumerable(Of Feature)) As Envelope
        ' Calculate extent of FeatureQueryResult
        Dim extent As Envelope
        If fragments.Count = 0 Then
            Throw New InvalidOperationException("Cannot compute aggregate for an empty set.")
        End If
        extent = fragments.First.Geometry.Extent
        For i = 1 To fragments.Count - 1
            extent = GeometryEngine.CombineExtents(extent, fragments(i).Geometry)
        Next
        Return extent
    End Function

    <Extension()>
    Public Function AggrExtent(Of T)(ByVal fragments As IEnumerable(Of T),
                      ByVal selector As Func(Of T, Feature)) As Envelope
        Return (From element In fragments Select selector(element)).AggrExtent()
    End Function
End Module
