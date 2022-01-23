Imports System.Diagnostics.Contracts
Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry
Public Class DataSet
    ' Encapsulation of a DataSet
    ''' <summary>Name of this data set type</summary>
    Public ReadOnly Property Name As String
    ''' <summary>Name of field containing area. Applies to both shp and dbf</summary>
    Public ReadOnly Property AreaField As String
    ''' <summary>Path of shapefile</summary>
    Public ReadOnly Property shpFileName As String
    Public Property shpShapeFileTable As ShapefileFeatureTable
    ''' <summary>KML style</summary>

    Public ReadOnly Property shpStyle As String
    ''' <summary>Collection of Features</summary>
    Public Property shpFragments As FeatureQueryResult
    ''' <summary>Copyright message for shapefile data</summary>
    Public ReadOnly Property shpCopyright As String
    ''' <summary>Name of field containing park ID</summary>
    Public ReadOnly Property shpIDField As String
    ''' <summary>Name of field containing park name</summary>
    Public ReadOnly Property shpNameField As String
    ''' <summary>Name of field containing park type</summary>
    Public ReadOnly Property shpTypeField As String
    Public ReadOnly Property shpAuthorityField As String
    Public ReadOnly Property dbfConnection As System.Data.OleDb.OleDbConnection ' dbf database connection
    Public ReadOnly Property dbfTableName As String                    ' name of table in dbf database
    ''' <summary>Scale divisor for hectares</summary>
    Public ReadOnly Property dbfAreaScale As Integer
    Public ReadOnly Property HasXYmetaData As Boolean                 ' true if shapefile has X & Y metadata

    ''' <summary>Create a new dataset</summary>
    ''' <param name="Name">The name of dataset</param>
    ''' <param name="FileName">The name of the shapefile</param>
    Public Sub New(ByVal Name As String, ByVal FileName As String)
        _Name = Name
        _shpFileName = FileName
        ' open shapefile
        OpenShapeFile()

        Select Case _Name
            Case "CAPAD_T"
                _shpStyle = "terrestrial"
                _shpCopyright = "Collaborative Australian Protected Areas Database (CAPAD) 2020, Commonwealth of Australia 2021"
                _shpAuthorityField = "AUTHORITY"
                _dbfTableName = _shpShapeFileTable.TableName
                _AreaField = "GIS_AREA"
                _shpNameField = "NAME"
                _shpIDField = "PA_ID"
                _shpTypeField = "TYPE_ABBR"
                _dbfAreaScale = 1
                _HasXYmetaData = True
            Case "CAPAD_M"
                _shpStyle = "marine"
                _shpCopyright = "Collaborative Australian Protected Areas Database (CAPAD) 2020, Commonwealth of Australia 2021"
                _shpAuthorityField = "AUTHORITY"
                _dbfTableName = _shpShapeFileTable.TableName
                _AreaField = "GIS_AREA"
                _shpNameField = "NAME"
                _shpIDField = "PA_ID"
                _shpTypeField = "TYPE_ABBR"
                _dbfAreaScale = 1
                _HasXYmetaData = True
            Case "VIC_PARKS"
                _shpStyle = "terrestrial"
                _shpCopyright = "Victorian Department of Environment, Land, Water and Planning (DELWP)"
                _shpAuthorityField = "MANAGER"
                _dbfTableName = "parkres"
                _AreaField = "HECTARES"
                _shpNameField = "NAME"
                _shpIDField = "PRIMS_ID"
                _shpTypeField = "AREA_TYPE"
                _dbfAreaScale = 1
                _HasXYmetaData = False
            Case "ZL"
                _shpStyle = "terrestrial"
                _shpCopyright = "New Zealand Department of Conservation"
                _shpAuthorityField = "Legislatio"
                _dbfTableName = "DOC_PublicConservationAreas_2017_06_01"
                _AreaField = "Shape_Area"
                _shpNameField = "Name"
                _shpIDField = "Conservati"
                _shpTypeField = "Type"
                _dbfAreaScale = 10000
                _HasXYmetaData = False
            Case Else
                Dim unused = MessageBox.Show($"Unrecognised DataSet: {_Name }", "Unrecognised DataSet name", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Stop
        End Select

        ' open dBase IV file
        Dim BaseConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Mode=Read;Extended Properties=dBase IV"
        'Dim BaseConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Mode=Read;Extended Properties=""dBase IV"""
        Dim ConnectionString = $"{BaseConnectionString };Data Source={IO.Path.GetDirectoryName(_shpFileName) };"
        _dbfConnection = New System.Data.OleDb.OleDbConnection(ConnectionString)
        _dbfConnection.Open()
        'Dim userTables As DataTable = _dbfConnection.GetSchema("Tables")
    End Sub

    Private Async Sub OpenShapeFile()
        ' Open the Shape File
        _shpShapeFileTable = Await ShapefileFeatureTable.OpenAsync(_shpFileName).ConfigureAwait(False)
    End Sub

    Public Function BuildWhere(ByVal GISID As String) As String
        ' Build a database WHERE clause for query
        ' GISID is a comma separated list of quoted ID's

        Contract.Requires(Not String.IsNullOrEmpty(GISID), "Bad GISID")
        Dim where As String = ""
        Select Case _Name
            Case "CAPAD_T", "CAPAD_M", "ZL"
                where = $"{_shpIDField } IN ({GISID })"
            Case "VIC_PARKS"
                where = $"{_shpIDField } IN ({GISID.Replace("'", "") })"    ' remove quotes if any
        End Select
        Contract.Ensures(Not String.IsNullOrEmpty(where))
        Return where
    End Function

    ''' <summary>Calculate centroid of collection of fragments. Result is WGS84</summary>
    Public ReadOnly Property Centroid As Esri.ArcGISRuntime.Geometry.MapPoint

        Get
            Contract.Requires(_shpFragments IsNot Nothing And _shpFragments.Any, "No fragment collection")
            Dim X_COORD As Double = 0, Y_COORD As Double = 0
            Dim env As Envelope = Nothing, result As MapPoint, point As MapPoint

            Select Case _Name
                Case "CAPAD_T", "CAPAD_M"
                    For Each frag As Feature In _shpFragments
                        X_COORD += CDbl(frag.GetAttributeValue("LONGITUDE"))
                        Y_COORD += CDbl(frag.GetAttributeValue("LATITUDE"))
                    Next
                    X_COORD /= _shpFragments.Count
                    Y_COORD /= _shpFragments.Count
                Case "VIC_PARKS"
                    For Each frag As Feature In _shpFragments
                        If frag.Geometry IsNot Nothing Then
                            env = Form1.EnvelopeUnion(env, frag.Geometry.Extent)
                        End If
                    Next
                    ' get center of envelope
                    X_COORD = env.GetCenter.X
                    Y_COORD = env.GetCenter.Y
                Case "ZL"
                    For Each frag As Feature In _shpFragments
                        env = Form1.EnvelopeUnion(env, frag.Geometry.Extent)
                    Next
                    ' get center of envelope
                    X_COORD = env.GetCenter.X
                    Y_COORD = env.GetCenter.Y
            End Select
            Contract.Ensures(X_COORD <> 0 And Y_COORD <> 0)
            point = New MapPoint(X_COORD, Y_COORD, _shpFragments.SpatialReference)   ' create the point
            result = GeometryEngine.Project(point, SpatialReferences.Wgs84)     ' convert to wgs84
            Return result
        End Get
    End Property

    ''' <summary>Calculate area of collection of fragments in hectares</summary>
    Public ReadOnly Property Area As Double

        Get
            Contract.Requires(_shpFragments IsNot Nothing And _shpFragments.Any, "No fragment collection")
            Dim GIS_AREA As Double = 0

            Select Case _Name
                Case "CAPAD_T", "CAPAD_M"
                    For Each frag As Feature In _shpFragments
                        GIS_AREA += CDbl(frag.GetAttributeValue("GIS_AREA"))
                    Next
                Case "VIC_PARKS"
                    For Each frag As Feature In _shpFragments
                        GIS_AREA += CDbl(frag.GetAttributeValue("TOTAL_AREA"))
                    Next
                Case "ZL"
                    GIS_AREA = _shpFragments(0).GetAttributeValue("Shape_Area").ToString / 10000
            End Select
            Contract.Ensures(GIS_AREA <> 0, "0 area")
            Return GIS_AREA
        End Get
    End Property

    ''' <summary>Form name of park</summary>
    Public ReadOnly Property ParkName As String

        Get
            Contract.Requires(_shpFragments IsNot Nothing And _shpFragments.Any, "No fragment collection")

            Return _shpFragments(0).GetAttributeValue(_shpNameField)
        End Get
    End Property

    ''' <summary>Form type of park</summary>
    Public ReadOnly Property ParkType As String

        Get
            Contract.Requires(_shpFragments IsNot Nothing And _shpFragments.Any, "No fragment collection")

            Return _shpFragments(0).GetAttributeValue(_shpTypeField)
        End Get
    End Property

    Public ReadOnly Property GeoData As (Area As Double, Center As MapPoint, Env As Envelope)
        ' Iterating a fragments collection is very expensive, so calculate Area, Centroid and Envelope all at once
        Get
            Contract.Requires(_shpFragments IsNot Nothing And _shpFragments.Any, "No fragment collection")
            Dim GIS_AREA As Double = 0
            Dim X_COORD As Double = 0, Y_COORD As Double = 0, Center As MapPoint
            Dim env As Envelope = Nothing

            ' Do Area
            Select Case _Name
                Case "CAPAD_T", "CAPAD_M"
                    For Each frag As Feature In _shpFragments
                        GIS_AREA += CDbl(frag.GetAttributeValue("GIS_AREA"))
                        env = Form1.EnvelopeUnion(env, frag.Geometry.Extent)
                        X_COORD += CDbl(frag.GetAttributeValue("LONGITUDE"))
                        Y_COORD += CDbl(frag.GetAttributeValue("LATITUDE"))
                    Next
                    X_COORD /= _shpFragments.Count
                    Y_COORD /= _shpFragments.Count
                Case "VIC_PARKS"
                    For Each frag As Feature In _shpFragments
                        GIS_AREA += CDbl(frag.GetAttributeValue("TOTAL_AREA"))
                        env = Form1.EnvelopeUnion(env, frag.Geometry.Extent)
                    Next
                    ' get center of envelope
                    X_COORD = env.GetCenter.X
                    Y_COORD = env.GetCenter.Y
                Case "ZL"
                    GIS_AREA = _shpFragments(0).GetAttributeValue("Shape_Area").ToString / 10000
                    For Each frag As Feature In _shpFragments
                        env = Form1.EnvelopeUnion(env, frag.Geometry.Extent)
                    Next
                    ' get center of envelope
                    X_COORD = env.GetCenter.X
                    Y_COORD = env.GetCenter.Y
            End Select
            Contract.Ensures(GIS_AREA <> 0, "0 area")
            Contract.Ensures(X_COORD <> 0 And Y_COORD <> 0, "0 X or Y")

            Center = New MapPoint(X_COORD, Y_COORD, _shpFragments.SpatialReference)   ' create the point
            Dim result As (Area As Double, Center As MapPoint, Env As Envelope) = (GIS_AREA, Center, env)
            Return result
        End Get
    End Property
End Class
