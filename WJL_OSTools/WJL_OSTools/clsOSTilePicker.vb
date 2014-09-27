#Region "Namespaces"

Imports System.Math
Imports System.IO

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.Windows

#End Region

<Assembly: CommandClass(GetType(WJL1_OS_Tile_Picker.Commands))> 
<Assembly: ExtensionApplication(GetType(WJL1_OS_Tile_Picker.Initialisation))> 

Namespace WJL1_OS_Tile_Picker

    Public Class Initialisation
        Implements Autodesk.AutoCAD.Runtime.IExtensionApplication
        Public Sub Initialize() Implements IExtensionApplication.Initialize
            Dim myDWG As Document
            myDWG = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument

            Dim myEd As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            myEd.WriteMessage(vbLf)
            myEd.WriteMessage(vbLf & "OS Tile Picker by Will Lamb, lambwill@gmail.com")
            myEd.WriteMessage(vbLf & "Initialising commands...")
            myEd.WriteMessage(vbLf & "   'OSTilePick' command loads OS mapping tiles")
            myEd.WriteMessage(vbLf)
            ''Set OSTilePick to Attach mode
            'My.Settings.MapDetach = False
            'My.Settings.Save()
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate

        End Sub
    End Class

    Public Class Commands
        ''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        ''Not sure this is the best way to go about it, make these Shared?
        Dim OSTilePicker As New OSTilePicker
        Dim RasterPicker As New OSTilePicker.Raster
        Dim MasterMapPicker As New OSTilePicker.MasterMap
        Dim LinkConverter As New OSTilePicker.LinkConverter
        ''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

        <CommandMethod("OSTilePickSettings")> _
        Public Sub OSTilePickSettings_Method()
            Dim SettingsForm As New frmSettings()
            SettingsForm.ShowDialog()
        End Sub

        <CommandMethod("OSTileCommands", "OSTilePick", Nothing, CommandFlags.Modal, Nothing, "OS_Map_Picker", "Command1")> _
        Public Sub OSTilePick_Method()
            If My.Settings.FirstRun Then
                My.Settings.FirstRun = False
                Dim SettingsForm As New frmSettings()
                SettingsForm.ShowDialog()
            End If

            'Set OSTilePick to Attach mode
            My.Settings.MapDetach = False
            My.Settings.Save()

            ''Split the string stored in the SearchPaths settings into arrays of strings for the filesearch
            Dim SearchPaths_MasterMap() As String = Split(My.Settings.SearchPathsMM500, My.Settings.StringDelimiter)
            Dim SearchPaths_Raster() As String = Split(My.Settings.SearchPathsSV, My.Settings.StringDelimiter)

            ''Get the setting for asking to load a new copy of a tile if one is already attached
            Dim AskIfAttached As Boolean = My.Settings.AskIfAttached

            ''Get the prefix & Suffix for the placeholder block name
            Dim PlaceholderBlockPrefix As String = ""
            Dim PlaceholderBlockSufix As String = "-placeholder"

            '' Get the current database and start the Transaction Manager
            Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim myDB As Database = myDWG.Database
            Dim myED As Editor = Application.DocumentManager.MdiActiveDocument.Editor
NextTile:
            ''Set up lists of MapTiles for Raster and Mastermap MapTiles
            Dim MapTiles_MasterMap As New List(Of WJL1_OS_Tile_Picker.OSTilePicker.MapTile)
            Dim MapTiles_Raster As New List(Of WJL1_OS_Tile_Picker.OSTilePicker.MapTile)

            ''Set up a new TileJig class for ''jigging'' the tile picker
            Dim TileJig As New OSTilePicker.TileJig()

            ''Set up a new MapTile for the jig
            Dim JigMapTile As New OSTilePicker.MapTile
            TileJig.JigTile = JigMapTile

            'Get the current map type setting and set the TileJig type 
            TileJig.JigMapType = My.Settings.CurrentMapType


            Dim TileJigPromptOptions As New JigPromptPointOptions
            Dim PromptString As String = vbLf & "Select location for OS tile by mouse or enter co-ordinates or "

            PromptString = PromptString & "[map Type/Attach/Detach/Link/Settings] : "
            Dim Keywords As String = "Type Attach Detach Link Settings"

            TileJigPromptOptions.SetMessageAndKeywords(PromptString, Keywords)
            TileJigPromptOptions.UserInputControls = UserInputControls.NullResponseAccepted

            TileJig.JigPromptPointOptions = TileJigPromptOptions

            ''Use the TileJig to pick the tile (or get a keyword)
            Dim JigPointResult As PromptPointResult = myED.Drag(TileJig)
            Select Case JigPointResult.Status

                Case PromptStatus.OK
                    ''*********************************************************
                    ''Need to test if user input coordinates as text, the cursor location currectly overides the user input
                    ''*********************************************************
TileFromLink:
                    If JigMapTile.Initialised = True Then
                        ''If the user picked a valid point...

                        ''Get a copy of the MapTile from the TileJig
                        If JigMapTile.IsRaster Then
                            MapTiles_Raster.Add(JigMapTile)
                        Else
                            MapTiles_MasterMap.Add(JigMapTile)
                        End If
                    Else
                        ''If the user has picked a non-valid point, inform user that point is outside OS Grid
                        myED.WriteMessage(vbLf & "Point outside OS Grid.")
                    End If

                    If Not My.Settings.MapDetach Then
                        ''If OS Picker is in Attach mode

                        ''Set up an object collection for & insert the raster tile entities (object collection is held on to for commiting later)
                        Dim RasterEntitiesToBeCommited As DBObjectCollection
                        RasterEntitiesToBeCommited = RasterPicker.InsertRasterTiles(MapTiles_Raster, SearchPaths_Raster, AskIfAttached)

                        'Set up an object collection for & insert the mastermap tile entities (object collection is held on to for commiting later)
                        Dim MasterMapEntitiesToBeCommited As DBObjectCollection
                        MasterMapEntitiesToBeCommited = MasterMapPicker.InsertMasterMapTiles(MapTiles_MasterMap, SearchPaths_MasterMap, AskIfAttached)

                        'Set up an object collection for the entities to be commited at the end
                        Dim UnfoundEntitiesToBeCommited As New DBObjectCollection()

                        ''Ad the tiles that haven't been found to a new collection
                        Dim UnfoundTiles As New List(Of OSTilePicker.MapTile)
                        For Each MapTile In MapTiles_Raster
                            If MapTile.FilePath = "" Then UnfoundTiles.Add(MapTile)
                        Next
                        For Each MapTile In MapTiles_MasterMap
                            If MapTile.FilePath = "" Then UnfoundTiles.Add(MapTile)
                        Next

                        ''Process the unfound tiles
                        For Each MapTile In UnfoundTiles
                            If MapTile.FilePath = "" Then
                                ''If the MapTile file hasn't been found
                                ''Tell the user
                                myED.WriteMessage(vbLf & MapTile.Name & " not found.")

                                ''Create the place holder block 
                                Dim BlockId As ObjectId = OSTilePicker.CreatePlaceholderBlock(MapTile, PlaceholderBlockPrefix, PlaceholderBlockSufix, AskIfAttached)

                                If Not BlockId = Nothing Then
                                    ''If the block is defined

                                    Dim BlockRef As New BlockReference(Point3d.Origin, BlockId)
                                    UnfoundEntitiesToBeCommited.Add(BlockRef)
                                End If

                            End If
                        Next
                        ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                        ''Combine these?
                        OSTilePicker.CommitEntities(UnfoundEntitiesToBeCommited)
                        OSTilePicker.CommitEntities(RasterEntitiesToBeCommited)
                        OSTilePicker.CommitEntities(MasterMapEntitiesToBeCommited)
                        ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    Else
                        ''If OS Picker is in Detach mode

                        Dim DeleteIfInserted As Boolean = True

                        ''Detach the raster tiles (raster images)
                        RasterPicker.DetachRaster(MapTiles_Raster, DeleteIfInserted)
                        ''Detach the MasterMap tiles (xrefs)
                        MasterMapPicker.DetachMasterMap(MapTiles_MasterMap, DeleteIfInserted)

                        Dim UnfoundTiles As New List(Of OSTilePicker.MapTile)
                        For Each MapTile In MapTiles_Raster
                            If MapTile.FilePath = "" Then
                                Dim TilePlaceholderName As String = PlaceholderBlockPrefix & MapTile.Name & PlaceholderBlockSufix
                                MasterMapPicker.DeleteBlockRef(TilePlaceholderName)
                            End If
                        Next
                        For Each MapTile In MapTiles_MasterMap
                            If MapTile.FilePath = "" Then
                                If MapTile.FilePath = "" Then
                                    Dim TilePlaceholderName As String = PlaceholderBlockPrefix & MapTile.Name & PlaceholderBlockSufix
                                    MasterMapPicker.DeleteBlockRef(TilePlaceholderName)
                                End If
                            End If

                        Next


                    End If
                    ''Go back to the jig to get the next tile
                    GoTo NextTile

                Case PromptStatus.Keyword
                    ''If the user entered a keyword...
                    Select Case JigPointResult.StringResult
                        Case "Type"
                            ''If the keyword was Type

                            Dim msgMM500 As String = "Mastermap (500m)"
                            Dim msgMM1000 As String = "Mastermap (1km)"
                            Dim msgSV As String = "Streetview (1:10 000)"
                            Dim msg50kNew As String = "Raster 1:50 000 (10km)"
                            Dim msg50kOld As String = "Raster 1:50 000 (20km)"

                            myED.WriteMessage(vbLf & _
                                    "Map Types:" & vbLf & _
                                    "  1 - " & msgMM500 & vbLf & _
                                    "  2 - " & msgMM1000 & vbLf & _
                                    "  3 - " & msgSV & vbLf & _
                                    "  4 - " & msg50kNew & vbLf & _
                                    "  5 - " & msg50kOld)

                            Dim pIntOpts As New PromptIntegerOptions(vbLf & "Set map type: ")
                            pIntOpts.LowerLimit = 1
                            pIntOpts.UpperLimit = 5
                            pIntOpts.DefaultValue = My.Settings.CurrentMapType
                            pIntOpts.UseDefaultValue = True
                            Dim pIntRes As PromptIntegerResult = myDWG.Editor.GetInteger(pIntOpts)

                            If pIntRes.Status = PromptStatus.OK Then
                                My.Settings.CurrentMapType = pIntRes.Value
                                ''Save the changes to settings
                                My.Settings.Save()
                            End If
                            Dim MapTypeMessage As String

                            Select Case My.Settings.CurrentMapType
                                Case 1
                                    MapTypeMessage = msgMM500
                                Case 2
                                    MapTypeMessage = msgMM1000
                                Case 3
                                    MapTypeMessage = msgSV
                                Case 4
                                    MapTypeMessage = msg50kNew
                                Case 5
                                    MapTypeMessage = msg50kOld
                                Case Else
                                    MapTypeMessage = "UNKNOWN"
                            End Select

                            myED.WriteMessage(vbLf & "Map type set to " & MapTypeMessage)

                        Case "Attach"
                            My.Settings.MapDetach = False
                            ''Save the changes to settings
                            My.Settings.Save()
                            myED.WriteMessage(vbLf & "Attach mode")

                        Case "Detach"
                            My.Settings.MapDetach = True
                            ''Save the changes to settings
                            My.Settings.Save()
                            myED.WriteMessage(vbLf & "Detach mode")

                        Case "Link"
                            JigMapTile = LinkConverter.GetTileFromLink(My.Settings.CurrentMapType)
                            If JigMapTile IsNot Nothing Then
                                GoTo TileFromLink
                            End If

                        Case "Settings"
                            Dim SettingsForm As New frmSettings()
                            SettingsForm.ShowDialog()
                    End Select

                    'Go back to the jig to get the next tile
                    GoTo NextTile
                Case Else
                    ''If the user did something else (e.g. hit escape), quit
                    GoTo quit
            End Select
quit:
        End Sub
    End Class

    Public Class OSTilePicker
        
        Public Class MapTile
            Private Const LabelMTextScale As Double = 0.1
            Public Enum MapType
                MasterMap500 = 1
                MasterMap1000 = 2
                StreetView = 3
                FiftyThousandNew = 4
                FiftyThousandOld = 5
            End Enum
            Public Const NumberOfMapTypes As Integer = 5

            Property PickedPoint As Point3d
            Property Type As MapType

            Property Easting As Double
            Property Northing As Double
            Property LandRangerLetters As String

            Property IsRaster As Boolean
            Property TileWidth As Double
            Property Scale As Double
            Property Size As Double

            Property Quadrant As String
            Property Name As String

            Property Outline As Polyline
            Property LabelMText As MText

            Property InsPoint As Point3d
            Property Initialised As Boolean = False

            Property FilePath As String
            Property SearchPaths As String()

            Public Sub Initialise(ByVal PickedPoint As Point3d, ByVal Type As MapType)
                ''Populate the properties for the Maptile

                ''Round down the X & Y componants of the picked point for Easting & Northing
                Dim PickedEasting As Double = Floor(PickedPoint.X)
                Dim PickedNorthing As Double = Floor(PickedPoint.Y)

                ''Use the Easting & Northing to get the LandRanger letters, if GetLRLetters returns "" the point is outside OS grid
                LandRangerLetters = GetLRLetters(PickedEasting, PickedNorthing)
                If LandRangerLetters <> "" Then
                    Dim QuadrantX As Integer
                    Dim QuadrantY As Integer

                    Dim EastingString As String = ""
                    Dim NorthingString As String = ""

                    Select Case Type
                        Case MapType.MasterMap500
                            ''Set the defaults for MasterMap500
                            IsRaster = False
                            Scale = 1
                            Size = 500
                            InsPoint = New Point3d(0, 0, 0)

                            Easting = RoundDown(PickedEasting, Size)
                            Northing = RoundDown(PickedNorthing, Size)

                            ''Convert the coordinates to strings
                            EastingString = CStr(Easting)
                            NorthingString = CStr(Northing)

                            ''Get the 4th number in the Eastings & Northings to get the quadrant for the tile
                            QuadrantX = Mid(EastingString, 4, 1)
                            QuadrantY = Mid(NorthingString, 4, 1)
                            Quadrant = GetQuadrant(QuadrantX, QuadrantY)

                            ''get the 2nd & 3rd numbers in the Eastings & Northings for the tile name
                            EastingString = Mid(EastingString, 2, 2)
                            NorthingString = Mid(NorthingString, 2, 2)

                            SearchPaths() = Split(My.Settings.SearchPathsMM500, My.Settings.StringDelimiter)

                        Case MapType.MasterMap1000
                            ''Set the defaults for MasterMap1000
                            IsRaster = False
                            Scale = 1
                            Size = 1000
                            InsPoint = New Point3d(0, 0, 0)

                            Easting = RoundDown(PickedEasting, Size)
                            Northing = RoundDown(PickedNorthing, Size)

                            ''MasterMap1000 has no quadrant in the name
                            Quadrant = ""

                            ''Convert the coordinates to strings
                            EastingString = CStr(Easting)
                            NorthingString = CStr(Northing)

                            ''get the 2nd & 3rd numbers in the Eastings & Northings for the tile name
                            EastingString = Mid(EastingString, 2, 2)
                            NorthingString = Mid(NorthingString, 2, 2)

                            SearchPaths() = Split(My.Settings.SearchPathsMM1000, My.Settings.StringDelimiter)

                        Case MapType.StreetView
                            ''Set the defaults for StreetView
                            IsRaster = True
                            TileWidth = 5000
                            'Scale = 10000 'This should get set here
                            Size = 5000

                            Easting = RoundDown(PickedEasting, Size)
                            Northing = RoundDown(PickedNorthing, Size)

                            InsPoint = New Point3d(Easting, Northing, 0)

                            ''Convert the coordinates to strings
                            EastingString = CStr(Easting)
                            NorthingString = CStr(Northing)

                            ''Get the 3rd number in the Eastings & Northings to get the quadrant for the tile
                            QuadrantX = Mid(EastingString, 3, 1)
                            QuadrantY = Mid(NorthingString, 3, 1)
                            Quadrant = GetQuadrant(QuadrantX, QuadrantY)

                            ''get the 2nd number in the Eastings & Northings for the tile name
                            EastingString = Mid(EastingString, 2, 1)
                            NorthingString = Mid(NorthingString, 2, 1)

                            SearchPaths() = Split(My.Settings.SearchPathsSV, My.Settings.StringDelimiter)

                        Case MapType.FiftyThousandNew
                            ''Set the defaults for StreetView
                            IsRaster = True
                            TileWidth = 2000

                            Size = 10000

                            Easting = RoundDown(PickedEasting, Size)
                            Northing = RoundDown(PickedNorthing, Size)

                            InsPoint = New Point3d(Easting, Northing, 0)

                            Quadrant = ""

                            ''Convert the coordinates to strings
                            EastingString = CStr(Easting)
                            NorthingString = CStr(Northing)

                            ''get the 2nd number in the Eastings & Northings for the tile name
                            EastingString = Mid(EastingString, 2, 1)
                            NorthingString = Mid(NorthingString, 2, 1)

                            SearchPaths() = Split(My.Settings.SearchPaths50kNew, My.Settings.StringDelimiter)

                        Case MapType.FiftyThousandOld
                            ''Set the defaults for StreetView
                            IsRaster = True
                            TileWidth = 4000

                            Size = 20000

                            Easting = RoundDown(PickedEasting, Size)
                            Northing = RoundDown(PickedNorthing, Size)

                            InsPoint = New Point3d(Easting, Northing, 0)

                            Quadrant = ""

                            ''Convert the coordinates to strings
                            EastingString = CStr(Easting)
                            NorthingString = CStr(Northing)

                            ''get the 2nd number in the Eastings & Northings for the tile name
                            EastingString = Mid(EastingString, 2, 1)
                            NorthingString = Mid(NorthingString, 2, 1)

                            SearchPaths() = Split(My.Settings.SearchPaths50kOld, My.Settings.StringDelimiter)

                    End Select
                    ''Assemble the tile name
                    Name = LandRangerLetters & EastingString & NorthingString & Quadrant
                    Outline = GetOutline(Easting, Northing, Size)
                    LabelMText = GetLabelMText(Name, Easting, Northing, Size)
                    Initialised = True
                End If

            End Sub

            ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            ''Move this to LambTools?
            Private Function RoundDown(ByVal Number As Double, ByVal RoundTo As Double) As Double
                Dim RoundedNumber As Double
                If RoundTo > 0 Then
                    ''Rounds down to nearest RoundTo
                    RoundedNumber = Number / RoundTo
                    RoundedNumber = Floor(RoundedNumber)
                    RoundedNumber = RoundedNumber * RoundTo
                End If
                Return RoundedNumber
            End Function
            ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

            Private Function GetLRLetters(ByVal Easting As Double, ByVal Northing As Double) As String
                ''Turn the E&N doubles into strings for processing
                Dim strEasting As String = CStr(Easting)
                Dim strNorthing As String = CStr(Northing)

                Dim strLandRanger As String = ""

                If strEasting.Length = 6 And strNorthing.Length = 6 Then
                    ''get the first number of the Easting & Northings and convert them from strings to integers
                    Dim intE1 As Integer = CInt(Left(strEasting, 1))
                    Dim intN1 As Integer = CInt(Left(strNorthing, 1))

                    ''Set up string array with OS Landranger letters
                    Dim Letters(7, 9) As String
                    Letters(0, 0) = "SV" : Letters(1, 0) = "SW" : Letters(2, 0) = "SX" : Letters(3, 0) = "SY" : Letters(4, 0) = "SZ" : Letters(5, 0) = "TV" : Letters(6, 0) = ""
                    Letters(0, 1) = "SQ" : Letters(1, 1) = "SR" : Letters(2, 1) = "SS" : Letters(3, 1) = "ST" : Letters(4, 1) = "SU" : Letters(5, 1) = "TQ" : Letters(6, 1) = "TR"
                    Letters(0, 2) = "" : Letters(1, 2) = "SM" : Letters(2, 2) = "SN" : Letters(3, 2) = "SO" : Letters(4, 2) = "SP" : Letters(5, 2) = "TL" : Letters(6, 2) = "TM"
                    Letters(0, 3) = "" : Letters(1, 3) = "SG" : Letters(2, 3) = "SH" : Letters(3, 3) = "SJ" : Letters(4, 3) = "SK" : Letters(5, 3) = "TF" : Letters(6, 3) = "TG"
                    Letters(0, 4) = "" : Letters(1, 4) = "SB" : Letters(2, 4) = "SC" : Letters(3, 4) = "SD" : Letters(4, 4) = "SE" : Letters(5, 4) = "TA" : Letters(6, 4) = "TB"
                    Letters(0, 5) = "" : Letters(1, 5) = "NW" : Letters(2, 5) = "NX" : Letters(3, 5) = "NY" : Letters(4, 5) = "NZ" : Letters(5, 5) = "OV" : Letters(6, 5) = "OW"
                    Letters(0, 6) = "NQ" : Letters(1, 6) = "NR" : Letters(2, 6) = "NS" : Letters(3, 6) = "NT" : Letters(4, 6) = "NU" : Letters(5, 6) = "OQ" : Letters(6, 6) = ""
                    Letters(0, 7) = "NL" : Letters(1, 7) = "NM" : Letters(2, 7) = "NN" : Letters(3, 7) = "NO" : Letters(4, 7) = "NP" : Letters(5, 7) = "OL" : Letters(6, 7) = ""
                    Letters(0, 8) = "NF" : Letters(1, 8) = "NG" : Letters(2, 8) = "NH" : Letters(3, 8) = "NJ" : Letters(4, 8) = "NK" : Letters(5, 8) = "OF" : Letters(6, 8) = ""
                    Letters(0, 9) = "NA" : Letters(1, 9) = "NB" : Letters(2, 9) = "NC" : Letters(3, 9) = "ND" : Letters(4, 9) = "NE" : Letters(5, 9) = "OA" : Letters(6, 9) = ""

                    ''Pick out the right letters
                    strLandRanger = Letters(intE1, intN1)
                End If
                ''Return the letters
                Return strLandRanger
            End Function

            Private Function GetQuadrant(ByVal x As Integer, ByVal y As Integer) As String
                ''Get the quadrant for the tile name (NW, NE, SW or SE)
                Dim VertQuad As Char
                If y < 5 Then
                    VertQuad = "S"
                Else
                    VertQuad = "N"
                End If

                Dim HorzQuad As Char
                If x < 5 Then
                    HorzQuad = "W"
                Else
                    HorzQuad = "E"
                End If

                Dim Quadrant As String = VertQuad & HorzQuad
                Return Quadrant
            End Function

            Private Function GetOutline(ByVal Easting As Double, ByVal Northing As Double, ByVal Size As Double) As Polyline
                '' Create a polyline with two segments (3 points)
                Dim Outline As Polyline = New Polyline()

                Dim BottomLeft As New Point2d(Easting, Northing)
                Outline.AddVertexAt(0, BottomLeft, 0, 0, 0)

                Dim TopLeft As New Point2d(Easting, (Northing + Size))
                Outline.AddVertexAt(0, TopLeft, 0, 0, 0)

                Dim TopRight As New Point2d((Easting + Size), (Northing + Size))
                Outline.AddVertexAt(0, TopRight, 0, 0, 0)

                Dim BottomRight As New Point2d((Easting + Size), Northing)
                Outline.AddVertexAt(0, BottomRight, 0, 0, 0)

                Outline.Closed = True

                Return Outline
            End Function

            Private Function GetLabelMText(ByVal Name As String, ByVal Easting As Double, ByVal Northing As Double, ByVal Size As Double) As MText
                'Create the MText for the MapTile placeholder 
                Dim LabelMText As MText = New MText()

                LabelMText.Attachment = AttachmentPoint.MiddleCenter

                Dim LabelLocation As New Point3d((Easting + (Size / 2)), (Northing + (Size / 2)), 0)
                LabelMText.Location = LabelLocation
                LabelMText.TextHeight = Size * LabelMTextScale
                LabelMText.Width = 0
                LabelMText.Contents = Name

                Return LabelMText
            End Function
        End Class

        Public Class TileJig
            Inherits DrawJig

            Property JigTile As MapTile
            Property JigMapType As MapTile.MapType
            Property JigPromptPointOptions As JigPromptPointOptions
            Dim UserPoint As Point3d

            Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus
                Dim JigPropmtPointResult As PromptPointResult
                ''££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££
                ''Needs to take computed point rather than crosshair position, ie honour snaps or command line entry
                JigPropmtPointResult = prompts.AcquirePoint(JigPromptPointOptions)

                ''££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££££

                Select Case JigPropmtPointResult.Status
                    Case PromptStatus.None
                        ''User input keyword, tell WorldDraw it doesn't need to do anything
                        Return SamplerStatus.NoChange
                    Case PromptStatus.Keyword
                        ''User input keyword, tell WorldDraw it doesn't need to do anything
                        Return SamplerStatus.NoChange
                    Case PromptStatus.OK

                        UserPoint = JigPropmtPointResult.Value
                        'Return SamplerStatus.OK
                        ''Test if the user is hovering over a new tile so that the jig only regens when it needs to
                        ''(doesn't work, seems to cause jig to return previous point or some reason)
                        If UserPoint.X < JigTile.Easting Or UserPoint.X > (JigTile.Easting + JigTile.Size) Then
                            Return SamplerStatus.OK
                        ElseIf UserPoint.Y < JigTile.Northing Or UserPoint.Y > (JigTile.Northing + JigTile.Size) Then
                            Return SamplerStatus.OK
                        Else
                            Return SamplerStatus.NoChange
                        End If
                    Case Else
                        ''If something else happens (e.g. user hits escape), return Cancel
                        Return SamplerStatus.Cancel
                End Select

            End Function

            Protected Overrides Function WorldDraw(ByVal Draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean
                ''(Re)Initialise the JigTile to update it for the new user point
                JigTile.Initialise(UserPoint, JigMapType)
                ''Draw the outline of the tile
                If JigTile.Initialised = True Then
                    Draw.Geometry.Draw(JigTile.Outline)
                    Draw.Geometry.Draw(JigTile.LabelMText)
                End If
                Return True
            End Function

        End Class

        Public Class Raster
            ''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            ''Not sure this is the best way to go about it, make these Shared?
            Dim OSTilePicker As New OSTilePicker
            ''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

            Public Function InsertRasterTiles(ByRef MapTiles_Raster As List(Of MapTile), ByVal SearchPaths_Raster() As String, Optional ByVal AskIfAttached As Boolean = True) As DBObjectCollection
                OSTilePicker.FindTiles(MapTiles_Raster) ', SearchPaths_Raster)

                'Set up an object collection for the entities to be commited at the end
                Dim EntitiesToBeCommited As New DBObjectCollection()

                For Each MapTile In MapTiles_Raster

                    If MapTile.FilePath <> "" Then
                        ''If the MapTile file has been found

                        ''Set the image insertion Scale
                        'MapTile.Scale = MapTile.Size / MapTile.ImageWidth
                        Dim ImageLayer As String = ""      'Set the layer for the Raster tiles ("" = use current layer)
                        Dim RasterFade As Integer = 65     'Set the fade for the raster tiles
                        'Dim AskIfAttached As Boolean = True  'Ask if the user wans to load the tile in if it's already loaded
                        ''***********************************************************
                        ''Add_Raster commits the images, would be better if they get commited with everything else below
                        Add_Raster(MapTile.FilePath, MapTile.InsPoint, MapTile.Size, RasterFade, ImageLayer, AskIfAttached)
                        ''***********************************************************
                    End If
                Next
                Return EntitiesToBeCommited

            End Function

            Public Function Add_Raster(ByVal RasterFile As String, ByVal InsPoint As Point3d, ByVal TileWidth As Double, Optional ByVal RasterFade As Integer = 65, Optional ByVal ImageLayer As String = "", Optional ByVal AskIfAttached As Boolean = True) As ObjectId
                Dim RasterID As ObjectId

                Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
                Dim myEd As Editor = myDWG.Editor

                Dim RasterEnt As New RasterImage
                Dim RasterDef As New RasterImageDef

                Dim ImageDefID As ObjectId
                Dim ImageDic As DBDictionary
                Dim ImageDicID As ObjectId

                Dim RasterName As String = ""
                RasterName = RasterFile.Substring(RasterFile.LastIndexOf("\") + 1)
                RasterName = RasterName.Substring(0, RasterName.IndexOf("."))

                'Exit if there's no active drawing (is this necessary?)
                If myDWG Is Nothing Then Return Nothing

                Try

                    Using Trans As Transaction = myDWG.TransactionManager.StartTransaction()
                        'RasterEnt = New RasterImage
                        'RasterEnt.Dispose() '' force loading of RasterImage.dbx module (needed for 2009 and earlier)

                        'Try to get the Image Dictionary 
                        ImageDicID = RasterImageDef.GetImageDictionary(myDWG.Database)
                        If ImageDicID.IsNull Then
                            'If no Image Dictionary is found, try to create one
                            RasterImageDef.CreateImageDictionary(myDWG.Database)
                            ImageDicID = RasterImageDef.GetImageDictionary(myDWG.Database)
                        End If

                        If ImageDicID.IsNull Then
                            'If the Image Dictionary ID is still not set to anything, the Image Dictionary couldn't be created. Exit the function, returning Nothing.
                            myEd.WriteMessage(vbLf & "Add_Raster could not create image dictionary.")
                            Return Nothing
                        End If

                        'Open the Image Dictionary to check if the raster image is already defined
                        ImageDic = Trans.GetObject(ImageDicID, OpenMode.ForRead)
                        If ImageDic Is Nothing Then
                            myEd.WriteMessage(vbLf & "Add_Raster could not open image dictionary")
                            Return Nothing
                        Else
                            If ImageDic.Contains(RasterName) Then

                                If AskIfAttached Then
                                    Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                                    pKeyOpts.Message = vbLf & "Tile " & RasterName & " aready exists. Do you wish to load in another copy?"
                                    pKeyOpts.Keywords.Add("Yes")
                                    pKeyOpts.Keywords.Add("No")
                                    pKeyOpts.Keywords.Default = "No"
                                    pKeyOpts.AllowNone = True
                                    Dim pKeyRes As PromptResult = myEd.GetKeywords(pKeyOpts)

                                    Select Case pKeyRes.StringResult
                                        Case "Yes"
                                            'If the raster image is already in the dictionary, find the entry for it

                                            '******************************************************
                                            ''Can I just use ImageDic.GetAt("RasterName") instead?
                                            For Each dictEntry In ImageDic
                                                'Get the object from the dictionary entry
                                                Dim obj As DBObject = Trans.GetObject(dictEntry.Value(), OpenMode.ForWrite, False)

                                                'Check that object is a RasterImageDef
                                                If obj.GetRXClass.IsDerivedFrom(RXClass.GetClass(GetType(RasterImageDef))) Then
                                                    ''Create a temporary RasterImageDef wrapper for the object
                                                    Dim RasterDefTemp As RasterImageDef = DisposableWrapper.Create(GetType(RasterImageDef), obj.UnmanagedObject, False)
                                                    If RasterDefTemp.SourceFileName = RasterFile Then
                                                        ''If the image edfinition's source file matches the raster source file, get the definition & ID
                                                        RasterDef = RasterDefTemp
                                                        RasterDef.Load()
                                                        ImageDefID = RasterDef.Id
                                                    End If
                                                End If
                                            Next
                                            '******************************************************

                                            If RasterDef Is Nothing Then
                                                'If the RasterDef's name was in the Image Dictionary, but it has a different source file, exit the function
                                                myEd.WriteMessage(vbLf & "Image name " & RasterName & "is already in use with another source file.")
                                                Return Nothing
                                            End If
                                        Case Else
                                            Return Nothing
                                    End Select

                                End If
                            Else
                                'If the raster image is not in the Image Dictionary, add it in
                                RasterDef = New RasterImageDef
                                RasterDef.SourceFileName = RasterFile
                                RasterDef.ActiveFileName = RasterFile
                                RasterDef.Load()

                                ImageDic.UpgradeOpen()
                                ImageDefID = ImageDic.SetAt(RasterName, RasterDef)
                                Trans.AddNewlyCreatedDBObject(RasterDef, True)
                            End If
                            RasterEnt = New RasterImage
                            RasterEnt.SetDatabaseDefaults(myDWG.Database)



                            ''Set the image to use for the image entity
                            RasterEnt.ImageDefId = ImageDefID

                            ''Set the layer (if supplied)
                            If ImageLayer <> "" Then
                                RasterEnt.Layer = ImageLayer
                            End If

                            If RasterFade > 0 And RasterFade <= 100 Then
                                RasterEnt.Fade = RasterFade
                            End If

                            ''Get the BTR for the current space
                            Dim CurrentSpaceBTR As BlockTableRecord = Trans.GetObject(myDWG.Database.CurrentSpaceId, OpenMode.ForWrite)

                            ''Add the image to the BTR
                            CurrentSpaceBTR.AppendEntity(RasterEnt)
                            Trans.AddNewlyCreatedDBObject(RasterEnt, True)
                            RasterEnt.AssociateRasterDef(RasterDef)

                            Dim Scale As Double
                            Scale = TileWidth / RasterEnt.Width
                            ''Scale the image entity
                            Dim ScaleMatrix As Matrix3d = Matrix3d.Scaling(Scale, New Point3d(0, 0, 0))
                            RasterEnt.TransformBy(ScaleMatrix)

                            ''Move the image entity
                            Dim DispaceMatrix As Matrix3d = Matrix3d.Displacement(New Vector3d(InsPoint.X, InsPoint.Y, InsPoint.Z))
                            RasterEnt.TransformBy(DispaceMatrix)

                            Trans.Commit()
                            RasterID = RasterEnt.Id
                        End If
                    End Using
                    Return RasterID
                Catch ex As Exception
                    myEd.WriteMessage(vbLf & "Error inserting image " & RasterName & vbLf & ex.ToString)
                    Return Nothing
                End Try

            End Function

            Public Sub DetachRaster(ByRef MapTiles_Raster As List(Of MapTile), Optional ByVal DeleteIfInserted As Boolean = True)
                Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
                Dim myDB As Database = myDWG.Database
                Dim ImageDefRemoved As Boolean
                For Each MapTile In MapTiles_Raster

                    If DeleteIfInserted Then
                        Dim ImageDeleteCount As Integer = EraseRasterImage(MapTile.Name)  'This throws up the 'Detach image?' dialogue, it would be nice to suppress it and detach automatically...
                        ImageDefRemoved = RemoveImageDef(MapTile.Name)
                        If ImageDeleteCount > 0 Or ImageDefRemoved = True Then
                            MapTile.FilePath = "DETACHED"
                        End If

                    Else
                        Dim CountOnly As Boolean = True
                        Dim ImageCount As Integer = EraseRasterImage(MapTile.Name, CountOnly)
                        If ImageCount = 0 Then
                            ImageDefRemoved = RemoveImageDef(MapTile.Name)
                            If ImageDefRemoved = True Then
                                MapTile.FilePath = "DETACHED"
                            End If
                        End If
                    End If

                Next
            End Sub

            Private Function EraseRasterImage(ByVal RasterName As String, Optional ByVal CountOnly As Boolean = False) As Integer
                Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
                Dim myDB As Database = myDWG.Database
                Dim ImageDeleteCount As Integer = 0

                Using Trans As Transaction = myDB.TransactionManager.StartTransaction()
                    Dim myBT As BlockTable = TryCast(Trans.GetObject(myDB.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim myBTR As BlockTableRecord = TryCast(Trans.GetObject(myBT(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                    For Each ObjectId As ObjectId In myBTR
                        If ObjectId.ObjectClass.DxfName = "IMAGE" Then
                            Dim RasterEnt As RasterImage = TryCast(Trans.GetObject(ObjectId, OpenMode.ForRead), RasterImage)
                            If RasterEnt.Name = RasterName Then
                                RasterEnt.UpgradeOpen()
                                RasterEnt.Erase() 'This throws up the 'Detach imge?' dialogue, it would be nice to suppress it and detach automatically...
                                ImageDeleteCount += 1
                            End If
                        End If
                    Next
                    Trans.Commit()
                End Using
                Return ImageDeleteCount
            End Function

            Function RemoveImageDef(ByVal RasterName As String) As Boolean
                Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
                Dim myEd As Editor = myDWG.Editor

                Dim ImageDic As DBDictionary
                Dim ImageDicID As ObjectId

                'Exit if there's no active drawing (is this necessary?)
                If myDWG Is Nothing Then Return False
                Try
                    Using Trans As Transaction = myDWG.TransactionManager.StartTransaction()

                        'Try to get the Image Dictionary 
                        ImageDicID = RasterImageDef.GetImageDictionary(myDWG.Database)
                        If ImageDicID.IsNull Then
                            'No Image Dictionary is found
                            Return Nothing
                        End If

                        'Open the Image Dictionary to check if the raster image is already defined
                        ImageDic = Trans.GetObject(ImageDicID, OpenMode.ForRead)
                        If ImageDic Is Nothing Then
                            myEd.WriteMessage(vbLf & "DeleteImageDef could not open image dictionary")
                            Return Nothing
                        Else
                            If ImageDic.Contains(RasterName) Then
                                ''If there's a raster image defined with the right name, remove it from the dictionary
                                ImageDic.UpgradeOpen()
                                ImageDic.Remove(RasterName)
                                Trans.Commit()
                                Return True
                            Else
                                Return False
                            End If


                        End If
                    End Using

                Catch ex As Exception
                    myEd.WriteMessage(vbLf & "Error detaching image definition: " & RasterName & vbLf & ex.ToString)
                    Return False
                End Try

            End Function

        End Class

        Public Class MasterMap

            ''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            ''Not sure this is the best way to go about it, make these Shared?
            Dim OSTilePicker As New OSTilePicker
            ''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

            Public Sub SetLayerProperties(ByVal LayerFreezeList() As String, ByVal XRefName As String, Optional ByVal LayerColour As Color = Nothing)

                '' Get the current document and database 
                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                Dim acCurDb As Database = acDoc.Database
                '' Start a transaction 
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    '' Open the Layer table for read 
                    Dim acLyrTbl As LayerTable
                    acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                    ''Set up a container for the layer being evaluated
                    Dim acLyrTblRec As LayerTableRecord


                    ''##############################################################
                    ''Set the colour for all the XRef layers

                    If LayerColour IsNot Nothing Then
                        ''Cycle through all the layers
                        For Each layerId As ObjectId In acLyrTbl
                            acLyrTblRec = TryCast(acTrans.GetObject(layerId, OpenMode.ForWrite), LayerTableRecord)
                            If InStr(acLyrTblRec.Name, (XRefName & "|")) > 0 Then
                                ''If the layer is part of the XRef, set the colour
                                acLyrTblRec.Color = LayerColour
                            End If
                        Next
                    End If
                    ''##############################################################

                    Dim sLayerName As String
                    For Each LayerToFreeze As String In LayerFreezeList
                        sLayerName = XRefName & "|" & LayerToFreeze
                        If acLyrTbl.Has(sLayerName) Then
                            acLyrTblRec = acTrans.GetObject(acLyrTbl(sLayerName), OpenMode.ForWrite)
                            '' Freeze the layer 
                            acLyrTblRec.IsFrozen = True
                        End If
                    Next
                    '' Save the changes and dispose of the transaction 
                    acTrans.Commit()
                End Using
            End Sub

            Public Function InsertMasterMapTiles(ByRef MapTiles_MasterMap As List(Of MapTile), ByVal SearchPaths_MasterMap() As String, Optional ByVal AskIfAttached As Boolean = True) As DBObjectCollection
                '' Get the current database and start the Transaction Manager
                Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
                Dim myDB As Database = myDWG.Database
                Dim myEd As Editor = myDWG.Editor

                OSTilePicker.FindTiles(MapTiles_MasterMap) ', SearchPaths_MasterMap)

                Dim XRefID As ObjectId

                'Set up an object collection for the entities to be commited at the end
                Dim EntitiesToBeCommited As New DBObjectCollection()

                Dim XRefAttached As Boolean = False
                For Each MapTile In MapTiles_MasterMap

                    If MapTile.FilePath <> "" Then
                        ''If the MapTile file has been found
                        Dim TileXRef As BlockReference
                        Using tr As Transaction = myDB.TransactionManager.StartTransaction
                            Dim myBT As BlockTable = myDB.BlockTableId.GetObject(OpenMode.ForRead)
                            If myBT.Has(MapTile.Name) Then
                                Dim btr As BlockTableRecord = myBT(MapTile.Name).GetObject(OpenMode.ForRead)
                                XRefID = btr.ObjectId
                                XRefAttached = True
                            End If
                        End Using


                        If XRefAttached Then
                            'myEd.WriteMessage(vbLf & "Tile " & MapTile.Name & " aready exists.")
                            If AskIfAttached Then
                                Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                                pKeyOpts.Message = vbLf & "Tile " & MapTile.Name & " already exists. Do you wish to load in another copy?"
                                pKeyOpts.Keywords.Add("Yes")
                                pKeyOpts.Keywords.Add("No")
                                pKeyOpts.Keywords.Default = "No"
                                pKeyOpts.AllowNone = True
                                Dim pKeyRes As PromptResult = myEd.GetKeywords(pKeyOpts)

                                Select Case pKeyRes.StringResult
                                    Case "Yes"
                                        TileXRef = New BlockReference(New Point3d(0, 0, 0), XRefID)
                                    Case Else
                                        TileXRef = Nothing
                                End Select
                            Else
                                TileXRef = Nothing
                            End If
                        Else
                            ''Attach the Xref
                            XRefID = myDB.AttachXref(MapTile.FilePath, MapTile.Name)
                            TileXRef = New BlockReference(New Point3d(0, 0, 0), XRefID)
                        End If

                        ''Get the LayerFreezeList array by splitting the LayerFreezeList setting string
                        Dim LayerFreezeList() As String = Split(My.Settings.LayerFreezeListMM500, My.Settings.StringDelimiter)

                        ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                        ''Would be better if LayerColourIndex setting held the color object rather than the colour index
                        Dim LayerColour As Color = Color.FromColorIndex(ColorMethod.ByAci, My.Settings.LayerColourIndexMM500)
                        ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

                        ''Set properties for xref layers
                        SetLayerProperties(LayerFreezeList, MapTile.Name, LayerColour)

                        If TileXRef IsNot Nothing Then
                            ''Add the XRef to the collection of entities to be commited
                            EntitiesToBeCommited.Add(TileXRef)
                        End If

                    End If
                Next
                Return EntitiesToBeCommited
            End Function

            Public Sub DetachMasterMap(ByRef MapTiles_MasterMap As List(Of MapTile), Optional ByVal DeleteIfInserted As Boolean = True)
                Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
                Dim myDB As Database = myDWG.Database
                For Each MapTile In MapTiles_MasterMap

                    If DeleteIfInserted Then
                        Dim BlockDeleteCount As Integer = DeleteBlockRef(MapTile.Name)
                        If DetachXref(myDB, MapTile.Name) = True Then
                            MapTile.FilePath = "DETACHED"
                        End If
                    Else
                        Dim CountOnly As Boolean = True
                        Dim BlockCount As Integer = DeleteBlockRef(MapTile.Name, CountOnly)
                        If BlockCount = 0 Then
                            DetachXref(myDB, MapTile.Name)
                            If DetachXref(myDB, MapTile.Name) = True Then
                                MapTile.FilePath = "DETACHED"
                            End If
                        End If
                    End If

                Next
            End Sub

            Public Function DeleteBlockRef(ByVal Blockname As String, Optional ByVal CountOnly As Boolean = False) As Integer
                Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
                Dim myDB As Database = myDWG.Database
                Dim BlockDeleteCount As Integer = 0
                Using Trans As Transaction = myDB.TransactionManager.StartTransaction()
                    Dim myBT As BlockTable = DirectCast(Trans.GetObject(myDB.BlockTableId, OpenMode.ForRead), BlockTable)
                    For Each ObjectId As ObjectId In myBT
                        Dim btr As BlockTableRecord = DirectCast(Trans.GetObject(ObjectId, OpenMode.ForWrite), BlockTableRecord)
                        Dim BlockRefIDs As ObjectIdCollection = btr.GetBlockReferenceIds(True, False)
                        For i As Integer = 0 To BlockRefIDs.Count - 1
                            Dim blkRef As BlockReference = DirectCast(Trans.GetObject(BlockRefIDs(i), OpenMode.ForRead), BlockReference)

                            If blkRef.Name = Blockname Then
                                BlockRefIDs.RemoveAt(i)
                                blkRef.UpgradeOpen()
                                blkRef.[Erase]()
                                BlockDeleteCount += 1
                            End If

                        Next
                    Next
                    Trans.Commit()
                End Using
                Return BlockDeleteCount
            End Function

            Private Shared Function DetachXref(ByVal db As Database, ByVal xrefName As String) As Boolean
                Dim XRefDetached As Boolean = False
                Using trans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim xrefGraph As XrefGraph = db.GetHostDwgXrefGraph(True)
                    Dim xrefCount As Integer = xrefGraph.NumNodes
                    For i As Integer = 0 To xrefCount - 1
                        Dim xrefNode As XrefGraphNode = xrefGraph.GetXrefNode(i)
                        If xrefNode.Name.ToLower() = xrefName.ToLower() Then
                            Dim xrefId As ObjectId = xrefNode.BlockTableRecordId
                            db.DetachXref(xrefId)
                            XRefDetached = True
                            Exit For
                        End If
                    Next
                    trans.Commit()
                End Using
                Return XRefDetached
            End Function

        End Class

        Public Class LinkConverter
            Public Function GetTileFromLink(ByVal MapType As WJL1_OS_Tile_Picker.OSTilePicker.MapTile.MapType) As WJL1_OS_Tile_Picker.OSTilePicker.MapTile
                Dim MapTile As WJL1_OS_Tile_Picker.OSTilePicker.MapTile = Nothing
                Dim LinkString As String
                Dim LinkStringLCase As String
                Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
                Dim pStrOpts As New PromptStringOptions(vbLf & "Paste map hyperlink here: ")
                pStrOpts.AllowSpaces = True
                Dim pStrRes As PromptResult = myDWG.Editor.GetString(pStrOpts)

                Select Case pStrRes.Status
                    Case PromptStatus.OK
                        ''Make the link text all lower case
                        LinkString = pStrRes.StringResult
                        LinkStringLCase = LCase(LinkString)
                        ''Get the insertion Point from the link using InPointFromLink
                        Dim InPoint As Point3d = InPointFromLink(LinkStringLCase)
                        If InPoint = Nothing Then
                            ''If InPointFromLink doesn't return a Point3d
                            myDWG.Editor.WriteMessage(vbLf & "Hyperlink not recognised.")
                            MapTile = Nothing
                        Else
                            ''If InPointFromLink returns an insertion point
                            'myDWG.Editor.WriteMessage(vbLf & "Point = " & InPoint.X & "," & InPoint.Y & "," & InPoint.Z)

                            ''Get the map tile for that point
                            MapTile = New WJL1_OS_Tile_Picker.OSTilePicker.MapTile
                            MapTile.Initialise(InPoint, MapType)
                        End If
                    Case Else
                        ''User cancelled (or did something wrong)
                        MapTile = Nothing
                End Select
                Return MapTile
            End Function

            Public Function InPointFromLink(ByVal LinkString As String) As Point3d
                Dim InPoint As Point3d
                Dim SourceWebSites As New List(Of String)
                SourceWebSites.Add("streetmap.co.uk")
                'SourceWebSites.Add("maps.google")

                Dim SourceWebSite As String = ""
                For Each WebSite As String In SourceWebSites
                    If InStr(LinkString, WebSite) <> 0 Then
                        SourceWebSite = WebSite
                    End If
                Next

                Select Case SourceWebSite
                    Case "streetmap.co.uk"
                        InPoint = PointFromStreetmap(LinkString)
                        'Case "maps.google"
                        'InPoint = PointFromGoogleMaps(LinkString)
                    Case Else
                        InPoint = Nothing
                End Select
                Return InPoint
            End Function

            Public Function PointFromStreetmap(ByVal LinkString As String) As Point3d
                Dim InPoint As Point3d
                Dim EastingsStartPoint As Integer
                Dim NorthingsStartPoint As Integer

                Dim strEastings As String
                Dim strNorthings As String

                Dim dblEastings As Double = 0
                Dim dblNorthings As Double = 0

                EastingsStartPoint = InStr(LinkString, "x=")
                If EastingsStartPoint > 0 Then
                    EastingsStartPoint += 2
                    strEastings = Mid(LinkString, EastingsStartPoint, 6)
                    If strEastings Like "######" Then
                        dblEastings = CDbl(strEastings)
                    End If
                End If


                NorthingsStartPoint = InStr(LinkString, "y=")
                If NorthingsStartPoint > 0 Then
                    NorthingsStartPoint += 2
                    strNorthings = Mid(LinkString, NorthingsStartPoint, 6)
                    If strNorthings Like "######" Then
                        dblNorthings = CDbl(strNorthings)
                    End If
                End If

                If dblEastings > 0 And dblNorthings > 0 Then
                    InPoint = New Point3d(dblEastings, dblNorthings, 0)
                Else
                    InPoint = Nothing
                End If


                Return InPoint
            End Function
        End Class

        Sub CommitEntities(ByVal EntitiesToBeCommited As DBObjectCollection)
            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim myED As Editor = Application.DocumentManager.MdiActiveDocument.Editor

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                '' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)


                For Each ent As Entity In EntitiesToBeCommited
                    Try
                        acBlkTblRec.AppendEntity(ent)
                        acTrans.AddNewlyCreatedDBObject(ent, True)
                    Catch ex As Exception
                        myED.WriteMessage(ex.Message)
                    End Try

                Next

                '' Save the new object to the database
                acTrans.Commit()

            End Using
        End Sub

        Sub FindTiles(ByRef Maptiles As List(Of MapTile)) ', ByVal SearchPaths() As String)
            ''This subroutine cycles through each SearchPath and searches for each MapTile in MapTiles, appending the found file path to MapTile.FilePath (or not if it can't find it)


            ''String for holding the current search path
            Dim strSearchPath As String

            ''List of file extensions for raster mapping
            Dim RasterExtensions As New List(Of String)
            RasterExtensions.Add(".jpg")
            RasterExtensions.Add(".png")
            RasterExtensions.Add(".tif")

            ''List of file extensions for vector (dwg) mapping 
            Dim DwgExtensions As New List(Of String)
            DwgExtensions.Add(".dwg")
            'DwgExtensions.Add(".dxf")



            'Cycle through each MapTile in MapTiles
            For Each MapTile In Maptiles
                If MapTile.FilePath = "" Then
                    ''Cycle through the filepaths
                    ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    ''This could take a while if there's a few search paths, maybe index the mapping directories to narrow down the search?
                    For Each SearchPath In MapTile.SearchPaths
                        strSearchPath = SearchPath & "\"
                        'If the MapTile hasn't been found yet, test if it IsRaster mapping
                        If MapTile.IsRaster Then
                            ''If MapTile IsRaster, try to find the file in this searchpath (for each RasterExtension)
                            For Each RasterExtension In RasterExtensions
                                If File.Exists(strSearchPath & MapTile.Name & RasterExtension) Then
                                    MapTile.FilePath = strSearchPath & MapTile.Name & RasterExtension
                                    Exit For
                                End If
                            Next
                        Else
                            ''If MapTile isn't raster mapping, try to find the file in this searchpath (for each DwgExtension)
                            For Each DwgExtension In DwgExtensions
                                If File.Exists(strSearchPath & MapTile.Name & DwgExtension) Then
                                    MapTile.FilePath = strSearchPath & MapTile.Name & DwgExtension
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                End If
            Next

            ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

        End Sub

        Public Function CreatePlaceholderBlock(ByVal MapTile As WJL1_OS_Tile_Picker.OSTilePicker.MapTile, Optional ByVal BlockPrefix As String = "", Optional ByVal BlockSufix As String = "-placeholder", Optional ByVal AskIfAttached As Boolean = True) As ObjectId
            Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim myDB As Database = myDWG.Database
            Dim myEd As Editor = myDWG.Editor
            Dim BlockID As ObjectId

            Dim Trans As Transaction = myDB.TransactionManager.StartTransaction()
            Using Trans
                '' Get the block table from the drawing
                Dim myBT As BlockTable = DirectCast(Trans.GetObject(myDB.BlockTableId, OpenMode.ForRead), BlockTable)
                '' Check the block name, to see whether it's already in use
                Dim blkName As String
                blkName = BlockPrefix & MapTile.Name & BlockSufix
                Try
                    '' Check that the block name is valid
                    SymbolUtilityServices.ValidateSymbolName(blkName, False)
                    If myBT.Has(blkName) Then

                        If AskIfAttached Then
                            Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                            pKeyOpts.Message = vbLf & "A placeholder for tile " & MapTile.Name & " already exists. Do you wish to load in another copy?"
                            pKeyOpts.Keywords.Add("Yes")
                            pKeyOpts.Keywords.Add("No")
                            pKeyOpts.Keywords.Default = "No"
                            pKeyOpts.AllowNone = True
                            Dim pKeyRes As PromptResult = myEd.GetKeywords(pKeyOpts)

                            Select Case pKeyRes.StringResult
                                Case "Yes"
                                    BlockID = myBT.Item(blkName)
                                    Return BlockID
                                Case Else
                                    Return Nothing
                            End Select
                        Else
                            Return Nothing
                        End If

                    End If
                Catch
                    ''If the block name is invalid
                    myEd.WriteMessage(vbLf & blkName & " is not a valid block name.")
                    Return Nothing
                End Try

                '' Create the new block definition
                Dim BlockDef As New BlockTableRecord()
                ' ... and set its properties
                BlockDef.Name = blkName
                ' Add the new block to the block table
                myBT.UpgradeOpen()
                BlockID = myBT.Add(BlockDef)
                Trans.AddNewlyCreatedDBObject(BlockDef, True)

                ''Create the block contents and add them to the block definition
                Dim BlockContents As DBObjectCollection = CreateBlockContents(MapTile)
                For Each Item As DBObject In BlockContents
                    BlockDef.AppendEntity(Item)
                    Trans.AddNewlyCreatedDBObject(Item, True)
                Next

                ''Commit the transaction
                Trans.Commit()

            End Using
            Return BlockID

        End Function

        Public Function CreateBlockContents(ByVal MapTile As MapTile) As DBObjectCollection
            Dim BlockContents As New DBObjectCollection()
            Dim BlockContentsColour As Color = Color.FromColorIndex(ColorMethod.ByBlock, 0)
            Dim BlockContentsLayer As String = "0"

            ''Add the MText label with the tile name for the middle of the block
            Dim LabelMText As Entity = MapTile.LabelMText
            BlockContents.Add(LabelMText)

            ''Add the polyline outline for the tile
            Dim Outline As Entity = MapTile.Outline
            BlockContents.Add(Outline)

            ' ''###########################################################################
            ' ''This doesn't work, blocks need ATTSYNCing?!
            ' ''Add the Attributes to the block
            'Dim AttDef As New AttributeDefinition
            'AttDef.Position = New Point3d(MapTile.Easting, MapTile.Northing, 0)
            'AttDef.HorizontalMode = TextHorizontalMode.TextLeft
            'AttDef.VerticalMode = TextVerticalMode.TextBottom
            'AttDef.AlignmentPoint = AttDef.Position
            ''AttDef.Visible = False
            ''AttDef.Constant = True

            'AttDef.Tag = "MAP_TYPE"
            'AttDef.Prompt = "Map Tile Type:"
            'AttDef.TextString = MapTile.Type
            'Dim AttDefMapType As AttributeDefinition = AttDef
            'BlockContents.Add(AttDefMapType)
            ' ''###########################################################################

            For Each Item As Entity In BlockContents
                Item.Layer = BlockContentsLayer
                Item.Color = BlockContentsColour
            Next

            Return BlockContents
        End Function
    End Class

End Namespace
