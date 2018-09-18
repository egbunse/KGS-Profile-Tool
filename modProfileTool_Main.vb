Imports System.Math
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports System.Threading
Imports ArcGIS.Core.Geometry
Imports ArcGIS.Core.Data
Imports ArcGIS.Desktop.Core.Geoprocessing
Imports ArcGIS.Desktop.Mapping
Imports ArcGIS.Core.CIM
Imports ArcGIS.Desktop.Editing
Imports ArcGIS.Desktop.Editing.Attributes
Imports ArcGIS.Desktop.Framework.Threading.Tasks
Imports System.Threading.CancellationTokenSource
Imports ArcGIS.Desktop.Core

Module modProfileTool_Main


    '* Created 24Jul07 by:
    '* James L. Poelstra
    '* Email: james.arcscripts.help@nym.hush.com
    '*
    '* This code is directly based on work by:
    '* Michael Moex Maxelon
    '* Geologisches Institut, ETH Zentrum,8092 Zürich, Switzerland
    '* Dec 14 2004
    '* http://e-collection.ethbib.ethz.ch/show?type=bericht&nr=377
    '*
    '* My thanks to Michael for allowing this version of his work to be
    '* shared with others via ESRI.

    'Originally written in VBA
    'Converted to VB.net, April 2015 by:
    'Kristen Jordan
    'Kansas Data Access and Support Center
    'Email: kristen@kgs.ku.edu

    'Converted for ArcGIS Pro, April 2018 by: 
    'Emily Bunse
    'Kansas Geological Survey- Cartographic Services GRA
    'Email: egbunse@ku.edu; egbunse@gmail.com

    'General
    Private m_iTOCSurface As Integer 'the TOC position number for the geology layer
    Private m_iTOCElevation As Integer 'the TOC position number for the DEM

    'For grabbing map units for profile start and end clicks
    Private m_dStartXstore As Double = 0
    Private m_dStartYstore As Double = 0
    Private m_fLastClickTime As Date 'the time of last click
    Private m_dStartX As Double = 0
    Private m_dStartY As Double = 0
    Private m_dEndX As Double = 0
    Private m_dEndY As Double = 0

    'For Intersection Points
    Private m_adIntersectionPoints(0, 0) As Double '([intersection x],[intersection y],[distance from start of profile line])
    Private m_iIntersectionPointsCount As Integer 'the total number of intersections (duplicates removed)

    'For the profile tool
    Private m_dMapUnitsToUserUnits As Double
    Private m_dUserUnitsToMapUnits As Double
    Private m_dElevationLayerUnitsToMapUnits As Double
    Private m_dSubsurfaceTableUnitsToMapUnits As Double

    Private m_dElevDist As Double 'distance for elevation points (meters)
    Private m_dXstart As Double 'X coord of the profile's startpoint
    Private m_dYstart As Double 'Y coord of the profile's startpoint
    Private m_dXend As Double 'X coord of the profile's endpoint
    Private m_dYend As Double 'Y coord of the profile's endpoint
    Private m_sRasterLayerName As String 'stores the name of the esriDataSourcesRaster.Raster Layer
    Private m_sSurfaceLayerName As String 'stores the name of the esriGeometry.Polygon Feature Layer
    Private m_sSurfaceUnitID As String 'stores the name of the Geology Unit-ID esriGeoDatabase.Field
    Private m_sWorkspaceFolder As String 'stores the name of the esriGeoDatabase.Workspace Path
    Private m_sProfileName As String 'stores the name of the Polyline Shape (i.e. profile-line)
    Private m_sProfileLineName As String 'stores the name for the profile line (line on the map)
    Private m_sIntersectionPointsName As String 'Stores the name for the profile intersection points
    Private m_sBoxMajorName As String
    Private m_sBoxMinorName As String
    Private m_sReadmeFileName As String
    Private m_sWellsName As String
    Private m_sProfileDataFileName As String 'name of text file to store the coords for the profile

    'Delegate Sub StringArgReturningVoidDelegate([text] As String)

    'these are used for calling the profile box, and also reported in the readme file
    Private m_dBoxL As Double 'length of the profile box, if it is created
    Private m_dVTickMajor As Double 'spacing of the major vertical tick for the profile box, if it is created
    Private m_dVTickMinor As Double 'spacing of the minor vertical tick for the profile box, if it is created

    Private m_iVerticalExagg As Integer 'stores the vertical exaggeration to use
    Private m_adProfileData(0, 0) As Double '([rdX,rdY,rdZ,wrtX,wrtY,pdType],ArrayNumber)
    Private m_dProfileDataRdXStep As Double 'the step for rdX in the m_adProfileData array
    Private m_dProfileDataWrtXStep As Double 'the step for wrtX in the m_adProfileData array
    Private m_pSpatialReference As SpatialReference 'the spatial reference to use for all output files

    Private Const PI As Double = 3.14159265358979 'number used for PI
    '--Vars needed for running geoprocessing tools
    Dim tool_path As String
    Dim args As IReadOnlyList(Of String)
    Dim result As IGPResult
    Dim intersectionPoints As List(Of MapPoint) = New List(Of MapPoint)
    Dim subsetIntersectionPts As List(Of MapPoint) = New List(Of MapPoint)
    Dim profileOverlay As IDisposable
    Dim fullPolyline As List(Of Coordinate2D) = New List(Of Coordinate2D)

    Public startX As Double = 0
    Public startY As Double = 0
    Public endX As Double = 0
    Public endY As Double = 0

    Dim frmPT As frmProfileTool = New frmProfileTool()

    Public Sub SetProgBarLabelText(progBarLabel As Label, textToAdd As String)
        If progBarLabel.InvokeRequired Then
            progBarLabel.Invoke(Sub() SetProgBarLabelText(progBarLabel, textToAdd))
        Else
            progBarLabel.Text = textToAdd
            progBarLabel.Refresh()
            frmPT.Refresh()
        End If
    End Sub

    Public Sub RefreshProgBarValue(progBar As ProgressBar, val As Long)
        If progBar.InvokeRequired Then
            progBar.Invoke(Sub() RefreshProgBarValue(progBar, val))
        Else
            progBar.Value = val
            progBar.Refresh()
            frmPT.Refresh()
        End If
    End Sub

    Public Sub SetProgBarMax(progBar As ProgressBar, val As Integer)
        If progBar.InvokeRequired Then
            progBar.Invoke(Sub() SetProgBarMax(progBar, val))
        Else
            Debug.Print("This is the progress bar maximum value: " + Str(val))
            progBar.Minimum = 0
            progBar.Maximum = val
            progBar.Refresh()
            frmPT.Refresh()
        End If
    End Sub

    Public Sub StartGeoProfile(ByVal pOrigPoint As MapPoint)
        ' Get the map units for the start and end points of the profile
        ' and show the user form.

        ' This grabs the map units for the location of the user's click.
        Dim pPoint As MapPoint = pOrigPoint 'the location of the mouse click

        ' This will allow the user's click for the start point to expire
        ' if an end point click is not detected within the specified number
        ' of seconds.  The default is 10 seconds.

        Dim Seconds_to_wait As Integer
        Seconds_to_wait = 10

        Dim dtNow As Date = Now()

        If m_fLastClickTime = Nothing Then
            m_fLastClickTime = My.Computer.Clock.LocalTime
        End If

        If m_fLastClickTime < dtNow.AddSeconds(Seconds_to_wait * -1) Then
            m_dStartXstore = 0
            m_fLastClickTime = dtNow
        End If

        If m_dStartXstore = 0 Then
            'this is the first user click (start of the profile)
            'store this point and wait for second user click (end of the profile)
            m_dStartXstore = pPoint.X
            m_dStartYstore = pPoint.Y

        ElseIf pPoint.X <= m_dStartXstore Then
            'this is the second user click (end of the profile),
            'but the profile is east-west or vertical,
            'which is not allowed

            'reset user clicks
            m_dStartXstore = 0
            'inform user of the problem
            MsgBox("Profiles must be West to East." & Chr(13) & "Please redefine the profile.", vbOKOnly, "Error")
        Else
            'this is the second user click (end of the profile)
            'save the profile start and end coords
            m_dEndX = pPoint.X
            m_dEndY = pPoint.Y
            m_dStartX = m_dStartXstore
            m_dStartY = m_dStartYstore
            m_dStartXstore = 0
            m_dStartYstore = 0

            startX = m_dStartX
            startY = m_dStartY
            endX = m_dEndX
            endY = m_dEndY

            Debug.Print("Start: " + (CStr(Round(startX, 5))).ToString + "," + (CStr(Round(startY, 5))).ToString)
            Debug.Print("End: " + (CStr(Round(endX, 5))).ToString + "," + (CStr(Round(endY, 5))).ToString)

            Try
                PopulateTextBoxes(frmPT.txtStartX, (CStr(Round(startX, 5))).ToString)
                PopulateTextBoxes(frmPT.txtStartY, (CStr(Round(startY, 5))).ToString)
                PopulateTextBoxes(frmPT.txtEndX, (CStr(Round(endX, 5))).ToString)
                PopulateTextBoxes(frmPT.txtEndY, (CStr(Round(endY, 5))).ToString)
                ShowForm()
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
        End If
    End Sub

    Public Sub PopulateTextBoxes(textBox As TextBox, coordVal As String)
        If textBox.InvokeRequired Then
            Debug.Print("Invoking " + textBox.Name)
            textBox.Invoke(Sub() PopulateTextBoxes(textBox, coordVal))
        Else
            Debug.Print("Setting Value for " + textBox.Name)
            textBox.Text = coordVal
        End If
    End Sub

    'Opens the main form
    Public Sub ShowForm()
        If frmPT.InvokeRequired Then
            Debug.Print("Invoking the form.")
            frmPT.Invoke(Sub() ShowForm())
        Else
            Debug.Print("Showing the form.")
            frmPT.ShowDialog()
        End If
    End Sub

    'Closes the main form
    Public Sub CloseForm()
        If frmPT.InvokeRequired Then
            frmPT.Invoke(Sub() CloseForm())
        Else
            Debug.Print("Closing Form...")
            frmPT.Close()

        End If
    End Sub

    Public Async Sub GpMain(ByVal v_iTOCSurfaceNumber As Integer, ByVal v_iTOCElevationNumber As Integer,
        ByVal v_dUserUnitsToMapUnits As Double, ByVal v_dMapUnitsToUserUnits As Double,
        ByVal v_dElevationLayerUnitsToMapUnits As Double, ByVal v_dSubsurfaceTableUnitsToMapUnits As Double,
        ByVal frmProfileTool As FrmProfileTool)
        ' This sub controls everything by calling subfunctions for each major step of the analysis
        m_iTOCSurface = v_iTOCSurfaceNumber 'the TOC position for the surface layer
        m_iTOCElevation = v_iTOCElevationNumber 'the TOC position for the elevation layer
        m_dUserUnitsToMapUnits = v_dUserUnitsToMapUnits
        m_dMapUnitsToUserUnits = v_dMapUnitsToUserUnits
        m_dElevationLayerUnitsToMapUnits = v_dElevationLayerUnitsToMapUnits
        m_dSubsurfaceTableUnitsToMapUnits = v_dSubsurfaceTableUnitsToMapUnits

        'store path for output file directory
        m_sWorkspaceFolder = frmProfileTool.txtOutputPath.Text

        'see if storage is a geodatabase
        Dim shp As String
        If m_sWorkspaceFolder.Contains(".gdb") Then
            shp = ""
        Else
            shp = ".shp"
        End If

        'store names for output files
        Dim sProfileTag As String
        sProfileTag = frmProfileTool.txtProfileName.Text

        m_sProfileName = sProfileTag & "_profile" & shp
        m_sProfileLineName = sProfileTag & "_line" & shp
        m_sIntersectionPointsName = sProfileTag & "_intersection_points" & shp
        m_sBoxMajorName = sProfileTag & "_box_major" & shp
        m_sBoxMinorName = sProfileTag & "_box_minor" & shp
        m_sReadmeFileName = sProfileTag & "_readme"
        m_sWellsName = sProfileTag & "_wells" & shp
        m_sProfileDataFileName = sProfileTag & "_profile_coords"

        'assign variables based on the user form
        Call VariableAssigner(frmProfileTool)

        'get spatial reference from the map document
        Dim pMap As Map = MapView.Active.Map
        m_pSpatialReference = pMap.SpatialReference
        'set multiplier for elevation grid values
        If frmProfileTool.optElevationFt.Checked = True Then 'ft
            'Changed to find the units from the Active Map's Spatial Reference- EGB
            If pMap.SpatialReference.Unit.Equals(LinearUnit.Feet) Then  'pMap.MapUnits = 3 Then 'ft
                m_dElevationLayerUnitsToMapUnits = 1
                m_dUserUnitsToMapUnits = 1
            Else 'meters
                m_dElevationLayerUnitsToMapUnits = 381 / 1250 'ft to m conversion
                m_dUserUnitsToMapUnits = 381 / 1250 'ft to m conversion
            End If
        Else 'm
            If pMap.SpatialReference.Unit.Equals(LinearUnit.Feet) Then 'pMap.MapUnits = 3 Then 'ft
                m_dElevationLayerUnitsToMapUnits = 1250 / 381 'm to ft conversion
                m_dUserUnitsToMapUnits = 1250 / 381 'm to ft conversion
            Else 'meters
                m_dElevationLayerUnitsToMapUnits = 1
                m_dUserUnitsToMapUnits = 1
            End If
        End If
        Dim gp_result As Boolean

        Try
            'build the wells output file, if nessesary
            If frmProfileTool.chkAddWells.Checked = True Then
                'build the wells output file
                If m_dSubsurfaceTableUnitsToMapUnits = 0 Then
                    If frmProfileTool.optWellUnitsFt.Checked = True Then 'ft
                        If pMap.SpatialReference.Unit.FactoryCode = 9002 Then                  'pMap.MapUnits = 3 Then 'ft
                            m_dSubsurfaceTableUnitsToMapUnits = 1
                        Else 'meters
                            m_dSubsurfaceTableUnitsToMapUnits = 381 / 1250 'ft to m conversion
                        End If
                    Else 'm
                        If pMap.SpatialReference.Unit.FactoryCode = 9002 Then                   'pMap.MapUnits = 3 Then 'ft
                            m_dSubsurfaceTableUnitsToMapUnits = 1250 / 381 'm to ft conversion
                        Else 'meters
                            m_dSubsurfaceTableUnitsToMapUnits = 1
                        End If
                    End If
                End If
                pMap = Nothing

                frmProfileTool.cps = New CancelableProgressorSource()
                frmProfileTool.cps.Progressor.Max = 2
                Await GetSelectedWellPointsForProfileAsync(
                                m_dXstart, m_dYstart, m_dXend,
                                m_dYend, m_iVerticalExagg,
                                m_dSubsurfaceTableUnitsToMapUnits,
                                m_pSpatialReference, m_sWellsName,
                                m_sWorkspaceFolder, frmProfileTool, m_sRasterLayerName, frmProfileTool.cps)
                frmProfileTool.cps = Nothing
                'Since the wells shapefile gets added to the table of contents, the geology polygon shapefile is pushed further down the list
                m_iTOCSurface += 1
            End If

            'create the temp profile line layer
            Try
                SetProgBarLabelText(frmProfileTool.lblPrgBr, "Saving Profile Line...")
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
            Debug.Print(Str(startX) + "," + Str(startY))
            Debug.Print(Str(m_dXstart) + "," + Str(m_dYstart))
            Debug.Print(Str(endX) + "," + Str(endY))
            Debug.Print(Str(m_dXend) + "," + Str(m_dYend))
            If Not ((startX = m_dXstart And startY = m_dYstart) And (endX = m_dXend And endY = m_dYend)) Then
                gp_result = Await ModifySketch()
            End If

            frmProfileTool.cps = New CancelableProgressorSource()
            frmProfileTool.cps.Progressor.Max = 2

            Await CreateTempProfileLineAsync(frmProfileTool.cps)
            frmProfileTool.cps = Nothing

            If frmProfileTool.optProfileDelineationOn.Checked = True Then
                Dim transectLine As New List(Of Coordinate2D)() From {
                                        New Coordinate2D(m_dXstart, m_dYstart),
                                        New Coordinate2D(m_dXend, m_dYend)
                                        }
                Dim profileLineGeom As Geometry = Await QueuedTask.Run(Function()
                                                                           Return New PolylineBuilder(transectLine).ToGeometry()
                                                                       End Function)
                frmProfileTool.cps = New CancelableProgressorSource()
                frmProfileTool.cps.Progressor.Max = 2
                gp_result = Await CreateLineFeatureAsync(m_sWorkspaceFolder, m_sProfileLineName, profileLineGeom, frmProfileTool.cps)
                frmProfileTool.cps = Nothing
                Await AddLayerToMapAsync(m_sWorkspaceFolder, m_sProfileLineName)
                Debug.Print("Result CreateLineFeatureAsync: " + gp_result.ToString)
                'select the temp profile line
                gp_result = Await SelectFirstFeature(m_sWorkspaceFolder, m_sProfileLineName)
                Debug.Print("Result SelectFirstFeature: " + gp_result.ToString)
                'finds points where the temp profile line intersects the surface layer
                Try
                    SetProgBarLabelText(frmProfileTool.lblPrgBr, "Finding geologic surface intersections...")
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try
                frmProfileTool.cps = New CancelableProgressorSource()
                frmProfileTool.cps.Progressor.Max = 2
                gp_result = Await FindIntersections(frmProfileTool.cps)
                frmProfileTool.cps = Nothing
                Debug.Print("Result FindIntersections: " + gp_result.ToString)
                'remove the temporary profile line from the map
                Await RemoveTempProfileLine()
                'output the intersection points to a shapefile
                Try
                    SetProgBarLabelText(frmProfileTool.lblPrgBr, "Saving intersections...")
                Catch ex As Exception
                    Debug.Print(ex.Message)
                End Try
                frmProfileTool.cps = New CancelableProgressorSource()
                frmProfileTool.cps.Progressor.Max = 2
                Await SaveIntersectionPointsShp(frmProfileTool.cps)
                frmProfileTool.cps = Nothing
                If profileOverlay IsNot Nothing Then profileOverlay.Dispose()
                If ProfileTool.graphic IsNot Nothing Then ProfileTool.graphic.Dispose()
                Await MappingExtensions.ClearSketchAsync(MapView.Active)
                Await AddLayerToMapAsync(m_sWorkspaceFolder, m_sIntersectionPointsName)

            End If

            Try
                SetProgBarLabelText(frmProfileTool.lblPrgBr, "Preparing Profile Feature Class...")
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
            'create the necessary output shapefiles
            frmProfileTool.cps = New CancelableProgressorSource()
            frmProfileTool.cps.Progressor.Max = 2
            Await ShapeConstruct(frmProfileTool.cps)
            frmProfileTool.cps = Nothing
            'calculate the coordinates for the profile
            Call CoordinateCalcProfile()
            'sort the m_adProfileData array
            Call BubbleSort()

            'retrieve the profile elevation values from the DEM
            gp_result = Await ElevationGripProfile(frmProfileTool)
            Debug.Print("Elevation Grip Profile result: " + gp_result.ToString)

            'Calculate Intersection Point Elevations
            Call CalculateIntersectionPointElevations()

            'Calculate Output Profile Elevations
            Call CalculateOutputProfileElevations()

            'write the coordinates in the line shapefile and
            'internally call a function to check the points' ID assignments
            gp_result = Await FillPolyline(frmProfileTool)
            Await AddLayerToMapAsync(m_sWorkspaceFolder, m_sProfileName)
            'Run Smooth Profile Tool 
            If frmProfileTool.smoothProfile.Checked = True Then
                Dim profileFeatLayer As String = ""
                Dim outputSmoothName As String = ""
                Dim profileNameNoShp As String = ""
                profileFeatLayer = m_sWorkspaceFolder + "\" + m_sProfileName
                If shp IsNot "" Then
                    profileNameNoShp = m_sProfileName.Remove(m_sProfileName.Length - 4, 4)
                Else
                    profileNameNoShp = m_sProfileName
                End If
                outputSmoothName = m_sWorkspaceFolder + "\" + profileNameNoShp + "_smooth" + shp
                Call SmoothProfileLine(profileFeatLayer, outputSmoothName)
            End If
            'prepare to build the profile box, and build it if nessesary
            gp_result = Await SetupProfileBox(frmProfileTool)

            'Write the profile parameters to a readme file
            Call WriteReadmeFile(frmProfileTool)
            'write the m_adProfileData array to a text file
            Call WhatIsInProfileDataArray()
            Call CloseTool("The profile was successfully created. Files can be found at: " + m_sWorkspaceFolder + Environment.NewLine +
                    "Deactivate the Sketch Tool by clicking the Explore button." + Environment.NewLine +
                   "Have a nice day.", vbInformation + vbOKOnly, "That was it ...")

        Catch ex As TaskCanceledException
            CloseForm()
        End Try
    End Sub

    Private Sub VariableAssigner(ByVal frmProfileTool As frmProfileTool)
        ' load all variable values from "frmProfileTool" into this module

        With frmProfileTool

            m_dXstart = CDbl(.txtStartX.Text)
            m_dYstart = CDbl(.txtStartY.Text)
            m_dXend = CDbl(.txtEndX.Text)
            m_dYend = CDbl(.txtEndY.Text)

            'use the most accurate breaks for the type of profile requested

            'If .optPrecisionParts.Text <> "" Or .optPrecisionParts.Text <> " " Then 'parts
            If .optPrecisionParts.Checked Then
                'round to whole number
                .txtPrecisionParts.Text = CStr(CInt(.txtPrecisionParts.Text))

                'if the user requested only 1 part, change it to two parts
                If .txtPrecisionParts.Text = "1" Then
                    .txtPrecisionParts.Text = "2"
                End If

                m_dElevDist = Sqrt(((CDbl(m_dYend) - CDbl(m_dYstart)) ^ 2 + (CDbl(m_dXend) - CDbl(m_dXstart)) ^ 2)) / CDbl(.txtPrecisionParts.Text)

            Else 'distance

                m_dElevDist = CDbl(.txtPrecisionMeasure.Text) 'KJjuneedit

            End If

            m_sRasterLayerName = .lstElevationLayer.Text

            If frmProfileTool.optProfileDelineationOn.Checked = True Then
                m_sSurfaceLayerName = .lstSurfaceUnitLayer.Text
                m_sSurfaceUnitID = .lstSurfaceUnitField.Text
            End If

            m_iVerticalExagg = CInt(.txtVerticalExagg.Text)

        End With
    End Sub

    Private Async Function ModifySketch() As Task(Of Boolean)
        Debug.Print("modify sketch")
        Dim newLine As LineSegment
        profileOverlay = Await QueuedTask.Run(Function()
                                                  newLine = LineBuilder.CreateLineSegment(MapPointBuilder.CreateMapPoint(m_dXstart, m_dYstart),
                                                                         MapPointBuilder.CreateMapPoint(m_dXend, m_dYend), MapView.Active.Map.SpatialReference)
                                                  Return MappingExtensions.AddOverlay(MapView.Active, New PolylineBuilder(newLine).ToGeometry(), ProfileTool.pLineSymbolRef)
                                              End Function)
        Return True
    End Function

    'https://github.com/Esri/arcgis-pro-sdk/wiki/ProSnippets-Geodatabase#creating-a-feature
    Private Async Function CreateLineFeatureAsync(ByVal v_sFolder As String, ByVal v_sLayerName As String, ByVal geometryFeatToAdd As Polyline, ByVal cps As CancelableProgressorSource) As Task(Of Boolean)
        'Private Async Function createFeature(ByVal layerToModify As String, ByVal geometryFeatToAdd As List(Of Coordinate2D), ByVal geometryType As String) As Task(Of Boolean)
        Debug.Print("Create Line Feature Function")
        'Private Async Function createFeature(ByVal featureClass As String, ByVal gdbURI As String, ByVal geometryFeatToAdd As List(Of Coordinate2D), ByVal geometryType As String) As Task
        Dim message As String = String.Empty
        Dim pFeatureClass_Current As FeatureClass
        Dim spatialRef As SpatialReference = Nothing
        Dim creationResult As Boolean = False
        'Dim fsConnectionPath As FileSystemConnectionPath
        'Dim shapefile As FileSystemDatastore
        Dim geomToAdd As Geometry = Nothing
        Dim saveResult As Boolean = False
        Await QueuedTask.Run(Function()
                                 'Attempt to add in handling of the Exit button, but doesn't reset the tool
                                 'Code example: https://github.com/Esri/arcgis-pro-sdk-community-samples/blob/b76a6625dc3fa2e4287c1b39ab4b95e3a8f2b04b/Framework/ProgressDialog/ProgressDialogModule.cs
                                 'While Not cps.CancellationTokenSource.Token.IsCancellationRequested
                                 '    cps.Progressor.Value += 1
                                 '    cps.Progressor.Status = (cps.Progressor.Value * 100 / cps.Progressor.Max) & " % Completed"
                                 '    cps.Progressor.Message = "Message " & cps.Progressor.Value

                                 '    If Debugger.IsAttached Then
                                 '        Debug.WriteLine(String.Format("RunCancelableProgress Loop{0}", cps.Progressor.Value))
                                 '    End If

                                 'Open the new Feature Class
                                 If v_sFolder.Contains(".gdb") Then
                                     Dim fgdbConnectionPath As FileGeodatabaseConnectionPath = New FileGeodatabaseConnectionPath(New Uri(v_sFolder))
                                     Dim filegdb As Geodatabase = New Geodatabase(fgdbConnectionPath)
                                     pFeatureClass_Current = filegdb.OpenDataset(Of FeatureClass)(v_sLayerName)
                                 Else
                                     Dim fsConnectionPath As FileSystemConnectionPath = New FileSystemConnectionPath(New Uri(v_sFolder), FileSystemDatastoreType.Shapefile)
                                     Dim shapefile As FileSystemDatastore = New FileSystemDatastore(fsConnectionPath)
                                     pFeatureClass_Current = shapefile.OpenDataset(Of FeatureClass)(v_sLayerName) '--May need to grab the last part of this string 
                                 End If
                                 Debug.Print("Create Line FC: " + pFeatureClass_Current.GetName())
                                 Using pFeatureClass_Current    'Should already be open
                                     'Set up tools needed to add Polyline to Feature Class as a Feature 
                                     Dim rowBuff As RowBuffer = Nothing
                                     Dim featToAdd As Feature = Nothing
                                     Try
                                         Dim editOperation As New EditOperation With {
                                    .Name = "Create Feature in " + pFeatureClass_Current.GetName(),
                                    .SelectNewFeatures = True,
                                    .SelectModifiedFeatures = True
                                 }
                                         editOperation.Callback(Function(context)
                                                                    Dim fcDef As FeatureClassDefinition = pFeatureClass_Current.GetDefinition()
                                                                    rowBuff = pFeatureClass_Current.CreateRowBuffer()
                                                                    rowBuff("Shape") = geometryFeatToAdd
                                                                    featToAdd = pFeatureClass_Current.CreateRow(rowBuff)
                                                                    'To Indicate that the attribute table has to be updated
                                                                    context.Invalidate(featToAdd)
                                                                    Return True
                                                                End Function, CType(pFeatureClass_Current, Dataset))
                                         creationResult = editOperation.Execute()
                                         Try
                                             Dim objectID As Long = featToAdd.GetObjectID()
                                             Debug.Print("Creation result for " + objectID.ToString + " " + creationResult.ToString)
                                         Catch ref As NullReferenceException
                                             Debug.Print(ref.Message)
                                         End Try
                                         saveResult = Project.Current.SaveEditsAsync().Result
                                         Debug.Print("Save Result: " + saveResult.ToString)
                                         Return saveResult
                                     Catch ex As GeodatabaseEditingException
                                         Console.WriteLine("Gdb error in createAddIntersections: " + ex.Message)
                                     Finally
                                         If rowBuff IsNot Nothing Then
                                             rowBuff.Dispose()
                                         End If
                                         If featToAdd IsNot Nothing Then
                                             featToAdd.Dispose()
                                         End If
                                     End Try
                                 End Using

                                 '    If cps.Progressor.Value = cps.Progressor.Max Then Exit While
                                 'End While
                                 Return True
                             End Function, cps.Progressor)


        Return True
    End Function

    Private Async Function CreatePointFeatureAsync(ByVal v_sFolder As String, ByVal v_sLayerName As String, ByVal geometryFeatToAdd As List(Of Coordinate2D), ByVal cps As CancelableProgressorSource) As Task(Of Boolean)
        Dim creationResult As Boolean = False
        Dim message As String = String.Empty
        Dim pFeatureClass_Current As FeatureClass
        Dim spatialRef As SpatialReference = Nothing
        Dim geomToAdd As Geometry = Nothing
        Dim saveResult As Boolean = False
        Await QueuedTask.Run(Function()
                                 'Open the new Feature Class
                                 If v_sFolder.Contains(".gdb") Then
                                     Dim fgdbConnectionPath As FileGeodatabaseConnectionPath = New FileGeodatabaseConnectionPath(New Uri(v_sFolder))
                                     Dim filegdb As Geodatabase = New Geodatabase(fgdbConnectionPath)
                                     pFeatureClass_Current = filegdb.OpenDataset(Of FeatureClass)(v_sLayerName)
                                 Else
                                     Dim fsConnectionPath As FileSystemConnectionPath = New FileSystemConnectionPath(New Uri(v_sFolder), FileSystemDatastoreType.Shapefile)
                                     Dim shapefile As FileSystemDatastore = New FileSystemDatastore(fsConnectionPath)
                                     pFeatureClass_Current = shapefile.OpenDataset(Of FeatureClass)(v_sLayerName) '--May need to grab the last part of this string 
                                 End If
                                 Using pFeatureClass_Current    'Should already be open
                                     'Set up tools needed to add Polyline to Feature Class as a Feature 
                                     Dim rowBuff As RowBuffer = Nothing
                                     Dim featToAdd As Feature = Nothing
                                     Try
                                         Dim editOperation As New EditOperation With {
                                        .Name = "Create Feature in " + pFeatureClass_Current.GetName(),
                                        .SelectNewFeatures = True,
                                        .SelectModifiedFeatures = True
                                     }
                                         editOperation.Callback(Function(context)
                                                                    Dim fcDef As FeatureClassDefinition = pFeatureClass_Current.GetDefinition()
                                                                    rowBuff = pFeatureClass_Current.CreateRowBuffer()
                                                                    For Each coord In geometryFeatToAdd
                                                                        Debug.Print(coord.X.ToString + ", " + coord.Y.ToString)
                                                                        geomToAdd = New MapPointBuilder(coord.X, coord.Y).ToGeometry()
                                                                        rowBuff(fcDef.GetShapeField()) = geomToAdd
                                                                        featToAdd = pFeatureClass_Current.CreateRow(rowBuff)
                                                                    Next
                                                                    'To Indicate that the attribute table has to be updated
                                                                    context.Invalidate(featToAdd)
                                                                    Return True
                                                                End Function, CType(pFeatureClass_Current, Dataset))
                                         creationResult = editOperation.Execute()
                                         Try
                                             Dim objectID As Long = featToAdd.GetObjectID()
                                             Debug.Print("Creation result for " + objectID.ToString + " " + creationResult.ToString)
                                         Catch ref As NullReferenceException
                                             Debug.Print(ref.Message)
                                         End Try
                                         saveResult = Project.Current.SaveEditsAsync().Result

                                         Debug.Print("Save result: " + saveResult.ToString)
                                     Catch ex As GeodatabaseEditingException
                                         Console.WriteLine("Gdb error in createAddIntersections: " + ex.Message)
                                     Finally
                                         If rowBuff IsNot Nothing Then
                                             rowBuff.Dispose()
                                         End If
                                         If featToAdd IsNot Nothing Then
                                             featToAdd.Dispose()
                                         End If
                                     End Try
                                 End Using
                                 Return saveResult
                             End Function, cps.Progressor)
        Return saveResult
    End Function

    Private Async Function CreateTempProfileLineAsync(ByVal cps As CancelableProgressorSource) As Task
        'Create a temporary polyline shapefile that contains the profile line.
        Debug.Print("Create Temp Profile Line")
        'setup name and location for the temporary shapefile
        Dim type As String = "Shape"
        '--As per this post: https://geonet.esri.com/thread/189471-is-it-possible-to-specify-fields-to-create-feature-class-with
        '  the Pro API does not include interfaces to perform DDL actions, only DML. 
        '  Therefore, they suggest using Geoprocessing Tools to do any data creation, deletion, etc. 

        'While Not cps.CancellationTokenSource.Token.IsCancellationRequested
        '    cps.Progressor.Value += 1
        '    cps.Progressor.Status = (cps.Progressor.Value * 100 / cps.Progressor.Max) & " % Completed"
        '    cps.Progressor.Message = "Message " & cps.Progressor.Value

        '    If Debugger.IsAttached Then
        '        Debug.WriteLine(String.Format("RunCancelableProgress Loop{0}", cps.Progressor.Value))
        '    End If

        'Create the new Feature Class
        '--Run the Create Feature Class Tool
        tool_path = "management.CreateFeatureClass"
            args = Geoprocessing.MakeValueArray(m_sWorkspaceFolder, m_sProfileLineName, "POLYLINE", Nothing, "DISABLED", "DISABLED", m_pSpatialReference)
            result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, cps.CancellationTokenSource.Token, Nothing, GPExecuteToolFlags.AddToHistory)
            Debug.Print("Executing create new feature class for " + m_sProfileLineName)
            If result.IsFailed Then
                MsgBox("Error creating: " & m_sProfileLineName & vbNewLine _
                                            & "Profile creation terminated.", vbCritical, "Error creating profile...")
                CloseForm()
            End If

            If Not m_sWorkspaceFolder.Contains(".gdb") Then
                tool_path = "management.AddSpatialIndex"
                Dim tempProfile As String = m_sWorkspaceFolder + "\" + m_sProfileLineName
                args = Geoprocessing.MakeValueArray(tempProfile, Nothing, Nothing, Nothing)
                result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, cps.CancellationTokenSource.Token, Nothing, GPExecuteToolFlags.AddToHistory)
                Debug.Print("Adding Spatial Index to: " + m_sProfileLineName)
                If result.IsFailed Then
                    MsgBox("Error creating: " & m_sProfileLineName & vbNewLine _
                                            & "Profile creation terminated.", vbCritical, "Try again in a file geodatabase...")
                    CloseForm()
                End If
            End If

        '    If cps.Progressor.Value = cps.Progressor.Max Then Exit While
        'End While
        tool_path = Nothing
        args = Nothing
        result = Nothing
    End Function

    Private Async Function AddLayerToMapAsync(ByVal v_sFolder As String, ByVal v_sLayerName As String) As Task
        ' This adds a layer to the map
        Debug.Print("Add layer to map")
        Dim featClassURI As Uri = New Uri(v_sFolder + "\" + v_sLayerName)
        Dim newLineLayer As Layer = Nothing
        Await QueuedTask.Run(Function()
                                 newLineLayer = LayerFactory.Instance.CreateLayer(featClassURI, MapView.Active.Map, 0)
                                 Return True
                             End Function)
    End Function

    Private Async Function SelectFirstFeature(ByVal v_sFolder As String, ByVal v_sLayerName As String) As Task(Of Boolean)
        Debug.Print("Select First Feature")
        Dim fc As FeatureClass = Nothing
        Dim shapefile As FileSystemDatastore
        Dim selectionResult As Boolean = False

        Await QueuedTask.Run(Function()
                                 If v_sFolder.Contains(".gdb") Then
                                     Dim fgdbConnectionPath As FileGeodatabaseConnectionPath = New FileGeodatabaseConnectionPath(New Uri(v_sFolder))
                                     Dim filegdb As Geodatabase = New Geodatabase(fgdbConnectionPath)
                                     fc = filegdb.OpenDataset(Of FeatureClass)(v_sLayerName)
                                 Else
                                     Dim fsConnectionPath As FileSystemConnectionPath = New FileSystemConnectionPath(New Uri(v_sFolder), FileSystemDatastoreType.Shapefile)
                                     shapefile = New FileSystemDatastore(fsConnectionPath)
                                     fc = shapefile.OpenDataset(Of FeatureClass)(v_sLayerName) '--May need to grab the last part of this string 
                                 End If
                                 Using fc
                                     Dim selectAll As Selection
                                     selectAll = fc.Select(Nothing, SelectionType.ObjectID, SelectionOption.Normal)
                                     'Debug.Print("Selection Count from SelectFirstFeature: " + selectAll.GetCount().ToString() + MapView.Active.Map.SelectionCount.ToString())
                                 End Using
                                 If MapView.Active.Map.SelectionCount > 0 Then
                                     Debug.Print(MapView.Active.Map.SelectionCount.ToString)
                                     selectionResult = True
                                 End If

                                 Return selectionResult
                             End Function)
        'free memory
        fc = Nothing
        Return True
    End Function

    Private Async Function FindIntersections(ByVal cps As CancelableProgressorSource) As Task(Of Boolean)
        Debug.Print("Find Intersections")
        'find points where a line layer intersects a polygon layer
        'must select a line in the line layer before running script.
        Dim pFeatureLayer As FeatureLayer
        Dim fcpFeatureLayer As FeatureClass
        Dim pFeatureLayer2 As FeatureLayer
        Dim pFeatureLayer2SpatRef As SpatialReference
        Dim fcpFeatureLayer2 As FeatureClass
        Dim pSpatialRef As SpatialReference = MapView.Active.Map.SpatialReference
        Dim selectionCount As Integer = MapView.Active.Map.SelectionCount
        Dim pGeometry As Geometry = Nothing
        Dim row As Row = Nothing
        Dim shape As Geometry = Nothing
        Dim intersectPts As Multipoint = Nothing
        Dim intersectionResult As Boolean = False

        'The line segment feature class's layer
        Await QueuedTask.Run(Function()
                                 'While Not cps.CancellationTokenSource.Token.IsCancellationRequested
                                 '    cps.Progressor.Value += 1
                                 '    cps.Progressor.Status = (cps.Progressor.Value * 100 / cps.Progressor.Max) & " % Completed"
                                 '    cps.Progressor.Message = "Message " & cps.Progressor.Value

                                 '    If Debugger.IsAttached Then
                                 '        Debug.WriteLine(String.Format("RunCancelableProgress Loop{0}", cps.Progressor.Value))
                                 '    End If

                                 Dim getXPointsResult As Boolean
                                     pFeatureLayer = MapView.Active.Map.Layers(0)
                                     fcpFeatureLayer = pFeatureLayer.GetFeatureClass()
                                     Dim selectAll As Selection
                                     Debug.Print("Layer selection: " + fcpFeatureLayer.GetName())
                                     Using fcpFeatureLayer
                                         selectAll = fcpFeatureLayer.Select(Nothing, SelectionType.ObjectID, SelectionOption.Normal)
                                     End Using
                                     'The intersection layer- denoted by user
                                     pFeatureLayer2 = MapView.Active.Map.Layers(m_iTOCSurface + 1)
                                     pFeatureLayer2SpatRef = pFeatureLayer2.GetSpatialReference()
                                     fcpFeatureLayer2 = pFeatureLayer2.GetFeatureClass()
                                     selectionCount = pFeatureLayer.GetSelection().GetCount()

                                     If selectAll.GetCount() > 0 Then
                                         Dim rowCursor As RowCursor = selectAll.Search(Nothing, False)
                                         Try
                                             Using rowCursor
                                                 While rowCursor.MoveNext() = True
                                                     Using feat As Feature = DirectCast(rowCursor.Current(), Feature)
                                                         shape = feat.GetShape()
                                                         getXPointsResult = GetXPoints(shape, fcpFeatureLayer2, pFeatureLayer2SpatRef).Result
                                                     End Using
                                                 End While
                                             End Using
                                         Catch ex As GeodatabaseEditingException
                                             Debug.Print("Gdb error in createAddIntersections: " + ex.Message)
                                         Finally
                                             If rowCursor IsNot Nothing Then
                                                 rowCursor.Dispose()
                                             End If
                                         End Try
                                     End If

                                     If intersectionPoints.Count > 0 Then
                                         intersectionResult = True
                                     Else
                                         MsgBox("There was a problem finding intersections.  " & vbNewLine _
                                            & "Profile creation terminated.", vbCritical, "Error creating profile...")
                                         CloseForm()
                                     End If

                                     'Save the intersection point coords for later use

                                     Dim n As Integer
                                     Dim iRawPointCount As Integer
                                     Dim adRawPoints(0, 0) As Double '(index,[x,y,trash flag])

                                     'iRawPointCount = intersectPts.PointCount
                                     iRawPointCount = intersectionPoints.Count
                                     Debug.Print("Number of Intersection Points: " + iRawPointCount.ToString)
                                     'if there are intersections, then prepare to process them, else exit this sub
                                     If iRawPointCount > 0 Then
                                         ReDim adRawPoints(iRawPointCount - 1, 2) '(index,[x,y,trash flag])
                                     Else 'there are no intersections
                                         m_iIntersectionPointsCount = 0
                                         ReDim m_adProfileData(5, 0) 'redim m_adProfileData array for the first time
                                         Exit Function
                                     End If

                                     'GetXPoints may return duplicate intersections points if polygons touch
                                     'so populate a temporary array with all of the intersection
                                     'points, then identify the duplicates
                                     'fill the temporary array
                                     For i = 0 To iRawPointCount - 1
                                         adRawPoints(i, 0) = intersectionPoints.Item(i).X
                                         adRawPoints(i, 1) = intersectionPoints.Item(i).Y
                                     Next

                                     Dim iNumberOfDuplicates As Integer
                                     iNumberOfDuplicates = 0

                                     'flag the duplicates
                                     For i = 0 To iRawPointCount - 1 'for each point
                                         For n = 0 To iRawPointCount - 1 'compare it to all other points
                                             If Not i = n And Not adRawPoints(i, 2) = 1 _
                                    And (Round(adRawPoints(i, 0), 5) = Round(adRawPoints(n, 0), 5) And
                                    Round(adRawPoints(i, 1), 5) = Round(adRawPoints(n, 1), 5)) Then 'this is a duplicate     
                                                 'For testing limiting by varying precisions use Format(Val adRawPoints('i' or 'n', 0), "0.000") 
                                                 adRawPoints(n, 2) = 1 ' using the value 1 to indicate a duplicate record
                                                 iNumberOfDuplicates = iNumberOfDuplicates + 1
                                             End If
                                         Next
                                     Next

                                     m_iIntersectionPointsCount = iRawPointCount - iNumberOfDuplicates
                                     Debug.Print("Number of intersection points going into FC: " + Str(m_iIntersectionPointsCount))

                                     ReDim m_adIntersectionPoints(m_iIntersectionPointsCount - 1, 2)

                                     'redim m_adProfileData array for the second time
                                     ReDim m_adProfileData(5, m_iIntersectionPointsCount - 1)

                                     n = 0
                                     Dim usedPoints As StringBuilder = New StringBuilder()
                                     For i = 0 To m_iIntersectionPointsCount - 1
                                         If Not adRawPoints(n, 2) = 1 Then 'add this point
                                             m_adIntersectionPoints(i, 0) = adRawPoints(n, 0) 'x
                                             m_adIntersectionPoints(i, 1) = adRawPoints(n, 1) 'y
                                             subsetIntersectionPts.Add(MapPointBuilder.CreateMapPoint(adRawPoints(n, 0), adRawPoints(n, 1), MapView.Active.Map.SpatialReference))
                                             m_adProfileData(0, i) = adRawPoints(n, 0) 'rdx
                                             m_adProfileData(1, i) = adRawPoints(n, 1) 'rdy
                                             m_adProfileData(5, i) = 1 'mark as intersection point
                                             n = n + 1
                                         Else 'this is a duplicate, so ignore this point
                                             i = i - 1
                                             n = n + 1
                                         End If
                                     Next
                                 '    If cps.Progressor.Value = cps.Progressor.Max Then Exit While
                                 'End While
                                 Return intersectionResult
                             End Function, cps.Progressor)
        'Conversion Inspired by: http://pro.arcgis.com/en/pro-app/sdk/api-reference/index.html#topic7640.html
        Return intersectionResult
    End Function

    Private Async Function RemoveTempProfileLine() As Task
        Debug.Print("Remove Temp Profile Line")
        ' This will remove the temporary profile line from the map
        Dim profileLayer As Layer = MapView.Active.Map.Layers(0)

        Await QueuedTask.Run(Sub() MapView.Active.Map.RemoveLayer(profileLayer))    'Will not delete from the disk... 
        'free memory
        profileLayer = Nothing
    End Function

    Private Async Function SaveIntersectionPointsShp(ByVal cps As CancelableProgressorSource) As Task
        Debug.Print("Save Intersection Points Shp")
        'This will create a temporary point shapefile that contains the profile intersection points.
        'Open the folder to contain the shapefile as a workspace
        Dim fc As FeatureClass = Nothing
        'While Not cps.CancellationTokenSource.Token.IsCancellationRequested
        '    cps.Progressor.Value += 1
        '    cps.Progressor.Status = (cps.Progressor.Value * 100 / cps.Progressor.Max) & " % Completed"
        '    cps.Progressor.Message = "Message " & cps.Progressor.Value

        '    If Debugger.IsAttached Then
        '        Debug.WriteLine(String.Format("RunCancelableProgress Loop{0}", cps.Progressor.Value))
        '    End If

        'Create the new Feature Class
        '--Run the Create Feature Class Tool
        tool_path = "management.CreateFeatureClass"
            args = Geoprocessing.MakeValueArray(m_sWorkspaceFolder, m_sIntersectionPointsName, "POINT", Nothing, "DISABLED", "DISABLED", m_pSpatialReference)
            result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, cps.CancellationTokenSource.Token, Nothing, GPExecuteToolFlags.AddToHistory)
            If result.IsFailed Then
                MsgBox("Error creating: " & m_sIntersectionPointsName & vbNewLine _
                                      & "Profile creation terminated.", vbCritical, "Error creating profile...")
                CloseForm()
            End If

            If Not m_sWorkspaceFolder.Contains(".gdb") Then
                tool_path = "management.AddSpatialIndex"
                Dim tempProfile As String = m_sWorkspaceFolder + "\" + m_sIntersectionPointsName
                args = Geoprocessing.MakeValueArray(tempProfile, Nothing, Nothing, Nothing)
                result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, cps.CancellationTokenSource.Token, Nothing, GPExecuteToolFlags.AddToHistory)
                If result.IsFailed Then
                    MsgBox("Error creating: " & m_sProfileLineName & vbNewLine _
                                       & "Profile creation terminated.", vbCritical, "Try again in a file geodatabase...")
                    CloseForm()
                End If
            End If
            tool_path = Nothing
            args = Nothing
            result = Nothing

            Try
                SetProgBarMax(frmPT.prgBr, m_iIntersectionPointsCount)
            Catch ex As Exception
                Debug.Print("SetProgBarMax error in Save Intersections: " + ex.Message)
            End Try
            subsetIntersectionPts = subsetIntersectionPts.OrderBy(Function(mpPt) mpPt.X).ToList   'Order in ascending order based on the value of X
            Debug.Print("Points in intersectionPoints: " + subsetIntersectionPts.Count.ToString)

            For Each pt In subsetIntersectionPts
                Dim newCoords As New List(Of Coordinate2D)() From {
                                       New Coordinate2D(pt.X, pt.Y)
                                   }
                Await CreatePointFeatureAsync(m_sWorkspaceFolder, m_sIntersectionPointsName, newCoords, cps)
                Try
                    RefreshProgBarValue(frmPT.prgBr, subsetIntersectionPts.IndexOf(pt))
                Catch ex As Exception
                    Debug.Print("RefreshProgBarValue error in SaveIntersections: " + ex.Message)
                End Try
            Next
        '    If cps.Progressor.Value = cps.Progressor.Max Then Exit While
        'End While
    End Function

    Private Async Function ShapeConstruct(ByVal ct As CancelableProgressorSource) As Task
        Debug.Print("Shape Construct")
        ' constructs a Polyline shapefile (for the profile line)
        'Create the new Feature Class
        '--Run the Create Feature Class Tool 
        'tool_path = "management.CreateFeatureClass"
        args = Geoprocessing.MakeValueArray(m_sWorkspaceFolder, m_sProfileName, "POLYLINE", Nothing, "DISABLED", "DISABLED", m_pSpatialReference)
        tool_path = "management.CreateFeatureClass"
        result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, ct.CancellationTokenSource.Token, Nothing, GPExecuteToolFlags.AddToHistory)
        If result.IsFailed Then
            MsgBox("Error creating: " & m_sProfileName & vbNewLine _
                & "Profile creation terminated.", vbCritical, "Error creating profile...")
            CloseForm()
        End If

        'Add Fields to the new Feature Class 
        tool_path = "management.AddField"
        Dim fcPath As String = m_sWorkspaceFolder + "\" + m_sProfileName
        Dim fieldList = New Object(0, 3) {{"Surface", "TEXT", 25, "surface"}}
        For i = 0 To fieldList.GetUpperBound(0)
            Dim fieldName As String = fieldList(i, 0)
            Dim fieldType As String = fieldList(i, 1)
            Dim fieldPrec As Double = fieldList(i, 2)
            Dim fieldAlias As String = fieldList(i, 3)
            Debug.Print(fieldName + fieldType + Str(fieldPrec) + fieldAlias)
            args = Geoprocessing.MakeValueArray(fcPath, fieldName, fieldType, fieldPrec, "", "", fieldAlias, "NULLABLE", "NON_REQUIRED", "")
            result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, ct.CancellationTokenSource.Token, Nothing, GPExecuteToolFlags.AddToHistory)
            If result.IsFailed Then
                For Each msg In result.Messages
                    Debug.Print(msg.ToString)
                Next
                MsgBox("Error adding fields to: " & m_sProfileName & vbNewLine _
                        & "Profile creation terminated.", vbCritical, "Error creating profile...")
                CloseForm()
            End If
        Next

        If Not m_sWorkspaceFolder.Contains(".gdb") Then
            tool_path = "management.AddSpatialIndex"
            Dim tempProfile As String = m_sWorkspaceFolder + "\" + m_sProfileName
            args = Geoprocessing.MakeValueArray(tempProfile, Nothing, Nothing, Nothing)
            result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, ct.CancellationTokenSource.Token, Nothing, GPExecuteToolFlags.AddToHistory)
            If result.IsFailed Then
                MsgBox("Error creating: " & m_sProfileName & vbNewLine _
                & "Profile creation terminated.", vbCritical, "Try again in a file geodatabase...")
                CloseForm()
            End If
        End If

        'free memory
        tool_path = Nothing
        args = Nothing
        result = Nothing
    End Function

    Private Sub CoordinateCalcProfile()
        Debug.Print("coordinate calc profile")
        ' calculates the points for the new polyline assemblage
        Dim sRelationship As String = Nothing 'indicates the relationship between the profile's start-/endpoint
        Dim dAlpha As Double 'angle between a horizontal or vertical line and profile trend
        Dim dProfileLength As Double 'length of the profile
        Dim dTotalNumberOfPoints As Double 'total number of points
        'Assign the case indicator including cases of equal x- or y-coords
        If m_dXstart < m_dXend And m_dYstart > m_dYend Then 'down
            sRelationship = "-11"
        ElseIf m_dXstart <= m_dXend And m_dYstart <= m_dYend Then 'up
            sRelationship = "-1-1"
        End If

        'Prevent division by zero in coordinate's angle calculation
        If Not (m_dXend = m_dXstart Or m_dYend = m_dYstart) Then
            dAlpha = Atan((Abs(m_dYend - m_dYstart) / Abs(m_dXend - m_dXstart)) ^
                                Sign((m_dYend - m_dYstart) / (m_dXend - m_dXstart)))
        Else
            If m_dXend = m_dXstart Then
                dAlpha = PI / 2
            Else
                dAlpha = 0
            End If
        End If

        'Calculate the length of the profile
        dProfileLength = Sqrt(((CDbl(m_dYend) - CDbl(m_dYstart)) ^ 2 + (CDbl(m_dXend) - CDbl(m_dXstart)) ^ 2))

        'Calculate the number of profile points and resize the coordinate array
        Debug.Print("m_dElevDist: " + Str(m_dElevDist))
        If Not (dProfileLength / m_dElevDist) = Int(dProfileLength / m_dElevDist) Then
            dTotalNumberOfPoints = Int(dProfileLength / m_dElevDist) + 2 + m_iIntersectionPointsCount
        Else
            dTotalNumberOfPoints = Int(dProfileLength / m_dElevDist) + 1 + m_iIntersectionPointsCount
        End If

        ReDim Preserve m_adProfileData(5, dTotalNumberOfPoints - 1)

        Dim lArrayCounter As Long

        Select Case sRelationship
            Case "-11" 'down
                For lArrayCounter = m_iIntersectionPointsCount To dTotalNumberOfPoints - 2
                    m_adProfileData(0, lArrayCounter) = CDbl(m_dXstart) + CDbl(lArrayCounter - m_iIntersectionPointsCount) * m_dElevDist * Sin(dAlpha) 'rdx
                    m_adProfileData(1, lArrayCounter) = CDbl(m_dYstart) - CDbl(lArrayCounter - m_iIntersectionPointsCount) * m_dElevDist * Cos(dAlpha) 'rdy
                    m_adProfileData(3, lArrayCounter) = CDbl(lArrayCounter - m_iIntersectionPointsCount) * m_dElevDist 'wrtx
                Next
            Case "-1-1" 'up
                For lArrayCounter = m_iIntersectionPointsCount To dTotalNumberOfPoints - 2
                    m_adProfileData(0, lArrayCounter) = CDbl(m_dXstart) + CDbl(lArrayCounter - m_iIntersectionPointsCount) * m_dElevDist * Cos(dAlpha) 'rdx
                    m_adProfileData(1, lArrayCounter) = CDbl(m_dYstart) + CDbl(lArrayCounter - m_iIntersectionPointsCount) * m_dElevDist * Sin(dAlpha) 'rdy
                    m_adProfileData(3, lArrayCounter) = CDbl(lArrayCounter - m_iIntersectionPointsCount) * m_dElevDist 'wrtx
                Next
        End Select

        m_adProfileData(0, lArrayCounter) = CDbl(m_dXend) 'rdx
        m_adProfileData(1, lArrayCounter) = CDbl(m_dYend) 'rdy
        m_adProfileData(3, lArrayCounter) = CDbl(lArrayCounter - m_iIntersectionPointsCount) * m_dElevDist 'wrtx

        'now calculate the m_dProfileDataRdXStep and m_dProfileDataWrtXStep
        m_dProfileDataRdXStep = m_adProfileData(0, m_iIntersectionPointsCount + 1) - m_adProfileData(0, m_iIntersectionPointsCount)
        m_dProfileDataWrtXStep = m_dElevDist


    End Sub

    Private Sub BubbleSort()
        Debug.Print("Bubble Sort")
        ' This code originaly from:
        ' Chris Rae's VBA Code Archive - http://chrisrae.com/vba
        ' By Chris Rae, 19/5/99. Chris thanks
        ' Will Rickards and Roemer Lievaart for some fixes.
        ' James Poelstra customized Chris's code for this application.

        Dim bAnyChanges As Boolean
        Dim lBubbleSort As Long
        Dim dSwapRdX As Double
        Dim dSwapRdY As Double
        Dim dSwapWrtX As Double
        Dim dSwapPdType As Double

        Do
            bAnyChanges = False
            For lBubbleSort = LBound(m_adProfileData, 2) To UBound(m_adProfileData, 2) - 1
                If m_adProfileData(0, lBubbleSort) > m_adProfileData(0, lBubbleSort + 1) Then
                    ' These two need to be swapped
                    'store old values
                    dSwapRdX = m_adProfileData(0, lBubbleSort)
                    dSwapRdY = m_adProfileData(1, lBubbleSort)
                    dSwapWrtX = m_adProfileData(3, lBubbleSort)
                    dSwapPdType = m_adProfileData(5, lBubbleSort)
                    'replace with new values
                    m_adProfileData(0, lBubbleSort) = m_adProfileData(0, lBubbleSort + 1)
                    m_adProfileData(1, lBubbleSort) = m_adProfileData(1, lBubbleSort + 1)
                    m_adProfileData(3, lBubbleSort) = m_adProfileData(3, lBubbleSort + 1)
                    m_adProfileData(5, lBubbleSort) = m_adProfileData(5, lBubbleSort + 1)
                    'restore old values to next location
                    m_adProfileData(0, lBubbleSort + 1) = dSwapRdX
                    m_adProfileData(1, lBubbleSort + 1) = dSwapRdY
                    m_adProfileData(3, lBubbleSort + 1) = dSwapWrtX
                    m_adProfileData(5, lBubbleSort + 1) = dSwapPdType
                    bAnyChanges = True
                End If
            Next lBubbleSort
        Loop Until Not bAnyChanges

    End Sub

    Private Async Function ElevationGripProfile(ByVal frmProfileTool As frmProfileTool) As Task(Of Boolean)
        Debug.Print("elevation grip profile")
        ' Identifies the raster containing the elevation data and
        ' assigns the data to the respective points ProfileData array
        Dim pRasterLayer As RasterLayer
        Dim pRaster As Raster.Raster
        'look for the specified raster layer
        Dim pMap As Map = MapView.Active.Map
        For Each pLayer In pMap.Layers
            If pLayer.Name = m_sRasterLayerName Then
                pRasterLayer = pLayer
                pRaster = Await QueuedTask.Run(Function() pRasterLayer.GetRaster)
            End If
        Next

        Try
            SetProgBarMax(frmPT.prgBr, UBound(m_adProfileData, 2))
        Catch ex As Exception
            Debug.Print("SetProgBarMax error in ElevationGripProfile: " + ex.Message)
        End Try

        'loop through all points
        Dim pPoint As MapPoint
        Dim ptCounter As Integer
        For lPointCounter = 0 To UBound(m_adProfileData, 2)
            ptCounter = lPointCounter
            If m_adProfileData(5, lPointCounter) = 1 Then
                'this is an intersection point, so it will NOT have a raster value
                'allowing the profile line to maintain a consistent precision
            Else 'get the raster value for this point
                'this must be reset each time to prevent an error in the raster identify function
                pPoint = Await QueuedTask.Run(Function() MapPointBuilder.CreateMapPoint(m_adProfileData(0, ptCounter), m_adProfileData(1, ptCounter), m_pSpatialReference))
                m_adProfileData(2, lPointCounter) = Await QueuedTask.Run(Function()
                                                                             Return GetRasterValueForThisPoint(pPoint, m_sRasterLayerName).Result
                                                                         End Function)
                pPoint = Nothing
            End If
            'update progress bar
            Try
                RefreshProgBarValue(frmPT.prgBr, lPointCounter)
            Catch ex As Exception
                Debug.Print("RefreshProgBarValue error in ElevationGripProfile: " + ex.Message)
            End Try
        Next lPointCounter
        Return frmPT.prgBr.Value = frmPT.prgBr.Maximum
    End Function

    Private Sub CalculateIntersectionPointElevations()
        Debug.Print("calculate intersection point elevations")
        ' This will calculate the elevations for intersection points assuming a linear
        ' relationship between the previous and next precision point.  This will allow
        ' the precision of the output profile to remain constant over different geology.
        Dim l As Long
        Dim n As Long
        Dim dPrevX As Double
        Dim dPrevY As Double
        Dim dNextX As Double
        Dim dNextY As Double
        Dim dPrevRdX As Double

        For l = 0 To UBound(m_adProfileData, 2)
            If m_adProfileData(5, l) = 1 Then 'this has been flagged as an intersection point
                'get the dNextX and dNextY values
                For n = l To UBound(m_adProfileData, 2)
                    If Not m_adProfileData(5, n) = 1 Then
                        dNextX = m_adProfileData(3, n) 'wrtx
                        dNextY = m_adProfileData(2, n) 'rdz
                        Exit For
                    End If
                Next
                'calculate the wrtx value
                m_adProfileData(3, l) = m_dProfileDataWrtXStep * ((m_adProfileData(0, l) - dPrevRdX) / m_dProfileDataRdXStep) + dPrevX
                'calculate the elevation for this intersection point
                m_adProfileData(2, l) = fnLinear_yBetweenPoints(dPrevX, dPrevY, dNextX, dNextY, m_adProfileData(3, l))
            Else 'save this rdX and rdY for possible later use
                dPrevRdX = m_adProfileData(0, l) 'rdx (used to calculate the wrtx for intersection points)
                dPrevX = m_adProfileData(3, l) 'wrtx
                dPrevY = m_adProfileData(2, l) 'rdz
            End If
        Next
    End Sub

    Private Sub CalculateOutputProfileElevations()
        Debug.Print("calculate output profile elevations")
        ' This calculates the elevations for the output profile
        Dim l As Long
        For l = 0 To UBound(m_adProfileData, 2)
            m_adProfileData(4, l) = m_adProfileData(2, l) * m_iVerticalExagg
        Next
    End Sub

    Private Async Function CreateAddIntersections(ByVal geoToAdd As Geometry,
                                                  ByVal id As Integer, ByVal geoUnit As String) As Task(Of Boolean)
        Debug.Print("create add intersections")
        Dim creationResult As Boolean = False
        Dim message As String = String.Empty
        Dim saveResult As Boolean = False
        Dim fc As FeatureClass = Nothing
        Dim fsConnectionPath As FileSystemConnectionPath
        Dim shapefile As FileSystemDatastore
        Dim filegdb As Geodatabase
        Dim fgdbConnectionPath As FileGeodatabaseConnectionPath
        Await QueuedTask.Run(Function()
                                 If m_sWorkspaceFolder.Contains(".gdb") Then
                                     fgdbConnectionPath = New FileGeodatabaseConnectionPath(New Uri(m_sWorkspaceFolder))
                                     filegdb = New Geodatabase(fgdbConnectionPath)
                                     fc = filegdb.OpenDataset(Of FeatureClass)(m_sProfileName)
                                 Else
                                     fsConnectionPath = New FileSystemConnectionPath(New Uri(m_sWorkspaceFolder), FileSystemDatastoreType.Shapefile)
                                     shapefile = New FileSystemDatastore(fsConnectionPath)
                                     fc = shapefile.OpenDataset(Of FeatureClass)(m_sProfileName)
                                 End If
                                 Using fc
                                     Dim rowBuff As RowBuffer = Nothing
                                     Dim featToAdd As Feature = Nothing
                                     Try
                                         'Set up tools needed to add Polyline to Feature Class as a Feature 
                                         Dim editOperation As New EditOperation With {
                                            .Name = "Create Feature in " + fc.GetName(),
                                            .SelectNewFeatures = True,
                                            .SelectModifiedFeatures = True
                                         }
                                         editOperation.Callback(Function(context)
                                                                    Dim fcDef As FeatureClassDefinition = fc.GetDefinition()
                                                                    rowBuff = fc.CreateRowBuffer()
                                                                    'Add fields info
                                                                    rowBuff("ID") = id
                                                                    rowBuff("Surface") = geoUnit
                                                                    rowBuff(fcDef.GetShapeField()) = geoToAdd
                                                                    featToAdd = fc.CreateRow(rowBuff)
                                                                    'To Indicate that the attribute table has to be updated
                                                                    context.Invalidate(featToAdd)
                                                                    Return True
                                                                End Function, CType(fc, Dataset))
                                         creationResult = editOperation.Execute()
                                         Try
                                             Dim objectID As Long = featToAdd.GetObjectID()
                                             Debug.Print("Creation result for " + objectID.ToString + " " + creationResult.ToString)
                                         Catch ref As NullReferenceException
                                             Debug.Print(ref.Message)
                                         End Try
                                         saveResult = Project.Current.SaveEditsAsync().Result
                                         Debug.Print("Save result: " + saveResult.ToString)
                                     Catch ex As GeodatabaseEditingException
                                         Console.WriteLine("Gdb error in createAddIntersections: " + ex.Message)
                                     Finally
                                         If rowBuff IsNot Nothing Then
                                             rowBuff.Dispose()
                                         End If
                                         If featToAdd IsNot Nothing Then
                                             featToAdd.Dispose()
                                         End If
                                     End Try
                                 End Using
                                 Return saveResult
                             End Function)
        Return saveResult
    End Function

    Private Async Function FillPolyline(ByVal frmProfileTool As frmProfileTool) As Task(Of Boolean)
        'Adds line segments to the profile shapefile/feature class 

        Debug.Print("Fill polyline")
        Dim diff As Integer = 0
        Dim iCaseSelector As Integer 'specifies what the function "CheckPointID" is used for
        Dim lSegmentCounter As Long 'counter for 'for...next-loops'
        Dim bPointsSameId As Boolean 'specifies if 2 points are in the same polygon
        Dim pLine As LineSegment = Nothing 'a single 2-point-line
        Dim pPoint_IdLookupPoint As MapPoint = Nothing 'point to be passed to the function that looks up id values
        Dim geoVal As String = "" 'The surface geology value- derived in the GetGeologyValueForThisPoint Sub and used in the createAddIntersections Function
        Dim iID As Integer 'counter for the ID
        Dim geomToAdd As Geometry = Nothing
        Dim profileSegGeo As List(Of Coordinate2D) = New List(Of Coordinate2D)
        Dim ckGeoSegment As LineSegment
        Dim startReached As Boolean = False
        iID = 0
        'Set Progress Label Message
        Try
            SetProgBarLabelText(frmPT.lblPrgBr, "Writing profile results ...")
        Catch ex As Exception
            Debug.Print("SetProbBarLabelText error in FillPolyline: " + ex.Message)
        End Try
        'Reset Progress Bar Value 
        Try
            SetProgBarMax(frmPT.prgBr, UBound(m_adProfileData, 2) - 1)
        Catch ex As Exception
            Debug.Print("SetProgBarMax error in FillPolyline: " + ex.Message)
        End Try
        Dim lineCreation As Boolean = False
        Dim startPointSet As Boolean = False

        'loop through all points on the polyline
        For lSegmentCounter = 0 To UBound(m_adProfileData, 2) - 1

            iCaseSelector = 1
            profileSegGeo.Add(New Coordinate2D(m_adProfileData(3, lSegmentCounter), m_adProfileData(4, lSegmentCounter)))
            'check if the two points of a segment are in the same geology polygon
            If Not m_adProfileData(5, lSegmentCounter + 1) = 1 Then 'no intersection
                bPointsSameId = True
            Else 'there is an intersection at end of segment or this is the last segment
                bPointsSameId = False
            End If

            'at boundary or at end of profile ...
            If bPointsSameId = False _
                Or lSegmentCounter = UBound(m_adProfileData, 2) - 1 Then

                'add the current line segment to the collection
                'assign the ID and geology of this profile part
                iID = iID + 1
                'Get the Geology Info to Add to Feature
                profileSegGeo.Add(New Coordinate2D(m_adProfileData(3, lSegmentCounter + 1), m_adProfileData(4, lSegmentCounter + 1)))

                Await QueuedTask.Run(Function()
                                         If lSegmentCounter = UBound(m_adProfileData, 2) - 1 Then
                                             'this is the end of the profile, so get the geology
                                             'using the end point for the profile
                                             ckGeoSegment = LineBuilder.CreateLineSegment(MapPointBuilder.CreateMapPoint(m_adProfileData(0, lSegmentCounter), m_adProfileData(1, lSegmentCounter), MapView.Active.Map.SpatialReference),
                                                                                          MapPointBuilder.CreateMapPoint(m_adProfileData(0, lSegmentCounter + 1), m_adProfileData(1, lSegmentCounter + 1), MapView.Active.Map.SpatialReference))
                                             pPoint_IdLookupPoint = GeometryEngine.Instance.Centroid(New PolylineBuilder(ckGeoSegment).ToGeometry())
                                             Debug.Print("End Point: " + Str(pPoint_IdLookupPoint.X) + ", " + Str(pPoint_IdLookupPoint.Y))
                                         ElseIf startPointSet = False Then
                                             ckGeoSegment = LineBuilder.CreateLineSegment(MapPointBuilder.CreateMapPoint(m_dXstart, m_dYstart, MapView.Active.Map.SpatialReference),
                                                                                          MapPointBuilder.CreateMapPoint(m_adProfileData(0, lSegmentCounter + 1), m_adProfileData(1, lSegmentCounter + 1), MapView.Active.Map.SpatialReference))
                                             pPoint_IdLookupPoint = GeometryEngine.Instance.Centroid(New PolylineBuilder(ckGeoSegment).ToGeometry())
                                             Debug.Print("Start Point: " + Str(pPoint_IdLookupPoint.X) + ", " + Str(pPoint_IdLookupPoint.Y))
                                             startPointSet = True
                                         Else
                                             'get the geology using a median point
                                             'Find the next intersection and record x, y
                                             Dim nextX As Double
                                             Dim nextY As Double
                                             Dim nextPoint As MapPoint
                                             nextPoint = subsetIntersectionPts.Item(iID - 2)
                                             nextX = nextPoint.X
                                             nextY = nextPoint.Y
                                             Debug.Print("Current Point: " + Str(m_adProfileData(0, lSegmentCounter + 1)) + "," + Str(m_adProfileData(1, lSegmentCounter + 1)))
                                             Debug.Print("Previous Point: " + Str(nextX) + "," + Str(nextY))

                                             ckGeoSegment = LineBuilder.CreateLineSegment(MapPointBuilder.CreateMapPoint(m_adProfileData(0, lSegmentCounter + 1), m_adProfileData(1, lSegmentCounter + 1), MapView.Active.Map.SpatialReference),
                                                                                          MapPointBuilder.CreateMapPoint(nextX, nextY, MapView.Active.Map.SpatialReference))
                                             pPoint_IdLookupPoint = GeometryEngine.Instance.Centroid(New PolylineBuilder(ckGeoSegment).ToGeometry())
                                             Debug.Print("Mid Point: " + Str(pPoint_IdLookupPoint.X) + ", " + Str(pPoint_IdLookupPoint.Y))
                                         End If
                                         Return True
                                     End Function)

                If frmProfileTool.optProfileDelineationOn.Checked = True Then
                    geoVal = Await QueuedTask.Run(Function()
                                                      Return GetGeologyValueForThisPoint(pPoint_IdLookupPoint, m_iTOCSurface + 1).Result
                                                  End Function)
                Else
                    geoVal = "unknown"
                End If

                'Add the line to the Feature Class
                Try
                    Debug.Print("ID: " + Str(iID))
                    Debug.Print("Geology Value: " + geoVal)
                    'Check if any points have already been added
                    Await QueuedTask.Run(Function()
                                             geomToAdd = New PolylineBuilder(profileSegGeo).ToGeometry()
                                             Return True
                                         End Function)
                    Await createAddIntersections(geomToAdd, iID, geoVal)
                    fullPolyline.AddRange(profileSegGeo)
                    'Clear the list of Coordinates to start collecting for the next segment
                    geomToAdd = Nothing
                    profileSegGeo.Clear()
                Catch ex As AggregateException
                    For Each x In ex.InnerExceptions
                        Debug.Print("Aggregate Ex. error in FillPolyline after createAddIntersections: " + x.Message)
                    Next
                End Try
            End If
            Try
                RefreshProgBarValue(frmPT.prgBr, lSegmentCounter)
            Catch ex As Exception
                Debug.Print("RefreshProgBarValue error in FillPolyline: " + ex.Message)
            End Try
        Next lSegmentCounter

        'Clear out last selected features
        Await QueuedTask.Run(Function()
                                 Dim pLayer As BasicFeatureLayer = Nothing
                                 Dim pLayers As IEnumerable(Of Layer) = MapView.Active.Map.Layers
                                 pLayer = pLayers(m_iTOCSurface + 1)
                                 pLayer.ClearSelection()
                                 Return True
                             End Function)
        Return frmPT.prgBr.Value = frmPT.prgBr.Maximum
    End Function

    Private Sub SmoothProfileLine(ByVal profile As String, ByVal smoothProfile As String)
        'Run Smooth Line Tool From Cartography Toolbox
        tool_path = "cartography.SmoothLine"
        args = Geoprocessing.MakeValueArray(profile, smoothProfile, "PAEK", "Meters", "NO_FIXED", "NO_CHECK")
        'When tool is opened the tolerance doesn't get reflected?? How to format the parameter??
        Geoprocessing.OpenToolDialog(tool_path, args)
    End Sub

    Private Async Function GetGeologyValueForThisPoint(ByVal v_pPoint_Lookup As MapPoint, ByVal v_iTOCLayerNumber As Integer) As Task(Of String)
        'This will return the surface unit for the specified point
        Debug.Print("get geology value for this point")
        Dim surfaceUnit As String = ""
        Await QueuedTask.Run(Function()
                                 Try
                                     Dim pLayer As BasicFeatureLayer
                                     Dim pMap As Map = MapView.Active.Map
                                     Dim pLayers As IEnumerable(Of Layer) = pMap.Layers
                                     pLayer = pLayers(v_iTOCLayerNumber)
                                     Dim ptFeatures As Dictionary(Of BasicFeatureLayer, List(Of Long)) = MapView.Active.GetFeatures(v_pPoint_Lookup)
                                     MapView.Active.SelectFeatures(v_pPoint_Lookup)
                                     Dim ptOID As List(Of Long)
                                     Dim insp As Inspector = New Inspector()
                                     If ptFeatures.Count > 0 Then
                                         For Each kvp In ptFeatures
                                             Debug.Print(kvp.Key.Name)
                                             If kvp.Key.Name.Equals(pLayer.Name) Then    'Find the layer containing geologic surfaces in the set of features
                                                 ptOID = kvp.Value
                                                 Debug.Print("Possible geologic unit values for Segment: ")
                                                 'Working with the inspector: http://pro.arcgis.com/en/pro-app/sdk/api-reference/index.html#topic9684.html
                                                 'Old Code (ProfileTool_VB) retrieved first geologic unit in 'GetFeatures' (used Identify) array
                                                 For Each value In kvp.Value
                                                     Debug.Print(value.ToString)
                                                     insp.Load(pLayer, value)
                                                     If GeometryEngine.Instance.Contains(insp(insp.GeometryAttribute.FieldIndex), v_pPoint_Lookup) Then
                                                         surfaceUnit = insp(m_sSurfaceUnitID)
                                                         Debug.Print("Geologic Unit in get geoval: " + surfaceUnit)
                                                         Exit For
                                                     End If
                                                 Next
                                             End If
                                         Next
                                     Else
                                         surfaceUnit = "error"
                                     End If
                                 Catch ex As AggregateException
                                     For Each x In ex.InnerExceptions
                                         Debug.Print("Aggregate Ex. error in GetGeologyValueForThisPoint: " + x.Message)
                                     Next
                                 End Try

                                 Return surfaceUnit
                             End Function)
        Return surfaceUnit
    End Function

    Private Async Function GetXPoints(ByVal v_pGeometry As Geometry, ByVal v_pFeatureClass As FeatureClass, v_pFC_SpatRef As SpatialReference) As Task(Of Boolean)
        Debug.Print("get x points")
        ' This is called from within the FindIntersections sub
        ' It finds the intersections for the selected feature (the temp profile line)
        Dim fcFeatShape As Geometry = Nothing
        Dim intersectGeo As Geometry = Nothing
        Dim intersectCur As RowCursor = Nothing
        Dim pSpatialFilter As SpatialQueryFilter = New SpatialQueryFilter With {
                    .SpatialRelationship = SpatialRelationship.Intersects,
                    .OutputSpatialReference = v_pFC_SpatRef,
                    .FilterGeometry = v_pGeometry
        }
        Dim intersectionMultiPoints As List(Of Multipoint) = New List(Of Multipoint)
        Await QueuedTask.Run(Function()
                                 Try
                                     intersectCur = v_pFeatureClass.Search(pSpatialFilter, False)
                                     Using intersectCur
                                         While intersectCur.MoveNext() = True
                                             Using feat As Feature = DirectCast(intersectCur.Current(), Feature)
                                                 fcFeatShape = feat.GetShape()
                                                 'Need to get MultiPOint to save all points of intersection to a point collection???
                                                 'Before performing the intersection make sure to check if one exists http://pro.arcgis.com/en/pro-app/sdk/api-reference/index.html#topic8255.html
                                                 If GeometryEngine.Instance.Intersects(fcFeatShape, v_pGeometry) Then
                                                     intersectionMultiPoints.Add(GeometryEngine.Instance.Intersection(fcFeatShape, v_pGeometry, GeometryDimension.esriGeometry0Dimension))
                                                     For Each multi In intersectionMultiPoints
                                                         For Each coord In multi.Copy2DCoordinatesToList
                                                             intersectionPoints.Add(MapPointBuilder.CreateMapPoint(coord.X, coord.Y, v_pFC_SpatRef))
                                                         Next
                                                     Next
                                                 End If
                                             End Using
                                         End While
                                     End Using
                                 Catch ex As AggregateException
                                     For Each x In ex.InnerExceptions
                                         Debug.Print("GetXPoints Exceptions: " + x.Message)
                                     Next
                                 Finally
                                     If intersectCur IsNot Nothing Then
                                         intersectCur.Dispose()
                                     End If
                                 End Try
                                 If intersectionPoints.Count > 0 Then
                                     Return True
                                 Else
                                     MessageBox.Show("Error collecting intersecting points.")
                                 End If
                                 Return True
                             End Function)
        'free memory
        fcFeatShape = Nothing
        intersectGeo = Nothing
        v_pFeatureClass = Nothing
        v_pGeometry = Nothing

        Return True
    End Function

    Private Sub WhatIsInProfileDataArray()
        Debug.Print("what is in profile data array")
        ' This dumps the coords for the profile into a text file.
        ' Note that if user specified the profile be divided into X parts, there will be more point
        ' than that in this output.  The output file also includes intersection points.
        Dim wsFolder As String = m_sWorkspaceFolder
        If m_sWorkspaceFolder.Contains(".gdb") Then
            Dim intPos As Integer
            intPos = m_sWorkspaceFolder.LastIndexOfAny("\")
            intPos += 1
            wsFolder = m_sWorkspaceFolder.Substring(0, intPos)
        End If

        Dim path As String = wsFolder & "\" & m_sProfileDataFileName & ".txt"

        'remove the file if it already exists
        If Dir(path, vbDirectory) <> "" Then
            Kill(path)
        End If

        Dim i As Integer
        If Not File.Exists(path) Then
            Using fs As System.IO.FileStream = File.Create(path)
                Dim headerinfo As Byte() = New UTF8Encoding(True).GetBytes("PointNO,ProfileX,ProfileY,RawElev,InterPt" & vbCrLf)    'Changed to add in marked intersection pts
                fs.Write(headerinfo, 0, headerinfo.Length)
                For i = 0 To UBound(m_adProfileData, 2)
                    Dim moreInfo As Byte() = New UTF8Encoding(True).GetBytes(i & "," & m_adProfileData(3, i) & "," & m_adProfileData(4, i) & "," & m_adProfileData(2, i) & "," & m_adProfileData(5, i) & vbCrLf)
                    fs.Write(moreInfo, 0, moreInfo.Length)
                Next
            End Using
        End If

    End Sub

    Public Async Function GetRasterValueForThisPoint(ByVal v_iPoint_ClickPoint As MapPoint, ByVal v_sRasterName As String) As Task(Of Double)
        'Debug.Print("get raster value for this point")
        'This returns the value of the DEM at the specified point

        Dim pRasterLayer As RasterLayer
        Dim pRaster As Raster.Raster
        Dim pMap As Map
        Dim pSR As SpatialReference = Nothing
        Dim pPoint As MapPoint
        Dim pPointGeom As Geometry
        Dim colRow As Tuple(Of Integer, Integer)
        Dim pixelValue As Double
        Dim rasterValue As Double

        Await QueuedTask.Run(Function()
                                 Try
                                     pRaster = Nothing
                                     pMap = MapView.Active.Map
                                     For Each pLayer In pMap.Layers
                                         If pLayer.Name = m_sRasterLayerName Then
                                             pRasterLayer = pLayer
                                             pRaster = pRasterLayer.GetRaster()
                                             pSR = pRaster.GetSpatialReference()
                                         End If
                                     Next
                                     pPoint = MapPointBuilder.CreateMapPoint(v_iPoint_ClickPoint, m_pSpatialReference)
                                     If pSR.Name <> m_pSpatialReference.Name Then
                                         'to do the identify, the point (in map coordinates) needs to be projected into the raster coordinates
                                         pPointGeom = GeometryEngine.Instance.Project(pPoint, pSR)
                                         pPoint = CType(pPointGeom, MapPoint)
                                     End If

                                     'Use MapToPixel? http://pro.arcgis.com/en/pro-app/sdk/api-reference/index.html#topic15405.html
                                     colRow = pRaster.MapToPixel(pPoint.X, pPoint.Y)
                                     pixelValue = CDbl(pRaster.GetPixelValue(0, colRow.Item1, colRow.Item2))
                                 Catch ex As AggregateException
                                     For Each x In ex.InnerExceptions
                                         Debug.Print("Aggregate Ex. error in GetRasterValueForThisPoint: " + x.Message)
                                     Next
                                     MsgBox("Point x=" & Str(v_iPoint_ClickPoint.X) & " y=" & Str(v_iPoint_ClickPoint.Y) & " has value " & Str(pixelValue), vbCritical, "Error!")
                                 End Try
                                 rasterValue = pixelValue * m_dElevationLayerUnitsToMapUnits
                                 Return rasterValue
                             End Function)
        Return rasterValue
    End Function

    Private Function FnLinear_m(ByVal v_dX1 As Double, ByVal v_dY1 As Double, ByVal v_dX2 As Double, ByVal v_dY2 As Double) As Double
        'get slope of line from two points
        If Not v_dX2 - v_dX1 = 0 Then
            FnLinear_m = (v_dY2 - v_dY1) / (v_dX2 - v_dX1)
        Else
            FnLinear_m = 0
        End If
    End Function

    Private Function FnLinear_b(ByVal v_dX1 As Double, ByVal v_dY1 As Double, ByVal v_dX2 As Double, ByVal v_dY2 As Double) As Double
        'y intercept for line from two points
        If Not v_dX1 = v_dX2 Then
            FnLinear_b = -(v_dY2 - v_dY1) / (v_dX2 - v_dX1) * v_dX1 + v_dY1
        Else
            MsgBox("No y intercept for a vertical line!" & vbNewLine _
              & "First Point: (" & v_dX1 & "," & v_dY1 & ")" & vbNewLine _
              & "Second Point: (" & v_dX2 & "," & v_dY2 & ")")
        End If
    End Function

    Private Function FnLinear_yBetweenPoints(ByVal v_X1 As Double, ByVal v_Y1 As Double,
        ByVal v_X2 As Double, ByVal v_Y2 As Double, ByVal v_NewX As Double) As Double
        'Debug.Print("fn linear y between points")
        'y value for point at x between to points
        FnLinear_yBetweenPoints = FnLinear_m(v_X1, v_Y1, v_X2, v_Y2) * v_NewX + FnLinear_b(v_X1, v_Y1, v_X2, v_Y2)
    End Function

    Private Function CONV_Min(ByVal Min As Integer) As Single
        Debug.Print("conv min")
        'this converts minutes into microsoft time...
        'can use any number of minutes
        CONV_Min = Min / 60 * (1 / 24)
    End Function

    Private Function CONV_Hr(ByVal Hr As Integer) As Single
        Debug.Print("conv hr")
        'this converts hours into microsoft time...
        'can use any integer to represent hours.
        'remember that the reference date is 12/30/1899
        CONV_Hr = Hr * (1 / 24)
    End Function

    Private Function CONV_Sec(ByVal Sec As Integer) As Single
        Debug.Print("conv sec")
        'this converts minutes into microsoft time...
        CONV_Sec = Sec / 60 * (1 / 60) * (1 / 24)
    End Function

    Private Function DateTimeIDTag()
        Debug.Print("date time id tag")
        ' This was designed to attatch a unique string (based on the current date/time
        ' to new files that might otherwise have the same name.  Output looks like: 08-03-2004_133538
        Dim DateTimeID As String
        Dim dtTime As Date = Now()
        DateTimeID = Str(dtTime)
        MsgBox(DateTimeID) 'KJ this needs some tweaking most likely
        'correct the hour
        If Len(dtTime) = 10 Then
            DateTimeID = DateTimeID & "_0" & Left(dtTime, 1) & Mid(dtTime, 3, 2) & Mid(dtTime, 6, 2)
        Else
            DateTimeID = DateTimeID & "_" & Left(dtTime, 2) & Mid(dtTime, 4, 2) & Mid(dtTime, 7, 2)
        End If
        'change hour to military time
        If Right(dtTime, 2) = "PM" Then
            DateTimeID = Left(DateTimeID, 11) & CStr(12 + CInt(Mid(DateTimeID, 12, 2))) & Right(DateTimeID, 4)
        End If
        DateTimeIDTag = DateTimeID
    End Function

    Private Async Function SetupProfileBox(ByVal frmProfileTool As frmProfileTool) As Task(Of Boolean)
        Debug.Print("setup profile box")
        If frmProfileTool.chkBuildProfileBox.Checked = True Then
            'User wants at least one profile box.
            'Setup the call for the first box and call it.
            With frmProfileTool
                If .txtBoxL.Enabled = True Then
                    'use the length that the user entered
                    m_dBoxL = CDbl(.txtBoxL.Text) * m_dUserUnitsToMapUnits 'KJjuneedit
                Else
                    'use the length of the profile
                    m_dBoxL = CDbl(Sqrt(((CDbl(m_dYend) - CDbl(m_dYstart)) ^ 2 + (CDbl(m_dXend) - CDbl(m_dXstart)) ^ 2)))
                End If

                If .chkSquareGrid.Checked = True Then
                    'calculate the distance to get a square grid
                    m_dVTickMajor = CDbl(.txtHTick.Text) '/ CDbl(m_iVerticalExagg) * m_dUserUnitsToMapUnits 'KJjuneedit
                Else
                    'use the value the user entered
                    m_dVTickMajor = CDbl(.txtVTick.Text) '* m_dUserUnitsToMapUnits 'KJjuneedit
                End If
                Try
                    Await ProfileBox(frmProfileTool, True, m_sSurfaceLayerName, m_dBoxL, CDbl(.txtBoxT.Text) * m_dUserUnitsToMapUnits,
                    CDbl(.txtBoxB.Text) * m_dUserUnitsToMapUnits, m_dVTickMajor, CDbl(.txtHTick.Text) * m_dUserUnitsToMapUnits,
                    m_iVerticalExagg, .chkBuildGrid.Checked, .chkSquareGrid.Checked, m_sWorkspaceFolder,
                    m_sBoxMajorName, True, True, "Building the Major profile box...")
                Catch ex As Exception
                    Debug.Print("Error building the major profile box: " + ex.Message)
                End Try

                'check to see if the user wants an additional profile box
                If .chkMinorTick.Checked = True Then
                    'user wants an additional profile box
                    If .chkSquareGrid.Checked = True Then
                        'calculate the distance to get a square grid
                        m_dVTickMinor = CDbl(.txtHTickMinor.Text) / CDbl(m_iVerticalExagg) '* m_dUserUnitsToMapUnits 'KJjuneedit
                    Else
                        'use the value the user entered
                        m_dVTickMinor = CDbl(.txtVTickMinor.Text) '* m_dUserUnitsToMapUnits 'KJjuneedit 
                    End If

                    Try
                        Await ProfileBox(frmProfileTool, True, m_sSurfaceLayerName, m_dBoxL, CDbl(.txtBoxT.Text) * m_dUserUnitsToMapUnits,
                        CDbl(.txtBoxB.Text) * m_dUserUnitsToMapUnits, m_dVTickMinor, CDbl(.txtHTickMinor.Text) * m_dUserUnitsToMapUnits,
                        m_iVerticalExagg, .chkBuildGrid.Checked, .chkSquareGrid.Checked, m_sWorkspaceFolder,
                        m_sBoxMinorName, True, False, "Building the Minor profile box...")
                    Catch ex As Exception
                        Debug.Print("Error building the minor profile box: " + ex.Message)
                    End Try
                End If
            End With
        End If
        Return True
    End Function

    Private Async Sub WriteReadmeFile(ByVal frmProfileTool As frmProfileTool)
        'Write the profile parameters to a readme file
        'strip extra slashes off folder name

        If m_sWorkspaceFolder(Len(m_sWorkspaceFolder) - 1) = "\" Then
            While m_sWorkspaceFolder.Substring(Len(m_sWorkspaceFolder) - 1) = "\"
                m_sWorkspaceFolder = m_sWorkspaceFolder.Remove(Len(m_sWorkspaceFolder) - 1)
            End While
        End If

        Dim wsFolder As String = m_sWorkspaceFolder
        Debug.Print("wsFolder: " + wsFolder)

        If m_sWorkspaceFolder.Contains(".gdb") Then
            'get the root name of the folder containing the geodatabase where the text files were saved
            Dim intPos As Integer
            intPos = m_sWorkspaceFolder.LastIndexOfAny("\")
            intPos += 1
            wsFolder = m_sWorkspaceFolder.Substring(0, intPos)
        End If

        Dim path As String = wsFolder & "\" & m_sReadmeFileName & ".txt"
        Debug.Print("Path used: " + path)

        'remove the readme if it already exists
        If Dir(path, vbDirectory) <> "" Then
            Kill(path)
        End If

        If Not File.Exists(path) Then
            Using fs As FileStream = File.Create(path)
                Dim retrievedPath As String = ""
                Dim layerName As String = ""
                Dim info As Byte() = New UTF8Encoding(True).GetBytes("Profile generated " & Now & " by user " & myWhoAmI() & vbCrLf)
                fs.Write(info, 0, info.Length)
                Dim moreInfo As Byte() = New UTF8Encoding(True).GetBytes(vbCrLf & "PROFILE OPTIONS" & vbCrLf)
                fs.Write(moreInfo, 0, moreInfo.Length)
                If frmProfileTool.optProfileDelineationOn.Checked = True Then
                    layerName = frmProfileTool.lstSurfaceUnitLayer.Text
                    retrievedPath = Await QueuedTask.Run(Function()
                                                             Return GetRealNameAndPath(layerName).Result
                                                         End Function)

                    Dim surfInfo As Byte() = New UTF8Encoding(True).GetBytes("Surface unit source file:" & vbCrLf & "     " & retrievedPath & vbCrLf)
                    fs.Write(surfInfo, 0, surfInfo.Length)
                End If
                layerName = frmProfileTool.lstElevationLayer.Text
                retrievedPath = Await QueuedTask.Run(Function()
                                                         Return GetRealNameAndPath(layerName).Result
                                                     End Function)

                Dim moreInfo2 As Byte() = New UTF8Encoding(True).GetBytes("Elevation source data:" & vbCrLf & "     " & retrievedPath & vbCrLf)
                fs.Write(moreInfo2, 0, moreInfo2.Length)
                Dim dataUnitAbbr As String = " m"
                If frmProfileTool.optElevationFt.Checked = True Then
                    dataUnitAbbr = " ft"
                End If
                moreInfo2 = New UTF8Encoding(True).GetBytes("Elevation sampled every:" & vbCrLf & "     " & frmProfileTool.txtPrecisionMeasure.Text & dataUnitAbbr & vbCrLf)
                fs.Write(moreInfo2, 0, moreInfo2.Length)
                moreInfo2 = New UTF8Encoding(True).GetBytes("Vertical exaggeration:" & vbCrLf & "     " & m_iVerticalExagg & "x" & vbCrLf)
                fs.Write(moreInfo2, 0, moreInfo2.Length)
                moreInfo2 = New UTF8Encoding(True).GetBytes("Spatial Reference:" & vbCrLf & "     " & m_pSpatialReference.Name & vbCrLf)
                fs.Write(moreInfo2, 0, moreInfo2.Length)
                moreInfo2 = New UTF8Encoding(True).GetBytes("Profile coordinates:" & vbCrLf & "     Start x: " & m_dXstart & vbCrLf & "     Start y: " & m_dYstart & vbCrLf & "     End x: " & m_dXend & vbCrLf & "     End y: " & m_dYend & vbCrLf & vbCrLf)
                fs.Write(moreInfo2, 0, moreInfo2.Length)

                If frmProfileTool.chkBuildProfileBox.Checked = True Then
                    Dim boxUnit As String = "meters"
                    If frmProfileTool.optBoxUnitsFT.Checked = True Then
                        boxUnit = "feet"
                    End If

                    moreInfo2 = New UTF8Encoding(True).GetBytes("PROFILE BOX OPTIONS:" & vbCrLf & "Box dimensions (" & boxUnit & ")" & vbCrLf & "     Top: " & frmProfileTool.txtBoxT.Text &
                                                                vbCrLf & "     Bottom: " & frmProfileTool.txtBoxB.Text & vbCrLf & "     Length: " & m_dBoxL & vbCrLf & "Horizontal ticks (" & boxUnit & ")" & vbCrLf &
                                                                "     Major: " & frmProfileTool.txtHTick.Text & vbCrLf)
                    fs.Write(moreInfo2, 0, moreInfo2.Length)

                    If frmProfileTool.chkMinorTick.Checked = True Then
                        moreInfo2 = New UTF8Encoding(True).GetBytes("     Minor: " & frmProfileTool.txtHTickMinor.Text & vbCrLf)
                        fs.Write(moreInfo2, 0, moreInfo2.Length)
                    End If

                    moreInfo2 = New UTF8Encoding(True).GetBytes("Vertical ticks (" & boxUnit & ")" & vbCrLf & "     Major: " & m_dVTickMajor & vbCrLf)
                    fs.Write(moreInfo2, 0, moreInfo2.Length)

                    If frmProfileTool.chkMinorTick.Checked = True Then
                        moreInfo2 = New UTF8Encoding(True).GetBytes("     Minor: " & m_dVTickMinor & vbCrLf & vbCrLf)
                        fs.Write(moreInfo2, 0, moreInfo2.Length)
                    End If
                End If

                If frmProfileTool.chkAddWells.Checked = True Then
                    layerName = frmProfileTool.lstWellLayer.Text
                    retrievedPath = Await QueuedTask.Run(Function()
                                                             Return GetRealNameAndPath(layerName).Result
                                                         End Function)
                    moreInfo2 = New UTF8Encoding(True).GetBytes("WELL OVERLAY PARAMETERS" & vbCrLf & "Point layer options" & vbCrLf & "     Point layer source file:" & vbCrLf &
                                                                "          " & retrievedPath & vbCrLf &
                                                                "     PointID field = " & frmProfileTool.lstWellsPointID.Text & vbCrLf)
                    fs.Write(moreInfo2, 0, moreInfo2.Length)

                    If frmProfileTool.chkGetElevFromDEM.Checked = True Then
                        moreInfo2 = New UTF8Encoding(True).GetBytes("     Top elevations from elevation source" & vbCrLf)
                        fs.Write(moreInfo2, 0, moreInfo2.Length)
                    Else
                        moreInfo2 = New UTF8Encoding(True).GetBytes("     Elevation field = " & frmProfileTool.lstWellsElev.Text & vbCrLf)
                        fs.Write(moreInfo2, 0, moreInfo2.Length)
                    End If

                    layerName = frmProfileTool.lstTables.Text
                    Debug.Print("Layer Name: " + layerName)
                    retrievedPath = Await QueuedTask.Run(Function()
                                                             Return GetRealNameAndPath(layerName).Result
                                                         End Function)

                    moreInfo2 = New UTF8Encoding(True).GetBytes("vbCrLfSUBSURFACE DATA TABLE OPTIONS" & vbCrLf & "     Subsurface data source file:" & vbCrLf & "          " & retrievedPath &
                                                                vbCrLf & "     PointID field = " & frmProfileTool.lstTablePointID.Text & vbCrLf &
                                                                "     Layer top depth field = " & frmProfileTool.lstTableTop.Text & vbCrLf &
                                                                "     Layer bottom depth field = " & frmProfileTool.lstTableBottom.Text & vbCrLf &
                                                                "     Subsurface unit field = " & frmProfileTool.lstTableUnit.Text & vbCrLf & vbCrLf)
                    fs.Write(moreInfo2, 0, moreInfo2.Length)
                End If

                moreInfo2 = New UTF8Encoding(True).GetBytes("LOCATION OF OUTPUT FILES:" & vbCrLf & "     " & wsFolder & vbCrLf &
                                            "          " & m_sReadmeFileName & ".txt" & vbCrLf &
                                            "          " & m_sProfileDataFileName & ".txt" & vbCrLf)
                fs.Write(moreInfo2, 0, moreInfo2.Length)

                If wsFolder <> m_sWorkspaceFolder Then
                    moreInfo2 = New UTF8Encoding(True).GetBytes("     " & m_sWorkspaceFolder & vbCrLf)
                    fs.Write(moreInfo2, 0, moreInfo2.Length)
                End If

                moreInfo2 = New UTF8Encoding(True).GetBytes("          " & m_sProfileName & vbCrLf &
                                            "          " & m_sProfileLineName & vbCrLf)
                fs.Write(moreInfo2, 0, moreInfo2.Length)

                If frmProfileTool.optProfileDelineationOn.Checked = True Then
                    moreInfo2 = New UTF8Encoding(True).GetBytes("          " & m_sIntersectionPointsName & vbCrLf)
                    fs.Write(moreInfo2, 0, moreInfo2.Length)
                End If

                If frmProfileTool.chkBuildProfileBox.Checked = True Then
                    moreInfo2 = New UTF8Encoding(True).GetBytes("          " & m_sBoxMajorName & vbCrLf)
                    fs.Write(moreInfo2, 0, moreInfo2.Length)
                    If frmProfileTool.chkMinorTick.Checked = True Then
                        moreInfo2 = New UTF8Encoding(True).GetBytes("          " & m_sBoxMinorName & vbCrLf)
                        fs.Write(moreInfo2, 0, moreInfo2.Length)
                    End If
                End If
                If frmProfileTool.chkAddWells.Checked = True Then
                    moreInfo2 = New UTF8Encoding(True).GetBytes("          " & m_sWellsName & vbCrLf)
                    fs.Write(moreInfo2, 0, moreInfo2.Length)
                End If
            End Using
        End If
    End Sub

    Private Async Function GetRealNameAndPath(ByVal v_sLayerName As String) As Task(Of String)
        Debug.Print("get real name and path")
        'This will return the full path and file name for the passed layer name

        Dim pMap As Map = MapView.Active.Map
        Dim pLayers As IEnumerable(Of Layer) = pMap.Layers
        Dim returnURI As String = ""
        Dim dataConn As CIMStandardDataConnection
        Dim ws As String = ""
        Await QueuedTask.Run(Function()
                                 Try
                                     For Each layer In pLayers
                                         If TypeOf layer Is FeatureLayer Or TypeOf layer Is RasterLayer Then
                                             If layer.Name.Equals(v_sLayerName) Then
                                                 dataConn = layer.GetDataConnection()
                                                 ws = dataConn.WorkspaceConnectionString
                                                 ws = ws.Substring(ws.IndexOf("=") + 1)
                                                 If Not (ws.EndsWith("\") Or ws.EndsWith("/")) Then
                                                     ws += "\"
                                                 End If
                                                 returnURI = ws + v_sLayerName
                                                 Return returnURI
                                                 Exit For
                                             End If
                                         Else
                                             For Each pStandaloneTable In pMap.StandaloneTables
                                                 If pStandaloneTable.Name.Equals(v_sLayerName) Then
                                                     dataConn = pStandaloneTable.GetDataConnection()
                                                     ws = dataConn.WorkspaceConnectionString
                                                     ws = ws.Substring(ws.IndexOf("=") + 1)
                                                     If Not (ws.EndsWith("\") Or ws.EndsWith("/")) Then
                                                         ws += "\"
                                                     End If
                                                     returnURI = ws + v_sLayerName
                                                     Return returnURI
                                                     Exit For
                                                 End If
                                             Next
                                             Exit For
                                         End If
                                     Next
                                 Catch ex As AggregateException
                                     For Each x In ex.InnerExceptions
                                         Debug.Print("Aggregate Ex. error in GetRealNameAndPath: " + x.Message)
                                     Next
                                 End Try
                                 Debug.Print("Layer Source URI: " + returnURI)
                                 Return returnURI
                             End Function)
        Return returnURI
    End Function

    Private Function MyWhoAmI() As String
        Debug.Print("who am I")
        'get the user's logon name so it can be recorded in the readme file
        Dim sVar As String
        Dim i As Integer
        Dim username As String = "Unknown User"
        For i = 1 To 250
            sVar = Environ(i)
            If LCase(Left(sVar, 9)) = "username=" Then
                username = Right(sVar, Len(sVar) - 9)
                Exit For
            End If
        Next
        Return username
    End Function

    Public Sub CloseTool(ByVal msg As String, ByVal buttons As MsgBoxStyle, ByVal title As String)
        'Close the Form
        Try
            CloseForm()
            'notify the user about the files' locations
            MsgBox(msg, buttons, title)
        Catch ex As Exception

            Debug.Print(ex.Message)
        End Try
    End Sub
End Module



