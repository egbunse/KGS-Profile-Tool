Imports System.Math
Imports ArcGIS.Core.Geometry
Imports ArcGIS.Core.Data
Imports ArcGIS.Desktop.Mapping
Imports ArcGIS.Desktop.Core
Imports ArcGIS.Desktop.Core.Geoprocessing
Imports ArcGIS.Desktop.Editing
Imports ArcGIS.Desktop.Framework.Threading.Tasks
Imports System.Threading

Module modProfileWells


    '* Created 24Jul07 by:
    '* James L. Poelstra
    '* Email: james.arcscripts.help@nym.hush.com

    'Originally written in VBA
    'Converted to VB.net, April 2015 by:
    'Kristen Jordan
    'Kansas Data Access and Support Center
    'Email: kristen@kgs.ku.edu

    'Converted for ArcGIS Pro, April 2018 by: 
    'Emily Bunse
    'Kansas Geological Survey- Cartographic Services GRA
    'Email: egbunse@ku.edu; egbunse@gmail.com

    Private m_adInPoints(0, 0) As Double
    Private m_asPointIDs(0, 0) As String
    Private m_bThereWereErrors As Boolean
    Private m_adSubsurfaceData(0, 0) As Double
    Private m_asSubsurfaceData(0, 0) As String

    Private m_pPoint_Start As MapPoint  'IPoint 'start point of a segment
    Private m_pPoint_End As MapPoint    'IPoint 'end point of a segment
    Private m_pLine As LineSegment     'ILine 'a single 2-point-line
    Private m_pPolyline As Polyline     'IPolyline 'the polyline made of several Paths with unique IDs
    Private m_pFeatureClass_Current As FeatureClass  'IFeatureClass
    Private m_pFeature As Feature                     'IFeature
    Private m_pFeatureClass As FeatureClass         'IFeatureClass
    'Private m_apPoint(4) As MapPoint               'IPoint

    Private m_dSubsurfaceTableUnitsToMapUnits As Double

    Dim tool_path As String
    Dim args As IReadOnlyList(Of String)
    Dim result As IGPResult

    Public Async Function GetSelectedWellPointsForProfileAsync(
        ByVal v_dStartX As Double,
        ByVal v_dStartY As Double,
        ByVal v_dEndX As Double,
        ByVal v_dEndY As Double,
        ByVal v_iVerticalExagg As Integer,
        ByVal v_dSubsurfaceTableUnitsToMapUnits As Double,
        ByVal v_pSpatialReference As SpatialReference,
        ByVal v_sWellName As String,
        ByVal v_sFolderName As String,
        ByVal frmProfileTool As FrmProfileTool,
        ByVal v_sRasterLayerName As String,
        ByVal cps As CancelableProgressorSource) As Task(Of Boolean)
        'This module will draw the user-selected points and associated subsurface data along the profile.
        'Points are located along the profile at their perpendicular intersection with the profile line.
        'Requirements:
        'A point layer with attributes containing a unique pointed, and (optionally) the elevation of the well
        'A subsurface table containing pointed, depth to layer top, depth to layer bottom, and material_code
        m_bThereWereErrors = False
        m_dSubsurfaceTableUnitsToMapUnits = v_dSubsurfaceTableUnitsToMapUnits
        Debug.Print(Str(v_dSubsurfaceTableUnitsToMapUnits))
        Dim v_sWellPointID As String = frmProfileTool.lstWellsPointID.Text
        Dim v_sWellsElev As String = frmProfileTool.lstWellsElev.Text

        Await GetSelectedPointCoords(frmProfileTool.lstWellLayer.Text, v_sWellPointID, v_sWellsElev, frmProfileTool)

        If m_bThereWereErrors = False Then
            Call CalculateIntersections(v_dStartX, v_dStartY, v_dEndX, v_dEndY)

            'get the elevation for the selected points
            If frmProfileTool.chkGetElevFromDEM.Checked = True Then
                Await GetElevationsFromDEM(v_sRasterLayerName)
            End If

            Await ReadSubsurfaceData(frmProfileTool, v_iVerticalExagg)

            Await BuildProfileWellsSHP(v_iVerticalExagg, v_pSpatialReference, v_sWellName, v_sFolderName, frmProfileTool)

        End If

        Await ClearSelection(frmProfileTool.lstWellLayer.Text)

        Return True
    End Function

    Private Async Function GetSelectedPointCoords(ByVal v_sPointsName As String, ByVal pointID As String, ByVal wellsElev As String, ByVal frmProfileTool As FrmProfileTool) As Task(Of Boolean)
        'Setup objects
        Dim pFeature As Feature    'IFeature
        Dim pLayer As Layer   'ILayer
        Dim pMap As Map = MapView.Active.Map
        Dim pLayers As IEnumerable(Of Layer) = pMap.Layers
        Dim pFeatureLayer As FeatureLayer
        Dim pSelection As Selection
        Dim shape As Geometry = Nothing
        Dim fieldList As IReadOnlyList(Of Field)
        'Get coordinates for the selected points in the points layer
        Dim lCount As Long 'selected points counter
        lCount = 0

        Dim pPoint As MapPoint
        Await QueuedTask.Run(Function()
                                 For Each layer In pLayers
                                     If layer.Name = v_sPointsName Then
                                         Debug.Print(v_sPointsName)
                                         pFeatureLayer = layer
                                         pSelection = pFeatureLayer.GetSelection()
                                         Debug.Print("Selected wells: " + Str(pSelection.GetCount()))
                                         Using rowCursor As RowCursor = pSelection.Search(Nothing, False)
                                             If pSelection.GetCount() = 0 Then
                                                 MsgBox("No points were selected in the well points layer." _
                                                 & vbNewLine & "A wells overlay will not be created.",
                                                 vbCritical + vbOKOnly, "Error ...")
                                                 m_bThereWereErrors = True
                                                 'Exit Sub
                                                 Exit For
                                             End If

                                             ReDim m_adInPoints(pSelection.GetCount() - 1, 6) 'index,[point coord x, point coord y, intersection coord x, intersection coord y, distance from profile_line, elevation, profile x coord]
                                             ReDim m_asPointIDs(pSelection.GetCount() - 1, 0)

                                             Dim i As Integer

                                             While rowCursor.MoveNext()
                                                 Using feat As Feature = DirectCast(rowCursor.Current(), Feature)
                                                     pPoint = feat.GetShape()
                                                     Debug.Print("Wells Points: " + Str(pPoint.X) + " " + Str(pPoint.Y))
                                                     m_adInPoints(lCount, 0) = pPoint.X
                                                     m_adInPoints(lCount, 1) = pPoint.Y

                                                     'store the PointID and (optionally) the Elevation
                                                     fieldList = feat.GetFields()
                                                     For i = 0 To fieldList.Count - 1  'for each field in the layer
                                                         If fieldList(i).Name = pointID Then
                                                             Debug.Print("Values: " + Str(feat.GetOriginalValue(i)))
                                                             Debug.Print("lCount: " + lCount.ToString)
                                                             m_asPointIDs(lCount, 0) = feat.GetOriginalValue(i)
                                                         End If

                                                         If frmProfileTool.chkGetElevFromDEM.Checked = False _
                                                             And fieldList(i).Name = wellsElev Then
                                                             m_adInPoints(lCount, 5) = feat.GetOriginalValue(i) * m_dSubsurfaceTableUnitsToMapUnits
                                                         End If
                                                     Next
                                                 End Using
                                                 lCount += 1
                                             End While
                                         End Using
                                     End If
                                 Next
                                 Return True
                             End Function)


        'Dim pFeatureLayer As IFeatureLayer

        'pMxDocument = My.ArcMap.Application.Document
        'pEnumLayer = pMxDocument.FocusMap.Layers

        'pEnumLayer.Reset()
        'pLayer = pEnumLayer.Next

        ''Get coordinates for the selected points in the points layer
        'Dim lCount As Long 'selected points counter
        'lCount = 0

        'Dim pPoint As MapPoint 'IPoint

        'Do Until pLayer Is Nothing

        '    If pLayer.Name = v_sPointsName Then 'this is the points layer
        '        pFeatureLayer = pLayer
        '        pFeatureSelection = pFeatureLayer
        '        pFeatureSelection.SelectionSet.Search(Nothing, False, pFeatureCursor)

        '        'if no points are selected, then exit sub
        '        If pFeatureSelection.SelectionSet.Count = 0 Then
        '            MsgBox("No points were selected in the well points layer." _
        '            & vbNewLine & "A wells overlay will Not be created.",
        '            vbCritical + vbOKOnly, "Error ...")
        '            m_bThereWereErrors = True
        '            Exit Sub
        '        End If

        '        ReDim m_adInPoints(pFeatureSelection.SelectionSet.Count - 1, 6) 'index,[point coord x, point coord y, intersection coord x, intersection coord y, distance from profile_line, elevation, profile x coord]
        '        ReDim m_asPointIDs(pFeatureSelection.SelectionSet.Count - 1, 0)

        '        pFeature = pFeatureCursor.NextFeature

        '        Dim i As Integer

        '        Do While Not pFeature Is Nothing 'for each selected feature...
        '            'store the point coords
        '            pPoint = pFeature.Shape
        '            m_adInPoints(lCount, 0) = pPoint.X
        '            m_adInPoints(lCount, 1) = pPoint.Y

        '            'store the PointID and (optionally) the Elevation
        '            For i = 0 To pFeature.Fields.FieldCount - 1 'for each field in the layer
        '                If pFeature.Fields.Field(i).Name = frmProfileTool.lstWellsPointID.Text Then
        '                    m_asPointIDs(lCount, 0) = pFeature.Value(i)
        '                End If
        '                If frmProfileTool.chkGetElevFromDEM.Checked = False _
        '                    And pFeature.Fields.Field(i).Name = frmProfileTool.lstWellsElev.Text Then
        '                    m_adInPoints(lCount, 5) = pFeature.Value(i) * m_dSubsurfaceTableUnitsToMapUnits
        '                End If
        '            Next

        '            pFeature = pFeatureCursor.NextFeature
        '            lCount = lCount + 1
        '        Loop
        '        Exit Do
        '    End If
        '    pLayer = pEnumLayer.Next
        'Loop

        'Free memory
        pPoint = Nothing
        pFeatureLayer = Nothing
        pLayer = Nothing
        pFeature = Nothing
        pMap = Nothing
        pLayers = Nothing
        pSelection = Nothing
        shape = Nothing
        fieldList = Nothing
        Return True
    End Function

    Private Sub CalculateIntersections(
        ByVal v_dStartX As Double,
        ByVal v_dStartY As Double,
        ByVal v_dEndX As Double,
        ByVal v_dEndY As Double)
        Debug.Print("calculate intersections wells")

        'Calculate the slope for the source line
        Dim dM As Double
        dM = (v_dEndY - v_dStartY) / (v_dEndX - v_dStartX)

        'calculate the y-intercept for the source line
        Dim dB As Double
        dB = v_dStartY - (dM * v_dStartX)

        'Calculate the perpendicular slope
        Dim dMp As Double
        If dM = 0 Then 'prevent division by zero
            dMp = 1.0E+15
        Else
            dMp = -1 / dM
        End If

        'enter loop to calculate the intersection point coords for all selected points
        Dim i As Integer
        For i = 0 To UBound(m_adInPoints, 1)

            'calculate the y-intercept for the Perp. line
            Dim dBp As Double
            dBp = m_adInPoints(i, 1) - (dMp * m_adInPoints(i, 0))

            'calculate the x-intercept for the intersection point of the two lines
            Dim dIx As Double
            dIx = (dBp - dB) / (dM - dMp)

            'calculate the y-intercept for the intersection point of the two lines
            Dim dIy As Double
            dIy = dM * dIx + dB

            'store the intersection point coords
            m_adInPoints(i, 2) = dIx
            m_adInPoints(i, 3) = dIy

            'store the point's distance from the profile line
            m_adInPoints(i, 4) = Sqrt(((m_adInPoints(i, 3) - m_adInPoints(i, 1)) ^ 2 _
                + (m_adInPoints(i, 2) - m_adInPoints(i, 0)) ^ 2))

            'store the x coordinate for the profile
            m_adInPoints(i, 6) = Sqrt(((dIy - v_dStartY) ^ 2 + (dIx - v_dStartX) ^ 2))
        Next
    End Sub

    Private Async Function GetElevationsFromDEM(ByVal v_sRasterLayerName As String) As Task(Of Boolean)
        ' Identifies the raster containing the elevation data and
        ' stores it in the array
        Dim pMap As Map
        Dim pEnumLayer As IEnumerable(Of Layer)
        Dim pLayer As Layer
        Dim pRasterLayer As RasterLayer
        Dim pRaster As Raster.Raster
        Dim pPoint As MapPoint
        Dim lPointCounter As Long
        Debug.Print("get elevations from DEM wells")
        'get the specified raster layer
        pMap = MapView.Active.Map
        pEnumLayer = pMap.Layers
        For Each pLayer In pEnumLayer
            If pLayer.Name = v_sRasterLayerName Then
                pRasterLayer = pLayer
                pRaster = Await QueuedTask.Run(Function() pRasterLayer.GetRaster)
            End If
        Next

        'loop through all points
        For lPointCounter = 0 To UBound(m_adInPoints, 1)

            'this must be reset each time to prevent an error in the raster identify function
            pPoint = Await QueuedTask.Run(Function() MapPointBuilder.CreateMapPoint(m_adInPoints(lPointCounter, 2), m_adInPoints(lPointCounter, 3)))

            'call public sub from the modProfileTool_Main module
            m_adInPoints(lPointCounter, 5) = Await QueuedTask.Run(Function()
                                                                      Return GetRasterValueForThisPoint(pPoint, v_sRasterLayerName).Result
                                                                  End Function)
            Debug.Print(Str(m_adInPoints(lPointCounter, 5)))
            pPoint = Nothing
        Next lPointCounter

        'free memory
        pPoint = Nothing
        pRaster = Nothing
        pRasterLayer = Nothing
        pLayer = Nothing
        pEnumLayer = Nothing
        pMap = Nothing
        Return True
    End Function

    Private Async Function ReadSubsurfaceData(ByVal frmProfileTool As FrmProfileTool, Optional ByVal v_iVerticalExagg As Integer = 0) As Task(Of Boolean)
        Dim pMap As Map
        Dim pStandaloneTable As StandaloneTable = Nothing
        pMap = MapView.Active.Map
        Dim pEnumLayer As IEnumerable(Of Layer)
        pEnumLayer = pMap.Layers
        Dim fieldList As IReadOnlyList(Of Field)
        Debug.Print("read subsurface data wells")
        Dim alFieldIDs(3, 1) As Long '(FieldNoForField,FieldType)

        Try
            'hook the correct table and read data from it
            For Each pTable In pMap.StandaloneTables
                If pTable.Name = frmProfileTool.lstTables.Text Then 'this is the table to read
                    pStandaloneTable = pTable
                    fieldList = Await QueuedTask.Run(Function() pStandaloneTable.GetTable().GetDefinition().GetFields())
                    'loop through all fields and store IDs and field type for the selected fields
                    'field types: 0-smallinteger;1-integer;2-single;3-double;4-string;5-date;6-oid;7-geometry;8-blob
                    Dim lFieldCounter As Long

                    For lFieldCounter = 0 To fieldList.Count - 1
                        If fieldList(lFieldCounter).Name = frmProfileTool.lstTablePointID.Text Then
                            alFieldIDs(0, 0) = lFieldCounter
                            alFieldIDs(0, 1) = fieldList(lFieldCounter).FieldType
                        ElseIf fieldList(lFieldCounter).Name = frmProfileTool.lstTableUnit.Text Then
                            alFieldIDs(1, 0) = lFieldCounter
                            alFieldIDs(1, 1) = fieldList(lFieldCounter).FieldType
                        ElseIf fieldList(lFieldCounter).Name = frmProfileTool.lstTableTop.Text Then
                            alFieldIDs(2, 0) = lFieldCounter
                            alFieldIDs(2, 1) = fieldList(lFieldCounter).FieldType
                        ElseIf fieldList(lFieldCounter).Name = frmProfileTool.lstTableBottom.Text Then
                            alFieldIDs(3, 0) = lFieldCounter
                            alFieldIDs(3, 1) = fieldList(lFieldCounter).FieldType
                        End If
                    Next
                    Exit For
                End If
            Next
        Catch ex As Exception
            Debug.Print("Profile Wells- error reading table: " + ex.Message)
            MsgBox("There was a problem reading data from the subsurface table.", vbCritical, "Error creating profile...")
        End Try

        'prepare to loop throught all selected points in the points layer
        Dim pQueryFilter As QueryFilter
        Dim pSelection As Selection
        Dim k As Long
        Dim lSubsurfaceDataCounter As Long

        Try
            pQueryFilter = New QueryFilter With {
                            .SubFields = frmProfileTool.lstTablePointID.Text & "," _
                            & frmProfileTool.lstTableUnit.Text & "," _
                            & frmProfileTool.lstTableTop.Text & "," _
                            & frmProfileTool.lstTableBottom.Text '"POINTID,geo,layertopde,layerbotde"
                           }
            lSubsurfaceDataCounter = 0

            'for each selected point (well)
            Await QueuedTask.Run(Function()
                                     Debug.Print("Rows in table: " + Str(pStandaloneTable.GetTable().GetCount()))
                                     ReDim m_asSubsurfaceData(1, pStandaloneTable.GetTable().GetCount())
                                     ReDim m_adSubsurfaceData(3, pStandaloneTable.GetTable().GetCount())
                                     Return True
                                 End Function)
            For k = 0 To UBound(m_asPointIDs, 1)
                Debug.Print("k: " + k.ToString)
                'finish building the query filter for this point
                If alFieldIDs(0, 1) = 4 Then 'point IDs are in a text field
                    pQueryFilter.WhereClause = frmProfileTool.lstTablePointID.Text & " = '" & m_asPointIDs(k, 0) & "'"
                Else 'point IDs are numeric
                    pQueryFilter.WhereClause = frmProfileTool.lstTablePointID.Text & " = " & m_asPointIDs(k, 0)
                End If
                'get all subsurface records (rows in the subsurface table) for this point and save data
                Await QueuedTask.Run(Function()
                                         Using rowCursor As RowCursor = pStandaloneTable.GetTable().Search(pQueryFilter, False)
                                             While rowCursor.MoveNext()
                                                 Using pRow As Row = rowCursor.Current
                                                     'redim the array
                                                     ' ReDim Preserve m_asSubsurfaceData(1, lSubsurfaceDataCounter)  '[PointID,Geo],index
                                                     'ReDim Preserve m_adSubsurfaceData(3, lSubsurfaceDataCounter)  '[Top,Bottom,TopAdjusted,BottomAdjusted],index
                                                     m_asSubsurfaceData(0, lSubsurfaceDataCounter) = CStr(pRow.GetOriginalValue(alFieldIDs(0, 0))) 'pointID
                                                     Debug.Print("Assigned Point IDs: " + Str(m_asSubsurfaceData(0, lSubsurfaceDataCounter)))
                                                     'm_asSubsurfaceData(0, lSubsurfaceDataCounter) = CStr(pRow.GetOriginalValue(rowCursor.FindField(frmProfileTool.lstTablePointID.Text)))
                                                     'geo is the only field that will alow a null value
                                                     If Not (pRow.GetOriginalValue(alFieldIDs(1, 0))) Is Nothing Or pRow.GetOriginalValue(alFieldIDs(1, 0)) <> "" Then
                                                         m_asSubsurfaceData(1, lSubsurfaceDataCounter) = CStr(pRow.GetOriginalValue(alFieldIDs(1, 0))) 'geo
                                                     Else
                                                         m_asSubsurfaceData(1, lSubsurfaceDataCounter) = "not entered"
                                                     End If

                                                     m_adSubsurfaceData(0, lSubsurfaceDataCounter) = CDbl(pRow.GetOriginalValue(alFieldIDs(2, 0))) * m_dSubsurfaceTableUnitsToMapUnits 'top elevation
                                                     m_adSubsurfaceData(2, lSubsurfaceDataCounter) = CDbl(m_adSubsurfaceData(0, lSubsurfaceDataCounter) * CDbl(v_iVerticalExagg)) 'top elevation adjusted for shp file
                                                     m_adSubsurfaceData(1, lSubsurfaceDataCounter) = CDbl(pRow.GetOriginalValue(alFieldIDs(3, 0))) * m_dSubsurfaceTableUnitsToMapUnits 'bottom elevation
                                                     m_adSubsurfaceData(3, lSubsurfaceDataCounter) = CDbl(m_adSubsurfaceData(1, lSubsurfaceDataCounter) * CDbl(v_iVerticalExagg)) 'bottom elevation adjusted for shp file

                                                     lSubsurfaceDataCounter = lSubsurfaceDataCounter + 1
                                                     Debug.Print(Str(lSubsurfaceDataCounter))
                                                 End Using
                                             End While
                                         End Using
                                         Return True
                                     End Function)
            Next
        Catch ex As Exception
            Debug.Print("Profile Wells error getting points: " + ex.Message)
            MsgBox("There was a problem with the Point IDs in the wells location layer.", vbCritical, "Error creating profile...")
        End Try

        'free memory
        pMap = Nothing
        pStandaloneTable = Nothing
        pQueryFilter = Nothing
        pSelection = Nothing
    End Function

    'https://github.com/Esri/arcgis-pro-sdk/wiki/ProSnippets-Geodatabase#creating-a-feature
    Private Async Function CreateFeature(ByVal workspaceFolder As String, ByVal wellsName As String, ByVal geometryFeatToAdd As List(Of Coordinate2D), ByVal geometryType As String,
                                         ByVal v_sPointID As String, ByVal v_dDist As Double, ByVal v_dStart As Double, ByVal v_dEnd As Double, ByVal v_sSubsurfaceUnit As String) As Task
        Dim creationResult As Boolean = False
        Dim message As String = String.Empty
        Dim saveResult As Boolean = False
        Dim fc As FeatureClass = Nothing
        Dim fsConnectionPath As FileSystemConnectionPath
        Dim shapefile As FileSystemDatastore
        Dim filegdb As Geodatabase
        Dim fgdbConnectionPath As FileGeodatabaseConnectionPath
        Await QueuedTask.Run(Function()
                                 If workspaceFolder.Contains(".gdb") Then
                                     fgdbConnectionPath = New FileGeodatabaseConnectionPath(New Uri(workspaceFolder))
                                     filegdb = New Geodatabase(fgdbConnectionPath)
                                     fc = filegdb.OpenDataset(Of FeatureClass)(wellsName)
                                 Else
                                     fsConnectionPath = New FileSystemConnectionPath(New Uri(workspaceFolder), FileSystemDatastoreType.Shapefile)
                                     shapefile = New FileSystemDatastore(fsConnectionPath)
                                     fc = shapefile.OpenDataset(Of FeatureClass)(wellsName)
                                     'm_sWorkspaceFolder = m_sWorkspaceFolder & "\"
                                 End If
                                 Using fc
                                     'Set up tools needed to add Polyline to Feature Class as a Feature 
                                     Dim rowBuff As RowBuffer = Nothing
                                     Dim newLine As Feature = Nothing
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
                                                                    If geometryType.Equals("Polyline") Then
                                                                        rowBuff(fcDef.GetShapeField()) = New PolylineBuilder(geometryFeatToAdd).ToGeometry()
                                                                    Else   'Polygon
                                                                        rowBuff(fcDef.GetShapeField()) = New PolygonBuilder(geometryFeatToAdd).ToGeometry()
                                                                    End If
                                                                    rowBuff("PointID") = v_sPointID
                                                                    rowBuff("Distance") = v_dDist
                                                                    rowBuff("DepthStart") = v_dStart
                                                                    rowBuff("DepthEnd") = v_dEnd
                                                                    rowBuff("Units") = v_sSubsurfaceUnit
                                                                    newLine = fc.CreateRow(rowBuff)
                                                                    context.Invalidate(newLine)
                                                                    Return True
                                                                End Function, fc)
                                         creationResult = editOperation.Execute()
                                         Try
                                             Dim objectID As Long = newLine.GetObjectID()
                                             Debug.Print("Creation result for " + objectID.ToString + " " + creationResult.ToString)
                                         Catch ref As NullReferenceException
                                             Debug.Print(ref.Message)
                                         End Try
                                         saveResult = Project.Current.SaveEditsAsync().Result
                                         Debug.Print("Save result: " + saveResult.ToString)
                                     Catch ex As GeodatabaseEditingException
                                         Console.WriteLine("Profile Wells- create feature error: " + ex.Message)
                                     Finally
                                         If rowBuff IsNot Nothing Then
                                             rowBuff.Dispose()
                                         End If
                                         If newLine IsNot Nothing Then
                                             newLine.Dispose()
                                         End If
                                     End Try
                                 End Using
                                 Return saveResult
                             End Function)
    End Function

    Private Async Function BuildProfileWellsSHP(ByVal v_iVerticalExagg As Integer,
        ByVal v_pSpatialReference As SpatialReference, ByVal v_sWellsName As String,
        ByVal v_sFolder As String, ByVal frmProfileTool As FrmProfileTool) As Task(Of Boolean)
        Debug.Print("build profile wells shp")
        'specify if wells will be creates as polylines or polygons
        Dim bWellsAsLines As Boolean
        bWellsAsLines = frmProfileTool.optLines.Checked

        Dim lPolygonWidth As Double
        'setup polygon width if nessesary
        If bWellsAsLines = False Then
            lPolygonWidth = CDbl(frmProfileTool.txtPolyW.Text)
        End If

        'Create the new Feature Class
        '--Run the Create Feature Class Tool
        tool_path = "management.CreateFeatureClass"

        If bWellsAsLines = True Then
            args = Geoprocessing.MakeValueArray(v_sFolder, v_sWellsName, "POLYLINE", Nothing, "DISABLED", "DISABLED", v_pSpatialReference)
        Else
            args = Geoprocessing.MakeValueArray(v_sFolder, v_sWellsName, "POLYGON", Nothing, "DISABLED", "DISABLED", v_pSpatialReference)
        End If

        result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, Nothing, Nothing, GPExecuteToolFlags.None)
        If result.IsFailed Then
            MsgBox("Error creating: " & v_sWellsName & vbNewLine _
                & "Profile creation terminated.", vbCritical, "Error adding wells...")
            CloseForm()
        End If

        'Add Fields to the new Feature Class 
        tool_path = "management.AddField"
        Dim fcPath As String = v_sFolder + "\" + v_sWellsName
        Dim fieldsList = New Object(4, 3) {{"PointID", "TEXT", 25, "pointid"}, {"Distance", "DOUBLE", 15, "distance"},
        {"DepthStart", "DOUBLE", 15, "depthstart"}, {"DepthEnd", "DOUBLE", 15, "depthend"}, {"Units", "TEXT", 25, "units"}}
        For i = 0 To fieldsList.GetUpperBound(0)
            Dim fieldName As String = fieldsList(i, 0)
            Dim fieldType As String = fieldsList(i, 1)
            Dim fieldPrec As Double = fieldsList(i, 2)
            Dim fieldAlias As String = fieldsList(i, 3)
            args = Geoprocessing.MakeValueArray(fcPath, fieldName, fieldType, fieldPrec, "", "", fieldAlias, "NULLABLE", "NON_REQUIRED", "")
            result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, Nothing, Nothing, GPExecuteToolFlags.None)
            If result.IsFailed Then
                For Each msg In result.Messages
                    Debug.Print(msg.ToString)
                Next
                MsgBox("Error adding fields to: " & v_sWellsName & vbNewLine _
                    & "Profile creation terminated.", vbCritical, "Error adding wells...")
                CloseForm()
            End If
        Next

        'prepare to add features to shapefile
        Dim lPointCounter As Long
        Dim lLayerCounter As Long
        Dim dX As Double
        Dim dElevationAdjusted As Double
        Dim sCurentPointID As String
        Dim lPointCount As Long
        Dim lLayerCount As Long
        Dim dVerticalExagg As Double

        dVerticalExagg = CDbl(v_iVerticalExagg)
        lPointCount = UBound(m_adInPoints, 1)
        lLayerCount = UBound(m_adSubsurfaceData, 2)
        Debug.Print("Subsurface Layer Count: " + Str(lLayerCount))

        'setup the progress bar counter
        Try
            SetProgBarLabelText(frmProfileTool.lblPrgBr, "Building the profile wells shapefile ...")
            SetProgBarMax(frmProfileTool.prgBr, lPointCount)
        Catch ex As Exception
            Debug.Print("Error updating progress bars in Profile Wells.")
        End Try

        Try
            If bWellsAsLines = True Then
                'add lines to shapefile
                For lPointCounter = 0 To lPointCount 'for each point        '0
                    Debug.Print("In Create Features, incrementing: " + Str(lPointCounter))
                    dX = m_adInPoints(lPointCounter, 6) 'x coord for profile well
                    dElevationAdjusted = m_adInPoints(lPointCounter, 5) * dVerticalExagg 'y coord for profile well top
                    sCurentPointID = m_asPointIDs(lPointCounter, 0)
                    For lLayerCounter = 0 To lLayerCount 'for each subsurface layer
                        If m_asSubsurfaceData(0, lLayerCounter) = sCurentPointID Then
                            Dim newCoords As New List(Of Coordinate2D)() From {
                                New Coordinate2D(dX, dElevationAdjusted - m_adSubsurfaceData(2, lLayerCounter)),
                                New Coordinate2D(dX, dElevationAdjusted - m_adSubsurfaceData(3, lLayerCounter))
                            }
                            Debug.Print("Line: " + Str(dX) + Str(dElevationAdjusted - m_adSubsurfaceData(2, lLayerCounter)))
                            Debug.Print("Line: " + Str(dX) + Str(dElevationAdjusted - m_adSubsurfaceData(3, lLayerCounter)))
                            Debug.Print("Creating the features for the wells data...")
                            Await CreateFeature(v_sFolder, v_sWellsName, newCoords, "Polyline",
                                                sCurentPointID, m_adInPoints(lPointCounter, 4),
                                                m_adSubsurfaceData(0, lLayerCounter), m_adSubsurfaceData(1, lLayerCounter),
                                                m_asSubsurfaceData(1, lLayerCounter))
                        End If
                    Next
                    Try
                        RefreshProgBarValue(frmProfileTool.prgBr, lPointCounter)
                    Catch ex As Exception
                        Debug.Print("Error refreshing Prog Bar in Profile Wells Creation: " + ex.Message)
                    End Try
                Next
            Else
                'add polygons to shapefile
                For lPointCounter = 0 To lPointCount 'for each point
                    dX = m_adInPoints(lPointCounter, 6) 'profile well center
                    dElevationAdjusted = m_adInPoints(lPointCounter, 5) * dVerticalExagg 'profile well top
                    sCurentPointID = m_asPointIDs(lPointCounter, 0)
                    For lLayerCounter = 0 To lLayerCount 'for each subsurface layer
                        If m_asSubsurfaceData(0, lLayerCounter) = sCurentPointID Then
                            Dim newCoords As New List(Of Coordinate2D)() From {
                                New Coordinate2D(dX - (lPolygonWidth * 0.5), dElevationAdjusted - m_adSubsurfaceData(2, lLayerCounter)),
                                New Coordinate2D(dX + (lPolygonWidth * 0.5), dElevationAdjusted - m_adSubsurfaceData(2, lLayerCounter)),
                                New Coordinate2D(dX + (lPolygonWidth * 0.5), dElevationAdjusted - m_adSubsurfaceData(3, lLayerCounter)),
                                New Coordinate2D(dX - (lPolygonWidth * 0.5), dElevationAdjusted - m_adSubsurfaceData(3, lLayerCounter)),
                                New Coordinate2D(dX - (lPolygonWidth * 0.5), dElevationAdjusted - m_adSubsurfaceData(2, lLayerCounter)) 'Same as the start
                            }
                            Await CreateFeature(v_sFolder, v_sWellsName, newCoords, "Polygon", sCurentPointID, m_adInPoints(lPointCounter, 4),
                                                m_adSubsurfaceData(0, lLayerCounter), m_adSubsurfaceData(1, lLayerCounter), m_asSubsurfaceData(1, lLayerCounter))
                        End If
                    Next
                    Try
                        RefreshProgBarValue(frmProfileTool.prgBr, lPointCounter)
                    Catch ex As Exception
                        Debug.Print("Error refreshing Prog Bar in Profile Wells Creation: " + ex.Message)
                    End Try
                Next
            End If
        Catch ex As AggregateException
            For Each x In ex.InnerExceptions
                Debug.Print("Aggregate Exception in Profile Wells Creation: " + x.Message)
                MsgBox("Could not build the wells overlay shapefile for the profile.", vbCritical, "Error creating profile...")
            Next
        End Try

        Dim newWellsLayer As Layer
        Dim wellsPath As Uri = New Uri(v_sFolder + "\" + v_sWellsName)
        Await QueuedTask.Run(Function()
                                 newWellsLayer = LayerFactory.Instance.CreateLayer(wellsPath, MapView.Active.Map, 0)
                                 Return True
                             End Function)

        m_pFeature = Nothing
        m_pPolyline = Nothing
        m_pLine = Nothing
        m_pPoint_Start = Nothing
        m_pPoint_End = Nothing
        m_pFeatureClass_Current = Nothing
        m_pFeatureClass = Nothing
        v_pSpatialReference = Nothing
        Return True
    End Function

    Private Async Function ClearSelection(ByVal wellsFCname) As Task(Of Boolean)
        Dim pMap As Map = MapView.Active.Map
        Dim pLayers As IEnumerable(Of Layer) = pMap.Layers
        Dim pFeatureLayer As FeatureLayer
        Await QueuedTask.Run(Function()
                                 For Each layer In pLayers
                                     If layer.Name = wellsFCname Then
                                         Debug.Print(wellsFCname)
                                         pFeatureLayer = layer
                                         pFeatureLayer.ClearSelection()
                                     End If
                                 Next
                                 Return True
                             End Function)
        Return True
    End Function
End Module
