Imports ArcGIS.Core.Geometry
Imports ArcGIS.Core.Data
Imports System.Math
Imports ArcGIS.Desktop.Core.Geoprocessing
Imports ArcGIS.Desktop.Mapping
Imports ArcGIS.Desktop.Editing
Imports ArcGIS.Desktop.Framework.Threading.Tasks
Imports ArcGIS.Desktop.Core

Module modProfileBox

    '* Created 24Jul07 by:
    '* James L. Poelstra
    '* Email: james.arcscripts.help@nym.hush.com

    'Originally written in VBA
    'Converted to VB.net, April 2015 by:
    'Kristen Jordan
    'Kansas Data Access and Support Center
    'Email: kristen@kgs.ku.edu

    'Converted for ArcGIS Pro, August 2017 by: 
    'Emily Bunse
    'Kansas Geological Survey- Cartographic Services GRA
    'Email: egbunse@ku.edu; egbunse@gmail.com

    Private m_pPoint_Start As MapPoint 'start point of a segment
    Private m_pPoint_End As MapPoint 'end point of a segment
    Private m_pLine As LineSegment 'a single 2-point-line
    Private m_pSegmentCollection As PolylineBuilder 'a collection of lines
    Private m_pPolyline As Polyline 'the polyline made of several Paths with unique IDs
    Private m_pFeatureClass As FeatureClass

    Private m_lCounter As Long 'A counter that can be used to update a progress bar
    Private m_bUpdateProgressBar As Boolean 'Will the progress bar in the profile tool be updated?

    '--Vars needed for running geoprocessing tools 
    Dim tool_path As String
    Dim args As IReadOnlyList(Of String)
    Dim result As IGPResult
    Dim sWorkspace As String 'Output workspace
    Dim sPolylineName As String 'Output filename
    Dim ticLabel As String = "Tic_m"

    Public Async Function ProfileBox(ByVal frmProfileTool As frmProfileTool,
        Optional ByVal v_bHasPassedVariables As Boolean = False,
        Optional ByVal v_sInSurfaceLayer As String = "",
        Optional ByVal v_dInBoxL As Double = 1,
        Optional ByVal v_dInBoxT As Double = 1,
        Optional ByVal v_dInBoxB As Double = 1,
        Optional ByVal v_dInVTic As Double = 1,
        Optional ByVal v_dInHTic As Double = 1,
        Optional ByVal v_iInProfileExagg As Integer = 1,
        Optional ByVal v_bInAddGrid As Boolean = False,
        Optional ByVal v_bInSquareGrid As Boolean = False,
        Optional ByVal v_sInWorkspace As String = "",
        Optional ByVal v_sInPolylineName As String = "",
        Optional ByVal v_bUpdateProfileStatus As Boolean = False,
        Optional ByVal v_IsMajorTick As Boolean = True,
        Optional ByVal v_PrgBarMsg As String = "") As Task(Of Boolean)
        'This sub will create a box with (labeled) exterior tic marks and optionally
        'an interior grid to help geologists draw subsurface cross sections.
        'There is an option to draw a square grid, which is useful if having a true
        'square grid is more important than the vertical tic distance.
        'Note: this sub can be called independently to create boxes;
        'it is not dependent on the profile tool.

        Dim sSurfaceLayerName As String 'Name of the layer that has the spatial reference
        Dim vntInputBox As Object
        Dim dBoxL As Double 'Box Length
        Dim dBoxT As Long 'Box Top Elevation
        Dim dBoxB As Long 'Box Bottom Elevation
        Dim iProfileExagg As Integer 'The Profile Exaggeration
        Dim dVTic As Double 'the vertical tic spacing
        Dim dHTic As Double 'the spacing between horizontal tic marks
        Dim bAddGrid As Boolean 'Build a grid inside of the box?
        Dim bSquareGrid As Boolean 'Produce a square grid, i.e. calculate the vertical tics

        'see about unit differences
        Dim userUnits As String
        Dim dataUnits As String
        Dim cu As Double = 1 'conversion unit for tic labels, set to 1 in case no conversion is necessary
        Dim boolConvert As Boolean = False

        If frmProfileTool.optBoxUnitsFT.Checked = True Then
            userUnits = "ft" 'user wants grid in feet
            boolConvert = True
        Else
            userUnits = "m" 'user wants grid in meters
        End If

        'see what units the data is in
        If frmProfileTool.optElevationFt.Checked = True Then
            dataUnits = "ft" 'data is in feet
        Else
            dataUnits = "m" 'data is in meters
        End If

        If userUnits <> dataUnits Then
            If userUnits = "ft" And dataUnits = "m" Then
                'convert meters to feet
                cu = 3.28084
            ElseIf userUnits = "m" And dataUnits = "ft" Then
                'convert feet to meters
                cu = 0.3048
            End If
        End If

        'Setup variables
        If v_bHasPassedVariables = True Then
            sSurfaceLayerName = v_sInSurfaceLayer
            dBoxL = v_dInBoxL / cu 'box length in data units
            dBoxT = v_dInBoxT / cu 'box top in data units
            dBoxB = v_dInBoxB / cu 'box bottom in data units
            iProfileExagg = v_iInProfileExagg
            dVTic = v_dInVTic 'desired tic height in user units
            dHTic = v_dInHTic 'desired tic length in user units
            bAddGrid = v_bInAddGrid
            sWorkspace = v_sInWorkspace
            sPolylineName = v_sInPolylineName
            bSquareGrid = v_bInSquareGrid
            m_bUpdateProgressBar = v_bUpdateProfileStatus
        Else
            sSurfaceLayerName = InputBox("Enter the name of the layer to use for the spatial reference." _
                & vbNewLine & "Name should be typed exactly as it appears in the map TOC", "Input")
            If sSurfaceLayerName = "" Then Exit Function
            vntInputBox = InputBox("How long is the box?", "Input", "24000")
            If vntInputBox = "" Then Exit Function
            dBoxL = CDbl(vntInputBox)
            vntInputBox = InputBox("Enter the top elevation for the box:", "Input", "1000")
            If vntInputBox = "" Then Exit Function
            dBoxT = CLng(vntInputBox)
            vntInputBox = InputBox("Enter the bottom elevation for the box:", "Input", "-200")
            If vntInputBox = "" Then Exit Function
            dBoxB = CLng(vntInputBox)
            vntInputBox = InputBox("What is the Profile exageration?", "Input", "10")
            If vntInputBox = "" Then Exit Function
            iProfileExagg = CLng(vntInputBox)
            vntInputBox = InputBox("Enter the Horizontal tic interval", "Input", "2000")
            If vntInputBox = "" Then Exit Function
            dHTic = CDbl(vntInputBox)
            Dim intMsg As Integer
            intMsg = MsgBox("Build a grid inside of the box?", vbYesNo, "Input")
            If intMsg = 6 Then bAddGrid = True Else bAddGrid = False
            Select Case bAddGrid
                Case False
                    vntInputBox = InputBox("Enter the Vertical tic interval", "Input", "200")
                    If vntInputBox = "" Then Exit Function
                    dVTic = CDbl(vntInputBox)
                Case True
                    intMsg = MsgBox("Build a square grid?", vbYesNo, "Input")
                    If intMsg = 6 Then bSquareGrid = True Else bSquareGrid = False
                    If bSquareGrid = False Then
                        vntInputBox = InputBox("Enter the Vertical tic interval", "Input", "200")
                        If vntInputBox = "" Then Exit Function
                        dVTic = CDbl(vntInputBox)
                    End If
            End Select

            m_bUpdateProgressBar = False
            sWorkspace = InputBox("Enter the path to the temporary Directory:", "Input", "C:\WINNT\Temp")
            sPolylineName = "profile_box"
        End If

        Dim diff As Integer = 0
        Dim fc As FeatureClass = Nothing

        'Edit tic label based on data measurements
        If frmProfileTool.optBoxUnitsFT.Checked = True Then
            ticLabel = "Tic_ft"
        End If

        'Create the new Feature Class
        '--Run the Create Feature Class Tool
        tool_path = "management.CreateFeatureClass"
        args = Geoprocessing.MakeValueArray(sWorkspace, sPolylineName, "POLYLINE", Nothing, "DISABLED", "DISABLED", MapView.Active.Map.SpatialReference)
        result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, Nothing, Nothing, GPExecuteToolFlags.None)
        For Each msg In result.Messages
            Debug.Print(msg.Text)
        Next

        'Add Fields to the new Feature Class 
        tool_path = "management.AddField"
        Dim fcPath As String = sWorkspace + "\" + sPolylineName
        args = Geoprocessing.MakeValueArray(fcPath, ticLabel, "DOUBLE", "", "", 15, ticLabel, Nothing, "NON_REQUIRED", "")
        result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, Nothing, Nothing, GPExecuteToolFlags.None)
        For Each msg In result.Messages
            Debug.Print(msg.Text)
        Next

        tool_path = "management.AddField"
        args = Geoprocessing.MakeValueArray(fcPath, "Type", "TEXT", "", "", 25, "type", "NULLABLE", "NON_REQUIRED", "")
        result = Await Geoprocessing.ExecuteToolAsync(tool_path, args, Nothing, Nothing, Nothing, GPExecuteToolFlags.None)
        For Each msg In result.Messages
            Debug.Print(msg.Text)
        Next

        Dim lXstart As Long
        Dim lYstart As Long

        lXstart = 0
        lYstart = 0

        Dim dBoxH As Double
        Dim dDesiredHeight As Double
        Dim dDesiredLength As Double

        dBoxH = dBoxT - dBoxB 'height in data units
        dDesiredHeight = v_dInBoxT - v_dInBoxB 'height in desired units
        dDesiredLength = v_dInBoxL 'length in desired units
        'dBoxL is length in data units

        'setup vertical ticks
        If bSquareGrid = True Then 'calculate the vertical tick to produce a square grid
            dVTic = dHTic / CDbl(iProfileExagg)
        End If

        Dim lTicCount As Long 'the number of tick marks to add, should never add ticks above the box
        Dim lTicL As Long 'the length of the line representing the tic

        lTicCount = CLng(dDesiredHeight / dVTic) + 1  'the +1 is because the ticks start at a distance of 0
        lTicL = CLng(dVTic * CLng(iProfileExagg) / 2) 'this is the actual length of the tick itself
        If userUnits = "ft" Then 'make sure the tick line isn't ridiculously long if the grid is in feet
            lTicL = CLng(lTicL / cu)
        End If

        'convert tic height to data units 
        dVTic = dVTic / cu

        'setup horizontal ticks
        Dim lHTC As Long 'the number of horizontal tick marks to add
        lHTC = cu * (CLng(dDesiredLength / dHTic)) + 1

        'convert tic length to data units 
        dHTic = dHTic / cu

        Dim lHTL As Long 'the length of the horizontal line representing the tick

        lHTL = lTicL

        'make sure the box length unit works
        dBoxL = dBoxL * cu

        'setup counter (m_lCounter) for profile box
        m_lCounter = 0
        Dim progMax As Long = 0
        If m_bUpdateProgressBar = True Then
            Try
                SetProgBarLabelText(frmProfileTool.lblPrgBr, v_PrgBarMsg)
            Catch ex As Exception
                Debug.Print("Profile Box- error SetProgBarLabelText: " + ex.Message)
            End Try
            If bAddGrid = True Then
                Try
                    progMax = lTicCount * 3 + lHTC * 3 + 4
                    SetProgBarMax(frmProfileTool.prgBr, progMax)
                Catch ex As Exception
                    Debug.Print("Profile Box- error SetProgBarMax: " + ex.Message)
                End Try
            Else
                Try
                    progMax = lTicCount * 2 + lHTC * 2 + 4
                    SetProgBarMax(frmProfileTool.prgBr, progMax)
                Catch ex As Exception
                    Debug.Print("Profile Box- error SetProgBarMax: " + ex.Message)
                End Try
            End If
        End If

        'add box bottom
        Await QueuedTask.Run(Function()
                                 m_pPoint_Start = MapPointBuilder.CreateMapPoint(0, dBoxB * iProfileExagg)
                                 m_pPoint_End = MapPointBuilder.CreateMapPoint(dBoxL, dBoxB * iProfileExagg)
                                 m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                 Return True
                             End Function)
        Await AddLineToFile(frmProfileTool, diff, m_pLine, "LN")

        'add box top
        Await QueuedTask.Run(Function()
                                 m_pPoint_Start = MapPointBuilder.CreateMapPoint(0, ((dBoxB + dBoxH) * iProfileExagg))
                                 m_pPoint_End = MapPointBuilder.CreateMapPoint(dBoxL, ((dBoxB + dBoxH) * iProfileExagg))
                                 m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                 Return True
                             End Function)
        Await AddLineToFile(frmProfileTool, diff, m_pLine, "LN")

        'add box side left
        Await QueuedTask.Run(Function()
                                 m_pPoint_Start = MapPointBuilder.CreateMapPoint(0, dBoxB * iProfileExagg)
                                 m_pPoint_End = MapPointBuilder.CreateMapPoint(0, (dBoxB + dBoxH) * iProfileExagg)
                                 m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                 Return True
                             End Function)
        Await AddLineToFile(frmProfileTool, diff, m_pLine, "LN")

        'add box side right
        Await QueuedTask.Run(Function()
                                 m_pPoint_Start = MapPointBuilder.CreateMapPoint(dBoxL, dBoxB * iProfileExagg)
                                 m_pPoint_End = MapPointBuilder.CreateMapPoint(dBoxL, ((dBoxB + dBoxH) * iProfileExagg))
                                 m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                 Return True
                             End Function)
        Await AddLineToFile(frmProfileTool, diff, m_pLine, "LN")

        'add vertical tic marks
        Dim i As Integer
        Dim vertY As Double

        For i = 0 To lTicCount - 1
            'left tic
            Await QueuedTask.Run(Function()
                                     m_pPoint_Start = MapPointBuilder.CreateMapPoint(lTicL * -1, dVTic * iProfileExagg * i + dBoxB * iProfileExagg)
                                     m_pPoint_End = MapPointBuilder.CreateMapPoint(0, dVTic * iProfileExagg * i + dBoxB * iProfileExagg)
                                     m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                     Return True
                                 End Function)
            Await AddLineToFile(frmProfileTool, diff, m_pLine, "VT", (dBoxB + dVTic * i), boolConvert, cu)
            'right tic
            Await QueuedTask.Run(Function()
                                     m_pPoint_Start = MapPointBuilder.CreateMapPoint(dBoxL, dVTic * iProfileExagg * i + dBoxB * iProfileExagg)
                                     m_pPoint_End = MapPointBuilder.CreateMapPoint(dBoxL + lTicL, dVTic * iProfileExagg * i + dBoxB * iProfileExagg)
                                     m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                     Return True
                                 End Function)
            Await AddLineToFile(frmProfileTool, diff, m_pLine, "VT", (dBoxB + dVTic * i), boolConvert, cu)
            vertY = dVTic * iProfileExagg * i + dBoxB * iProfileExagg
            If bAddGrid = True Then
                Await QueuedTask.Run(Function()
                                         m_pPoint_Start = MapPointBuilder.CreateMapPoint(0, vertY)
                                         m_pPoint_End = MapPointBuilder.CreateMapPoint(dBoxL, vertY)
                                         m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                         Return True
                                     End Function)
                Await AddLineToFile(frmProfileTool, diff, m_pLine, "GD")
            End If
        Next

        Dim top As Double
        top = dVTic * iProfileExagg * (lTicCount - 1) + dBoxB * iProfileExagg
        'add horizontal tic marks
        For i = 0 To lHTC - 1
            If (dHTic * i) < dBoxL Then
                'top tic
                Await QueuedTask.Run(Function()
                                         m_pPoint_Start = MapPointBuilder.CreateMapPoint(dHTic * i, top)
                                         m_pPoint_End = MapPointBuilder.CreateMapPoint(dHTic * i, top + lHTL)
                                         m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                         Return True
                                     End Function)
                Await AddLineToFile(frmProfileTool, diff, m_pLine, "HT", (dHTic * i), boolConvert, cu)

                'bottom tic
                Await QueuedTask.Run(Function()
                                         m_pPoint_Start = MapPointBuilder.CreateMapPoint(dHTic * i, dBoxB * iProfileExagg)
                                         m_pPoint_End = MapPointBuilder.CreateMapPoint(dHTic * i, -1 * lHTL + dBoxB * iProfileExagg)
                                         m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                         Return True
                                     End Function)
                Await AddLineToFile(frmProfileTool, diff, m_pLine, "HT", (dHTic * i), boolConvert, cu)
                If bAddGrid = True Then
                    Await QueuedTask.Run(Function()
                                             m_pPoint_Start = MapPointBuilder.CreateMapPoint(dHTic * i, dBoxB * iProfileExagg)
                                             m_pPoint_End = MapPointBuilder.CreateMapPoint(dHTic * i, top)
                                             m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                             Return True
                                         End Function)
                    Await AddLineToFile(frmProfileTool, diff, m_pLine, "GD")
                End If
            End If
        Next

        'add box top
        Await QueuedTask.Run(Function()
                                 m_pPoint_Start = MapPointBuilder.CreateMapPoint(0, top)
                                 m_pPoint_End = MapPointBuilder.CreateMapPoint(dBoxL, top)
                                 m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                 Return True
                             End Function)
        Await AddLineToFile(frmProfileTool, diff, m_pLine, "LN")

        'add box side left
        Await QueuedTask.Run(Function()
                                 m_pPoint_Start = MapPointBuilder.CreateMapPoint(0, dBoxB * iProfileExagg)
                                 m_pPoint_End = MapPointBuilder.CreateMapPoint(0, top)
                                 m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                 Return True
                             End Function)
        Await AddLineToFile(frmProfileTool, diff, m_pLine, "LN")

        'add box side right
        Await QueuedTask.Run(Function()
                                 m_pPoint_Start = MapPointBuilder.CreateMapPoint(dBoxL, dBoxB * iProfileExagg)
                                 m_pPoint_End = MapPointBuilder.CreateMapPoint(dBoxL, top)
                                 m_pLine = LineBuilder.CreateLineSegment(m_pPoint_Start, m_pPoint_End)
                                 Return True
                             End Function)
        Await AddLineToFile(frmProfileTool, diff, m_pLine, "LN")

        If Not v_bHasPassedVariables Then
            Dim msg As String = "File location:" & vbNewLine & sWorkspace & sPolylineName
            If Not sWorkspace.Contains(".gdb") Then
                msg = msg & ".shp"
            End If
            MsgBox(msg)
        End If

        Dim newLineLayer As Layer
        Dim boxPath As Uri = New Uri(sWorkspace + "\" + sPolylineName)
        Await QueuedTask.Run(Function()
                                 newLineLayer = LayerFactory.Instance.CreateLayer(boxPath, MapView.Active.Map, 0)
                                 Return True
                             End Function)
        Return True
    End Function

    Private Async Function AddLineToFC(ByVal lineSegment As LineSegment,
                                       ByVal diff As Integer, ByVal v_sLineType As String, ByVal v_dTicValue As String) As Task(Of Boolean)
        Dim creationResult As Boolean = False
        Dim saveResult As Boolean = False
        Dim fc As FeatureClass = Nothing
        Dim fsConnectionPath As FileSystemConnectionPath
        Dim shapefile As FileSystemDatastore
        Dim filegdb As Geodatabase

        Dim message As String = String.Empty
        Await QueuedTask.Run(Function()
                                 If sWorkspace.Contains(".gdb") Then
                                     filegdb = New Geodatabase(New FileGeodatabaseConnectionPath(New Uri(sWorkspace)))
                                     fc = filegdb.OpenDataset(Of FeatureClass)(sPolylineName)
                                 Else
                                     fsConnectionPath = New FileSystemConnectionPath(New Uri(sWorkspace), FileSystemDatastoreType.Shapefile)
                                     shapefile = New FileSystemDatastore(fsConnectionPath)
                                     fc = shapefile.OpenDataset(Of FeatureClass)(sPolylineName)
                                 End If
                                 Using fc
                                     Dim rowBuff As RowBuffer = Nothing
                                     Dim newLine As Feature = Nothing
                                     Try
                                         Dim editOperation As New EditOperation()
                                         editOperation.Callback(Function(context)
                                                                    Dim fcDef As FeatureClassDefinition = fc.GetDefinition()
                                                                    rowBuff = fc.CreateRowBuffer()
                                                                    rowBuff(fcDef.GetShapeField()) = New PolylineBuilder(lineSegment).ToGeometry()
                                                                    rowBuff(ticLabel) = v_dTicValue
                                                                    rowBuff("Type") = v_sLineType
                                                                    newLine = fc.CreateRow(rowBuff)
                                                                    context.Invalidate(newLine)
                                                                    Return True
                                                                End Function, CType(fc, Dataset))
                                         creationResult = editOperation.Execute()
                                         Try
                                             Dim objectID As Long = newLine.GetObjectID()
                                             Debug.Print("Creation result for " + objectID.ToString + " " + creationResult.ToString)
                                         Catch ref As NullReferenceException
                                             Debug.Print(ref.Message)
                                         End Try
                                         saveResult = Project.Current.SaveEditsAsync().Result
                                     Catch ex As GeodatabaseEditingException
                                         Console.WriteLine("Gdb error in createAddIntersections: " + ex.Message)
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
        Return saveResult
    End Function

    Private Async Function AddLineToFile(ByVal frmProfileTool As frmProfileTool, ByVal diff As Integer, ByVal m_line As LineSegment,
        Optional ByVal v_sLineType As String = "",
        Optional ByVal v_dTicValue As Double = 1,
        Optional ByVal convert As Boolean = False,
        Optional ByVal cu As Double = 1) As Task

        'see what units the data is in
        If convert = True Then
            v_dTicValue = Round(v_dTicValue * cu)
        End If

        Try
            Await addLineToFC(m_line, diff, v_sLineType, v_dTicValue)
            m_lCounter = m_lCounter + 1
            If m_bUpdateProgressBar = True Then
                Try
                    RefreshProgBarValue(frmProfileTool.prgBr, m_lCounter)
                Catch ex As Exception
                    Debug.Print("RefreshProgBarValue error in AddLineToFile: " + ex.Message)
                End Try
            End If
        Catch ex As AggregateException
            For Each x In ex.InnerExceptions
                Debug.Print("Aggregate Ex. error in AddLineToFile" + x.Message)
            Next
        End Try
    End Function

End Module
