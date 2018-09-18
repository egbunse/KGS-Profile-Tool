Imports System.Math
Imports System.IO
Imports System.Windows.Forms
Imports ArcGIS.Desktop.Mapping
Imports ArcGIS.Core.Data
Imports ArcGIS.Desktop.Framework.Threading.Tasks
Imports System.ComponentModel
Imports System.Threading
Imports ArcGIS.Desktop.Framework

Public Class FrmProfileTool
    Inherits Form

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

    Private m_asTocInfo() As String 'array to hold TOC layer names and positions
    Private m_iTOCSurface As Integer 'the TOC position for the surface units layer
    Private m_iTOCElevation As Integer 'the TOC position for the elevation grid
    Private m_iTOCWells As Integer 'the TOC position for the wells layer

    Private m_dMapUnitsToUserUnits As Double
    Private m_dUserUnitsToMapUnits As Double
    Private m_dElevationLayerUnitsToMapUnits As Double
    Private m_dSubsurfaceTableUnitsToMapUnits As Double
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Private m_sUserInputErrors As String

    Public progressor As CancelableProgressor
    Public cps As CancelableProgressorSource = Nothing

    'Fills in ListBoxes with values
    Public Sub PopulateListBox(listBoxName As ListBox, itemName As String)
        If listBoxName.InvokeRequired Then
            listBoxName.Invoke(Sub() PopulateListBox(listBoxName, itemName))
        Else
            listBoxName.Items.Add(itemName)
        End If
    End Sub

    'Changes the tab which is viewed: 'Profile Options,' 'Profile Layers and Paths,' 'Profile Box,' 'Wells'
    Public Sub ChangeSelIndex(multiPg As TabControl, index As Integer)
        If multiPg.InvokeRequired Then
            multiPg.Invoke(Sub() ChangeSelIndex(multiPg, index))
        Else
            multiPg.SelectedIndex = index
        End If
    End Sub

    'Main Code
    Private Sub ChkGetElevFromDEM_Click(sender As Object, e As EventArgs) Handles chkGetElevFromDEM.CheckedChanged
        Select Case chkGetElevFromDEM.Checked
            Case True
                lstWellsElev.Enabled = False
                Dim gray As System.Drawing.Color = System.Drawing.Color.LightGray
                lstWellsElev.BackColor = gray
            Case Else
                lstWellsElev.Enabled = True
                Dim white As System.Drawing.Color = System.Drawing.Color.White
                lstWellsElev.BackColor = white
                'prevent hidden selections in the list boxes
                lstWellsElev.SelectedIndex = -1
        End Select
    End Sub

    Private Sub ChkSaveSettings_Click(sender As Object, e As EventArgs) Handles chkSaveSettings.CheckedChanged
        'forget settings if user uncheck the box
        'if checked, settings will be saved when the Create Profile button is clicked.
        If chkSaveSettings.Checked = False Then
            SaveSetting("ProfileTool", "Settings", "chkSaveSettings", CStr(chkSaveSettings.Checked))
        End If
    End Sub

    Private Sub MultiPage1_Change()
        'do nothing...
    End Sub

    Private Sub OptLines_Change(sender As Object, e As EventArgs) Handles optLines.CheckedChanged
        Select Case optLines.Checked
            Case True
                txtPolyW.Enabled = False
                txtPolyW.BackColor = System.Drawing.Color.LightGray
            Case Else
                txtPolyW.Enabled = True
                txtPolyW.BackColor = System.Drawing.Color.White
        End Select
    End Sub

    Private Sub OptProfileDelineationOff_Click(sender As Object, e As EventArgs) Handles optProfileDelineationOff.CheckedChanged
        'disable and gray out surface unit layer and field options
        lstSurfaceUnitLayer.Enabled = False
        lstSurfaceUnitField.Enabled = False
        lstSurfaceUnitLayer.BackColor = System.Drawing.Color.LightGray
        lstSurfaceUnitField.BackColor = System.Drawing.Color.LightGray
    End Sub

    Private Sub OptProfileDelineationOn_Click(sender As Object, e As EventArgs) Handles optProfileDelineationOn.CheckedChanged
        'enable surface unit layer and field options
        lstSurfaceUnitLayer.Enabled = True
        lstSurfaceUnitField.Enabled = True
        lstSurfaceUnitLayer.BackColor = System.Drawing.Color.White
        lstSurfaceUnitField.BackColor = System.Drawing.Color.White

        'prevent hidden selections in the list boxes
        lstSurfaceUnitLayer.SelectedIndex = -1
    End Sub

    Private Sub OptWellUnitsFt_Change(sender As Object, e As EventArgs) Handles optWellUnitsFt.CheckedChanged
        'change form lables
        If optWellUnitsFt.Checked = True Then
            lblWellsElev.Text = "Well surface elevation (ft) field:"
            lblWellTop.Text = "Layer top depth (ft) field:"
            lblwellBot.Text = "Layer bottom depth (ft) field:"
        Else
            lblWellsElev.Text = "Well surface elevation (m) field:"
            lblWellTop.Text = "Layer top depth (m) field:"
            lblwellBot.Text = "Layer bottom depth (m) field:"
        End If

        Dim pMap As Map = MapView.Active.Map
        'Based on this table: http://pro.arcgis.com/en/pro-app/sdk/api-reference/index.html#topic8349.html
        'http://pro.arcgis.com/en/pro-app/sdk/api-reference/index.html#topic8350.html

        If optWellUnitsFt.Checked = True Then 'ft
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
        Debug.Print("Subsurface Units: " + Str(m_dSubsurfaceTableUnitsToMapUnits))
        pMap = Nothing
        'pMxDocument = Nothing
    End Sub

    Private Sub TxtPolyW_Change(sender As Object, e As EventArgs) Handles txtPolyW.TextChanged
        'Make sure input is numaric
        txtPolyW.Text = CheckInput(txtPolyW.Text, "Well polygon width")
    End Sub

    Private Sub TxtVerticalExagg_Change(sender As Object, e As EventArgs) Handles txtVerticalExagg.TextChanged
        'Make sure input is numaric
        txtVerticalExagg.Text = CheckInput(txtVerticalExagg.Text, "Vertical exaggeration")

        'make sure input is >= 1
        Dim varToCheck As String
        varToCheck = txtVerticalExagg.Text

        If CDbl(varToCheck) < 1 Then
            Do Until CDbl(varToCheck) >= 1
                varToCheck = InputBox("Please enter a numeric (double) value >= 1", "Input value for: Vertical exaggeration")
            Loop
        End If

        txtVerticalExagg.Text = varToCheck
    End Sub

    Public Sub New()
        InitializeComponent()
        ''Call InitComponent()
        '' Setup the user form and display to user
        ''fill the profile data sources list boxes
        Call ListBoxFiller()

        ''set a better default state for the profile box options page
        chkSquareGrid.Checked = True
        ''Call chkSquareGrid_Click(sender, E)

        ''setup units and conversions
        Call SetupUnitMultipliers()
        ''Call optElevationFt_Change(sender, e)
        ''Call optWellUnitsFt_Change(sender, e)

        ''update form with saved settings, if user choose to do so
        If GetSetting("ProfileTool", "Settings", "chkSaveSettings") = "True" Then
            '    'update form with saved settings
            Call GetUserSettings()
            'update profile precision text boxes
            If optProfileDelineationOn.Checked = True Then
                Call CalculateDistBetweenPoints()
            Else
                Call CalculateNumberOfPoints()
            End If
        Else
            'update profile precision text boxes
            If optProfileDelineationOn.Checked = True Then
                Call CalculateDistBetweenPoints()
            Else
                Call CalculateNumberOfPoints()
            End If
        End If
        ''update profile length
        Call CalculateProfileLength()
    End Sub

    Private Sub SetupUnitMultipliers()

        Dim pMap As Map = MapView.Active.Map

        If pMap.SpatialReference.Unit.FactoryCode <> 9002 And pMap.SpatialReference.Unit.FactoryCode <> 9001 Then   'Not Feet or Meters
            'unknown or unsuported units, assume feet
            MsgBox("Could not determine the base units for the map," & Chr(13) _
                & "or the map base units are not supported by this tool." & Chr(13) _
                & "Assuming map base units are meters.", vbOKOnly, "Error")
        End If

        'Factory Codes from: http://pro.arcgis.com/en/pro-app/sdk/api-reference/index.html#topic8350.html
        Select Case pMap.SpatialReference.Unit.FactoryCode
            Case 9002, 109008, 9093, 9030, 109001     'ft, in, mi, nautical mi, yds     '1(in), 3(ft), 4(yds), 5(mi), 6(nautical mi) 'standard units

                'standard units, request form entries as feet (3)
                'lblPrecisionUnit.Text = "feet"
                'lblPolyW.Text = "Polygon width (ft):"
                'lblBoxTop.Text = "Top elevation (feet):"
                'lblBoxBottom.Text = "Bottom elevation (feet):"
                'lblBoxLength.Text = "Length (feet):"
                'lblVerticalTick.Text = "Vertical tick spacing (feet):"
                'lblHorizTick.Text = "Horizontal tick spacing (feet):"
                'lblMinorVertTick.Text = "Vertical tick spacing (feet):"
                'lblMinorHorizTick.Text = "Horizontal tick spacing (feet):"

                If pMap.SpatialReference.Unit.FactoryCode = 9002 Then 'ft
                    m_dUserUnitsToMapUnits = 1 'ft to ft conversion
                    m_dMapUnitsToUserUnits = 1 'ft to ft conversion
                Else
                    m_dUserUnitsToMapUnits = 381 / 1250 'ft to m conversion
                    m_dMapUnitsToUserUnits = 1250 / 381 'm to ft conversion
                End If

            Case 109006, 109005, 9036, 9001, 109007    'cm, dm, km, m, mm                 '7(mm), 8(cm), 9(m), 10(km), 12(dm) 'metric units

                'metric units, request form entries as meters (9)
                'lblPrecisionUnit.Text = "meters"
                'lblPolyW.Text = "Polygon width (m):"
                'lblBoxTop.Text = "Top elevation (meters):"
                'lblBoxBottom.Text = "Bottom elevation (meters):"
                'lblBoxLength.Text = "Length (meters):"
                'lblVerticalTick.Text = "Vertical tick spacing (meters):"
                'lblHorizTick.Text = "Horizontal tick spacing (meters):"
                'lblMinorVertTick.Text = "Vertical tick spacing (meters):"
                'lblMinorHorizTick.Text = "Horizontal tick spacing (meters):"

                If pMap.SpatialReference.Unit.FactoryCode = 9002 Then 'ft
                    m_dUserUnitsToMapUnits = 1250 / 381 'm to ft conversion
                    m_dMapUnitsToUserUnits = 381 / 1250 'ft to m conversion
                Else
                    m_dUserUnitsToMapUnits = 1 'm to m conversion
                    m_dMapUnitsToUserUnits = 1 'm to m conversion
                End If

            Case Else

                'other type of unit, request form entries as meters (9)
                lblPrecisionUnit.Text = "meters"
                lblPolyW.Text = "Polygon width (m):"
                lblBoxTop.Text = "Top elevation (meters):"
                lblBoxBottom.Text = "Bottom elevation (meters):"
                lblBoxLength.Text = "Length (feet):"
                lblVerticalTick.Text = "Vertical tick spacing (meters):"
                lblHorizTick.Text = "Horizontal tick spacing (meters):"
                lblMinorVertTick.Text = "Vertical tick spacing (meters):"
                lblMinorHorizTick.Text = "Horizontal tick spacing (meters):"

                If pMap.SpatialReference.Unit.FactoryCode = 9002 Then 'ft
                    m_dUserUnitsToMapUnits = 1250 / 381 'm to ft conversion
                    m_dMapUnitsToUserUnits = 381 / 1250 'ft to m conversion
                Else
                    m_dUserUnitsToMapUnits = 1 'm to m conversion
                    m_dMapUnitsToUserUnits = 1 'm to m conversion
                End If

        End Select

        pMap = Nothing
        'pMxDocument = Nothing
    End Sub

    Private Sub ChkVerticalExagg_Click(sender As Object, e As EventArgs) Handles chkVerticalExagg.CheckedChanged
        ' Determines if a vertical exaggeration can be specified
        If chkVerticalExagg.Checked = True Then
            txtVerticalExagg.Enabled = True
        Else
            txtVerticalExagg.Enabled = False
            txtVerticalExagg.Text = "1"
        End If
    End Sub

    Private Sub OptPrecisionMeasure_Click(sender As Object, e As EventArgs) Handles optPrecisionMeasure.CheckedChanged
        ' User chooses to specify the distance (meters) between points in the profile
        txtPrecisionParts.Enabled = False
        txtPrecisionMeasure.Enabled = True
        Call CalculateNumberOfPoints()
    End Sub

    Private Sub TxtPrecisionMeasure_Change(sender As Object, e As EventArgs) Handles txtPrecisionMeasure.TextChanged
        'If enabled, make sure input is numaric
        If txtPrecisionMeasure.Enabled = True Then
            If txtPrecisionMeasure.Text <> "" Or txtPrecisionMeasure.Text = " " Then
                txtPrecisionMeasure.Text = CheckInput(txtPrecisionMeasure.Text, "Precision: Measure")
                Call CalculateNumberOfPoints()
            End If
        End If
    End Sub

    Private Sub TxtPrecisionParts_Change(sender As Object, e As EventArgs) Handles txtPrecisionParts.TextChanged
        'If enabled, make sure input is numaric
        If txtPrecisionParts.Enabled = True Then
            If txtPrecisionParts.Text <> "" Or txtPrecisionParts.Text = " " Then
                txtPrecisionParts.Text = CheckInput(txtPrecisionParts.Text, "Precision: Parts")
                Call CalculateDistBetweenPoints()
            End If
        End If
    End Sub

    Private Sub CalculateDistBetweenPoints()
        ' This calculates the distance between points and displays info to user.
        Try
            txtPrecisionMeasure.Text = CStr(Round((Sqrt((CDbl(txtEndX.Text) - CDbl(txtStartX.Text)) ^ 2 _
                        + (CDbl(txtEndY.Text) - CDbl(txtStartY.Text)) ^ 2) / CDbl(txtPrecisionParts.Text)) * m_dMapUnitsToUserUnits, 2))
        Catch ex As Exception
            txtPrecisionMeasure.Text = ""
        End Try
    End Sub

    Private Sub CalculateNumberOfPoints()
        ' This calculates the number of points along the profile and displays info to user.
        Try
            txtPrecisionParts.Text = CStr(Round(Sqrt((CDbl(txtEndX.Text) - CDbl(txtStartX.Text)) ^ 2 +
                (CDbl(txtEndY.Text) - CDbl(txtStartY.Text)) ^ 2) / (CDbl(txtPrecisionMeasure.Text / m_dMapUnitsToUserUnits)), 0))
        Catch ex As Exception
            txtPrecisionParts.Text = ""
        End Try
    End Sub

    Private Sub CalculateProfileLength()
        ' This calculates the length of the profile to display on the profile box page of the user form
        Try
            If Not CStr(CLng((Sqrt((CDbl(txtEndX.Text) - CDbl(txtStartX.Text)) ^ 2 _
                    + (CDbl(txtEndY.Text) - CDbl(txtStartY.Text)) ^ 2)) * m_dMapUnitsToUserUnits)) = 0 Then
                'update textbox
                Dim expression As String

                If optBoxUnitsFT.Checked = True Then
                    'convert meters to feet
                    expression = CStr(CLng((Sqrt((CDbl(txtEndX.Text) - CDbl(txtStartX.Text)) ^ 2 _
                                + (CDbl(txtEndY.Text) - CDbl(txtStartY.Text)) ^ 2)) * m_dMapUnitsToUserUnits)) * 3.28084
                Else
                    'leave meters as meters
                    expression = CStr(CLng((Sqrt((CDbl(txtEndX.Text) - CDbl(txtStartX.Text)) ^ 2 _
                                + (CDbl(txtEndY.Text) - CDbl(txtStartY.Text)) ^ 2)) * m_dMapUnitsToUserUnits))
                End If

                txtBoxL.Text = expression
            Else
                'this hits on the form activate, when start and end coords are added, but
                'the unit multipliers have not yet been set.
                'ignore this state...
            End If
        Catch ex As Exception
            txtPrecisionMeasure.Text = ""
        End Try
    End Sub

    Private Sub TxtStartX_Change(sender As Object, e As EventArgs) Handles txtStartX.TextChanged
        If txtStartX.Text <> "X (start)" Then
            'Make sure input is numaric
            txtStartX.Text = CheckInput(txtStartX.Text, "Profile startpoint (X coordinate)")
            'Update the precision for the new coords
            Call UpdatePrecision()
            Call CalculateProfileLength()
        End If
    End Sub

    Private Sub TxtEndX_Change(sender As Object, e As EventArgs) Handles txtEndX.TextChanged
        If txtEndX.Text <> "X (end)" Then
            'Make sure input is numaric
            txtEndX.Text = CheckInput(txtEndX.Text, "Profile endpoint (X coordinate)")
            'Update the precision for the new coords
            Call UpdatePrecision()
            Call CalculateProfileLength()
        End If
    End Sub

    Private Sub TxtStartY_Change(sender As Object, e As EventArgs) Handles txtStartY.TextChanged
        If txtStartY.Text <> "Y (start)" Then
            'Make sure input is numaric
            txtStartY.Text = CheckInput(txtStartY.Text, "Profile startpoint (Y coordinate)")
            'Update the precision for the new coords
            Call UpdatePrecision()
            Call CalculateProfileLength()
        End If
    End Sub

    Private Sub TxtEndY_Change(sender As Object, e As EventArgs) Handles txtEndY.TextChanged
        If txtEndY.Text <> "Y (end)" Then
            'Make sure input is numaric
            txtEndY.Text = CheckInput(txtEndY.Text, "Profile endpoint (Y coordinate)")
            'Update the precision for the new coords
            Call UpdatePrecision()
            Call CalculateProfileLength()
        End If
    End Sub

    Private Sub UpdatePrecision()
        ' This updates the precision whenever the start or end point coords are changed
        If txtPrecisionParts.Enabled Then
            Call CalculateDistBetweenPoints()
        Else
            Call CalculateNumberOfPoints()
        End If
    End Sub

    Private Sub ChkBuildProfileBox_Click(sender As Object, e As EventArgs) Handles chkBuildProfileBox.CheckedChanged
        ' Show the form to setup the profile box
        Select Case chkBuildProfileBox.Checked
            Case True
                MultiPage1.TabPages.Add(Page4)
            Case Else
                MultiPage1.TabPages.Remove(Page4)
        End Select
        MultiPage1.Refresh()
    End Sub

    Private Sub ChkAddWells_Click(sender As Object, e As EventArgs) Handles chkAddWells.CheckedChanged
        ' Show the form to setup the wells data
        Select Case chkAddWells.Checked
            Case True
                MultiPage1.TabPages.Add(Page3)
            Case Else
                MultiPage1.TabPages.Remove(Page3)
        End Select
        Me.Refresh()
    End Sub

    Public Async Sub ListBoxFiller()
        ' Fill all list boxes based on layers in the TOC
        Dim pMap As Map 'IMap
        Dim pEnumLayer As IEnumerable(Of Layer) 'IEnumLayer
        Dim pFeatureLayer As FeatureLayer
        Dim pFeatureClass As FeatureClass
        pMap = MapView.Active.Map
        pEnumLayer = pMap.Layers 'This will see all layers, layer files, and layers w/in layer files
        If pEnumLayer.Count = 0 Then
            MsgBox("Could not list layers in the map's TOC.", vbCritical, "Error")
            CloseForm()
        End If

        Dim i As Integer
        ReDim m_asTocInfo(256)

        i = 0
        'Do While Not pLayer Is Nothing
        Dim shapeTypeString As String = ""
        For Each pLayer In pEnumLayer
            'add layer name and TOC position to the array m_asTocInfo array
            m_asTocInfo(i) = pLayer.Name
            If TypeOf pLayer Is FeatureLayer And pLayer.ConnectionStatus = ConnectionStatus.Connected Then 'And pLayer.Valid = True Then
                pFeatureLayer = pLayer
                pFeatureClass = Await QueuedTask.Run(Function() pFeatureLayer.GetFeatureClass())
                'pFeatureClass = pFeatureLayer.GetFeatureClass()
                shapeTypeString = Await QueuedTask.Run(Function() pFeatureClass.GetDefinition.GetShapeType().ToString)
                If shapeTypeString = "Polygon" Then     'If pFeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon Then
                    'add to surface polygon list
                    ChangeSelIndex(MultiPage1, 1)
                    PopulateListBox(lstSurfaceUnitLayer, pLayer.Name)
                    'MultiPage1.SelectedIndex = 1
                    'lstSurfaceUnitLayer.Items.Add(pLayer.Name)
                ElseIf shapeTypeString = "Point" Then   'ElseIf pFeatureClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPoint Then
                    'add to well point list
                    ChangeSelIndex(MultiPage1, 2)
                    PopulateListBox(lstWellLayer, pLayer.Name)
                    'MultiPage1.SelectedIndex = 2
                    'lstWellLayer.Items.Add(pLayer.Name)
                End If
            ElseIf TypeOf pLayer Is RasterLayer And pLayer.ConnectionStatus = ConnectionStatus.Connected Then 'And pLayer.Valid = True Then
                'add to elevation grid list
                'MultiPage1.SelectedIndex = 1
                ChangeSelIndex(MultiPage1, 1)
                PopulateListBox(lstElevationLayer, pLayer.Name)
                'lstElevationLayer.Items.Add(pLayer.Name)
            End If

            'pLayer = pEnumLayer.Next

            i = i + 1
        Next
        'Loop

        'add all standalone tables to the listbox
        For Each pStandaloneTable In pMap.StandaloneTables
            ChangeSelIndex(MultiPage1, 2)
            PopulateListBox(lstTables, pStandaloneTable.Name)
            'MultiPage1.SelectedIndex = 2
            'lstTables.Items.Add(pStandaloneTable.Name)
        Next

        ReDim Preserve m_asTocInfo(i - 1)

        'select the first layer in each list box
        'but do not select any fields as user will need to do this

        If lstElevationLayer.Items.Count > 0 Then
            lstElevationLayer.SelectedIndex = 0
        End If

        If lstSurfaceUnitLayer.Items.Count > 0 Then
            lstSurfaceUnitLayer.SelectedIndex = 0
        End If

        If lstWellLayer.Items.Count > 0 Then
            lstWellLayer.SelectedIndex = 0
        End If

        If lstTables.Items.Count > 0 Then
            lstTables.SelectedIndex = 0
        End If

        'now that the list boxs have been filled, hide the wells page
        Page3.Visible = False
        MultiPage1.SelectedIndex = 0

        'clear memory
        pMap = Nothing
        pEnumLayer = Nothing
        pFeatureLayer = Nothing
        pFeatureClass = Nothing

        Exit Sub
    End Sub

    Private Async Sub FirstSurfaceUnitLayer_Change(sender As Object, e As EventArgs) Handles lstSurfaceUnitLayer.SelectedValueChanged

        If lstSurfaceUnitLayer.SelectedIndex <> -1 Then
            'something is selected, okay to continue
            Dim pMap As Map
            Dim pEnumLayer As IEnumerable(Of Layer)
            Dim pFeatureLayer As FeatureLayer
            Dim pFeatureClass As FeatureClass
            Dim pFields As IReadOnlyList(Of Field)
            Dim pField As Field

            pMap = MapView.Active.Map
            pEnumLayer = pMap.Layers 'This will see all layers, layer files, and layers w/in layer files

            'switch to the profile options page
            MultiPage1.SelectedIndex = 1

            lstSurfaceUnitField.Items.Clear()

            Dim lFldCntr As Long
            ' loop through all feature layers to find the one
            ' selected feature class in the polygon layer list box
            For Each pLayer In pEnumLayer
                If pLayer.Name = lstSurfaceUnitLayer.SelectedItem.ToString Then
                    pFeatureLayer = pLayer
                    pFeatureClass = Await QueuedTask.Run(Function() pFeatureLayer.GetFeatureClass())
                    'loop through all fields of the feature class found above
                    pFields = Await QueuedTask.Run(Function() pFeatureClass.GetDefinition().GetFields())
                    If pFields.Count = 0 Then
                        MsgBox("Could not read field names in the selected surface units layer.", vbCritical, "Error")
                        CloseForm()
                    Else
                        For lFldCntr = 0 To (pFields.Count - 1)
                            pField = pFields.Item(lFldCntr)
                            If pField.Name <> "Shape" Then
                                PopulateListBox(lstSurfaceUnitField, pField.Name)
                                'lstSurfaceUnitField.Items.Add(pField.Name)
                            End If
                        Next lFldCntr
                        Exit For
                    End If
                End If
            Next

            pMap = Nothing
            pEnumLayer = Nothing
            pFeatureLayer = Nothing
            pFeatureClass = Nothing
            pFields = Nothing
            pField = Nothing
        Else
            'nothing is selected, so don't display any fields
            lstSurfaceUnitField.Items.Clear()
        End If
        Exit Sub
    End Sub

    Private Async Sub FirstWellLayer_Change(sender As Object, e As EventArgs) Handles lstWellLayer.SelectedValueChanged
        If lstWellLayer.SelectedIndex <> -1 Then
            'something is selected, okay to continue

            Dim pMap As Map
            Dim pEnumLayer As IEnumerable(Of Layer)
            Dim pFeatureLayer As FeatureLayer
            Dim pFeatureClass As FeatureClass
            Dim pFields As IReadOnlyList(Of Field)
            Dim pField As Field

            pMap = MapView.Active.Map
            pEnumLayer = pMap.Layers 'This will see all layers, layer files, and layers w/in layer files

            'switch to wells page
            MultiPage1.SelectedIndex = 2

            'clear field selections
            lstWellsPointID.Items.Clear()
            lstWellsElev.Items.Clear()

            Dim lFldCntr As Long
            'loop through all layers to find the selected feature class
            'Do While Not pLayer Is Nothing
            For Each pLayer In pEnumLayer
                If pLayer.Name = lstWellLayer.SelectedItem.ToString Then
                    pFeatureLayer = pLayer
                    pFeatureClass = Await QueuedTask.Run(Function() pFeatureLayer.GetFeatureClass())
                    'loop through all fields of the selected feature class
                    pFields = Await QueuedTask.Run(Function() pFeatureClass.GetDefinition().GetFields())
                    If pFields.Count = 0 Then
                        MsgBox("Could not read field names in the selected well layer.", vbCritical, "Error")
                        CloseForm()
                    Else
                        For lFldCntr = 0 To (pFields.Count - 1)
                            pField = pFields.Item(lFldCntr)
                            If pField.Name <> "Shape" Then
                                PopulateListBox(lstWellsPointID, pField.Name)
                                PopulateListBox(lstWellsElev, pField.Name)
                                'lstWellsPointID.Items.Add(pField.Name)
                                'lstWellsElev.Items.Add(pField.Name)
                            End If
                        Next lFldCntr
                        Exit For
                    End If
                End If
                'pLayer = pEnumLayer.Next
            Next
            'Loop

            'clear memory
            pMap = Nothing
            pEnumLayer = Nothing
            pFeatureLayer = Nothing
            pFeatureClass = Nothing
            pFields = Nothing
            pField = Nothing
        Else
            'nothing is selected, so don't display any fields
            lstWellsPointID.Items.Clear()
            lstWellsElev.Items.Clear()
        End If

        Exit Sub
    End Sub

    Private Async Sub FirstTables_Change(sender As Object, e As EventArgs) Handles lstTables.SelectedValueChanged
        If lstTables.SelectedIndex <> -1 Then
            'something is selected, okay to continue
            Dim pMap As Map
            Dim pEnumLayer As IEnumerable(Of Layer)
            Dim pFields As IReadOnlyList(Of Field)
            Dim pField As Field
            Dim pTable As Table

            pMap = MapView.Active.Map
            pEnumLayer = pMap.Layers 'This will see all layers, layer files, and layers w/in layer files

            'switch to wells page
            MultiPage1.SelectedIndex = 2

            'clear field selections
            lstTablePointID.Items.Clear()
            lstTableTop.Items.Clear()
            lstTableBottom.Items.Clear()
            lstTableUnit.Items.Clear()

            Dim lFldCntr As Long
            'loop through all standalone tables to find the selected table
            For Each pStandaloneTable In pMap.StandaloneTables
                If pStandaloneTable.Name = lstTables.Text And pStandaloneTable.ConnectionStatus = ConnectionStatus.Connected Then
                    pTable = Await QueuedTask.Run(Function() pStandaloneTable.GetTable())
                    pFields = Await QueuedTask.Run(Function() pTable.GetDefinition().GetFields())
                    'pTable = pStandaloneTable.GetTable()
                    'pFields = pTable.GetDefinition().GetFields()
                    If pFields.Count = 0 Then
                        MsgBox("Could not read field names in the selected standalone table.", vbCritical, "Error")
                        CloseForm()
                    Else
                        'add field names to the list box
                        For lFldCntr = 0 To (pFields.Count - 1)
                            pField = pFields.Item(lFldCntr)
                            If pField.Name <> "Shape" Then
                                PopulateListBox(lstTablePointID, pField.Name)
                                PopulateListBox(lstTableTop, pField.Name)
                                PopulateListBox(lstTableBottom, pField.Name)
                                PopulateListBox(lstTableUnit, pField.Name)
                            End If
                        Next lFldCntr
                    End If
                End If
            Next

            'clear memory
            pMap = Nothing
            pEnumLayer = Nothing
            pFields = Nothing
            pField = Nothing
        Else
            'nothing is selected, so don't display any fields
            lstTablePointID.Items.Clear()
            lstTableUnit.Items.Clear()
            lstTableTop.Items.Clear()
            lstTableBottom.Items.Clear()
        End If
        Exit Sub
    End Sub

    Private Sub ChkBuildGrid_Click(sender As Object, e As EventArgs) Handles chkBuildGrid.CheckedChanged
        Select Case chkBuildGrid.Checked
            Case True
                chkSquareGrid.Enabled = True
                txtVTick.Enabled = True
                txtVTick.BackColor = System.Drawing.Color.White
                If chkMinorTick.Checked = True Then
                    txtVTickMinor.Enabled = True
                    txtVTickMinor.BackColor = System.Drawing.Color.White
                End If
            Case Else
                chkSquareGrid.Enabled = False
                chkSquareGrid.Checked = False
                txtVTick.Enabled = True
                txtVTick.BackColor = System.Drawing.Color.White
                If chkMinorTick.Checked = True Then
                    txtVTickMinor.Enabled = True
                    txtVTickMinor.BackColor = System.Drawing.Color.White
                End If
        End Select
    End Sub

    Private Sub ChkMinorTick_Click(sender As Object, e As EventArgs) Handles chkMinorTick.CheckedChanged
        Select Case chkMinorTick.Checked
            Case True
                If chkSquareGrid.Checked = True Then
                    txtVTickMinor.Enabled = False
                    txtVTickMinor.BackColor = System.Drawing.Color.LightGray
                Else
                    txtVTickMinor.Enabled = True
                    txtVTickMinor.BackColor = System.Drawing.Color.White
                End If
                txtHTickMinor.Enabled = True
                txtHTickMinor.BackColor = System.Drawing.Color.White
            Case Else
                txtVTickMinor.Enabled = False
                txtVTickMinor.BackColor = System.Drawing.Color.LightGray
                txtHTickMinor.Enabled = False
                txtHTickMinor.BackColor = System.Drawing.Color.LightGray
        End Select
    End Sub

    Private Sub ChkSquareGrid_Click(sender As Object, e As EventArgs) Handles chkSquareGrid.CheckedChanged
        Select Case chkSquareGrid.Checked
            Case True
                'calculate the square grid dimension on the fly
                Dim vExagg As Integer = CInt(txtVerticalExagg.Text)
                If vExagg <> Nothing Then
                    If vExagg = 0 Then
                        txtVTick.Text = txtHTick.Text
                    Else
                        txtVTick.Text = txtHTick.Text / vExagg
                    End If
                Else
                    txtVTick.Text = ""
                End If
                txtVTick.Enabled = False
                txtVTick.BackColor = System.Drawing.Color.LightGray
                If chkMinorTick.Checked = True Then
                    If vExagg <> Nothing Then
                        If vExagg = 0 Then
                            txtVTickMinor.Text = txtHTickMinor.Text
                        Else
                            txtVTickMinor.Text = txtHTickMinor.Text / vExagg
                        End If
                    Else
                        txtVTickMinor.Text = ""
                    End If
                    txtVTickMinor.Enabled = False
                    txtVTickMinor.BackColor = System.Drawing.Color.LightGray
                End If
            Case Else
                txtVTick.Enabled = True
                txtVTick.BackColor = System.Drawing.Color.White
                If chkMinorTick.Checked = True Then
                    txtVTickMinor.Enabled = True
                    txtVTickMinor.BackColor = System.Drawing.Color.White
                End If
        End Select
    End Sub

    Private Sub ChkUseProfileL_Click(sender As Object, e As EventArgs) Handles chkUseProfileL.CheckedChanged
        Select Case chkUseProfileL.Checked
            Case True
                Call CalculateProfileLength()
                txtBoxL.Enabled = False
                txtBoxL.BackColor = System.Drawing.Color.LightGray
            Case Else
                txtBoxL.Enabled = True
                txtBoxL.BackColor = System.Drawing.Color.White
        End Select
    End Sub

    Private Function CheckInput(varToCheck As Object, strTxtName As String) As Object
        'check the input of the text boxes to make sure they are numeric
        If varToCheck = "-" Then
            CheckInput = varToCheck
            Exit Function
        ElseIf varToCheck = "0" Then
            CheckInput = varToCheck
            Exit Function
        End If
        If strTxtName = "Vertical tick spacing (major)" _
            Or strTxtName = "Vertical tick spacing (minor)" Then
            Do Until IsNumeric(varToCheck)
                varToCheck = InputBox("Please enter a numeric (double) value:", "Input value for: " & strTxtName)
            Loop
        Else
            Do Until IsNumeric(varToCheck)
                varToCheck = InputBox("Please enter a long integer value:", "Input value for: " & strTxtName)
            Loop
        End If
        CheckInput = varToCheck
    End Function

    Private Sub TxtBoxB_TextChanged(sender As Object, e As EventArgs) Handles txtBoxB.TextChanged
        'Make sure input is numaric
        txtBoxB.Text = CheckInput(txtBoxB.Text, "Box bottom elevation")
    End Sub

    Private Sub TxtBoxL_TextChanged(sender As Object, e As EventArgs) Handles txtBoxL.TextChanged
        'Make sure input is numaric
        txtBoxL.Text = CheckInput(txtBoxL.Text, "Box length")
    End Sub

    Private Sub TxtBoxT_TextChanged(sender As Object, e As EventArgs) Handles txtBoxT.TextChanged
        'Make sure input is numaric
        txtBoxT.Text = CheckInput(txtBoxT.Text, "Box top elevation")
    End Sub

    Private Sub TxtHTick_TextChanged(sender As Object, e As EventArgs) Handles txtHTick.TextChanged
        'Make sure input is numaric
        'txtHTick.Text = CheckInput(txtHTick.Text, "Horizontal tick spacing (major)")
        Dim vExagg As Integer = CInt(txtVerticalExagg.Text)

        If chkSquareGrid.Checked = True Then
            If vExagg <> Nothing Then
                If vExagg = 0 Then
                    txtVTick.Text = txtHTick.Text
                Else
                    txtVTick.Text = txtHTick.Text / vExagg
                End If
            End If
        End If
    End Sub

    Private Sub TxtHTickMinor_TextChanged(sender As Object, e As EventArgs) Handles txtHTickMinor.TextChanged
        'Make sure input is numaric
        'txtHTickMinor.Text = CheckInput(txtHTickMinor.Text, "Horizontal tick spacing (minor)")

        Dim vExagg As Integer = CInt(txtVerticalExagg.Text)
        If chkMinorTick.Checked = True And chkSquareGrid.Checked = True Then
            If vExagg <> Nothing Then
                If vExagg = 0 Then
                    txtVTickMinor.Text = txtHTickMinor.Text
                Else
                    txtVTickMinor.Text = txtHTickMinor.Text / vExagg
                End If

            End If
            'Else
            '    txtVTickMinor.Text = ""
        End If
    End Sub

    Private Sub TxtVTick_TextChanged(sender As Object, e As EventArgs) Handles txtVTick.TextChanged
        'Make sure input is numaric
        txtVTick.Text = CheckInput(txtVTick.Text, "Vertical tick spacing (major)")
    End Sub

    Private Sub TxtVTickMinor_TextChanged(sender As Object, e As EventArgs) Handles txtVTickMinor.TextChanged
        'Make sure input is numaric
        txtVTickMinor.Text = CheckInput(txtVTickMinor.Text, "Vertical tick spacing (minor)")
    End Sub

    Private Function CheckAllFormOptions() As Boolean
        ' Check user inputs on the form to determine if we are ready to run
        On Error GoTo eh

        Dim bPage1 As Boolean = True
        Dim bPage2 As Boolean = True
        Dim bPage3 As Boolean = True
        Dim bPage4 As Boolean = True

        Dim sPage1Errors As String = ""
        Dim sPage2Errors As String = ""
        Dim sPage3Errors As String = ""
        Dim sPage4Errors As String = ""

        'Prepare the error message lead-in text, just in case it is needed.
        m_sUserInputErrors = "You have not provided enough information to continue." & Chr(13) _
            & "Please correct the following profile options and try again."

        'check page1 options

        If IsNumeric(txtEndX.Text) = False Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: Endpoint x-coordinate must be numeric"
        End If

        If IsNumeric(txtEndY.Text) = False Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: Endpoint y-coordinate must be numeric"
        End If

        If IsNumeric(txtStartX.Text) = False Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: startpoint x-coordinate must be numeric"
        End If

        If IsNumeric(txtStartY.Text) = False Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: startpoint y-coordinate must be numeric"
        End If

        If CDbl(txtEndX.Text) <= CDbl(txtStartX.Text) Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: profile must be west-to-east and not vertical"
        End If

        If IsNumeric(txtPrecisionMeasure.Text) = False Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: precision distance must be numeric"
        End If

        If IsNumeric(txtPrecisionParts.Text) = False Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: precision parts must be numeric"
        End If

        If IsNumeric(txtVerticalExagg.Text) = False Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: vertical exaggeration must be numeric"
        End If

        If CDbl(txtVerticalExagg.Text) < 1 Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: vertical exaggeration must >= 1"
        End If

        If txtProfileName.Text = "" Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: profile name cannot be blank"
        End If

        If txtOutputPath.Text = "" Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: directory for output files cannot be blank"
        End If

        'check to make sure the directory for output files is valid
        If Directory.Exists(txtOutputPath.Text) <> True Then
            'something is wrong
            bPage1 = False
            sPage1Errors = sPage1Errors & Chr(13) _
                & "- Options: directory for output files is not valid"
        End If

        For Each file In Directory.GetFiles(txtOutputPath.Text)
            If file.Contains(txtProfileName.Text) Then
                bPage1 = False
                sPage1Errors = sPage1Errors & Chr(13) _
                    & "- Options: profile name already exists in directory"
                Exit For
            End If
        Next

        'check page2 options

        If optProfileDelineationOn.Checked = True Then

            If lstSurfaceUnitLayer.SelectedIndex = -1 Then
                'something is wrong
                bPage2 = False
                sPage2Errors = sPage2Errors & Chr(13) _
                    & "- Layers and Paths: a surface polygon layer must be selected"
            End If

            If lstSurfaceUnitField.SelectedIndex = -1 Then
                'something is wrong
                bPage2 = False
                sPage2Errors = sPage2Errors & Chr(13) _
                    & "- Layers and Paths: a surface units field must be selected"
            End If

            If lstElevationLayer.SelectedIndex = -1 Then
                'something is wrong
                bPage2 = False
                sPage2Errors = sPage2Errors & Chr(13) _
                    & "- Layers and Paths: an elevation layer must be selected"
            End If

        Else

            If lstElevationLayer.SelectedIndex = -1 Then
                'something is wrong
                bPage2 = False
                sPage2Errors = sPage2Errors & Chr(13) _
                    & "- Layers and Paths: an elevation layer must be selected"
            End If

        End If

        If optElevationFt.Checked = True Or optElevationM.Checked = True Then
            'this option group is okay
        Else
            'something is wrong
            bPage2 = False
            sPage2Errors = sPage2Errors & Chr(13) _
                & "- Layers and Paths: elevation units must be selected for the elevation grid"
        End If

        ' check page3 options

        If lstWellLayer.SelectedIndex = -1 Then
            'something is wrong
            bPage3 = False
            sPage3Errors = sPage3Errors & Chr(13) _
                & "- Wells: a wells layer must be selected"
        End If

        If lstWellsPointID.SelectedIndex = -1 Then
            'something is wrong
            bPage3 = False
            sPage3Errors = sPage3Errors & Chr(13) _
                & "- Wells: a point ID field must be selected for the wells layer"
        End If

        If lstTables.SelectedIndex = -1 Then
            'something is wrong
            bPage3 = False
            sPage3Errors = sPage3Errors & Chr(13) _
                & "- Wells: a subsurface data table must be selected"
        End If

        If lstTablePointID.SelectedIndex = -1 Then
            'something is wrong
            bPage3 = False
            sPage3Errors = sPage3Errors & Chr(13) _
                & "- Wells: a point ID field must be selected for the subsurface data table"
        End If

        If lstTableUnit.SelectedIndex = -1 Then
            'something is wrong
            bPage3 = False
            sPage3Errors = sPage3Errors & Chr(13) _
                & "- Wells: a subsurface units field must be selected for the subsurface data table"
        End If

        If lstTableTop.SelectedIndex = -1 Then
            'something is wrong
            bPage3 = False
            sPage3Errors = sPage3Errors & Chr(13) _
                & "- Wells: a layer top depth field must be selected for the subsurface data table"
        End If

        If lstTableBottom.SelectedIndex = -1 Then
            'something is wrong
            bPage3 = False
            sPage3Errors = sPage3Errors & Chr(13) _
                & "- Wells: a layer bottom depth field must be selected for the subsurface data table"
        End If

        If chkGetElevFromDEM.Checked = False Then
            If lstWellsElev.SelectedIndex = -1 Then
                'something is wrong
                bPage3 = False
                sPage3Errors = sPage3Errors & Chr(13) _
                    & "- Wells: a well elevation field must be selected for the wells layer"
            End If
        End If

        If optPolys.Checked = True Then
            If IsNumeric(txtPolyW.Text) = False Then
                'something is wrong
                bPage3 = False
                sPage3Errors = sPage3Errors & Chr(13) _
                    & "- Wells: the polygon width for wells must be numeric"
            End If
        End If

        If optWellUnitsFt.Checked = True Or optWellUnitsM.Checked = True Then
            'this option group is okay
        Else
            'something is wrong
            bPage3 = False
            sPage3Errors = sPage3Errors & Chr(13) _
                & "- Wells: subsurface depth units must be selected for the subsurface data table"
        End If

        'check page4 options

        If IsNumeric(txtBoxT.Text) = False Then
            'something is wrong
            bPage4 = False
            sPage4Errors = sPage4Errors & Chr(13) _
                & "- Box: the box top elevation must be numeric"
        End If

        If IsNumeric(txtBoxB.Text) = False Then
            'something is wrong
            bPage4 = False
            sPage4Errors = sPage4Errors & Chr(13) _
                & "- Box: the box bottom elevation must be numeric"
        End If

        If IsNumeric(txtHTick.Text) = False Then
            'something is wrong
            bPage4 = False
            sPage4Errors = sPage4Errors & Chr(13) _
                & "- Box: the major tick horizontal spacing must be numeric"
        End If

        If txtBoxL.Enabled = True Then
            If IsNumeric(txtBoxL.Text) = False Then
                'something is wrong
                bPage4 = False
                sPage4Errors = sPage4Errors & Chr(13) _
                    & "- Box: the box length must be numeric"
            End If
        End If

        If txtVTick.Enabled = True Then
            If IsNumeric(txtVTick.Text) = False Then
                'something is wrong
                bPage4 = False
                sPage4Errors = sPage4Errors & Chr(13) _
                    & "- Box: the major tick vertical spacing must be numeric"
            End If
        End If

        If txtVTickMinor.Enabled = True Then
            If IsNumeric(txtVTickMinor.Text) = False Then
                'something is wrong
                bPage4 = False
                sPage4Errors = sPage4Errors & Chr(13) _
                    & "- Box: the minor tick vertical spacing must be numeric"
            End If
        End If

        If txtHTickMinor.Enabled = True Then
            If IsNumeric(txtHTickMinor.Text) = False Then
                'something is wrong
                bPage4 = False
                sPage4Errors = sPage4Errors & Chr(13) _
                    & "- Box: the minor tick horizontal spacing must be numeric"
            End If
        End If

        If CDbl(txtBoxT.Text) <= CDbl(txtBoxB.Text) Then
            'something is wrong
            bPage4 = False
            sPage4Errors = sPage4Errors & Chr(13) _
                & "- Box: the box top elevation must be above the box bottom elevation"
        End If

        'return results

        CheckAllFormOptions = False

        'check status of first 2 pages
        If bPage1 = True And bPage2 = True Then
            'doing okay so far
            CheckAllFormOptions = True
        Else
            CheckAllFormOptions = False
        End If
        m_sUserInputErrors = m_sUserInputErrors & sPage1Errors & sPage2Errors

        'check status of 3rd page, if needed
        If Page3.Visible = True Then
            If bPage3 = True And CheckAllFormOptions = True Then
                'still doing okay
                CheckAllFormOptions = True
            Else
                CheckAllFormOptions = False
            End If
            m_sUserInputErrors = m_sUserInputErrors & sPage3Errors
        End If

        'check status of 4th page, if needed
        If Page4.Visible = True Then
            If bPage4 = True And CheckAllFormOptions = True Then
                'still doing okay
                CheckAllFormOptions = True
            Else
                CheckAllFormOptions = False
            End If
            m_sUserInputErrors = m_sUserInputErrors & sPage4Errors
        End If

        Exit Function

eh:
        MsgBox("An error occurred while trying to verify user inputs." & Chr(13) _
            & "You need to troubleshoot the CheckAllFormOptions function.", vbCritical, "Error")
        'Me.Close()
        CloseForm()
    End Function

    Private Function SelectFile(vntPath As String) As String
        ' This opens a dialog box to select a directory.
        'Folder Browser Dialog 

        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog With {
            .Description = "Select a Workspace"
        }
        Dim result As DialogResult = Me.FolderBrowserDialog1.ShowDialog()
        If result = DialogResult.OK Then
            vntPath = Me.FolderBrowserDialog1.SelectedPath
        End If

        SelectFile = vntPath
    End Function

    Private Sub SaveUserSettings()
        'save user settings

        On Error Resume Next 'prevents problems when null values are encountered

        'page1 options
        SaveSetting("ProfileTool", "Settings", "optPrecisionParts", CStr(optPrecisionParts.Checked))
        SaveSetting("ProfileTool", "Settings", "txtPrecisionParts", txtPrecisionParts.Text)
        SaveSetting("ProfileTool", "Settings", "optPrecisionMeasure", CStr(optPrecisionMeasure.Checked))
        SaveSetting("ProfileTool", "Settings", "txtPrecisionMeasure", txtPrecisionMeasure.Text)

        SaveSetting("ProfileTool", "Settings", "chkVerticalExagg", CStr(chkVerticalExagg.Checked))
        SaveSetting("ProfileTool", "Settings", "txtVerticalExagg", txtVerticalExagg.Text)
        SaveSetting("ProfileTool", "Settings", "chkBuildProfileBox", CStr(chkBuildProfileBox.Checked))
        SaveSetting("ProfileTool", "Settings", "chkAddWells", CStr(chkAddWells.Checked))
        SaveSetting("ProfileTool", "Settings", "smoothProfile", CStr(smoothProfile.Checked))
        SaveSetting("ProfileTool", "Settings", "chkSaveSettings", CStr(chkSaveSettings.Checked))
        SaveSetting("ProfileTool", "Settings", "txtProfileName", txtProfileName.Text)
        SaveSetting("ProfileTool", "Settings", "txtOutputPath", txtOutputPath.Text)

        'Page2 options
        SaveSetting("ProfileTool", "Settings", "optProfileDelineationOn", CStr(optProfileDelineationOn.Checked))
        SaveSetting("ProfileTool", "Settings", "optProfileDelineationOff", CStr(optProfileDelineationOff.Checked))
        SaveSetting("ProfileTool", "Settings", "lstSurfaceUnitLayer", lstSurfaceUnitLayer.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstSurfaceUnitField", lstSurfaceUnitField.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstElevationLayer", lstElevationLayer.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "optElevationM", CStr(optElevationM.Checked))
        SaveSetting("ProfileTool", "Settings", "optElevationFt", CStr(optElevationFt.Checked))

        'page3 options
        SaveSetting("ProfileTool", "Settings", "lstWellLayer", lstWellLayer.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstWellsPointID", lstWellsPointID.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstWellsElev", lstWellsElev.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstTables", lstTables.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstTablePointID", lstTablePointID.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstTableUnit", lstTableUnit.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstTableTop", lstTableTop.SelectedItem)
        SaveSetting("ProfileTool", "Settings", "lstTableBottom", lstTableBottom.SelectedItem)

        SaveSetting("ProfileTool", "Settings", "chkGetElevFromDEM", CStr(chkGetElevFromDEM.Checked))
        SaveSetting("ProfileTool", "Settings", "optLines", CStr(optLines.Checked))
        SaveSetting("ProfileTool", "Settings", "optPolys", CStr(optPolys.Checked))
        SaveSetting("ProfileTool", "Settings", "txtPolyW", txtPolyW.Text)
        SaveSetting("ProfileTool", "Settings", "optWellUnitsM", CStr(optWellUnitsM.Checked))
        SaveSetting("ProfileTool", "Settings", "optWellUnitsFt", CStr(optWellUnitsFt.Checked))

        'page4 options
        SaveSetting("ProfileTool", "Settings", "txtBoxT", txtBoxT.Text)
        SaveSetting("ProfileTool", "Settings", "txtBoxB", txtBoxB.Text)
        SaveSetting("ProfileTool", "Settings", "txtBoxL", txtBoxL.Text)
        SaveSetting("ProfileTool", "Settings", "txtVTick", txtVTick.Text)
        SaveSetting("ProfileTool", "Settings", "txtHTick", txtHTick.Text)
        SaveSetting("ProfileTool", "Settings", "txtVTickMinor", txtVTickMinor.Text)
        SaveSetting("ProfileTool", "Settings", "txtHTickMinor", txtHTickMinor.Text)
        SaveSetting("ProfileTool", "Settings", "chkUseProfileL", CStr(chkUseProfileL.Checked))
        SaveSetting("ProfileTool", "Settings", "chkMinorTick", CStr(chkMinorTick.Checked))
        SaveSetting("ProfileTool", "Settings", "chkBuildGrid", CStr(chkBuildGrid.Checked))
        SaveSetting("ProfileTool", "Settings", "chkSquareGrid", CStr(chkSquareGrid.Checked))
    End Sub

    Private Sub GetUserSettings()
        'get user settings and update the form

        On Error Resume Next 'prevents problems when null values are encountered

        'page1 options
        optPrecisionParts.Checked = GetSetting("ProfileTool", "Settings", "optPrecisionParts")
        optPrecisionMeasure.Checked = GetSetting("ProfileTool", "Settings", "optPrecisionMeasure")

        If optPrecisionParts.Checked = True Then
            txtPrecisionParts.Text = GetSetting("ProfileTool", "Settings", "txtPrecisionParts")
        Else
            txtPrecisionMeasure.Text = GetSetting("ProfileTool", "Settings", "txtPrecisionMeasure")
        End If

        chkVerticalExagg.Checked = GetSetting("ProfileTool", "Settings", "chkVerticalExagg")
        txtVerticalExagg.Text = GetSetting("ProfileTool", "Settings", "txtVerticalExagg")
        chkBuildProfileBox.Checked = GetSetting("ProfileTool", "Settings", "chkBuildProfileBox")
        chkAddWells.Checked = GetSetting("ProfileTool", "Settings", "chkAddWells")
        smoothProfile.Checked = GetSetting("ProfileTool", "Settings", "smoothProfile")
        chkSaveSettings.Checked = GetSetting("ProfileTool", "Settings", "chkSaveSettings")
        txtProfileName.Text = GetSetting("ProfileTool", "Settings", "txtProfileName")
        txtOutputPath.Text = GetSetting("ProfileTool", "Settings", "txtOutputPath")

        'Page2 options
        optProfileDelineationOn.Checked = GetSetting("ProfileTool", "Settings", "optProfileDelineationOn")
        optProfileDelineationOff.Checked = GetSetting("ProfileTool", "Settings", "optProfileDelineationOff")
        lstSurfaceUnitLayer.SelectedItem = GetSetting("ProfileTool", "Settings", "lstSurfaceUnitLayer")
        lstSurfaceUnitField.SelectedItem = GetSetting("ProfileTool", "Settings", "lstSurfaceUnitField")
        lstElevationLayer.SelectedItem = GetSetting("ProfileTool", "Settings", "lstElevationLayer")
        optElevationM.Checked = GetSetting("ProfileTool", "Settings", "optElevationM")
        optElevationFt.Checked = GetSetting("ProfileTool", "Settings", "optElevationFt")

        'page3 options
        lstWellLayer.SelectedItem = GetSetting("ProfileTool", "Settings", "lstWellLayer")
        lstWellsPointID.SelectedItem = GetSetting("ProfileTool", "Settings", "lstWellsPointID")
        lstWellsElev.SelectedItem = GetSetting("ProfileTool", "Settings", "lstWellsElev")
        lstTables.SelectedItem = GetSetting("ProfileTool", "Settings", "lstTables")
        lstTablePointID.SelectedItem = GetSetting("ProfileTool", "Settings", "lstTablePointID")
        lstTableUnit.SelectedItem = GetSetting("ProfileTool", "Settings", "lstTableUnit")
        lstTableTop.SelectedItem = GetSetting("ProfileTool", "Settings", "lstTableTop")
        lstTableBottom.SelectedItem = GetSetting("ProfileTool", "Settings", "lstTableBottom")
        chkGetElevFromDEM.Checked = GetSetting("ProfileTool", "Settings", "chkGetElevFromDEM")
        optLines.Checked = GetSetting("ProfileTool", "Settings", "optLines")
        optPolys.Checked = GetSetting("ProfileTool", "Settings", "optPolys")
        txtPolyW.Text = GetSetting("ProfileTool", "Settings", "txtPolyW")
        optWellUnitsM.Checked = GetSetting("ProfileTool", "Settings", "optWellUnitsM")
        optWellUnitsFt.Checked = GetSetting("ProfileTool", "Settings", "optWellUnitsFt")

        'page4 options
        txtBoxT.Text = GetSetting("ProfileTool", "Settings", "txtBoxT")
        txtBoxB.Text = GetSetting("ProfileTool", "Settings", "txtBoxB")
        txtBoxL.Text = GetSetting("ProfileTool", "Settings", "txtBoxL")
        txtVTick.Text = GetSetting("ProfileTool", "Settings", "txtVTick")
        txtHTick.Text = GetSetting("ProfileTool", "Settings", "txtHTick")
        txtVTickMinor.Text = GetSetting("ProfileTool", "Settings", "txtVTickMinor")
        txtHTickMinor.Text = GetSetting("ProfileTool", "Settings", "txtHTickMinor")
        chkUseProfileL.Checked = GetSetting("ProfileTool", "Settings", "chkUseProfileL")
        chkMinorTick.Checked = GetSetting("ProfileTool", "Settings", "chkMinorTick")
        chkBuildGrid.Checked = GetSetting("ProfileTool", "Settings", "chkBuildGrid")
        chkSquareGrid.Checked = GetSetting("ProfileTool", "Settings", "chkSquareGrid")
    End Sub

    Private Sub GetTOCForSelectedLayers()
        ' finds the TOC number for the selected geology and DEM layers

        Dim i As Integer
        Dim iUbound As Integer

        iUbound = UBound(m_asTocInfo)

        For i = 0 To iUbound
            'get surface unit layer position
            If m_asTocInfo(i) = lstSurfaceUnitLayer.SelectedItem Then
                m_iTOCSurface = i
            End If

            'get elevation grid position
            If m_asTocInfo(i) = lstElevationLayer.SelectedItem Then
                m_iTOCElevation = i
            End If

            'get well layer position
            If m_asTocInfo(i) = lstWellLayer.SelectedItem Then
                m_iTOCWells = i
            End If
        Next
    End Sub

    Private Sub UserForm_Terminate()
        Me.Close()
    End Sub

    Private Sub CmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        If cps IsNot Nothing Then
            Debug.Print("Canceling the task...")
            cps.CancellationTokenSource.Cancel()
        End If

    End Sub

    Private Sub CmdStart_Click(sender As Object, e As EventArgs) Handles cmdStart.Click
        If CheckAllFormOptions() = True Then
            'save settings, if user chose to do so
            If chkSaveSettings.Checked = True Then
                Call SaveUserSettings()
            Else
                SaveSetting("ProfileTool", "Settings", "chkSaveSettings", CStr(chkSaveSettings.Checked))
            End If

            'get TOC positions for the selected layers
            Call GetTOCForSelectedLayers()

            'setup the user form
            MultiPage1.SelectedIndex = 0
            MultiPage1.Enabled = False
            lblPrgBr.Visible = True
            prgBr.Visible = True
            cmdStart.Enabled = False

            'start building the profile
            Call GpMain(m_iTOCSurface, m_iTOCElevation, m_dUserUnitsToMapUnits, m_dMapUnitsToUserUnits,
            m_dElevationLayerUnitsToMapUnits, m_dSubsurfaceTableUnitsToMapUnits, Me)
        Else
            MsgBox(m_sUserInputErrors, vbOKOnly, "Incomplete Options")
        End If
    End Sub

    Private Sub CmdChangeTempPath_Click(sender As Object, e As EventArgs) Handles cmdChangeTempPath.Click
        Dim varPath As String = ""
        varPath = SelectFile(varPath)
        PopulateTextBoxes(txtOutputPath, varPath)
        'txtOutputPath.Text = varPath
    End Sub

    Private Sub OptPrecisionParts_CheckedChanged(sender As Object, e As EventArgs) Handles optPrecisionParts.CheckedChanged
        ' User chooses to specify the number of points to use in the profile
        txtPrecisionParts.Enabled = True
        txtPrecisionMeasure.Enabled = False
        Call CalculateDistBetweenPoints()
    End Sub

    Private Sub OptElevationM_CheckedChanged(sender As Object, e As EventArgs) Handles optElevationM.CheckedChanged
        Select Case optElevationM.Checked
            Case True
                optElevationFt.Checked = False
                lblPrecisionUnit.Text = "meters"
            Case False
                optElevationFt.Checked = True
                lblPrecisionUnit.Text = "feet"
        End Select
    End Sub


    Private Sub OptBoxUnitsFT_CheckedChanged(sender As Object, e As EventArgs) Handles optBoxUnitsFT.CheckedChanged
        Dim label As String = ""
        Select Case optBoxUnitsFT.Checked
            Case True
                optBoxUnitsM.Checked = False
                label = "feet"
            Case False
                optBoxUnitsM.Checked = True
                label = "meters"
        End Select

        CalculateProfileLength()

        Label1.Text = label
        Label2.Text = label
        Label3.Text = label
        Label5.Text = label
        Label6.Text = label
        Label7.Text = label
        Label9.Text = label
    End Sub

    Private Sub InitComponent() 'InitializeComponent()
        Me.SuspendLayout()
        '
        'frmProfileTool
        '
        Me.ClientSize = New System.Drawing.Size(592, 553)
        Me.Name = "frmProfileTool"
        Me.ResumeLayout(False)

    End Sub
End Class