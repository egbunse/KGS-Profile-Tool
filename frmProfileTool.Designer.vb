<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProfileTool
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MultiPage1 = New System.Windows.Forms.TabControl()
        Me.Page1 = New System.Windows.Forms.TabPage()
        Me.cmdChangeTempPath = New System.Windows.Forms.Button()
        Me.txtOutputPath = New System.Windows.Forms.TextBox()
        Me.txtProfileName = New System.Windows.Forms.TextBox()
        Me.txtEndY = New System.Windows.Forms.TextBox()
        Me.txtStartY = New System.Windows.Forms.TextBox()
        Me.txtEndX = New System.Windows.Forms.TextBox()
        Me.txtStartX = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtVerticalExagg = New System.Windows.Forms.TextBox()
        Me.lblParts1 = New System.Windows.Forms.Label()
        Me.txtPrecisionParts = New System.Windows.Forms.TextBox()
        Me.chkSaveSettings = New System.Windows.Forms.CheckBox()
        'Me.chkGraphicLine = New System.Windows.Forms.CheckBox()
        Me.smoothProfile = New System.Windows.Forms.CheckBox()
        Me.chkAddWells = New System.Windows.Forms.CheckBox()
        Me.chkBuildProfileBox = New System.Windows.Forms.CheckBox()
        Me.chkVerticalExagg = New System.Windows.Forms.CheckBox()
        Me.optPrecisionMeasure = New System.Windows.Forms.RadioButton()
        Me.optPrecisionParts = New System.Windows.Forms.RadioButton()
        Me.lblOutputPath = New System.Windows.Forms.Label()
        Me.lblProfileName = New System.Windows.Forms.Label()
        Me.lblYcoord = New System.Windows.Forms.Label()
        Me.lblXcoord = New System.Windows.Forms.Label()
        Me.lblOptions = New System.Windows.Forms.Label()
        Me.lblPrecision = New System.Windows.Forms.Label()
        Me.lblCoordEnd = New System.Windows.Forms.Label()
        Me.lblCoordStart = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblPrecisionUnit = New System.Windows.Forms.Label()
        Me.txtPrecisionMeasure = New System.Windows.Forms.TextBox()
        Me.Page2 = New System.Windows.Forms.TabPage()
        Me.lstElevationLayer = New System.Windows.Forms.ListBox()
        Me.lstSurfaceUnitField = New System.Windows.Forms.ListBox()
        Me.lstSurfaceUnitLayer = New System.Windows.Forms.ListBox()
        Me.lblDEMlist = New System.Windows.Forms.Label()
        Me.lblSurfaceUnitField = New System.Windows.Forms.Label()
        Me.lblSurfaceUnitLayer = New System.Windows.Forms.Label()
        Me.fraElevationUnits = New System.Windows.Forms.Panel()
        Me.optElevationM = New System.Windows.Forms.RadioButton()
        Me.optElevationFt = New System.Windows.Forms.RadioButton()
        Me.lblElevationUnits = New System.Windows.Forms.Label()
        Me.fraProfileDelineationOptions = New System.Windows.Forms.Panel()
        Me.optProfileDelineationOff = New System.Windows.Forms.RadioButton()
        Me.optProfileDelineationOn = New System.Windows.Forms.RadioButton()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Page4 = New System.Windows.Forms.TabPage()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.optBoxUnitsM = New System.Windows.Forms.RadioButton()
        Me.optBoxUnitsFT = New System.Windows.Forms.RadioButton()
        Me.lblBoxUnits = New System.Windows.Forms.Label()
        Me.txtHTickMinor = New System.Windows.Forms.TextBox()
        Me.txtVTickMinor = New System.Windows.Forms.TextBox()
        Me.txtHTick = New System.Windows.Forms.TextBox()
        Me.txtVTick = New System.Windows.Forms.TextBox()
        Me.chkUseProfileL = New System.Windows.Forms.CheckBox()
        Me.txtBoxL = New System.Windows.Forms.TextBox()
        Me.txtBoxB = New System.Windows.Forms.TextBox()
        Me.txtBoxT = New System.Windows.Forms.TextBox()
        Me.chkSquareGrid = New System.Windows.Forms.CheckBox()
        Me.chkBuildGrid = New System.Windows.Forms.CheckBox()
        Me.lblMinorVertTick = New System.Windows.Forms.Label()
        Me.lblMinorHorizTick = New System.Windows.Forms.Label()
        Me.chkMinorTick = New System.Windows.Forms.CheckBox()
        Me.lblHorizTick = New System.Windows.Forms.Label()
        Me.lblVerticalTick = New System.Windows.Forms.Label()
        Me.lblBoxLength = New System.Windows.Forms.Label()
        Me.lblBoxBottom = New System.Windows.Forms.Label()
        Me.lblBoxTop = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Page3 = New System.Windows.Forms.TabPage()
        Me.lstTableUnit = New System.Windows.Forms.ListBox()
        Me.lstWellsElev = New System.Windows.Forms.ListBox()
        Me.lstTableBottom = New System.Windows.Forms.ListBox()
        Me.lstTablePointID = New System.Windows.Forms.ListBox()
        Me.lstWellsPointID = New System.Windows.Forms.ListBox()
        Me.lstTableTop = New System.Windows.Forms.ListBox()
        Me.lstTables = New System.Windows.Forms.ListBox()
        Me.lstWellLayer = New System.Windows.Forms.ListBox()
        Me.chkGetElevFromDEM = New System.Windows.Forms.CheckBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.optWellUnitsM = New System.Windows.Forms.RadioButton()
        Me.optWellUnitsFt = New System.Windows.Forms.RadioButton()
        Me.Frame2 = New System.Windows.Forms.Panel()
        Me.txtPolyW = New System.Windows.Forms.TextBox()
        Me.lblPolyW = New System.Windows.Forms.Label()
        Me.optPolys = New System.Windows.Forms.RadioButton()
        Me.optLines = New System.Windows.Forms.RadioButton()
        Me.lblfraWellUnits = New System.Windows.Forms.Label()
        Me.lblFrame2 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lblWellsElev = New System.Windows.Forms.Label()
        Me.lblwellBot = New System.Windows.Forms.Label()
        Me.lblPlgnID = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.lblWellTop = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.lblPrgBr = New System.Windows.Forms.Label()
        Me.prgBr = New System.Windows.Forms.ProgressBar()
        Me.MultiPage1.SuspendLayout()
        Me.Page1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Page2.SuspendLayout()
        Me.fraElevationUnits.SuspendLayout()
        Me.fraProfileDelineationOptions.SuspendLayout()
        Me.Page4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Page3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.SuspendLayout()
        '
        'MultiPage1
        '
        Me.MultiPage1.Controls.Add(Me.Page1)
        Me.MultiPage1.Controls.Add(Me.Page2)
        Me.MultiPage1.Location = New System.Drawing.Point(3, 4)
        Me.MultiPage1.Margin = New System.Windows.Forms.Padding(4)
        Me.MultiPage1.Name = "MultiPage1"
        Me.MultiPage1.SelectedIndex = 0
        Me.MultiPage1.Size = New System.Drawing.Size(863, 733)
        Me.MultiPage1.TabIndex = 0
        '
        'Page1
        '
        Me.Page1.Controls.Add(Me.cmdChangeTempPath)
        Me.Page1.Controls.Add(Me.txtOutputPath)
        Me.Page1.Controls.Add(Me.txtProfileName)
        Me.Page1.Controls.Add(Me.txtEndY)
        Me.Page1.Controls.Add(Me.txtStartY)
        Me.Page1.Controls.Add(Me.txtEndX)
        Me.Page1.Controls.Add(Me.txtStartX)
        Me.Page1.Controls.Add(Me.Label4)
        Me.Page1.Controls.Add(Me.txtVerticalExagg)
        Me.Page1.Controls.Add(Me.lblParts1)
        Me.Page1.Controls.Add(Me.txtPrecisionParts)
        Me.Page1.Controls.Add(Me.chkSaveSettings)
        'Me.Page1.Controls.Add(Me.chkGraphicLine)
        Me.Page1.Controls.Add(Me.smoothProfile)
        Me.Page1.Controls.Add(Me.chkAddWells)
        Me.Page1.Controls.Add(Me.chkBuildProfileBox)
        Me.Page1.Controls.Add(Me.chkVerticalExagg)
        Me.Page1.Controls.Add(Me.optPrecisionMeasure)
        Me.Page1.Controls.Add(Me.optPrecisionParts)
        Me.Page1.Controls.Add(Me.lblOutputPath)
        Me.Page1.Controls.Add(Me.lblProfileName)
        Me.Page1.Controls.Add(Me.lblYcoord)
        Me.Page1.Controls.Add(Me.lblXcoord)
        Me.Page1.Controls.Add(Me.lblOptions)
        Me.Page1.Controls.Add(Me.lblPrecision)
        Me.Page1.Controls.Add(Me.lblCoordEnd)
        Me.Page1.Controls.Add(Me.lblCoordStart)
        Me.Page1.Controls.Add(Me.Panel1)
        Me.Page1.Controls.Add(Me.smoothProfile)
        Me.Page1.Location = New System.Drawing.Point(4, 25)
        Me.Page1.Margin = New System.Windows.Forms.Padding(4)
        Me.Page1.Name = "Page1"
        Me.Page1.Padding = New System.Windows.Forms.Padding(4)
        Me.Page1.Size = New System.Drawing.Size(855, 704)
        Me.Page1.TabIndex = 0
        Me.Page1.Text = "Profile Options"
        Me.Page1.UseVisualStyleBackColor = True
        '
        'cmdChangeTempPath
        '
        Me.cmdChangeTempPath.Location = New System.Drawing.Point(725, 534)
        Me.cmdChangeTempPath.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdChangeTempPath.Name = "cmdChangeTempPath"
        Me.cmdChangeTempPath.Size = New System.Drawing.Size(100, 28)
        Me.cmdChangeTempPath.TabIndex = 27
        Me.cmdChangeTempPath.Text = "Choose..."
        Me.cmdChangeTempPath.UseVisualStyleBackColor = True
        '
        'txtOutputPath
        '
        Me.txtOutputPath.Location = New System.Drawing.Point(8, 537)
        Me.txtOutputPath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtOutputPath.Name = "txtOutputPath"
        Me.txtOutputPath.Size = New System.Drawing.Size(692, 22)
        Me.txtOutputPath.TabIndex = 26
        '
        'txtProfileName
        '
        Me.txtProfileName.Location = New System.Drawing.Point(12, 474)
        Me.txtProfileName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtProfileName.Name = "txtProfileName"
        Me.txtProfileName.Size = New System.Drawing.Size(688, 22)
        Me.txtProfileName.TabIndex = 25
        Me.txtProfileName.Text = "Profile"
        '
        'txtEndY
        '
        Me.txtEndY.Location = New System.Drawing.Point(516, 123)
        Me.txtEndY.Margin = New System.Windows.Forms.Padding(4)
        Me.txtEndY.Name = "txtEndY"
        Me.txtEndY.Size = New System.Drawing.Size(132, 22)
        Me.txtEndY.TabIndex = 24
        Me.txtEndY.Text = "Y (end)"
        Me.txtEndY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtStartY
        '
        Me.txtStartY.Location = New System.Drawing.Point(224, 123)
        Me.txtStartY.Margin = New System.Windows.Forms.Padding(4)
        Me.txtStartY.Name = "txtStartY"
        Me.txtStartY.Size = New System.Drawing.Size(132, 22)
        Me.txtStartY.TabIndex = 23
        Me.txtStartY.Text = "Y (start)"
        Me.txtStartY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtEndX
        '
        Me.txtEndX.Location = New System.Drawing.Point(516, 82)
        Me.txtEndX.Margin = New System.Windows.Forms.Padding(4)
        Me.txtEndX.Name = "txtEndX"
        Me.txtEndX.Size = New System.Drawing.Size(132, 22)
        Me.txtEndX.TabIndex = 22
        Me.txtEndX.Text = "X (end)"
        Me.txtEndX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtStartX
        '
        Me.txtStartX.Location = New System.Drawing.Point(224, 82)
        Me.txtStartX.Margin = New System.Windows.Forms.Padding(4)
        Me.txtStartX.Name = "txtStartX"
        Me.txtStartX.Size = New System.Drawing.Size(132, 22)
        Me.txtStartX.TabIndex = 21
        Me.txtStartX.Text = "X (start)"
        Me.txtStartX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(444, 271)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(17, 17)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "X"
        '
        'txtVerticalExagg
        '
        Me.txtVerticalExagg.Location = New System.Drawing.Point(405, 266)
        Me.txtVerticalExagg.Margin = New System.Windows.Forms.Padding(4)
        Me.txtVerticalExagg.Name = "txtVerticalExagg"
        Me.txtVerticalExagg.Size = New System.Drawing.Size(37, 22)
        Me.txtVerticalExagg.TabIndex = 19
        Me.txtVerticalExagg.Text = "1"
        '
        'lblParts1
        '
        Me.lblParts1.AutoSize = True
        Me.lblParts1.Location = New System.Drawing.Point(483, 177)
        Me.lblParts1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblParts1.Name = "lblParts1"
        Me.lblParts1.Size = New System.Drawing.Size(40, 17)
        Me.lblParts1.TabIndex = 16
        Me.lblParts1.Text = "parts"
        '
        'txtPrecisionParts
        '
        Me.txtPrecisionParts.Location = New System.Drawing.Point(409, 174)
        Me.txtPrecisionParts.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrecisionParts.Name = "txtPrecisionParts"
        Me.txtPrecisionParts.Size = New System.Drawing.Size(63, 22)
        Me.txtPrecisionParts.TabIndex = 15
        Me.txtPrecisionParts.Text = "100"
        '
        'chkSaveSettings
        '
        Me.chkSaveSettings.AutoSize = True
        Me.chkSaveSettings.Location = New System.Drawing.Point(235, 383)
        Me.chkSaveSettings.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSaveSettings.Name = "chkSaveSettings"
        Me.chkSaveSettings.Size = New System.Drawing.Size(156, 21)
        Me.chkSaveSettings.TabIndex = 14
        Me.chkSaveSettings.Text = " Remember settings"
        Me.chkSaveSettings.UseVisualStyleBackColor = True
        '
        'chkGraphicLine
        '
        'Me.chkGraphicLine.AutoSize = True
        'Me.chkGraphicLine.Location = New System.Drawing.Point(235, 354)
        'Me.chkGraphicLine.Margin = New System.Windows.Forms.Padding(4)
        'Me.chkGraphicLine.Name = "chkGraphicLine"
        'Me.chkGraphicLine.Size = New System.Drawing.Size(207, 21)
        'Me.chkGraphicLine.TabIndex = 13
        'Me.chkGraphicLine.Text = " Draw graphic for profile line"
        'Me.chkGraphicLine.UseVisualStyleBackColor = True

        '
        'smoothProfile
        '
        Me.smoothProfile.AutoSize = True
        Me.smoothProfile.Location = New System.Drawing.Point(235, 354)
        Me.smoothProfile.Margin = New System.Windows.Forms.Padding(4)
        Me.smoothProfile.Name = "smoothProfile"
        Me.smoothProfile.Size = New System.Drawing.Size(180, 21)
        Me.smoothProfile.TabIndex = 13
        Me.smoothProfile.Text = " Set up Smooth Line Tool for Profile (see Geoprocessing Pane)"
        Me.smoothProfile.UseVisualStyleBackColor = True

        '
        'chkAddWells
        '
        Me.chkAddWells.AutoSize = True
        Me.chkAddWells.Location = New System.Drawing.Point(235, 326)
        Me.chkAddWells.Margin = New System.Windows.Forms.Padding(4)
        Me.chkAddWells.Name = "chkAddWells"
        Me.chkAddWells.Size = New System.Drawing.Size(230, 21)
        Me.chkAddWells.TabIndex = 12
        Me.chkAddWells.Text = " Overlay wells on the profile line"
        Me.chkAddWells.UseVisualStyleBackColor = True
        '
        'chkBuildProfileBox
        '
        Me.chkBuildProfileBox.AutoSize = True
        Me.chkBuildProfileBox.Location = New System.Drawing.Point(235, 298)
        Me.chkBuildProfileBox.Margin = New System.Windows.Forms.Padding(4)
        Me.chkBuildProfileBox.Name = "chkBuildProfileBox"
        Me.chkBuildProfileBox.Size = New System.Drawing.Size(290, 21)
        Me.chkBuildProfileBox.TabIndex = 11
        Me.chkBuildProfileBox.Text = " Build a box or grid around the profile line"
        Me.chkBuildProfileBox.UseVisualStyleBackColor = True
        '
        'chkVerticalExagg
        '
        Me.chkVerticalExagg.AutoSize = True
        Me.chkVerticalExagg.Location = New System.Drawing.Point(235, 270)
        Me.chkVerticalExagg.Margin = New System.Windows.Forms.Padding(4)
        Me.chkVerticalExagg.Name = "chkVerticalExagg"
        Me.chkVerticalExagg.Size = New System.Drawing.Size(164, 21)
        Me.chkVerticalExagg.TabIndex = 10
        Me.chkVerticalExagg.Text = "Vertical Exaggeration"
        Me.chkVerticalExagg.UseVisualStyleBackColor = True
        '
        'optPrecisionMeasure
        '
        Me.optPrecisionMeasure.AutoSize = True
        Me.optPrecisionMeasure.Location = New System.Drawing.Point(235, 206)
        Me.optPrecisionMeasure.Margin = New System.Windows.Forms.Padding(4)
        Me.optPrecisionMeasure.Name = "optPrecisionMeasure"
        Me.optPrecisionMeasure.Size = New System.Drawing.Size(213, 21)
        Me.optPrecisionMeasure.TabIndex = 9
        Me.optPrecisionMeasure.TabStop = True
        Me.optPrecisionMeasure.Text = "Pick an elevation value every"
        Me.optPrecisionMeasure.UseVisualStyleBackColor = True
        '
        'optPrecisionParts
        '
        Me.optPrecisionParts.AutoSize = True
        Me.optPrecisionParts.Location = New System.Drawing.Point(235, 177)
        Me.optPrecisionParts.Margin = New System.Windows.Forms.Padding(4)
        Me.optPrecisionParts.Name = "optPrecisionParts"
        Me.optPrecisionParts.Size = New System.Drawing.Size(162, 21)
        Me.optPrecisionParts.TabIndex = 8
        Me.optPrecisionParts.TabStop = True
        Me.optPrecisionParts.Text = "Divide the profile into"
        Me.optPrecisionParts.UseVisualStyleBackColor = True
        '
        'lblOutputPath
        '
        Me.lblOutputPath.AutoSize = True
        Me.lblOutputPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOutputPath.Location = New System.Drawing.Point(4, 511)
        Me.lblOutputPath.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblOutputPath.Name = "lblOutputPath"
        Me.lblOutputPath.Size = New System.Drawing.Size(222, 20)
        Me.lblOutputPath.TabIndex = 7
        Me.lblOutputPath.Text = "Directory for output files:"
        '
        'lblProfileName
        '
        Me.lblProfileName.AutoSize = True
        Me.lblProfileName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProfileName.Location = New System.Drawing.Point(8, 448)
        Me.lblProfileName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProfileName.Name = "lblProfileName"
        Me.lblProfileName.Size = New System.Drawing.Size(124, 20)
        Me.lblProfileName.TabIndex = 6
        Me.lblProfileName.Text = "Profile Name:"
        '
        'lblYcoord
        '
        Me.lblYcoord.AutoSize = True
        Me.lblYcoord.Location = New System.Drawing.Point(384, 127)
        Me.lblYcoord.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblYcoord.Name = "lblYcoord"
        Me.lblYcoord.Size = New System.Drawing.Size(89, 17)
        Me.lblYcoord.TabIndex = 5
        Me.lblYcoord.Text = "Y-coordinate"
        '
        'lblXcoord
        '
        Me.lblXcoord.AutoSize = True
        Me.lblXcoord.Location = New System.Drawing.Point(384, 86)
        Me.lblXcoord.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblXcoord.Name = "lblXcoord"
        Me.lblXcoord.Size = New System.Drawing.Size(89, 17)
        Me.lblXcoord.TabIndex = 4
        Me.lblXcoord.Text = "X-coordinate"
        '
        'lblOptions
        '
        Me.lblOptions.AutoSize = True
        Me.lblOptions.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOptions.Location = New System.Drawing.Point(89, 268)
        Me.lblOptions.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblOptions.Name = "lblOptions"
        Me.lblOptions.Size = New System.Drawing.Size(80, 20)
        Me.lblOptions.TabIndex = 3
        Me.lblOptions.Text = "Options:"
        '
        'lblPrecision
        '
        Me.lblPrecision.AutoSize = True
        Me.lblPrecision.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrecision.Location = New System.Drawing.Point(89, 177)
        Me.lblPrecision.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPrecision.Name = "lblPrecision"
        Me.lblPrecision.Size = New System.Drawing.Size(94, 20)
        Me.lblPrecision.TabIndex = 2
        Me.lblPrecision.Text = "Precision:"
        '
        'lblCoordEnd
        '
        Me.lblCoordEnd.AutoSize = True
        Me.lblCoordEnd.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCoordEnd.Location = New System.Drawing.Point(527, 34)
        Me.lblCoordEnd.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCoordEnd.Name = "lblCoordEnd"
        Me.lblCoordEnd.Size = New System.Drawing.Size(97, 25)
        Me.lblCoordEnd.TabIndex = 1
        Me.lblCoordEnd.Text = "Endpoint"
        '
        'lblCoordStart
        '
        Me.lblCoordStart.AutoSize = True
        Me.lblCoordStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCoordStart.Location = New System.Drawing.Point(229, 34)
        Me.lblCoordStart.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCoordStart.Name = "lblCoordStart"
        Me.lblCoordStart.Size = New System.Drawing.Size(105, 25)
        Me.lblCoordStart.TabIndex = 0
        Me.lblCoordStart.Text = "Startpoint"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lblPrecisionUnit)
        Me.Panel1.Controls.Add(Me.txtPrecisionMeasure)
        Me.Panel1.Location = New System.Drawing.Point(224, 155)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(575, 89)
        Me.Panel1.TabIndex = 28
        '
        'lblPrecisionUnit
        '
        Me.lblPrecisionUnit.AutoSize = True
        Me.lblPrecisionUnit.Location = New System.Drawing.Point(335, 55)
        Me.lblPrecisionUnit.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPrecisionUnit.Name = "lblPrecisionUnit"
        Me.lblPrecisionUnit.Size = New System.Drawing.Size(32, 17)
        Me.lblPrecisionUnit.TabIndex = 19
        Me.lblPrecisionUnit.Text = "feet"
        '
        'txtPrecisionMeasure
        '
        Me.txtPrecisionMeasure.Location = New System.Drawing.Point(239, 49)
        Me.txtPrecisionMeasure.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrecisionMeasure.Name = "txtPrecisionMeasure"
        Me.txtPrecisionMeasure.Size = New System.Drawing.Size(85, 22)
        Me.txtPrecisionMeasure.TabIndex = 18
        '
        'Page2
        '
        Me.Page2.Controls.Add(Me.lstElevationLayer)
        Me.Page2.Controls.Add(Me.lstSurfaceUnitField)
        Me.Page2.Controls.Add(Me.lstSurfaceUnitLayer)
        Me.Page2.Controls.Add(Me.lblDEMlist)
        Me.Page2.Controls.Add(Me.lblSurfaceUnitField)
        Me.Page2.Controls.Add(Me.lblSurfaceUnitLayer)
        Me.Page2.Controls.Add(Me.fraElevationUnits)
        Me.Page2.Controls.Add(Me.lblElevationUnits)
        Me.Page2.Controls.Add(Me.fraProfileDelineationOptions)
        Me.Page2.Controls.Add(Me.Label25)
        Me.Page2.Location = New System.Drawing.Point(4, 25)
        Me.Page2.Margin = New System.Windows.Forms.Padding(4)
        Me.Page2.Name = "Page2"
        Me.Page2.Padding = New System.Windows.Forms.Padding(4)
        Me.Page2.Size = New System.Drawing.Size(855, 704)
        Me.Page2.TabIndex = 1
        Me.Page2.Text = "Profile Layers and Paths"
        Me.Page2.UseVisualStyleBackColor = True
        '
        'lstElevationLayer
        '
        Me.lstElevationLayer.FormattingEnabled = True
        Me.lstElevationLayer.ItemHeight = 16
        Me.lstElevationLayer.Location = New System.Drawing.Point(572, 188)
        Me.lstElevationLayer.Margin = New System.Windows.Forms.Padding(4)
        Me.lstElevationLayer.Name = "lstElevationLayer"
        Me.lstElevationLayer.Size = New System.Drawing.Size(249, 388)
        Me.lstElevationLayer.TabIndex = 9
        '
        'lstSurfaceUnitField
        '
        Me.lstSurfaceUnitField.FormattingEnabled = True
        Me.lstSurfaceUnitField.ItemHeight = 16
        Me.lstSurfaceUnitField.Location = New System.Drawing.Point(279, 186)
        Me.lstSurfaceUnitField.Margin = New System.Windows.Forms.Padding(4)
        Me.lstSurfaceUnitField.Name = "lstSurfaceUnitField"
        Me.lstSurfaceUnitField.Size = New System.Drawing.Size(249, 388)
        Me.lstSurfaceUnitField.TabIndex = 8
        '
        'lstSurfaceUnitLayer
        '
        Me.lstSurfaceUnitLayer.FormattingEnabled = True
        Me.lstSurfaceUnitLayer.ItemHeight = 16
        Me.lstSurfaceUnitLayer.Location = New System.Drawing.Point(9, 186)
        Me.lstSurfaceUnitLayer.Margin = New System.Windows.Forms.Padding(4)
        Me.lstSurfaceUnitLayer.Name = "lstSurfaceUnitLayer"
        Me.lstSurfaceUnitLayer.Size = New System.Drawing.Size(249, 388)
        Me.lstSurfaceUnitLayer.TabIndex = 7
        '
        'lblDEMlist
        '
        Me.lblDEMlist.AutoSize = True
        Me.lblDEMlist.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDEMlist.Location = New System.Drawing.Point(568, 160)
        Me.lblDEMlist.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDEMlist.Name = "lblDEMlist"
        Me.lblDEMlist.Size = New System.Drawing.Size(130, 20)
        Me.lblDEMlist.TabIndex = 6
        Me.lblDEMlist.Text = "Elevation grid:"
        '
        'lblSurfaceUnitField
        '
        Me.lblSurfaceUnitField.AutoSize = True
        Me.lblSurfaceUnitField.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSurfaceUnitField.Location = New System.Drawing.Point(275, 160)
        Me.lblSurfaceUnitField.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSurfaceUnitField.Name = "lblSurfaceUnitField"
        Me.lblSurfaceUnitField.Size = New System.Drawing.Size(175, 20)
        Me.lblSurfaceUnitField.TabIndex = 5
        Me.lblSurfaceUnitField.Text = "Surface polygon layer:"
        '
        'lblSurfaceUnitLayer
        '
        Me.lblSurfaceUnitLayer.AutoSize = True
        Me.lblSurfaceUnitLayer.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSurfaceUnitLayer.Location = New System.Drawing.Point(9, 160)
        Me.lblSurfaceUnitLayer.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSurfaceUnitLayer.Name = "lblSurfaceUnitLayer"
        Me.lblSurfaceUnitLayer.Size = New System.Drawing.Size(197, 20)
        Me.lblSurfaceUnitLayer.TabIndex = 4
        Me.lblSurfaceUnitLayer.Text = "Surface polygon layer:"
        '
        'fraElevationUnits
        '
        Me.fraElevationUnits.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.fraElevationUnits.Controls.Add(Me.optElevationM)
        Me.fraElevationUnits.Controls.Add(Me.optElevationFt)
        Me.fraElevationUnits.Location = New System.Drawing.Point(572, 82)
        Me.fraElevationUnits.Margin = New System.Windows.Forms.Padding(4)
        Me.fraElevationUnits.Name = "fraElevationUnits"
        Me.fraElevationUnits.Size = New System.Drawing.Size(246, 62)
        Me.fraElevationUnits.TabIndex = 3
        '
        'optElevationM
        '
        Me.optElevationM.AutoSize = True
        Me.optElevationM.Location = New System.Drawing.Point(75, 17)
        Me.optElevationM.Margin = New System.Windows.Forms.Padding(4)
        Me.optElevationM.Name = "optElevationM"
        Me.optElevationM.Size = New System.Drawing.Size(72, 21)
        Me.optElevationM.TabIndex = 1
        Me.optElevationM.TabStop = True
        Me.optElevationM.Text = "Meters"
        Me.optElevationM.UseVisualStyleBackColor = True
        '
        'optElevationFt
        '
        Me.optElevationFt.AutoSize = True
        Me.optElevationFt.Location = New System.Drawing.Point(5, 17)
        Me.optElevationFt.Margin = New System.Windows.Forms.Padding(4)
        Me.optElevationFt.Name = "optElevationFt"
        Me.optElevationFt.Size = New System.Drawing.Size(57, 21)
        Me.optElevationFt.TabIndex = 0
        Me.optElevationFt.TabStop = True
        Me.optElevationFt.Text = "Feet"
        Me.optElevationFt.UseVisualStyleBackColor = True
        '
        'lblElevationUnits
        '
        Me.lblElevationUnits.AutoSize = True
        Me.lblElevationUnits.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblElevationUnits.Location = New System.Drawing.Point(573, 58)
        Me.lblElevationUnits.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblElevationUnits.Name = "lblElevationUnits"
        Me.lblElevationUnits.Size = New System.Drawing.Size(142, 20)
        Me.lblElevationUnits.TabIndex = 2
        Me.lblElevationUnits.Text = "Elevation Units:"
        '
        'fraProfileDelineationOptions
        '
        Me.fraProfileDelineationOptions.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.fraProfileDelineationOptions.Controls.Add(Me.optProfileDelineationOff)
        Me.fraProfileDelineationOptions.Controls.Add(Me.optProfileDelineationOn)
        Me.fraProfileDelineationOptions.Location = New System.Drawing.Point(13, 42)
        Me.fraProfileDelineationOptions.Margin = New System.Windows.Forms.Padding(4)
        Me.fraProfileDelineationOptions.Name = "fraProfileDelineationOptions"
        Me.fraProfileDelineationOptions.Size = New System.Drawing.Size(430, 103)
        Me.fraProfileDelineationOptions.TabIndex = 1
        '
        'optProfileDelineationOff
        '
        Me.optProfileDelineationOff.AutoSize = True
        Me.optProfileDelineationOff.Location = New System.Drawing.Point(5, 44)
        Me.optProfileDelineationOff.Margin = New System.Windows.Forms.Padding(4)
        Me.optProfileDelineationOff.Name = "optProfileDelineationOff"
        Me.optProfileDelineationOff.Size = New System.Drawing.Size(163, 21)
        Me.optProfileDelineationOff.TabIndex = 1
        Me.optProfileDelineationOff.TabStop = True
        Me.optProfileDelineationOff.Text = "No profile delineation"
        Me.optProfileDelineationOff.UseVisualStyleBackColor = True
        '
        'optProfileDelineationOn
        '
        Me.optProfileDelineationOn.AutoSize = True
        Me.optProfileDelineationOn.Location = New System.Drawing.Point(4, 15)
        Me.optProfileDelineationOn.Margin = New System.Windows.Forms.Padding(4)
        Me.optProfileDelineationOn.Name = "optProfileDelineationOn"
        Me.optProfileDelineationOn.Size = New System.Drawing.Size(280, 21)
        Me.optProfileDelineationOn.TabIndex = 0
        Me.optProfileDelineationOn.TabStop = True
        Me.optProfileDelineationOn.Text = "Delineate profile based on surface units"
        Me.optProfileDelineationOn.UseVisualStyleBackColor = True
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.Location = New System.Drawing.Point(9, 12)
        Me.Label25.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(171, 20)
        Me.Label25.TabIndex = 0
        Me.Label25.Text = "Profile Delineation:"
        '
        'Page4
        '
        Me.Page4.Controls.Add(Me.Label9)
        Me.Page4.Controls.Add(Me.Label7)
        Me.Page4.Controls.Add(Me.Label6)
        Me.Page4.Controls.Add(Me.Label5)
        Me.Page4.Controls.Add(Me.Label3)
        Me.Page4.Controls.Add(Me.Label2)
        Me.Page4.Controls.Add(Me.Label1)
        Me.Page4.Controls.Add(Me.Panel3)
        Me.Page4.Controls.Add(Me.lblBoxUnits)
        Me.Page4.Controls.Add(Me.txtHTickMinor)
        Me.Page4.Controls.Add(Me.txtVTickMinor)
        Me.Page4.Controls.Add(Me.txtHTick)
        Me.Page4.Controls.Add(Me.txtVTick)
        Me.Page4.Controls.Add(Me.chkUseProfileL)
        Me.Page4.Controls.Add(Me.txtBoxL)
        Me.Page4.Controls.Add(Me.txtBoxB)
        Me.Page4.Controls.Add(Me.txtBoxT)
        Me.Page4.Controls.Add(Me.chkSquareGrid)
        Me.Page4.Controls.Add(Me.chkBuildGrid)
        Me.Page4.Controls.Add(Me.lblMinorVertTick)
        Me.Page4.Controls.Add(Me.lblMinorHorizTick)
        Me.Page4.Controls.Add(Me.chkMinorTick)
        Me.Page4.Controls.Add(Me.lblHorizTick)
        Me.Page4.Controls.Add(Me.lblVerticalTick)
        Me.Page4.Controls.Add(Me.lblBoxLength)
        Me.Page4.Controls.Add(Me.lblBoxBottom)
        Me.Page4.Controls.Add(Me.lblBoxTop)
        Me.Page4.Controls.Add(Me.Label19)
        Me.Page4.Controls.Add(Me.Label24)
        Me.Page4.Controls.Add(Me.Label18)
        Me.Page4.Controls.Add(Me.Label17)
        Me.Page4.Location = New System.Drawing.Point(4, 25)
        Me.Page4.Margin = New System.Windows.Forms.Padding(4)
        Me.Page4.Name = "Page4"
        Me.Page4.Padding = New System.Windows.Forms.Padding(4)
        Me.Page4.Size = New System.Drawing.Size(855, 580)
        Me.Page4.TabIndex = 3
        Me.Page4.Text = "Profile Box"
        Me.Page4.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(386, 148)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(51, 17)
        Me.Label9.TabIndex = 31
        Me.Label9.Text = "meters"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(423, 400)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(51, 17)
        Me.Label7.TabIndex = 30
        Me.Label7.Text = "meters"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(423, 368)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 17)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "meters"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(423, 273)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 17)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "meters"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(423, 244)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 17)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "meters"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(386, 120)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 17)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "meters"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(386, 91)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "meters"
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.optBoxUnitsM)
        Me.Panel3.Controls.Add(Me.optBoxUnitsFT)
        Me.Panel3.Location = New System.Drawing.Point(243, 39)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(268, 37)
        Me.Panel3.TabIndex = 24
        '
        'optBoxUnitsM
        '
        Me.optBoxUnitsM.AutoSize = True
        Me.optBoxUnitsM.Location = New System.Drawing.Point(71, 12)
        Me.optBoxUnitsM.Margin = New System.Windows.Forms.Padding(4)
        Me.optBoxUnitsM.Name = "optBoxUnitsM"
        Me.optBoxUnitsM.Size = New System.Drawing.Size(72, 21)
        Me.optBoxUnitsM.TabIndex = 1
        Me.optBoxUnitsM.TabStop = True
        Me.optBoxUnitsM.Text = "Meters"
        Me.optBoxUnitsM.UseVisualStyleBackColor = True
        '
        'optBoxUnitsFT
        '
        Me.optBoxUnitsFT.AutoSize = True
        Me.optBoxUnitsFT.Location = New System.Drawing.Point(0, 12)
        Me.optBoxUnitsFT.Margin = New System.Windows.Forms.Padding(4)
        Me.optBoxUnitsFT.Name = "optBoxUnitsFT"
        Me.optBoxUnitsFT.Size = New System.Drawing.Size(57, 21)
        Me.optBoxUnitsFT.TabIndex = 0
        Me.optBoxUnitsFT.TabStop = True
        Me.optBoxUnitsFT.Text = "Feet"
        Me.optBoxUnitsFT.UseVisualStyleBackColor = True
        '
        'lblBoxUnits
        '
        Me.lblBoxUnits.AutoSize = True
        Me.lblBoxUnits.Location = New System.Drawing.Point(59, 58)
        Me.lblBoxUnits.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBoxUnits.Name = "lblBoxUnits"
        Me.lblBoxUnits.Size = New System.Drawing.Size(75, 17)
        Me.lblBoxUnits.TabIndex = 23
        Me.lblBoxUnits.Text = "Grid Units:"
        '
        'txtHTickMinor
        '
        Me.txtHTickMinor.Enabled = False
        Me.txtHTickMinor.Location = New System.Drawing.Point(282, 391)
        Me.txtHTickMinor.Margin = New System.Windows.Forms.Padding(4)
        Me.txtHTickMinor.Name = "txtHTickMinor"
        Me.txtHTickMinor.Size = New System.Drawing.Size(132, 22)
        Me.txtHTickMinor.TabIndex = 22
        Me.txtHTickMinor.Text = "0"
        Me.txtHTickMinor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtVTickMinor
        '
        Me.txtVTickMinor.Enabled = False
        Me.txtVTickMinor.Location = New System.Drawing.Point(282, 362)
        Me.txtVTickMinor.Margin = New System.Windows.Forms.Padding(4)
        Me.txtVTickMinor.Name = "txtVTickMinor"
        Me.txtVTickMinor.Size = New System.Drawing.Size(132, 22)
        Me.txtVTickMinor.TabIndex = 21
        Me.txtVTickMinor.Text = "0"
        Me.txtVTickMinor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtHTick
        '
        Me.txtHTick.Location = New System.Drawing.Point(282, 265)
        Me.txtHTick.Margin = New System.Windows.Forms.Padding(4)
        Me.txtHTick.Name = "txtHTick"
        Me.txtHTick.Size = New System.Drawing.Size(132, 22)
        Me.txtHTick.TabIndex = 20
        Me.txtHTick.Text = "0"
        Me.txtHTick.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtVTick
        '
        Me.txtVTick.Location = New System.Drawing.Point(282, 235)
        Me.txtVTick.Margin = New System.Windows.Forms.Padding(4)
        Me.txtVTick.Name = "txtVTick"
        Me.txtVTick.Size = New System.Drawing.Size(132, 22)
        Me.txtVTick.TabIndex = 19
        Me.txtVTick.Text = "0"
        Me.txtVTick.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'chkUseProfileL
        '
        Me.chkUseProfileL.AutoSize = True
        Me.chkUseProfileL.Location = New System.Drawing.Point(243, 176)
        Me.chkUseProfileL.Margin = New System.Windows.Forms.Padding(4)
        Me.chkUseProfileL.Name = "chkUseProfileL"
        Me.chkUseProfileL.Size = New System.Drawing.Size(141, 21)
        Me.chkUseProfileL.TabIndex = 18
        Me.chkUseProfileL.Text = "Use profile length"
        Me.chkUseProfileL.UseVisualStyleBackColor = True
        '
        'txtBoxL
        '
        Me.txtBoxL.Location = New System.Drawing.Point(243, 143)
        Me.txtBoxL.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBoxL.Name = "txtBoxL"
        Me.txtBoxL.Size = New System.Drawing.Size(132, 22)
        Me.txtBoxL.TabIndex = 17
        Me.txtBoxL.Text = "0"
        Me.txtBoxL.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtBoxB
        '
        Me.txtBoxB.Location = New System.Drawing.Point(243, 113)
        Me.txtBoxB.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBoxB.Name = "txtBoxB"
        Me.txtBoxB.Size = New System.Drawing.Size(132, 22)
        Me.txtBoxB.TabIndex = 16
        Me.txtBoxB.Text = "0"
        Me.txtBoxB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtBoxT
        '
        Me.txtBoxT.Location = New System.Drawing.Point(243, 84)
        Me.txtBoxT.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBoxT.Name = "txtBoxT"
        Me.txtBoxT.Size = New System.Drawing.Size(132, 22)
        Me.txtBoxT.TabIndex = 15
        Me.txtBoxT.Text = "0"
        Me.txtBoxT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'chkSquareGrid
        '
        Me.chkSquareGrid.AutoSize = True
        Me.chkSquareGrid.Location = New System.Drawing.Point(64, 492)
        Me.chkSquareGrid.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSquareGrid.Name = "chkSquareGrid"
        Me.chkSquareGrid.Size = New System.Drawing.Size(199, 21)
        Me.chkSquareGrid.TabIndex = 14
        Me.chkSquareGrid.Text = "Build a visually square grid"
        Me.chkSquareGrid.UseVisualStyleBackColor = True
        '
        'chkBuildGrid
        '
        Me.chkBuildGrid.AutoSize = True
        Me.chkBuildGrid.Checked = True
        Me.chkBuildGrid.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBuildGrid.Location = New System.Drawing.Point(64, 463)
        Me.chkBuildGrid.Margin = New System.Windows.Forms.Padding(4)
        Me.chkBuildGrid.Name = "chkBuildGrid"
        Me.chkBuildGrid.Size = New System.Drawing.Size(208, 21)
        Me.chkBuildGrid.TabIndex = 13
        Me.chkBuildGrid.Text = "Build a grid inside of the box"
        Me.chkBuildGrid.UseVisualStyleBackColor = True
        '
        'lblMinorVertTick
        '
        Me.lblMinorVertTick.AutoSize = True
        Me.lblMinorVertTick.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMinorVertTick.Location = New System.Drawing.Point(60, 366)
        Me.lblMinorVertTick.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMinorVertTick.Name = "lblMinorVertTick"
        Me.lblMinorVertTick.Size = New System.Drawing.Size(137, 17)
        Me.lblMinorVertTick.TabIndex = 12
        Me.lblMinorVertTick.Text = "Vertical tick spacing:"
        '
        'lblMinorHorizTick
        '
        Me.lblMinorHorizTick.AutoSize = True
        Me.lblMinorHorizTick.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMinorHorizTick.Location = New System.Drawing.Point(60, 395)
        Me.lblMinorHorizTick.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMinorHorizTick.Name = "lblMinorHorizTick"
        Me.lblMinorHorizTick.Size = New System.Drawing.Size(154, 17)
        Me.lblMinorHorizTick.TabIndex = 11
        Me.lblMinorHorizTick.Text = "Horizontal tick spacing:"
        '
        'chkMinorTick
        '
        Me.chkMinorTick.AutoSize = True
        Me.chkMinorTick.Location = New System.Drawing.Point(64, 336)
        Me.chkMinorTick.Margin = New System.Windows.Forms.Padding(4)
        Me.chkMinorTick.Name = "chkMinorTick"
        Me.chkMinorTick.Size = New System.Drawing.Size(146, 21)
        Me.chkMinorTick.TabIndex = 9
        Me.chkMinorTick.Text = "Include minor ticks"
        Me.chkMinorTick.UseVisualStyleBackColor = True
        '
        'lblHorizTick
        '
        Me.lblHorizTick.AutoSize = True
        Me.lblHorizTick.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHorizTick.Location = New System.Drawing.Point(59, 268)
        Me.lblHorizTick.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblHorizTick.Name = "lblHorizTick"
        Me.lblHorizTick.Size = New System.Drawing.Size(154, 17)
        Me.lblHorizTick.TabIndex = 8
        Me.lblHorizTick.Text = "Horizontal tick spacing:"
        '
        'lblVerticalTick
        '
        Me.lblVerticalTick.AutoSize = True
        Me.lblVerticalTick.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVerticalTick.Location = New System.Drawing.Point(59, 239)
        Me.lblVerticalTick.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblVerticalTick.Name = "lblVerticalTick"
        Me.lblVerticalTick.Size = New System.Drawing.Size(137, 17)
        Me.lblVerticalTick.TabIndex = 7
        Me.lblVerticalTick.Text = "Vertical tick spacing:"
        '
        'lblBoxLength
        '
        Me.lblBoxLength.AutoSize = True
        Me.lblBoxLength.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoxLength.Location = New System.Drawing.Point(59, 146)
        Me.lblBoxLength.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBoxLength.Name = "lblBoxLength"
        Me.lblBoxLength.Size = New System.Drawing.Size(56, 17)
        Me.lblBoxLength.TabIndex = 6
        Me.lblBoxLength.Text = "Length:"
        '
        'lblBoxBottom
        '
        Me.lblBoxBottom.AutoSize = True
        Me.lblBoxBottom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoxBottom.Location = New System.Drawing.Point(59, 117)
        Me.lblBoxBottom.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBoxBottom.Name = "lblBoxBottom"
        Me.lblBoxBottom.Size = New System.Drawing.Size(117, 17)
        Me.lblBoxBottom.TabIndex = 5
        Me.lblBoxBottom.Text = "Bottom elevation:"
        '
        'lblBoxTop
        '
        Me.lblBoxTop.AutoSize = True
        Me.lblBoxTop.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoxTop.Location = New System.Drawing.Point(59, 87)
        Me.lblBoxTop.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBoxTop.Name = "lblBoxTop"
        Me.lblBoxTop.Size = New System.Drawing.Size(98, 17)
        Me.lblBoxTop.TabIndex = 4
        Me.lblBoxTop.Text = "Top elevation:"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(19, 433)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(131, 17)
        Me.Label19.TabIndex = 3
        Me.Label19.Text = "Gridline Options:"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.Location = New System.Drawing.Point(19, 306)
        Me.Label24.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(155, 17)
        Me.Label24.TabIndex = 2
        Me.Label24.Text = "Minor Tick Intervals:"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(19, 209)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(155, 17)
        Me.Label18.TabIndex = 1
        Me.Label18.Text = "Major Tick Intervals:"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(19, 27)
        Me.Label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(127, 17)
        Me.Label17.TabIndex = 0
        Me.Label17.Text = "Box Dimensions:"
        '
        'Page3
        '
        Me.Page3.Controls.Add(Me.lstTableUnit)
        Me.Page3.Controls.Add(Me.lstWellsElev)
        Me.Page3.Controls.Add(Me.lstTableBottom)
        Me.Page3.Controls.Add(Me.lstTablePointID)
        Me.Page3.Controls.Add(Me.lstWellsPointID)
        Me.Page3.Controls.Add(Me.lstTableTop)
        Me.Page3.Controls.Add(Me.lstTables)
        Me.Page3.Controls.Add(Me.lstWellLayer)
        Me.Page3.Controls.Add(Me.chkGetElevFromDEM)
        Me.Page3.Controls.Add(Me.Panel2)
        Me.Page3.Controls.Add(Me.Frame2)
        Me.Page3.Controls.Add(Me.lblfraWellUnits)
        Me.Page3.Controls.Add(Me.lblFrame2)
        Me.Page3.Controls.Add(Me.Label13)
        Me.Page3.Controls.Add(Me.lblWellsElev)
        Me.Page3.Controls.Add(Me.lblwellBot)
        Me.Page3.Controls.Add(Me.lblPlgnID)
        Me.Page3.Controls.Add(Me.Label14)
        Me.Page3.Controls.Add(Me.lblWellTop)
        Me.Page3.Controls.Add(Me.Label8)
        Me.Page3.Controls.Add(Me.lbl1)
        Me.Page3.Location = New System.Drawing.Point(4, 22)
        Me.Page3.Name = "Page3"
        Me.Page3.Padding = New System.Windows.Forms.Padding(3)
        Me.Page3.Size = New System.Drawing.Size(639, 469)
        Me.Page3.TabIndex = 2
        Me.Page3.Text = "Wells"
        Me.Page3.UseVisualStyleBackColor = True
        '
        'lstTableUnit
        '
        Me.lstTableUnit.FormattingEnabled = True
        Me.lstTableUnit.ItemHeight = 16
        Me.lstTableUnit.Location = New System.Drawing.Point(219, 337)
        Me.lstTableUnit.Name = "lstTableUnit"
        Me.lstTableUnit.Size = New System.Drawing.Size(176, 100)
        Me.lstTableUnit.TabIndex = 20
        '
        'lstWellsElev
        '
        Me.lstWellsElev.FormattingEnabled = True
        Me.lstWellsElev.ItemHeight = 16
        Me.lstWellsElev.Location = New System.Drawing.Point(10, 335)
        Me.lstWellsElev.Name = "lstWellsElev"
        Me.lstWellsElev.Size = New System.Drawing.Size(176, 100)
        Me.lstWellsElev.TabIndex = 19
        '
        'lstTableBottom
        '
        Me.lstTableBottom.FormattingEnabled = True
        Me.lstTableBottom.ItemHeight = 16
        Me.lstTableBottom.Location = New System.Drawing.Point(432, 201)
        Me.lstTableBottom.Name = "lstTableBottom"
        Me.lstTableBottom.Size = New System.Drawing.Size(176, 100)
        Me.lstTableBottom.TabIndex = 18
        '
        'lstTablePointID
        '
        Me.lstTablePointID.FormattingEnabled = True
        Me.lstTablePointID.ItemHeight = 16
        Me.lstTablePointID.Location = New System.Drawing.Point(218, 201)
        Me.lstTablePointID.Name = "lstTablePointID"
        Me.lstTablePointID.Size = New System.Drawing.Size(176, 100)
        Me.lstTablePointID.TabIndex = 17
        '
        'lstWellsPointID
        '
        Me.lstWellsPointID.FormattingEnabled = True
        Me.lstWellsPointID.ItemHeight = 16
        Me.lstWellsPointID.Location = New System.Drawing.Point(10, 201)
        Me.lstWellsPointID.Name = "lstWellsPointID"
        Me.lstWellsPointID.Size = New System.Drawing.Size(176, 100)
        Me.lstWellsPointID.TabIndex = 16
        '
        'lstTableTop
        '
        Me.lstTableTop.FormattingEnabled = True
        Me.lstTableTop.ItemHeight = 16
        Me.lstTableTop.Location = New System.Drawing.Point(432, 28)
        Me.lstTableTop.Name = "lstTableTop"
        Me.lstTableTop.Size = New System.Drawing.Size(176, 132)
        Me.lstTableTop.TabIndex = 15
        '
        'lstTables
        '
        Me.lstTables.FormattingEnabled = True
        Me.lstTables.ItemHeight = 16
        Me.lstTables.Location = New System.Drawing.Point(218, 28)
        Me.lstTables.Name = "lstTables"
        Me.lstTables.Size = New System.Drawing.Size(176, 132)
        Me.lstTables.TabIndex = 14
        '
        'lstWellLayer
        '
        Me.lstWellLayer.FormattingEnabled = True
        Me.lstWellLayer.ItemHeight = 16
        Me.lstWellLayer.Location = New System.Drawing.Point(10, 28)
        Me.lstWellLayer.Name = "lstWellLayer"
        Me.lstWellLayer.Size = New System.Drawing.Size(176, 132)
        Me.lstWellLayer.TabIndex = 13
        '
        'chkGetElevFromDEM
        '
        Me.chkGetElevFromDEM.AutoSize = True
        Me.chkGetElevFromDEM.Location = New System.Drawing.Point(4, 449)
        Me.chkGetElevFromDEM.Name = "chkGetElevFromDEM"
        Me.chkGetElevFromDEM.Size = New System.Drawing.Size(258, 21)
        Me.chkGetElevFromDEM.TabIndex = 12
        Me.chkGetElevFromDEM.Text = "Get well surface elevation from DEM"
        Me.chkGetElevFromDEM.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.optWellUnitsM)
        Me.Panel2.Controls.Add(Me.optWellUnitsFt)
        Me.Panel2.Location = New System.Drawing.Point(432, 410)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(201, 42)
        Me.Panel2.TabIndex = 11
        '
        'optWellUnitsM
        '
        Me.optWellUnitsM.AutoSize = True
        Me.optWellUnitsM.Location = New System.Drawing.Point(73, 10)
        Me.optWellUnitsM.Name = "optWellUnitsM"
        Me.optWellUnitsM.Size = New System.Drawing.Size(72, 21)
        Me.optWellUnitsM.TabIndex = 1
        Me.optWellUnitsM.TabStop = True
        Me.optWellUnitsM.Text = "Meters"
        Me.optWellUnitsM.UseVisualStyleBackColor = True
        '
        'optWellUnitsFt
        '
        Me.optWellUnitsFt.AutoSize = True
        Me.optWellUnitsFt.Location = New System.Drawing.Point(7, 10)
        Me.optWellUnitsFt.Name = "optWellUnitsFt"
        Me.optWellUnitsFt.Size = New System.Drawing.Size(57, 21)
        Me.optWellUnitsFt.TabIndex = 0
        Me.optWellUnitsFt.TabStop = True
        Me.optWellUnitsFt.Text = "Feet"
        Me.optWellUnitsFt.UseVisualStyleBackColor = True
        '
        'Frame2
        '
        Me.Frame2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Frame2.Controls.Add(Me.txtPolyW)
        Me.Frame2.Controls.Add(Me.lblPolyW)
        Me.Frame2.Controls.Add(Me.optPolys)
        Me.Frame2.Controls.Add(Me.optLines)
        Me.Frame2.Location = New System.Drawing.Point(432, 333)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Size = New System.Drawing.Size(201, 54)
        Me.Frame2.TabIndex = 10
        '
        'txtPolyW
        '
        Me.txtPolyW.Enabled = False
        Me.txtPolyW.Location = New System.Drawing.Point(96, 24)
        Me.txtPolyW.Name = "txtPolyW"
        Me.txtPolyW.Size = New System.Drawing.Size(100, 22)
        Me.txtPolyW.TabIndex = 3
        Me.txtPolyW.Text = "10"
        Me.txtPolyW.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblPolyW
        '
        Me.lblPolyW.AutoSize = True
        Me.lblPolyW.Location = New System.Drawing.Point(4, 27)
        Me.lblPolyW.Name = "lblPolyW"
        Me.lblPolyW.Size = New System.Drawing.Size(121, 17)
        Me.lblPolyW.TabIndex = 2
        Me.lblPolyW.Text = "Polygon width (ft):"
        '
        'optPolys
        '
        Me.optPolys.AutoSize = True
        Me.optPolys.Location = New System.Drawing.Point(69, 3)
        Me.optPolys.Name = "optPolys"
        Me.optPolys.Size = New System.Drawing.Size(87, 21)
        Me.optPolys.TabIndex = 1
        Me.optPolys.TabStop = True
        Me.optPolys.Text = "Polygons"
        Me.optPolys.UseVisualStyleBackColor = True
        '
        'optLines
        '
        Me.optLines.AutoSize = True
        Me.optLines.Location = New System.Drawing.Point(3, 3)
        Me.optLines.Name = "optLines"
        Me.optLines.Size = New System.Drawing.Size(63, 21)
        Me.optLines.TabIndex = 0
        Me.optLines.TabStop = True
        Me.optLines.Text = "Lines"
        Me.optLines.UseVisualStyleBackColor = True
        '
        'lblfraWellUnits
        '
        Me.lblfraWellUnits.AutoSize = True
        Me.lblfraWellUnits.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblfraWellUnits.Location = New System.Drawing.Point(429, 390)
        Me.lblfraWellUnits.Name = "lblfraWellUnits"
        Me.lblfraWellUnits.Size = New System.Drawing.Size(116, 20)
        Me.lblfraWellUnits.TabIndex = 9
        Me.lblfraWellUnits.Text = "Well Units As:"
        '
        'lblFrame2
        '
        Me.lblFrame2.AutoSize = True
        Me.lblFrame2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFrame2.Location = New System.Drawing.Point(429, 312)
        Me.lblFrame2.Name = "lblFrame2"
        Me.lblFrame2.Size = New System.Drawing.Size(127, 20)
        Me.lblFrame2.TabIndex = 8
        Me.lblFrame2.Text = "Show Wells As:"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(215, 312)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(167, 20)
        Me.Label13.TabIndex = 7
        Me.Label13.Text = "Subsurface unit field:"
        '
        'lblWellsElev
        '
        Me.lblWellsElev.AutoSize = True
        Me.lblWellsElev.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWellsElev.Location = New System.Drawing.Point(7, 312)
        Me.lblWellsElev.Name = "lblWellsElev"
        Me.lblWellsElev.Size = New System.Drawing.Size(242, 20)
        Me.lblWellsElev.TabIndex = 6
        Me.lblWellsElev.Text = "Well surface elevation (ft) field:"
        '
        'lblwellBot
        '
        Me.lblwellBot.AutoSize = True
        Me.lblwellBot.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblwellBot.Location = New System.Drawing.Point(429, 175)
        Me.lblwellBot.Name = "lblwellBot"
        Me.lblwellBot.Size = New System.Drawing.Size(221, 20)
        Me.lblwellBot.TabIndex = 5
        Me.lblwellBot.Text = "Layer bottom depth (ft) field:"
        '
        'lblPlgnID
        '
        Me.lblPlgnID.AutoSize = True
        Me.lblPlgnID.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlgnID.Location = New System.Drawing.Point(215, 175)
        Me.lblPlgnID.Name = "lblPlgnID"
        Me.lblPlgnID.Size = New System.Drawing.Size(110, 20)
        Me.lblPlgnID.TabIndex = 4
        Me.lblPlgnID.Text = "Point ID field:"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(7, 175)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(110, 20)
        Me.Label14.TabIndex = 3
        Me.Label14.Text = "Point ID field:"
        '
        'lblWellTop
        '
        Me.lblWellTop.AutoSize = True
        Me.lblWellTop.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWellTop.Location = New System.Drawing.Point(429, 7)
        Me.lblWellTop.Name = "lblWellTop"
        Me.lblWellTop.Size = New System.Drawing.Size(193, 20)
        Me.lblWellTop.TabIndex = 2
        Me.lblWellTop.Text = "Layer top depth (ft) field:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(215, 7)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(208, 20)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "Subsurface Data Table:"
        '
        'lbl1
        '
        Me.lbl1.AutoSize = True
        Me.lbl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.Location = New System.Drawing.Point(7, 7)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(183, 20)
        Me.lbl1.TabIndex = 0
        Me.lbl1.Text = "Well Location Layer:"
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(749, 761)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(100, 28)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(614, 761)
        Me.cmdStart.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(127, 28)
        Me.cmdStart.TabIndex = 2
        Me.cmdStart.Text = "Create Profile"
        Me.cmdStart.UseVisualStyleBackColor = True
        '
        'lblPrgBr
        '
        Me.lblPrgBr.AutoSize = True
        Me.lblPrgBr.Location = New System.Drawing.Point(10, 742)
        Me.lblPrgBr.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPrgBr.Name = "lblPrgBr"
        Me.lblPrgBr.Size = New System.Drawing.Size(107, 17)
        Me.lblPrgBr.TabIndex = 3
        Me.lblPrgBr.Text = ""
        '
        'prgBr
        '
        Me.prgBr.Location = New System.Drawing.Point(13, 761)
        Me.prgBr.Margin = New System.Windows.Forms.Padding(4)
        Me.prgBr.Name = "prgBr"
        Me.prgBr.Size = New System.Drawing.Size(592, 28)
        Me.prgBr.TabIndex = 4
        '
        'frmProfileTool
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(868, 802)
        Me.Controls.Add(Me.prgBr)
        Me.Controls.Add(Me.lblPrgBr)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.MultiPage1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmProfileTool"
        Me.Text = "Profile Tool v3.0"
        Me.MultiPage1.ResumeLayout(False)
        Me.Page1.ResumeLayout(False)
        Me.Page1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Page2.ResumeLayout(False)
        Me.Page2.PerformLayout()
        Me.fraElevationUnits.ResumeLayout(False)
        Me.fraElevationUnits.PerformLayout()
        Me.fraProfileDelineationOptions.ResumeLayout(False)
        Me.fraProfileDelineationOptions.PerformLayout()
        Me.Page4.ResumeLayout(False)
        Me.Page4.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Page3.ResumeLayout(False)
        Me.Page3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MultiPage1 As System.Windows.Forms.TabControl
    Friend WithEvents Page1 As System.Windows.Forms.TabPage
    Friend WithEvents Page2 As System.Windows.Forms.TabPage
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents lblPrgBr As System.Windows.Forms.Label
    Friend WithEvents prgBr As System.Windows.Forms.ProgressBar
    Friend WithEvents Page3 As System.Windows.Forms.TabPage
    Friend WithEvents Page4 As System.Windows.Forms.TabPage
    Friend WithEvents cmdChangeTempPath As System.Windows.Forms.Button
    Friend WithEvents txtOutputPath As System.Windows.Forms.TextBox
    Friend WithEvents txtProfileName As System.Windows.Forms.TextBox
    Friend WithEvents txtEndY As System.Windows.Forms.TextBox
    Friend WithEvents txtStartY As System.Windows.Forms.TextBox
    Friend WithEvents txtEndX As System.Windows.Forms.TextBox
    Friend WithEvents txtStartX As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtVerticalExagg As System.Windows.Forms.TextBox
    Friend WithEvents txtPrecisionMeasure As System.Windows.Forms.TextBox
    Friend WithEvents lblParts1 As System.Windows.Forms.Label
    Friend WithEvents txtPrecisionParts As System.Windows.Forms.TextBox
    Friend WithEvents chkSaveSettings As System.Windows.Forms.CheckBox
    'Friend WithEvents chkGraphicLine As System.Windows.Forms.CheckBox
    Friend WithEvents smoothProfile As System.Windows.Forms.CheckBox
    Friend WithEvents chkAddWells As System.Windows.Forms.CheckBox
    Friend WithEvents chkBuildProfileBox As System.Windows.Forms.CheckBox
    Friend WithEvents chkVerticalExagg As System.Windows.Forms.CheckBox
    Friend WithEvents optPrecisionMeasure As System.Windows.Forms.RadioButton
    Friend WithEvents optPrecisionParts As System.Windows.Forms.RadioButton
    Friend WithEvents lblOutputPath As System.Windows.Forms.Label
    Friend WithEvents lblProfileName As System.Windows.Forms.Label
    Friend WithEvents lblYcoord As System.Windows.Forms.Label
    Friend WithEvents lblXcoord As System.Windows.Forms.Label
    Friend WithEvents lblOptions As System.Windows.Forms.Label
    Friend WithEvents lblPrecision As System.Windows.Forms.Label
    Friend WithEvents lblCoordEnd As System.Windows.Forms.Label
    Friend WithEvents lblCoordStart As System.Windows.Forms.Label
    Friend WithEvents lstElevationLayer As System.Windows.Forms.ListBox
    Friend WithEvents lstSurfaceUnitField As System.Windows.Forms.ListBox
    Friend WithEvents lstSurfaceUnitLayer As System.Windows.Forms.ListBox
    Friend WithEvents lblDEMlist As System.Windows.Forms.Label
    Friend WithEvents lblSurfaceUnitField As System.Windows.Forms.Label
    Friend WithEvents lblSurfaceUnitLayer As System.Windows.Forms.Label
    Friend WithEvents fraElevationUnits As System.Windows.Forms.Panel
    Friend WithEvents optElevationM As System.Windows.Forms.RadioButton
    Friend WithEvents optElevationFt As System.Windows.Forms.RadioButton
    Friend WithEvents lblElevationUnits As System.Windows.Forms.Label
    Friend WithEvents fraProfileDelineationOptions As System.Windows.Forms.Panel
    Friend WithEvents optProfileDelineationOff As System.Windows.Forms.RadioButton
    Friend WithEvents optProfileDelineationOn As System.Windows.Forms.RadioButton
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents lstTableUnit As System.Windows.Forms.ListBox
    Friend WithEvents lstWellsElev As System.Windows.Forms.ListBox
    Friend WithEvents lstTableBottom As System.Windows.Forms.ListBox
    Friend WithEvents lstTablePointID As System.Windows.Forms.ListBox
    Friend WithEvents lstWellsPointID As System.Windows.Forms.ListBox
    Friend WithEvents lstTableTop As System.Windows.Forms.ListBox
    Friend WithEvents lstTables As System.Windows.Forms.ListBox
    Friend WithEvents lstWellLayer As System.Windows.Forms.ListBox
    Friend WithEvents chkGetElevFromDEM As System.Windows.Forms.CheckBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents optWellUnitsM As System.Windows.Forms.RadioButton
    Friend WithEvents optWellUnitsFt As System.Windows.Forms.RadioButton
    Friend WithEvents Frame2 As System.Windows.Forms.Panel
    Friend WithEvents txtPolyW As System.Windows.Forms.TextBox
    Friend WithEvents lblPolyW As System.Windows.Forms.Label
    Friend WithEvents optPolys As System.Windows.Forms.RadioButton
    Friend WithEvents optLines As System.Windows.Forms.RadioButton
    Friend WithEvents lblfraWellUnits As System.Windows.Forms.Label
    Friend WithEvents lblFrame2 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lblWellsElev As System.Windows.Forms.Label
    Friend WithEvents lblwellBot As System.Windows.Forms.Label
    Friend WithEvents lblPlgnID As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lblWellTop As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents txtHTickMinor As System.Windows.Forms.TextBox
    Friend WithEvents txtVTickMinor As System.Windows.Forms.TextBox
    Friend WithEvents txtHTick As System.Windows.Forms.TextBox
    Friend WithEvents txtVTick As System.Windows.Forms.TextBox
    Friend WithEvents chkUseProfileL As System.Windows.Forms.CheckBox
    Friend WithEvents txtBoxL As System.Windows.Forms.TextBox
    Friend WithEvents txtBoxB As System.Windows.Forms.TextBox
    Friend WithEvents txtBoxT As System.Windows.Forms.TextBox
    Friend WithEvents chkSquareGrid As System.Windows.Forms.CheckBox
    Friend WithEvents chkBuildGrid As System.Windows.Forms.CheckBox
    Friend WithEvents lblMinorVertTick As System.Windows.Forms.Label
    Friend WithEvents lblMinorHorizTick As System.Windows.Forms.Label
    Friend WithEvents chkMinorTick As System.Windows.Forms.CheckBox
    Friend WithEvents lblHorizTick As System.Windows.Forms.Label
    Friend WithEvents lblVerticalTick As System.Windows.Forms.Label
    Friend WithEvents lblBoxLength As System.Windows.Forms.Label
    Friend WithEvents lblBoxBottom As System.Windows.Forms.Label
    Friend WithEvents lblBoxTop As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents optBoxUnitsM As System.Windows.Forms.RadioButton
    Friend WithEvents optBoxUnitsFT As System.Windows.Forms.RadioButton
    Friend WithEvents lblBoxUnits As System.Windows.Forms.Label
    Friend WithEvents lblPrecisionUnit As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
End Class
