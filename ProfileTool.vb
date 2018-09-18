Imports ArcGIS.Desktop.Mapping
Imports ArcGIS.Desktop.Framework
Imports ArcGIS.Core.Geometry
Imports ArcGIS.Desktop.Framework.Threading.Tasks
Imports System.Windows.Forms
Imports ArcGIS.Core.CIM


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
'Converted to ArcGIS VB.net Add-In, April 2015 by:
'Kristen Jordan
'Kansas Data Access and Support Center
'Email: kristen@kgs.ku.edu

'Tool conversion managed by:
'John W. Dunham, Ph.D.
'Cartographic Services Manager
'Kansas Geological Survey

'Converted for ArcGIS Pro, April 2018 by: 
'Emily Bunse
'Kansas Geological Survey- Cartographic Services GRA
'Email: egbunse@ku.edu; egbunse@gmail.com

Public Class ProfileTool
    Inherits MapTool
    Dim startPoint As MapPoint = Nothing
    Dim endPoint As MapPoint = Nothing
    Public Shared pLineSymbolRef As CIMSymbolReference = New CIMSymbolReference()
    Public Shared graphic As IDisposable
    Dim xmlGeometry As String = Nothing
    Dim profileLine As Polyline = Nothing

    Public Sub New()
        'Enables the Sketch Tool which will draw a Magenta (additional colors provided) line between two points that the user clicks
        Forms.Cursor.Current = Windows.Forms.Cursors.Cross
        'Dim pCIMColor As CIMColor = CIMColor.CreateRGBColor(255, 255, 0)    'Yellow
        Dim pCIMColor As CIMColor = CIMColor.CreateRGBColor(255, 0, 255)   'Magenta
        'Dim pCIMColor As CIMColor = CIMColor.CreateRGBColor(0, 255, 255)  'Aqua
        'Dim pCIMColor As CIMColor = CIMColor.CreateRGBColor(255, 0, 0)  'Red
        'Dim pCIMColor As CIMColor = CIMColor.CreateRGBColor(140, 140, 200) 'Original Color used 
        Dim pLineSymbol As CIMLineSymbol = SymbolFactory.Instance.ConstructLineSymbol(pCIMColor, 1.5, SimpleLineStyle.Solid)
        pLineSymbolRef.Symbol = pLineSymbol
        IsSketchTool = True
        SketchType = SketchGeometryType.Line
        SketchOutputMode = SketchOutputMode.Map
        SketchSymbol = pLineSymbolRef
    End Sub

    'When the sketch (two points connected by a line) is complete, go to the 'prepareGeometry' function
    Protected Overrides Async Function OnSketchCompleteAsync(geometry As Geometry) As Task(Of Boolean)
        Return Await PrepareGeometry(geometry)
    End Function

    Protected Overrides Function OnToolDeactivateAsync(hasMapViewChanged As Boolean) As Task
        Return MyBase.OnToolDeactivateAsync(hasMapViewChanged)
    End Function

    'Confirms the line is two points, gathers the geometry, and hands each point to the 'StartGeoProfile' sub in 'modProfileTool_Main'
    Public Async Function PrepareGeometry(geometry As Geometry) As Task(Of Boolean)
        If geometry.PointCount = 2 Then
            Dim buttons As MessageBoxButtons = MessageBoxButton.YesNo
            Dim result As DialogResult = MessageBox.Show("Continue with this sketch?", "Approve Sketch", buttons)
            If result = DialogResult.Yes Then
                xmlGeometry = geometry.ToXML()
                Debug.Print(xmlGeometry)
                profileLine = Await QueuedTask.Run(Function() PolylineBuilder.FromXML(xmlGeometry))
                startPoint = profileLine.Points.Item(0)
                endPoint = profileLine.Points.Item(1)
                graphic = Await QueuedTask.Run(Function()
                                                   pLineSymbolRef.Symbol = SymbolFactory.Instance.ConstructLineSymbol(CIMColor.CreateRGBColor(0, 255, 255), 1.5, SimpleLineStyle.Solid)
                                                   Return MappingExtensions.AddOverlayAsync(MapView.Active, geometry, pLineSymbolRef)
                                               End Function)
                Forms.Cursor.Current = Windows.Forms.Cursors.Arrow
                Call StartGeoProfile(startPoint)
                Call StartGeoProfile(endPoint)
            Else
                Forms.Cursor.Current = Windows.Forms.Cursors.Cross
            End If
        Else
            MessageBox.Show("The profile sketch must be a line containing exactly 2 points.")
            Forms.Cursor.Current = Windows.Forms.Cursors.Cross
        End If
        Return True
    End Function
End Class
