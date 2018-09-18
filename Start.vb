Imports ArcGIS.Desktop.Framework


Friend Class Start
    Inherits ArcGIS.Desktop.Framework.Contracts.Button
    'Set the cursor used to click the start and end points of the Profile to Crosshairs
    Protected Overrides Sub OnClick()
        Forms.Cursor.Current = Windows.Forms.Cursors.Cross

    End Sub
End Class

