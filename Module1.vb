Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Input
Imports ArcGIS.Desktop.Framework
Imports ArcGIS.Desktop.Framework.Contracts

Friend Class Module1
    Inherits ArcGIS.Desktop.Framework.Contracts.Module

    Private Shared Property _this As Object

    ''' <summary>
    ''' Retrieve the singleton instance to this module here
    ''' </summary>
    Public Shared ReadOnly Property Current() As Module1
        Get
            If (_this Is Nothing) Then
                _this = DirectCast(FrameworkApplication.FindModule("Profile_Tool_Pro_Module"), Module1)
            End If

            Return _this
        End Get
    End Property

#Region "Overrides"
    ''' <summary>
    ''' Called by Framework when ArcGIS Pro is closing
    ''' </summary>
    ''' <returns>False to prevent Pro from closing, otherwise True</returns>
    Protected Overrides Function CanUnload() As Boolean
        'TODO - add your business logic
        'return false to ~cancel~ Application close
        Return True
    End Function

#End Region

End Class
