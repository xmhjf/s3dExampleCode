Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Common.Middle

Public Class SizeEqual
    Inherits SupportPartSelectionRule

    Public Overrides Function SelectedPartFromPartClass(ByVal sPartName As String) As ReferenceData.Middle.Part
        Dim oSupport As SP3D.Support.Middle.Support
        Dim oRouteObject As PipeObjectInfo
        Dim oPart As Part

        oSupport = SupportHelper.Support
        oRouteObject = SupportedHelper.SupportedObjectInfo(ReferenceIndex)
        oPart = SupportPartBySize(sPartName, oRouteObject.NominalDiameter, NDComparisonOperatorType.NDFrom_EQUAL)

        Return oPart
    End Function

End Class
