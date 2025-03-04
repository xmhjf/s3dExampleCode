Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle.Services
Imports System.Math
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Route.Middle

Public Class SupportBOMDesc
    Inherits SupportBOMDescription

    Public Overrides Function BOMDescription(ByVal oSupportOrComponent As Common.Middle.BusinessObject) As String
        Dim oPipeSupp As PipeSupport = oSupportOrComponent

        'Get the Support Definition
        Dim oSupportDefinition As Part = SupportHelper.Support.SupportDefinition
        Dim oBeam1, oBeam2 As ConnectionComponent
        Dim dLength1, dLength2 As PropertyValueDouble
        Dim dictionary = SupportHelper.SupportComponentDictionary()
        oBeam1 = dictionary("VERT_SECTION")
        oBeam2 = dictionary("HOR_SECTION")
        dLength1 = oBeam1.GetPropertyValue("IJUAHgrOccLength", "Length")
        dLength2 = oBeam2.GetPropertyValue("IJUAHgrOccLength", "Length")

        'Get the Vertical Length from Route to Structure
        Dim dVLength As Double
        Dim sLength1, sLength2, sVLength As String
        dVLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical)

        'Format the distances into proper units
        sLength1 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dLength1.PropValue, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)
        sLength2 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dLength2.PropValue, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)
        sVLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dVLength, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)

        'Get The Primary Pipe Information and Specification
        Dim oPipeInfo As PipeObjectInfo = SupportedHelper.SupportedObjectInfo(1)
        Dim oPipeSpec As SpecificationBase = oPipeInfo.Spec

        'Generate the BOM Description
        Return "From ProgID: " & oSupportDefinition.PartDescription & ", " & oPipeInfo.NominalDiameter.Size & " " & oPipeInfo.NominalDiameter.Units & " NPD, " & oPipeSpec.SpecificationName & ", Vertical section length = " & sLength1 & ", Horizontal section length  = " & sLength2 & ", Vertical Length = " & sVLength
    End Function
End Class
