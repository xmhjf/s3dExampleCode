Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle.Services
Imports System.Math
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Route.Middle

Public Class BOMExample
    Implements ICustomHgrBOMDescription

    Public Function BOMDescription(ByVal oSupportOrComponent As Common.Middle.BusinessObject) As String Implements ICustomHgrBOMDescription.BOMDescription

        Dim oPipeSupp As PipeSupport = oSupportOrComponent

        'Get the Support Definition
        Dim oSupportDefinition As Part = oPipeSupp.SupportDefinition

        'Get the HgrBeam Support Components
        Dim oSupCompColl = oPipeSupp.AssemblyChildren

        Dim oBeam1 As ConnectionComponent
        Dim oBeam2 As ConnectionComponent
        Dim iBeamCount As Integer = 1
        For Each oSupComp As SupportComponent In oSupCompColl
            If oSupComp.IsTypeOfBOC("ConnectionComponent") Then
                If iBeamCount = 1 Then
                    oBeam1 = oSupComp
                    iBeamCount = iBeamCount + 1
                Else
                    oBeam2 = oSupComp
                End If
            End If
        Next

        'Get the Lengths of each HgrBeam
        Dim dLength1 As PropertyValueDouble = oBeam1.GetPropertyValue("IJUAHgrOccLength", "Length")
        Dim dLength2 As PropertyValueDouble = oBeam2.GetPropertyValue("IJUAHgrOccLength", "Length")

        'Format the distances into proper units
        Dim sLength1 As String
        Dim sLength2 As String
        sLength1 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dLength1.PropValue, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)
        sLength2 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dLength2.PropValue, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)

        'Get The Primary Pipe Information and Specification
        Dim oPipeFeat As PipeStraightFeature = oPipeSupp.SupportedObjects(0)
        Dim oPipeRun As PipeRun = oPipeFeat.SystemParent

        Dim oPipeND As NominalDiameter = oPipeFeat.NPD
        Dim oPipeSpec As SpecificationBase = oPipeRun.Specification

        'Generate the BOM Description
        Return "From ProgID: " & oSupportDefinition.PartDescription & ", " & oPipeND.Size & " " & oPipeND.Units & " NPD, " & oPipeSpec.SpecificationName & ", L1 = " & sLength1 & ", L2 = " & sLength2

    End Function

End Class
