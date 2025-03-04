Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle.Services
Imports System.Math
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Route.Middle

Public Class BOMUtilityPlate
    Implements ICustomHgrBOMDescription

    Public Function BOMDescription(ByVal oSupportOrComponent As Common.Middle.BusinessObject) As String Implements ICustomHgrBOMDescription.BOMDescription

        Dim oPart As IPart
        oPart = oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects(0)

        'Generate the BOM Description
        Return "From ProgId : " & oPart.PartDescription

    End Function

End Class
