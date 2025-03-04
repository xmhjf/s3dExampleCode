''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Copyright (C) 2016, Intergraph PPO. All rights reserved.
'File
'  StandardSupportSelRule.vb
'Abstract
'	This is an example .NET support selection rule.
'History :
'           Bathemma       06/08/2016          TR-CP-290861	RulePriority property is missing in the migrated databases
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Exceptions

'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Support.Content.Rules
'It is recommended that customers specify namespace of their classes to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this rule to suit their
'requirements, they should change the namespace/class name so the identity of the modified
'rule will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------

Public Class StandardSupportSelRule
    Inherits SupportSelectionRule

    Public Overrides ReadOnly Property Supports() As System.Collections.ObjectModel.Collection(Of ReferenceData.Middle.Part)
        Get
            'Define a collection to hold all the supports
            Dim oFilteredSupports As Collection(Of Part)
            'The discipline type for which this rule applies
            Dim eHgrDiscipline As DisciplineTypeFlags
            Dim eCrossSection As CrossSectionShape
            Dim dDepOrRad As Double
            'Get useful objects
            Dim oSupportedInfo = SupportedHelper.SupportedObjectInfo(1)     'Use the primary supported object

            'Get useful objects
            Dim soTypes(3) As SupportedObjectType, DiscFlags(3) As DisciplineTypeFlags
            soTypes(0) = SupportedObjectType.Pipe : DiscFlags(0) = DisciplineTypeFlags.Piping
            soTypes(1) = SupportedObjectType.Conduit : DiscFlags(1) = DisciplineTypeFlags.Conduit
            soTypes(2) = SupportedObjectType.HVAC : DiscFlags(2) = DisciplineTypeFlags.HVAC
            soTypes(3) = SupportedObjectType.CableTray : DiscFlags(3) = DisciplineTypeFlags.CableTray

            'Determine the discipine based on the supported object types
            Dim oSupportedObject As SupportedObjectInfo

            For i As Integer = 0 To 3
                If (oSupportedInfo.SupportedObjectType = soTypes(i)) Then
                    eHgrDiscipline = DiscFlags(i)
                    For Each oSupportedObject In SupportedHelper.SupportedObjectInfo(SupportedObjectType.All, FeatureType.All)
                        If oSupportedObject.SupportedObjectType <> soTypes(i) Then
                            eHgrDiscipline = DisciplineTypeFlags.Combined
                            Exit For
                        End If
                    Next
                End If
            Next

            'Filter the supports based on the size of the primary supported object
            If (oSupportedInfo.SupportedObjectType = SupportedObjectType.Pipe) Then
                Dim oPipeInfo As PipeObjectInfo = oSupportedInfo
                dDepOrRad = oPipeInfo.NominalDiameter.Size / 2.0
                dDepOrRad = ConvertToDBU(oPipeInfo.NominalDiameter.Units, dDepOrRad)
                oFilteredSupports = SupportsBySize(eHgrDiscipline, Nothing, oPipeInfo.NominalDiameter)
            ElseIf oSupportedInfo.SupportedObjectType = SupportedObjectType.Conduit Then
                Dim oConduitInfo As ConduitObjectInfo = oSupportedInfo
                dDepOrRad = oConduitInfo.NominalDiameter.Size / 2.0
                dDepOrRad = ConvertToDBU(oConduitInfo.NominalDiameter.Units, dDepOrRad)
                oFilteredSupports = SupportsBySize(eHgrDiscipline, Nothing, oConduitInfo.NominalDiameter)
            ElseIf oSupportedInfo.SupportedObjectType = SupportedObjectType.HVAC Then
                Dim oDuctInfo As DuctObjectInfo = oSupportedInfo
                dDepOrRad = oDuctInfo.Depth
                eCrossSection = oDuctInfo.CrossSectionShape
                oFilteredSupports = SupportsBySize(eHgrDiscipline, Nothing, oDuctInfo.Width, oDuctInfo.Depth, "m")
            ElseIf oSupportedInfo.SupportedObjectType = SupportedObjectType.CableTray Then
                Dim oCableWayInfo As CableTrayObjectInfo = oSupportedInfo
                dDepOrRad = oCableWayInfo.Depth
                oFilteredSupports = SupportsBySize(eHgrDiscipline, Nothing, oCableWayInfo.Width, oCableWayInfo.Depth, "m")
            End If

            'Filter supports based on input (SupportedFamily, Supported/Supporting Count, CommandType, etc...)
            oFilteredSupports = SupportsByInput(eHgrDiscipline, oFilteredSupports)

            If (oSupportedInfo.SupportedObjectType = SupportedObjectType.HVAC) Then
                oFilteredSupports = SupportsByCriteria(eHgrDiscipline, oFilteredSupports, "SupportedShapeType", ComparisonOperatorType.EQUAL, eCrossSection)
            End If

            If (SupportHelper.SupportingObjects.Count > 0) Then
                Dim dRouteStructMaxDistance, dRouteStructMinDistance As Double
                Dim eDistanceType As PortDistanceType

                If oSupportedInfo.SupportedObjectType = (SupportedObjectType.CableTray Or SupportedObjectType.HVAC) Then
                    eDistanceType = PortDistanceType.Direct
                Else
                    eDistanceType = PortDistanceType.Vertical
                End If

                'Filter supports based on TypeSelectionRule and the Configuration of the Primary Supported and Supporting Objects
                Dim oPortConfiguration = RefPortHelper.PortConfiguration(1, 1)
                If (oPortConfiguration = PortConfiguration.Intersecting) Then
                    dRouteStructMinDistance = 0
                Else
                    If (oPortConfiguration = PortConfiguration.Beside) Then
                        eDistanceType = PortDistanceType.Horizontal
                    End If
                End If

                If oSupportedInfo.FeatureType = FeatureType.Turn Then
                    eDistanceType = PortDistanceType.Vertical
                End If

                oFilteredSupports = SupportsByCriteria(eHgrDiscipline, oFilteredSupports, "TypeSelectionRule", ComparisonOperatorType.EQUAL, oPortConfiguration)   'Operator is not actually used with TypeSelectionRule

                'Filter supports based on Min and Max Assembly Length
                RefPortHelper.DistanceBetweenPortsCurves(eDistanceType, dRouteStructMinDistance, dRouteStructMaxDistance)

                dRouteStructMaxDistance = dRouteStructMaxDistance - dDepOrRad
                dRouteStructMinDistance = dRouteStructMinDistance - dDepOrRad

                oFilteredSupports = SupportsByRangeQuery(eHgrDiscipline, oFilteredSupports, "MinAssemblyLength", "MaxAssemblyLength", dRouteStructMinDistance, dRouteStructMaxDistance)
                If (SupportHelper.PlacementType <> PlacementType.PlaceByReference) Then
                    'Filter supports based on Supporting Face (Only filters if StrictFaceSelection is on)
                    If ((SupportHelper.SupportingObjects.Count > 1) And (eHgrDiscipline = DisciplineTypeFlags.Combined)) Then
                        Dim oFilteredSupportsTemp As Collection(Of Part)
                        For index As Integer = 2 To SupportHelper.SupportingObjects.Count
                            RefPortHelper.DistanceBetweenSpecifiedPortsCurves(1, index, eDistanceType, dRouteStructMinDistance, dRouteStructMaxDistance)
                            If IsNothing(oFilteredSupportsTemp) Then
                                oFilteredSupportsTemp = SupportsByRangeQuery(eHgrDiscipline, oFilteredSupports, "MinAssemblyLength", "MaxAssemblyLength", dRouteStructMinDistance, dRouteStructMaxDistance)
                            Else
                                Dim oFilteredSupportsTemp1 As Collection(Of Part) = SupportsByRangeQuery(eHgrDiscipline, oFilteredSupports, "MinAssemblyLength", "MaxAssemblyLength", dRouteStructMinDistance, dRouteStructMaxDistance)
                                oFilteredSupportsTemp = New Collection(Of Part)(oFilteredSupportsTemp.Union(oFilteredSupportsTemp1).ToList())
                            End If
                        Next
                        oFilteredSupports = oFilteredSupportsTemp
                    End If

                    oFilteredSupports = SupportsBySupportingFace(1, oFilteredSupports)

                End If
            End If

            'Finally, order the collection based on how you want the user to see it in the drop-down menu.
            'If OrderBy is not called, then the they will be ordered alphabetically based on Part Description
            Try
                oFilteredSupports = OrderBy(oFilteredSupports, "IJUAHgrRulePriority", "RulePriority", OrderDirection.ASCENDING)
            Catch ex As CmnFailedToGetInterfacesException
                'This is an expected exception when the requested interface is not available in metadata. Don't do anything here.
            End Try


            'Return the final collection
            Return oFilteredSupports

        End Get
    End Property
    '---------------- This is a Temperory Method to Convert Inch / Mm/ M to DBU units i.e metres   
    Private Function ConvertToDBU(ByVal StrUnitName As String, ByVal dValue As Double) As Double

        Dim oUOMManager As Ingr.SP3D.Common.Middle.Services.UOMManager = MiddleServiceProvider.UOMMgr
        Dim strinch, strmm, strM As String
        ConvertToDBU = 0
        strinch = oUOMManager.FormatUnit(UnitType.Distance, 0, UnitName.DISTANCE_INCH)
        strmm = oUOMManager.FormatUnit(UnitType.Distance, 0, UnitName.DISTANCE_MILLIMETER)
        strM = oUOMManager.FormatUnit(UnitType.Distance, 0, UnitName.DISTANCE_METER)
        If (strinch.EndsWith(StrUnitName)) Then
            ConvertToDBU = oUOMManager.ConvertUnitToDBU(UnitType.Distance, dValue, UnitName.DISTANCE_INCH)
        ElseIf (strmm.EndsWith(StrUnitName)) Then
            ConvertToDBU = oUOMManager.ConvertUnitToDBU(UnitType.Distance, dValue, UnitName.DISTANCE_MILLIMETER)
        ElseIf (strmm.EndsWith(StrUnitName)) Then
            ConvertToDBU = oUOMManager.ConvertUnitToDBU(UnitType.Distance, dValue, UnitName.DISTANCE_METER)
        End If

    End Function
End Class
