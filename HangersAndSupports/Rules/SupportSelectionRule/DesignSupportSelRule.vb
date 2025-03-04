''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Copyright (C) 2010, Intergraph PPO. All rights reserved.
'File
'  DesignSupportSelRule.vb
'Abstract
'	This is an example .NET support selection rule.
'05.Dec.2012     SVP    Modified the selection rule to support VB support ProgID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle.Services

'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Support.Content.Rules
'It is recommended that customers specify namespace of their classes to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this rule to suit their
'requirements, they should change the namespace/class name so the identity of the modified
'rule will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------

Public Class DesignSupportSelRule
    Inherits SupportSelectionRule

    Public Overrides ReadOnly Property Supports() As System.Collections.ObjectModel.Collection(Of ReferenceData.Middle.Part)
        Get
            'Define a collection to hold all the supports
            Dim oFilteredSupports As Collection(Of Part)
            'The discipline type for which this rule applies
            Dim eHgrDiscipline As DisciplineTypeFlags

            'Get useful objects
            Dim oSupportedInfo = SupportedHelper.SupportedObjectInfo(1)     'Use the primary supported object
            Dim soTypes(3) As SupportedObjectType, DiscFlags(3) As DisciplineTypeFlags
            soTypes(0) = SupportedObjectType.Pipe : DiscFlags(0) = DisciplineTypeFlags.Piping_Designed
            soTypes(1) = SupportedObjectType.Conduit : DiscFlags(1) = DisciplineTypeFlags.Conduit_Designed
            soTypes(2) = SupportedObjectType.HVAC : DiscFlags(2) = DisciplineTypeFlags.HVAC_Designed
            soTypes(3) = SupportedObjectType.CableTray : DiscFlags(3) = DisciplineTypeFlags.CableTray_Designed

            'Determine the discipine based on the supported object types
            Dim oSupportedObject As SupportedObjectInfo

            For i As Integer = 0 To 3
                If (oSupportedInfo.SupportedObjectType = soTypes(i)) Then
                    eHgrDiscipline = DiscFlags(i)
                    For Each oSupportedObject In SupportedHelper.SupportedObjectInfo(SupportedObjectType.All, FeatureType.All)
                        If oSupportedObject.SupportedObjectType <> soTypes(i) Then
                            eHgrDiscipline = DisciplineTypeFlags.Combined_Designed
                            Exit For
                        End If
                    Next
                End If
            Next

            'Filter supports based on input (SupportedFamily, Supported/Supporting Count, CommandType, etc...)
            oFilteredSupports = SupportsByInput(eHgrDiscipline, Nothing)

            'Filter supports based on designed support AssmInfoRule
            oFilteredSupports = SupportsByCriteria(eHgrDiscipline, oFilteredSupports, "AssmInfoRule", ComparisonOperatorType.EQUAL, "DesignSupport,Ingr.SP3D.Content.Support.Rules.GenericDesignAssm2")

            If (oFilteredSupports.Count = 0) Then
                oFilteredSupports = SupportsByInput(eHgrDiscipline, Nothing)
                oFilteredSupports = SupportsByCriteria(eHgrDiscipline, oFilteredSupports, "AssmInfoRule", ComparisonOperatorType.EQUAL, "HgrSupDesignSuppAIR.GenericDesignAssm2")
            End If

            'Return the final collection
            Return oFilteredSupports

        End Get
    End Property
End Class
