Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    '**********************************************************************************
    ' Copyright (C) 2010, Intergraph Corporation. All rights reserved
    '
    ' File:        CustomAsmForReferenceDataMetadataHelper.vb
    '
    ' Description: This is a dummy symbol test used for performing bulkload testing
    '              that is used to test the ReferenceDataMetadataHelper class.  This
    '              symbol is only partially implemented to support bulkload and is
    '              is not intended for any other uses.

    ' Declare:
    '   IJUATstDtNetMetadataHlprSphere interface declared as output
    '   IJDGeometry interface declared as output
    <OutputNotification("{743E231E-6FC6-4e80-938E-CD02A166BEC4}")> _
    <OutputNotification("IJDGeometry")> _
    Public Class CustomAsmForReferenceDataMetadataHelperBaseClass : Inherits CustomAssemblyDefinition
    End Class

    ' Declare:
    '   non-cached symbol
    '   symbol version
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmForReferenceDataMetadataHelper : Inherits CustomAsmForReferenceDataMetadataHelperBaseClass
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
            Public m_diameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        Private Const CONST_IJUATstDtNetMetadataHlprSphere As String = "IJUATstDtNetMetadataHlprSphere"
        Private Const CONST_PropertyDiameter As String = "Diameter"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of metadatahelper")> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(2, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()

        End Sub
    End Class
End Namespace


