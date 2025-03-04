''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright 1992 - 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'File
'  BlockSlabFndAsmDef.vb
'
'Abstract
'	BlockSlabFndAsmDef is a .NET custom assembly definition which creates a block and a slab component in the model.
'   This class subclasses from EquipmentFoundationCustomAssemblyDefinition.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Imports System
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.Structure.Middle.Services
Imports Ingr.SP3D.Structure.Exceptions

'===========================================================================================
'Namespace of this class is Ingr.SP3D.Content.Structure
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'===========================================================================================
Namespace Ingr.SP3D.Content.Structure

    ''' <summary>
    ''' Block and slab foundation CustomAssemblyDefinition.
    ''' It will update the IJUASPSBlockAndSlabFndnAsm interface for this definition and 
    ''' the assembly outputs will have their geometry (IJDGeometry), attributes (IJDAttributes) and material (IJStructMaterial) modified.
    ''' </summary>
    <SymbolVersion("1.1.0.0")> _
    <CacheOption(CacheOptionType.Cached)> _
    <OutputNotification(SPSSymbolConstants.IID_IJUASPSBlockAndSlabFndnAsm)> _
    <OutputNotification(SPSSymbolConstants.IID_IJDAttributes, True)> _
    <OutputNotification(SPSSymbolConstants.IID_IJStructMaterial, True)> _
    <OutputNotification(SPSSymbolConstants.IID_IJDGeometry, True)> _
    Public Class BlockSlabFndAsmDef : Inherits EquipmentFoundationCustomAssemblyDefinition

        '============================================================================================================================
        'DefinitionName/ProgID of this symbol is "EquipmentFoundationCustomAssemblies,Ingr.SP3D.Content.Structure.BlockSlabFndAsmDef"
        '============================================================================================================================

#Region "Definition of Inputs"

        'Following inputs are needed to create a block and a slab component, modification of these inputs will trigger the re-computation of block and slab component.
        <InputCatalogPart(1)> _
        Public m_oCatalogPart As InputCatalogPart

#End Region

#Region "Definitions of Aspects and their outputs"

        'the .NET custom assembly definition's assembly output will be re-evaluated on material, geometry and attribute changes.        
        <AssemblyOutput(1, SPSSymbolConstants.Block)> _
        Public m_oBlockAssemblyOutput As AssemblyOutput
        <AssemblyOutput(2, SPSSymbolConstants.Slab)> _
        Public m_oSlabAssemblyOutput As AssemblyOutput

#End Region

#Region "Public Override Functions and Methods"

        ''' <summary>
        ''' Evaluate the equipment foundation assembly.
        ''' Here we will do the following:
        ''' 1. Decide which assembly outputs are needed. 
        ''' 2. Create the ones which are needed, delete which are not needed now. 
        ''' 3. Evaluate the assembly to set the occurrence matrix on assembly. 
        ''' 4. Evaluate each component to set the occurrence matrix on the component and set its dimensions.
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            Try

                'The occurrence is equipment foundation. 
                Dim equipmentFoundation As EquipmentFoundation = DirectCast(Occurrence, EquipmentFoundation)

                Dim blockComponent As FoundationComponent = Nothing
                Dim slabComponent As FoundationComponent = Nothing

                'Construct the output objects for the assembly. Note that the size will be set later.
                'First the Block. Only construct the block if not generated yet and add it is as output
                If m_oBlockAssemblyOutput.Output Is Nothing Then
                    blockComponent = FoundationServices.CreateComponent(equipmentFoundation, FoundationComponentType.Block)
                    m_oBlockAssemblyOutput.Output = blockComponent
                Else
                    blockComponent = DirectCast(m_oBlockAssemblyOutput.Output, FoundationComponent)
                End If

                'Now construct the slab if required. Only construct the slab if not generated yet and add it as output
                Dim isSlabRequired As Boolean = StructHelper.GetBoolProperty(equipmentFoundation, SPSSymbolConstants.IJUASPSBlockAndSlabFndnAsm, SPSSymbolConstants.WithSlab)
                If isSlabRequired Then
                    If m_oSlabAssemblyOutput.Output Is Nothing Then
                        slabComponent = FoundationServices.CreateComponent(equipmentFoundation, FoundationComponentType.Slab)
                        m_oSlabAssemblyOutput.Output = slabComponent
                    Else
                        slabComponent = DirectCast(m_oSlabAssemblyOutput.Output, FoundationComponent)
                    End If
                Else
                    ' iF slab is not require dnow, delete it if it has been previously created
                    If Not m_oSlabAssemblyOutput.Output Is Nothing Then
                        m_oSlabAssemblyOutput.Output.Delete()
                    End If
                End If

                'Evaluate the Equipment foundation to set its occurrence matrix
                Dim length#, width#, height#
                MyBase.EvaluateFoundation(equipmentFoundation, width, length, height)

                'Now the evaluate cals
                'Evaluate the block to set the origin and orientation based on its supported/supporting objects
                'Based on these dimensions, appropriate foundation geometry needs to be constructed later in the component definition.
                MyBase.EvaluateFoundation(blockComponent, equipmentFoundation, width, length, height)

                'Now set the custom properties on the block
                'Get WithSlab property from the custom assembly parent -equipment foundation
                Dim hasSlab As Boolean = StructHelper.GetBoolProperty(equipmentFoundation, SPSSymbolConstants.IJUASPSBlockAndSlabFndnAsm, SPSSymbolConstants.WithSlab)
                Dim isSupportedBySurface As Boolean = MyBase.IsBoundedBySurface(equipmentFoundation)
                'For block, need to clip by supporting surface if this block is a custom assembly output and has no slab under it.
                Dim clipBySupportingSurface As Boolean = isSupportedBySurface And (Not hasSlab)
                Dim placedByPoint As Boolean = MyBase.IsPlacedByPoint(equipmentFoundation)
                FoundationServices.SetDimensionalProperties(blockComponent, FoundationComponentType.Block, length, width, height, placedByPoint, isSupportedBySurface, hasSlab)

                If Not slabComponent Is Nothing Then
                    'Evaluate the slab to set the origin and orientation based on its supported/supporting objects
                    'Based on these dimensions, appropriate foundation geometry needs to be constructed later in the component definition.
                    MyBase.EvaluateFoundation(slabComponent, equipmentFoundation, width, length, height)

                    'Now set the custom properties on the foundation
                    FoundationServices.SetDimensionalProperties(slabComponent, FoundationComponentType.Slab, length, width, height, placedByPoint, isSupportedBySurface)
                End If

            Catch ex As Exception
                If MyBase.ToDoListMessage Is Nothing Then 'Check ToDoListMessgae created already or not
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrEvaluateAssembly, "Error while evaluating Equipment Foundation Assembly in Block Slab Foundation Assembly Definition. Please check your custom code or contact S3D support."))
                End If
            End Try

        End Sub

#End Region

    End Class

End Namespace
