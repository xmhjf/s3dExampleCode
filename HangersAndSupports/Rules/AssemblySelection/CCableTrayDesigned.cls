VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCableTrayDesigned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CCombinedSupport.cls
'   ProgID:         HgrAssmSelRule.CCombinedSupport
'   Author:         Yan Ji
'   Creation Date:  12.Mar.2003
'   Description:
'    AssemblySelectionRule for CombinedSupport.
'    This ASR will be used to select Assembly based on SupportDiscipline
'
'   Change History:
'       23.Aug.2004             Yan Ji          Creation Date
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CCableTrayDesigned" 'Used for error messages

Implements IJHgrAssmSelectionRule

Public Function IJHgrAssmSelectionRule_GetAssemblies(ByVal pInputObjectInfo As Object) As HNGSUPPartSelectionHlpr.IJElements
Const METHOD = "IJHgrAssmSelectionRule_GetAssemblies"

    On Error GoTo ErrorHandler
    
    'Set the discipline type to be used while calling the services.
    Dim eHgrDiscipline As GSCADHgrSupFilterServices.HgrDisciplineType
    eHgrDiscipline = HgrCabletrayDesignedDisciplineType
    
    'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputObjectInfo As IJHgrInputObjectInfo
    Set my_IJHgrInputObjectInfo = pInputObjectInfo
    
    'Get all pipes on the support.
    Dim PipeCollection As Object
    On Error Resume Next
    Set PipeCollection = my_IJHgrInputObjectInfo.GetSupportedObjects
    On Error GoTo ErrorHandler
    
    'If there is no Pipes Collection exit.
    If PipeCollection Is Nothing Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If
    
    'If there are no Pipes in the returned collection exit.
    If PipeCollection.Count = 0 Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If
    
    '*********************************************************************************************
    'Filter based on Pipe Size, command type, discipline type
    '*********************************************************************************************
    Dim AssyCollFromPipeDia As Object
    Set AssyCollFromPipeDia = Nothing
    Dim AssmSelectionHelper As New AssySelectionHlpr

    Dim strFeatureType As String
    Dim varFeature() As Variant
'     'Get a Collection of catalog Parts based on PipeFeature
    varFeature = my_IJHgrInputObjectInfo.GetAttributeByRule("HgrSupFeatureType")
    strFeatureType = varFeature(0)

    Dim PartCollHlpr                As New HgrCatPartsCollHlpr
    Dim lngCommandType              As Long

    lngCommandType = my_IJHgrInputObjectInfo.GetCommandType

   '***************************************************************************************
   'Get the Collection of Assemblies based on the input count.
   ' Supporting count and Supported count
   '***************************************************************************************
    Dim AssemblyMergedCollection As Object

   'Get all structures on the support.  In particular we are looking for the
   'port on the input object that is being used by the support.
    Dim StructPortCollection As Object
    Set StructPortCollection = my_IJHgrInputObjectInfo.GetSupportingObjectPorts

    Dim lngSupportingCnt As Long
    If StructPortCollection Is Nothing Then
        lngSupportingCnt = 0
    Else
        lngSupportingCnt = StructPortCollection.Count
    End If

    Dim AssyFromPipeAndBeam As Object
    Set AssyFromPipeAndBeam = AssmSelectionHelper.GetAssembliesByInput(lngCommandType, eHgrDiscipline, strFeatureType, PipeCollection.Count, lngSupportingCnt)
    Set AssemblyMergedCollection = AssyFromPipeAndBeam 'PartCollHlpr.GetInterferencePartsColl(AssyFromPipeAndBeam, AssyCollFromPipeDia)

    ' TR#147246 ; check whether new type of design support (LCS similar to Std Support) is available or not
    ' if available, place DS of the new type. if not, then place DS of old type
    Dim AssyCollNew As Object
    Set AssyCollNew = AssmSelectionHelper.GetAssyCollByOperator("AssmInfoRule", "HgrSupDesignSuppAIR.GenericDesignAssm2", OPER_EQUAL, eHgrDiscipline)
    If Not AssyCollNew Is Nothing Then
        If AssyCollNew.Count > 0 Then
            Set AssemblyMergedCollection = PartCollHlpr.GetInterferencePartsColl(AssyFromPipeAndBeam, AssyCollNew)
        End If
        Set AssyCollNew = Nothing
    End If
    ' TR#147246 ends here

    ' Return the Collection of Assembly Definitions
    Set IJHgrAssmSelectionRule_GetAssemblies = AssemblyMergedCollection
    
    'AssyCollFromPipeDia.Clear
    'Set AssyCollFromPipeDia = Nothing
    'AssyFromPipeAndBeam.Clear
    Set AssyFromPipeAndBeam = Nothing
        
    Set AssmSelectionHelper = Nothing
    Set PartCollHlpr = Nothing
    Set AssemblyMergedCollection = Nothing
    Set StructPortCollection = Nothing
'    Set HgrComServ = Nothing
    Set PipeCollection = Nothing
    Set my_IJHgrInputObjectInfo = Nothing
    
    Exit Function
    
ErrorHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function








