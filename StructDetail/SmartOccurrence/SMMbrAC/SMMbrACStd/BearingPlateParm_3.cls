VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BearingPlateParm_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2014-15 Intergraph Corporation. All rights reserved.
'
'File : BearingPlateParm_3.cls
'
'Author : Alligators
'
'Description :
'   Control Parameters of the Member Items for the MbrBearingEndSel Selector Class
'
'
'History:

'   26/11/2014   Created
'   02/02/2015   svsmylav DI-263487(v2014R1 eCustomer)/DM-265485(v2016):
'                         Added code to handle inclined bounded member cases.
' 24/Mar/15  MDT/RPK  TR-269306 Handrail placed on builtup member, cannot create assembly connection, errors,
'                    modified the ParameterRuleLogic related to design/Builtup member as the bounding member
'   14/July/2015 -knukala
'           TR-CP-270890  Cannot Place Generic AC when Handrail Post member is bounded to a plate
'********************************************************************
Option Explicit

Const m_sClassName As String = "BearingPlateParm_3"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices


'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    pOH.SetOutput "Offset"
    pOH.SetOutput "Radius"
    pOH.SetOutput "FlangeThickness"
    pOH.SetOutput "WebThickness"
    

  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Defining Parameter Values"
    
    '***   We need to set four parameters  ***
    
    'Offset and Radius parameters are retained here for future use (incase of different dimensions for different size of crosssection):
    'As of now, below code gets catalog values and applies the same
    
    Dim oAttributes             As IJDAttributes
    Dim oSmartOccurrence As IJSmartOccurrence
    
    Set oSmartOccurrence = pPRL.SmartOccurrence
    Set oAttributes = oSmartOccurrence
    
    '***   First parameter:
    Dim dOffset As Double
    dOffset = oAttributes.CollectionOfAttributes("IJUAStdEndCutCornerParm").Item("Offset").Value
    pPRL.Add "Offset", dOffset
    
    '***   Second parameter:
    Dim dRadius As Double
    dRadius = oAttributes.CollectionOfAttributes("IJUAStdEndCutShapeParm").Item("Radius").Value
    pPRL.Add "Radius", dRadius
    
    ' To Get the flange thickness and web thickness from member
    Dim oPortBounded As IJPort
    Dim oBoundedObject As Object
    
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    Set oBoundedObject = oPortBounded.Connectable
    
    Dim dFlangeThickness As Double
    Dim dWebThickness As Double
    
    Dim oBoundedMemberPart As New StructDetailObjects.MemberPart
    Set oBoundedMemberPart.object = oBoundedObject
    Dim oBoundedAsPartCommon As ISPSMemberPartCommon
    Set oBoundedAsPartCommon = oBoundedObject
    dFlangeThickness = oBoundedMemberPart.flangeThickness
    dWebThickness = oBoundedMemberPart.webThickness
    
    'Check for inclined bounded member cases so that the effective web/flange thickness could be measured along the bounding surface
        
    'Step 1: Get bounded member axis and bounded location
    Dim sACName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName pPRL.SmartOccurrence, sACName, oACObject

    Dim oAppConn As IJAppConnection
    Set oAppConn = oACObject
    
    Dim oBoundingPort As IJPort
    GetAssemblyConnectionInputs oAppConn, Nothing, oBoundingPort
        
    ' ------------------------------------
    ' Get information about the connection
    ' ------------------------------------
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oBoundingMemberPart As ISPSMemberPartCommon
    Dim lStatus As Long
    Dim oBoundedCSMatrix As IJDT4x4
    Set oBoundedCSMatrix = New DT4x4
    Dim oBoundingCSMatrix As IJDT4x4
    Set oBoundingCSMatrix = New DT4x4
    oBoundedAsPartCommon.Rotation.GetTransform oBoundedCSMatrix
    InitMemberConnectionData oACObject, oBoundedData, oBoundingData, lStatus, sMsg
    
    If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
        Set oBoundingMemberPart = oBoundingPort.Connectable
    Else
        'Assumed as plate part for now.
    End If
    
    Dim oBoundedAxis_Vector As IJDVector
    Set oBoundedAxis_Vector = New dVector
    If oBoundedData.ePortId = SPSMemberAxisStart Then
        oBoundedAxis_Vector.Set oBoundedCSMatrix.IndexValue(0), _
                                oBoundedCSMatrix.IndexValue(1), _
                                oBoundedCSMatrix.IndexValue(2)
    ElseIf oBoundedData.ePortId = SPSMemberAxisEnd Then
        oBoundedAxis_Vector.Set -oBoundedCSMatrix.IndexValue(0), _
                                -oBoundedCSMatrix.IndexValue(1), _
                                -oBoundedCSMatrix.IndexValue(2)
    Else
        Exit Sub
    End If
    Dim oBoundedLocation As IJDPosition
    Set oBoundedLocation = New DPosition
    
    oBoundedLocation.Set oBoundedCSMatrix.IndexValue(12), _
                            oBoundedCSMatrix.IndexValue(13), _
                                oBoundedCSMatrix.IndexValue(14)
    

    Dim oBrgPltPortNormal As IJDVector
    Dim oV_vector As IJDVector
    Dim oU_vector As IJDVector
    If Not oBoundingPort Is Nothing Then
        If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
            'We need to know whether the U-vector or V-vector of bounding member need to be used as
            'the bounding surface normal: we are considering angles greater than 45 degrees
            oBoundingMemberPart.Rotation.GetTransform oBoundingCSMatrix
            Set oV_vector = New dVector
            oV_vector.Set oBoundingCSMatrix.IndexValue(8), oBoundingCSMatrix.IndexValue(9), oBoundingCSMatrix.IndexValue(10)
            
            Dim dBddAxisAndBdgV_DotProduct As Double
            dBddAxisAndBdgV_DotProduct = Abs(oBoundedAxis_Vector.Dot(oV_vector))
            
            Set oU_vector = New dVector
            oU_vector.Set oBoundingCSMatrix.IndexValue(4), oBoundingCSMatrix.IndexValue(5), oBoundingCSMatrix.IndexValue(6)
            
            Dim dBddAxisAndBdgU_DotProduct As Double
            dBddAxisAndBdgU_DotProduct = Abs(oBoundedAxis_Vector.Dot(oU_vector))
            
            Dim bIsWebBoundingSurface As Boolean
            If GreaterThan(dBddAxisAndBdgU_DotProduct, dBddAxisAndBdgV_DotProduct) Then
                bIsWebBoundingSurface = True
            Else
                bIsWebBoundingSurface = False
            End If
            
            'Considering bearing plate normal to be the same as that of the bounding surface normal: here,
            ' bIsWebBoundingSurface True/False values would correspond to use oU_vector/oV_vector of the bounding respectively
            If bIsWebBoundingSurface Then
                Set oBrgPltPortNormal = oU_vector '(of the bounding)
            Else
                Set oBrgPltPortNormal = oV_vector '(of the bounding)
            End If
            
        ElseIf (TypeOf oBoundingPort.Connectable Is SPSSlabEntity) Or (TypeOf oBoundingPort.Connectable Is SPSWallPart) _
                    Or (TypeOf oBoundingPort.Connectable Is IJPlate) Then
            If TypeOf oBoundingPort.Geometry Is IJSurfaceBody Then
                Dim oSurfaceBdy As IJSurfaceBody
                Set oSurfaceBdy = oBoundingPort.Geometry
                oSurfaceBdy.GetNormalFromPosition oBoundedLocation, oBrgPltPortNormal
            End If
        Else
            sMsg = "This Case is yet not handled"
            GoTo ErrorHandler
        End If
                        
        'Check if the bounded axis and the bearing plate normal are parallel (no need to adjust flange/web thickness),
        'otherwise compute these dimensions as measured in bearing plate plane
        Dim dDotProduct As Double
        dDotProduct = Abs(oBoundedAxis_Vector.Dot(oBrgPltPortNormal))
        
        If GreaterThan(dDotProduct, Cos(GetPI / 2 - degreeToRadian(89))) Then
            'Angle of bounded member axis is around 89 degrees or more with respect to bearing plate surface, so
            'use Flange/Web thickness values as is which is at bottom of this 'If' block
        ElseIf GreaterThan(dDotProduct, Cos(GetPI / 2 - degreeToRadian(20))) Then
            'Angle of bounded member axis is between 20 and 89 degrees (approximately)

            '***   Third parameter (adjusted flange thickness value):
            Set oV_vector = New dVector
            '*** Bounded V vector ***:
            oV_vector.Set oBoundedCSMatrix.IndexValue(8), oBoundedCSMatrix.IndexValue(9), oBoundedCSMatrix.IndexValue(10)
            
            Dim oResultantVec As IJDVector
            Set oResultantVec = oV_vector.Cross(oBrgPltPortNormal)
            oResultantVec.Length = 1#
            
            'Get a unit vector which is along projection of V-Vector on to bearing plate surface:
            Dim oProjectedVvecOnBrgPlt As IJDVector
            Set oProjectedVvecOnBrgPlt = oBrgPltPortNormal.Cross(oResultantVec)
            oProjectedVvecOnBrgPlt.Length = 1#
            
            dDotProduct = Abs(oProjectedVvecOnBrgPlt.Dot(oV_vector))
            dFlangeThickness = dFlangeThickness / dDotProduct
            
            '***   Fourth parameter (adjusted web thickness value):
            Set oU_vector = New dVector
            '*** Bounded U vector ***:
            oU_vector.Set oBoundedCSMatrix.IndexValue(4), oBoundedCSMatrix.IndexValue(5), oBoundedCSMatrix.IndexValue(6)

            Set oResultantVec = oU_vector.Cross(oBrgPltPortNormal)
            oResultantVec.Length = 1#

            'Get a unit vector which is along projection of U-Vector on to bearing plate surface:
            Dim oProjectedUvecOnBrgPlt As IJDVector
            Set oProjectedUvecOnBrgPlt = oBrgPltPortNormal.Cross(oResultantVec)
            oProjectedVvecOnBrgPlt.Length = 1#

            dDotProduct = Abs(oProjectedUvecOnBrgPlt.Dot(oU_vector))
            dWebThickness = dWebThickness / dDotProduct
        Else
            sMsg = "Bounded member inclination is less than limit 20 degrees (these cases are handled)"
            GoTo ErrorHandler
        End If
    End If

    '***   Third parameter:
    pPRL.Add "FlangeThickness", dFlangeThickness
    
    '***   Fourth parameter:
    pPRL.Add "WebThickness", dWebThickness

  Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

' ********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation

  Dim pDFact As New DefinitionFactory
  Dim pIH As IJDInputsHelper
  Dim pOH As IJDOutputsHelper
  
  pDFact.InitAbstractParameterRule pPR
  
  Set pIH = New InputHelper
  pIH.definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub


' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


