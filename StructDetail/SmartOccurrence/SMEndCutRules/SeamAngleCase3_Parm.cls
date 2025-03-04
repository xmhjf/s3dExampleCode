VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamAnglecase3_Parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "EndCutRules.SeamAngleCase3_Parm"
Const m_ParameterRuleName As String = CUSTOMERID + "EndCutRules.SeamAngleCase3_Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\SeamAngleCase3_Parm.cls"

Implements IJDUserSymbolServices
Implements IJEndCutDrawingType

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "CutDepth"
  pOH.SetOutput "WebSeamAngle"
  pOH.SetOutput "WebSeamPointOffset"
  pOH.SetOutput "FlangeSeamAngle"
  pOH.SetOutput "CutVang"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler

    Dim oPortBounded As IJPort
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
     
    'Get Bounded part CutDepth
    Dim dCuttingDepth As Double
    dCuttingDepth = GetBoundedCutDepth(oPortBounded, True)
    
    'Get Cut Angle in the Flange Plane (Sketching Plane V axis)
    Dim dCutVang As Double
    GetProjectionCutAngle pPRL, dCutVang
    
    ' get the part number;
    Dim sPartNumber As String
    'TR-169311 Trying to retrieve question's answer without
    'hardcoding the path
    GetSelectorAnswer pPRL, gsWeldPart, sPartNumber
    
    ' want the WebCut to also Cut the Flange
    ' set Cutting Depth based on Flange Length
    dCuttingDepth = dCuttingDepth * 3#
    pPRL.Add "CutDepth", dCuttingDepth
    
    pPRL.Add "WebSeamPointOffset", 0.00001
    
    'Second Symbol contains angles measured opposite of First Symbol
    'Therefore Input values are same for both First, Second Symbols
    Dim dPI As Double
    dPI = Atn(1#) * 4

    If sPartNumber = "First" Then
        pPRL.Add "WebSeamAngle", dPI / 3#
        pPRL.Add "FlangeSeamAngle", dPI / 2#
        pPRL.Add "CutVang", dCutVang
    Else
        pPRL.Add "WebSeamAngle", dPI / 3#
        pPRL.Add "FlangeSeamAngle", dPI / 2#
        pPRL.Add "CutVang", dCutVang
    End If
            
   
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
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
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation
  
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractParameterRule pPR
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
'**********************************************************
' Method Name : IJEndCutDrawingType_UpdateDrawingType
' Inputs      : feature occurence object
' Output      : string, returns the end cut type.
' Description : this method checks for the type of endcut, straight or sniped and offset or no-offset
'                and assigns the corresponding enum to a string variable.
'*********************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    Dim sDrawingtype As String
    Dim dOffset As Double
    Dim dAngle As Double
    
    dOffset = GetCustomAttribute(pSmartOccurrence, "IJUAAngledEndCuts", "WebSeamPointOffset")
    dAngle = GetCustomAttribute(pSmartOccurrence, "IJUAAngledEndCuts", "WebSeamAngle")
     
    If GreaterThan(dOffset, 0.004) Then
        ' offset
            If IsValueBetween(dAngle, 0, degreeToRadian(89.99)) Or IsValueBetween(dAngle, degreeToRadian(90.01), degreeToRadian(180)) Then
                sDrawingtype = Sniped_Offset_WebCuts
            Else
                sDrawingtype = Straight_Offset_WebCuts
            End If
        Else
            If IsValueBetween(dAngle, 0, degreeToRadian(89.99)) Or IsValueBetween(dAngle, degreeToRadian(90.01), degreeToRadian(180)) Then
                sDrawingtype = Sniped_No_Offset_WebCuts
            Else
                sDrawingtype = Straight_No_Offset_WebCuts
            End If
    End If
    
    
    
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.Primary, sDrawingtype
    IJEndCutDrawingType_UpdateDrawingType = sDrawingtype
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
Private Sub GetProjectionCutAngle(pPRL As IJDParameterLogic, _
                                  dCutVang As Double)
  On Error GoTo ErrorHandler
  
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    Dim dDot As Double
    
    Dim oMajorAxis As AutoMath.DVector
    Dim oAngNormal As AutoMath.DVector
    Dim oTopNormal As AutoMath.DVector
    Dim oWebNormal As AutoMath.DVector
    
    Dim oNothing As Object
    Dim oWebPlane As IJPlane
    Dim oTopPlane As IJPlane
    Dim oPosition As IJDPosition
    Dim oBoundedPort As IJPort
    Dim oBoundedObject As Object
    
    Dim eSDO_ObjectType As sdwObjectType
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDO_BeamPart As StructDetailObjects.BeamPart
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
     
    Set oBoundedPort = pPRL.InputObject(INPUT_BOUNDED)
    Set oBoundedObject = oBoundedPort.Connectable
    Set oSDO_Helper = New StructDetailObjects.Helper
    eSDO_ObjectType = oSDO_Helper.ObjectType(oBoundedObject)
     
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pPRL.SmartOccurrence
    Set oPosition = oSDO_WebCut.BoundedLocation
     
    'Get Bounded part
    If eSDO_ObjectType = SDOBJECT_STIFFENER Then
        Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart.object = oBoundedObject
        Set oTopPlane = oSDO_ProfilePart.EdgePlane(JXSEC_TOP, oPosition, _
                                                   oNothing)
        Set oWebPlane = oSDO_ProfilePart.EdgePlane(JXSEC_WEB_LEFT, oPosition, _
                                                   oNothing)
    ElseIf eSDO_ObjectType = SDOBJECT_BEAM Then
        Set oSDO_BeamPart = New StructDetailObjects.BeamPart
        Set oSDO_BeamPart.object = oBoundedObject
        Set oTopPlane = oSDO_BeamPart.EdgePlane(JXSEC_TOP, oPosition, _
                                                oNothing)
        Set oWebPlane = oSDO_BeamPart.EdgePlane(JXSEC_WEB_LEFT, oPosition, _
                                                oNothing)
    Else
        dCutVang = 0.00001
        Exit Sub
    End If
    
    ' Determine the Major Axis of the Web Plane
    ' the Projection Cut angle is based on the Web Plane Major Axis
    Set oWebNormal = New AutoMath.DVector
    oWebPlane.GetNormal dX, dY, dZ
    oWebNormal.Set dX, dY, dZ
    oWebNormal.Length = 1#
    oWebNormal.Get dX, dY, dZ
    
    Set oMajorAxis = New AutoMath.DVector
    If Abs(dX) > Abs(dY) Then
        If Abs(dX) > Abs(dZ) Then
            ' Major axis is X axis
            oMajorAxis.Set 1#, 0#, 0#
        Else
            ' Major axis is Z axis
            oMajorAxis.Set 0#, 0#, 1#
        End If
        
    ElseIf Abs(dY) > Abs(dZ) Then
        ' Major axis is Y axis
            oMajorAxis.Set 0#, 1#, 0#
        
    ElseIf Abs(dZ) > Abs(dX) Then
        ' Major axis is Z axis
            oMajorAxis.Set 0#, 0#, 1#
    
    Else
        ' No Major Axis: two or more axis are (almost) equal
        dCutVang = 0.00001
        Exit Sub
    End If
    
    ' check if Web Normal is Parallel to the Major Axis
    oMajorAxis.Length = 1#
    dDot = oMajorAxis.Dot(oWebNormal)
    If (1# - Abs(dDot)) < 0.000001 Then
        dCutVang = 0.00001
        Exit Sub
    ElseIf dDot < 0.000001 Then
        oMajorAxis.Length = -1#
    End If
    
    ' Get the Top Plane Normal: the Plane that the angle is calculated in
    Set oTopNormal = New AutoMath.DVector
    oTopPlane.GetNormal dX, dY, dZ
    oTopNormal.Set dX, dY, dZ
    oTopNormal.Length = 1#
    
    ' Project the Major Axis into the Top Plane
    Set oAngNormal = New AutoMath.DVector
    dDot = oMajorAxis.Dot(oTopNormal)
    oAngNormal.Set (oMajorAxis.x - (oTopNormal.x * dDot)), _
                   (oMajorAxis.y - (oTopNormal.y * dDot)), _
                   (oMajorAxis.z - (oTopNormal.z * dDot))
    
    ' Calculate the Angle between the Web Normal and the Major Axis
    dCutVang = oWebNormal.angle(oAngNormal, oTopNormal)
    If dCutVang > (Atn(1) * 4#) Then
        dCutVang = dCutVang - (Atn(1) * 8#)
    End If

  Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "GetProjectionCutAngle").Number
End Sub
