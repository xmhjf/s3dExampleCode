VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectStiffener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------
'  Copyright (C) 2003, Intergraph Corporation.  All rights reserved.
'
'  FILE:  ConnectRules.cls
'
'  DESCRIPTION:
'  Rule to compute the keypoint on the connecting entity that the tripping
'  stiffener should connect to and the offsets from that keypoint
'
'  AUTHOR:  Dan Pizzillo
'
'  HISTORY
'  030916    (DMP) created
'-------------------------------------------------------------------------------

Option Explicit
Implements IJStiffenerBetweenElementsRule


Private dCachedTSArea As Double
Private dCachedTSWebThickness As Double
Private dCachedTSFlangeThickness As Double
Private dCachedTSDiameter As Double
Private dCachedTSWidth As Double
Private dCachedTSDepth As Double
Private sCachedTSCrossSectionType As String
Private sCachedTSCrossSectionSize As String
Private dCachedCEArea As Double
Private dCachedCEWebThickness As Double
Private dCachedCEFlangeThickness As Double
Private dCachedCEDiameter As Double
Private dCachedCEWidth As Double
Private dCachedCEDepth As Double
Private sCachedCECrossSectionType As String
Private sCachedCECrossSectionSize As String
Private m_eCEMountingFace As StiffenerRuleLib.ProfileFaceName


Private Sub IJStiffenerBetweenElementsRule_ComputeEndPointOffsets _
        (ByVal sTSCrossSectionType As String, _
        ByVal sTSCrossSectionSize As String, _
        ByVal eTSMountingFace As StiffenerRuleLib.ProfileFaceName, _
        ByVal lTSMountingPoint As Long, _
        ByVal sCECrossSectionType As String, _
        ByVal sCECrossSectionSize As String, _
        ByVal bCEIsStiffener As Boolean, _
        ByVal eCEMountingFace As StiffenerRuleLib.ProfileFaceName, _
        ByVal lCEMountingPoint As Long, _
        ByVal eConnectType As StiffenerRuleLib.LandingCrvAttachmentMethod, _
        ByVal dApproachAngle As Double, _
        ByVal pCE As Object, _
        ByVal TSSecOrientation As StructMoldedOrientation, _
        plAttachmentLoadPoint As Long, _
        pdUOffset As Double, _
        pdVOffset As Double)

Const METHOD = "IJStiffenerBetweenElementsRule_ComputeEndPointOffsets"

Dim dRadius As Double

On Error GoTo ErrorHandler

'Initialize outputs
plAttachmentLoadPoint = 1
pdUOffset = 0#
pdVOffset = 0#

If dApproachAngle < 0# Or dApproachAngle >= 360# Then
  Err.Raise E_INVALIDARG
End If

m_eCEMountingFace = eCEMountingFace 'used by CalcFilletOffset

RefreshCache sTSCrossSectionType, sTSCrossSectionSize, eTSMountingFace, sCECrossSectionType, _
  sCECrossSectionSize, eCEMountingFace

dRadius = 0.5 * dCachedCEDiameter  'get the radius for circular profiles


'''''''''''''''Attached Connection''''''''''''''''''''''''
If bCEIsStiffener And ((eConnectType = LCA_ATTACHED) Or (eConnectType = LCA_TWISTED)) Then
  Select Case sCECrossSectionType
    Case "B", "EA", "UA", "BUTL3", "B_TSN"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            plAttachmentLoadPoint = 11
            pdUOffset = CalcFilletOffset()
          ElseIf dApproachAngle < 90# Then
            plAttachmentLoadPoint = 11
            pdVOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 270# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If
        
        Case TopSideOfSuperiorFlange
          If dApproachAngle < 90# Or dApproachAngle >= 270# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 270# Then
            plAttachmentLoadPoint = 3
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case LeftWeb
          If dApproachAngle < 180# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 360# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case RightSideOfSuperiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            plAttachmentLoadPoint = 11
            pdUOffset = CalcFilletOffset()
          ElseIf dApproachAngle < 90# Then
            plAttachmentLoadPoint = 11
            pdVOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 180# Then
            plAttachmentLoadPoint = 3
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = lCEMountingPoint
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case Else
          Err.Raise E_INVALIDARG
    
      End Select
      
    Case "BUT", "BUTL2", "TSType", "T_XType"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            Select Case lCEMountingPoint
              Case 3
                plAttachmentLoadPoint = 12
              Case 23
                plAttachmentLoadPoint = 14
              Case 1
                plAttachmentLoadPoint = 13
            End Select
          ElseIf dApproachAngle < 90# Then
            plAttachmentLoadPoint = 11
          ElseIf dApproachAngle < 270# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If
        
        Case TopSideOfSuperiorFlange
          If dApproachAngle < 90# Or dApproachAngle >= 270# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 270# Then
            Select Case lCEMountingPoint
              Case 13
                plAttachmentLoadPoint = 1
              Case Else
                plAttachmentLoadPoint = 3
            End Select
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case LeftSideOfSuperiorFlange
          If dApproachAngle < 180# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 360# Then
            Select Case lCEMountingPoint
              Case 11
                plAttachmentLoadPoint = 15
              Case 9
                plAttachmentLoadPoint = 17
            End Select
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case Else
          Err.Raise E_INVALIDARG
    
      End Select
      
    Case "I", "ISType"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            Select Case lCEMountingPoint
              Case 3
                plAttachmentLoadPoint = 11
                pdUOffset = CalcFilletOffset()
              Case 2
                plAttachmentLoadPoint = 12
              Case 1
                plAttachmentLoadPoint = 13
            End Select
          ElseIf dApproachAngle < 90# Then
            plAttachmentLoadPoint = 11
          ElseIf dApproachAngle < 270# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case LeftSideOfSuperiorFlange
          If dApproachAngle < 180# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 225# Then
            plAttachmentLoadPoint = 23
            pdUOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 270# Then
            Select Case lCEMountingPoint
              Case 11
                plAttachmentLoadPoint = 15
              Case 9
                plAttachmentLoadPoint = 17
              Case 25
                plAttachmentLoadPoint = 23
            End Select
          ElseIf dApproachAngle < 315# Then
            Select Case lCEMountingPoint
              Case 9
                plAttachmentLoadPoint = 17
              Case Else
                plAttachmentLoadPoint = 15
            End Select
          ElseIf dApproachAngle < 360# Then
            plAttachmentLoadPoint = 15
            pdUOffset = -1 * CalcFilletOffset()
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case Else
          Err.Raise E_INVALIDARG
    
      End Select
      
    Case "CSType", "C_SS"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            plAttachmentLoadPoint = 11
            pdUOffset = CalcFilletOffset()
          ElseIf dApproachAngle < 90# Then
            plAttachmentLoadPoint = 11
            pdVOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 270# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case LeftWeb
          If dApproachAngle < 180# Then
            plAttachmentLoadPoint = lCEMountingPoint
          ElseIf dApproachAngle < 225# Then
            plAttachmentLoadPoint = 23
            pdUOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 270# Then
            Select Case lCEMountingPoint
              Case 11
                plAttachmentLoadPoint = 15
              Case 9
                plAttachmentLoadPoint = 17
              Case 25
                plAttachmentLoadPoint = 23
            End Select
          ElseIf dApproachAngle < 315# Then
            Select Case lCEMountingPoint
              Case 9
                plAttachmentLoadPoint = 17
              Case Else
                plAttachmentLoadPoint = 15
            End Select
          ElseIf dApproachAngle < 360# Then
            plAttachmentLoadPoint = 15
            pdUOffset = -1 * CalcFilletOffset()
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case RightSideOfSuperiorFlange
          If dApproachAngle < 45# Then
            plAttachmentLoadPoint = 11
            pdUOffset = CalcFilletOffset()
          ElseIf dApproachAngle < 135# Then
            Select Case lCEMountingPoint
              Case 15
                plAttachmentLoadPoint = 11
                pdVOffset = -1 * CalcFilletOffset()
              Case 17
                plAttachmentLoadPoint = 9
              Case 26
                plAttachmentLoadPoint = 25
            End Select
          ElseIf dApproachAngle < 180# Then
            plAttachmentLoadPoint = 3
            pdUOffset = CalcFilletOffset()
          ElseIf dApproachAngle < 360# Then
            plAttachmentLoadPoint = lCEMountingPoint
          Else
            Err.Raise E_INVALIDARG
          End If
          
        Case Else
          Err.Raise E_INVALIDARG
    
      End Select
      
    Case "FB"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 90# Or dApproachAngle >= 270# Then
            Select Case lCEMountingPoint
              Case 3
                plAttachmentLoadPoint = 11
              Case 1
                plAttachmentLoadPoint = 13
              Case 23
                plAttachmentLoadPoint = 15
            End Select
          Else
            plAttachmentLoadPoint = lCEMountingPoint
          End If
        Case LeftWeb
            'DM331715  / TR310420  (ER toward plate case) & TR214913
            If dApproachAngle < 45# Or dApproachAngle >= 315# Then
                plAttachmentLoadPoint = 11
            ElseIf dApproachAngle >= 45# And dApproachAngle < 135 Then
                plAttachmentLoadPoint = lCEMountingPoint
            ElseIf dApproachAngle >= 135# And dApproachAngle < 225# Then
                plAttachmentLoadPoint = 3
            ElseIf dApproachAngle >= 225# And dApproachAngle < 315# Then
                plAttachmentLoadPoint = 19
            End If
            
'Original rules before 'TR214913
'            If dApproachAngle < 180.000001 Then '# Then
'                plAttachmentLoadPoint = lCEMountingPoint
'            Else
'                Select Case lCEMountingPoint
'                    Case 11
'                        plAttachmentLoadPoint = 15
'                        pdVOffset = -1 * CalcFilletOffset()
'                    Case 25
'                        plAttachmentLoadPoint = 26
'                    Case 3
'                        plAttachmentLoadPoint = 23
'                        pdVOffset = CalcFilletOffset
'                End Select
'            End If
          
        Case Else
      End Select
      
      
    Case "HalfR"
      If dApproachAngle > 90# And dApproachAngle < 270# Then
        plAttachmentLoadPoint = 1
      Else
        plAttachmentLoadPoint = 1
        pdUOffset = dRadius * Sin(180 + dApproachAngle)
        pdVOffset = dRadius * Cos(dApproachAngle)
      End If
      
    
    Case "P", "R"
      plAttachmentLoadPoint = 27
      pdUOffset = dRadius * Sin(180 + dApproachAngle)
      pdVOffset = dRadius * Cos(dApproachAngle)
    
    
    Case Else
      
      
  End Select
  
  
  
''''''''''Lapped Connection''''''''''''''''''''''''
ElseIf eConnectType = LCA_LAPPED Then 'Lapped
  
  Select Case sCECrossSectionType
    Case "B", "EA", "UA", "BUTL3", "B_TSN"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          If (dApproachAngle < 1# Or dApproachAngle > 359#) And lCEMountingPoint = 3 Then
            plAttachmentLoadPoint = 11
            pdVOffset = -1 * CalcLapOffset("Web")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case TopSideOfSuperiorFlange
          If dApproachAngle > 179# And dApproachAngle < 181# And lCEMountingPoint = 11 Then
            plAttachmentLoadPoint = 23
            pdVOffset = CalcLapOffset("Web")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case LeftWeb
          If dApproachAngle > 269# And dApproachAngle < 271# And lCEMountingPoint = 11 Then
            plAttachmentLoadPoint = 15
            pdUOffset = -1 * CalcLapOffset("Flange")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case RightSideOfSuperiorFlange
          If dApproachAngle > 89# And dApproachAngle < 91# And lCEMountingPoint = 15 _
            And Not sCECrossSectionType = "BUTL3" Then
            plAttachmentLoadPoint = 11
            pdUOffset = CalcLapOffset("Flange")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case Else
          Err.Raise E_INVALIDARG
    
      End Select
      
    Case "BUT", "BUTL2", "TSType", "T_XType"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          Err.Raise SBR_E_NO_SOLUTION
        Case TopSideOfSuperiorFlange
          If dApproachAngle > 179# And dApproachAngle < 181# And lCEMountingPoint = 12 Then
            plAttachmentLoadPoint = 3
            pdVOffset = CalcLapOffset("Web")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case LeftSideOfSuperiorFlange
          If dApproachAngle > 269# And dApproachAngle < 271# And lCEMountingPoint = 11 Then
            plAttachmentLoadPoint = 15
            pdUOffset = -1 * CalcLapOffset("Flange")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case Else
          Err.Raise E_INVALIDARG
      End Select
            
    Case "I", "ISType"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          Err.Raise SBR_E_NO_SOLUTION
        Case LeftSideOfSuperiorFlange
          If dApproachAngle > 269# And dApproachAngle < 271# And lCEMountingPoint = 11 Then
            plAttachmentLoadPoint = 15
            pdUOffset = -1 * CalcLapOffset("Flange")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case Else
          Err.Raise E_INVALIDARG
      End Select
        
        
    Case "CSType", "C_SS"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 1# Or dApproachAngle > 359# And lCEMountingPoint = 3 Then
            plAttachmentLoadPoint = 11
            pdVOffset = -1 * CalcLapOffset("Web")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case LeftWeb
          If dApproachAngle > 269# And dApproachAngle < 271# And lCEMountingPoint = 11 Then
            plAttachmentLoadPoint = 15
            pdUOffset = -1 * CalcLapOffset("Flange")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case RightSideOfSuperiorFlange
          If dApproachAngle > 89# And dApproachAngle < 91# And lCEMountingPoint = 15 Then
            plAttachmentLoadPoint = 11
            pdUOffset = CalcLapOffset("Flange")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case Else
          Err.Raise E_INVALIDARG
      End Select
    
    Case "FB"
      Select Case eCEMountingFace
        Case BottomSideOfInferiorFlange
          If (dApproachAngle < 1# Or dApproachAngle > 359#) And lCEMountingPoint = 3 Then
            plAttachmentLoadPoint = 11
            pdVOffset = -1 * CalcLapOffset("Web")
          Else
            Err.Raise SBR_E_NO_SOLUTION
          End If
        Case LeftWeb
          Err.Raise SBR_E_NO_SOLUTION
        Case Else
          Err.Raise E_INVALIDARG
      End Select
    Case Else
      Err.Raise E_INVALIDARG
      
      
  End Select

End If

Exit Sub

ErrorHandler:
  Err.Raise Err.Number

End Sub


Private Sub Class_Initialize()
  dCachedTSArea = 0#
  dCachedTSWebThickness = 0#
  dCachedTSFlangeThickness = 0#
  dCachedTSDiameter = 0#
  dCachedTSWidth = 0#
  dCachedTSDepth = 0#
  sCachedTSCrossSectionType = ""
  sCachedTSCrossSectionSize = ""
  dCachedCEArea = 0#
  dCachedCEWebThickness = 0#
  dCachedCEFlangeThickness = 0#
  dCachedCEDiameter = 0#
  dCachedCEWidth = 0#
  dCachedCEDepth = 0#
  sCachedCECrossSectionType = ""
  sCachedCECrossSectionSize = ""
End Sub

Private Sub RefreshCache _
        (sTSCrossSectionType As String, _
         sTSCrossSectionSize As String, _
         eTSMountingFace As StiffenerRuleLib.ProfileFaceName, _
         sCECrossSectionType As String, _
         sCECrossSectionSize As String, _
         eCEMountingFace As StiffenerRuleLib.ProfileFaceName)
            
  Dim oRefDataService As RefDataMiddleServices.RefdataSOMMiddleServices
  Dim oXSection As IJCrossSection
  Dim oAttrColl As IJDAttributes
  
  On Error GoTo ErrorHandler
  Set oRefDataService = New RefDataMiddleServices.RefdataSOMMiddleServices
  
  If Not sTSCrossSectionType = sCachedTSCrossSectionType Or Not sTSCrossSectionSize = sCachedTSCrossSectionSize Then
      
    Set oXSection = oRefDataService.GetCrossSection("ShipShapes", sTSCrossSectionSize)
    Set oAttrColl = oXSection
    
    ' For some reason, the attributes manager changes an E_NOINTERFACE to OLE_CANTBINDTOSOURCE
    ' if we ask for an attributes collection this particular cross section does not support.
    ' When we ask, for example, for the flange thickness of a flat bar we get this error.
    ' Cache the Default values, if these calls fail for any reason
    On Error Resume Next

    dCachedTSWebThickness = oAttrColl.CollectionOfAttributes("IJUAXSectionWeb").Item("WebThickness").Value
    dCachedTSFlangeThickness = oAttrColl.CollectionOfAttributes("IJUAXSectionFlange").Item("FlangeThickness").Value
    dCachedTSDiameter = oAttrColl.CollectionOfAttributes("IJUAXSectionFlange").Item("Diameter").Value
    dCachedTSArea = oAttrColl.CollectionOfAttributes("IStructCrossSectionDimensions").Item("Area").Value
    dCachedTSWidth = oAttrColl.CollectionOfAttributes("IStructCrossSectionDimensions").Item("Width").Value
    If dCachedTSWidth < 0.00001 Then
        ' If Width is not specified, then use FlangeLength
        dCachedTSWidth = oAttrColl.CollectionOfAttributes("IJUAXSectionFlange").Item("FlangeLength").Value
    End If
    dCachedTSDepth = oAttrColl.CollectionOfAttributes("IStructCrossSectionDimensions").Item("Depth").Value
    If dCachedTSDepth < 0.00001 Then
        ' If Depth is not specified, then use WebLength
        dCachedTSDepth = oAttrColl.CollectionOfAttributes("IJUAXSectionWeb").Item("WebLength").Value
    End If
    
    sCachedTSCrossSectionType = sTSCrossSectionType
    sCachedTSCrossSectionSize = sTSCrossSectionSize
    Err.Clear
  End If
  
  On Error GoTo ErrorHandler
  
  If Not sCECrossSectionType = sCachedCECrossSectionType Or Not sCECrossSectionSize = sCachedCECrossSectionSize Then
  
    Set oXSection = oRefDataService.GetCrossSection("ShipShapes", sCECrossSectionSize)
    Set oAttrColl = oXSection
    
    On Error Resume Next

    dCachedCEWebThickness = oAttrColl.CollectionOfAttributes("IJUAXSectionWeb").Item("WebThickness").Value
    dCachedCEFlangeThickness = oAttrColl.CollectionOfAttributes("IJUAXSectionFlange").Item("FlangeThickness").Value
    dCachedCEDiameter = oAttrColl.CollectionOfAttributes("IJUAXSectionFlange").Item("Diameter").Value
    dCachedCEArea = oAttrColl.CollectionOfAttributes("IStructCrossSectionDimensions").Item("Area").Value
    dCachedCEWidth = oAttrColl.CollectionOfAttributes("IStructCrossSectionDimensions").Item("Width").Value
    If dCachedCEWidth < 0.00001 Then
        ' If Width is not specified, then use FlangeLength
        dCachedCEWidth = oAttrColl.CollectionOfAttributes("IJUAXSectionFlange").Item("FlangeLength").Value
    End If
    dCachedCEDepth = oAttrColl.CollectionOfAttributes("IStructCrossSectionDimensions").Item("Depth").Value
    If dCachedCEDepth < 0.00001 Then
        ' If Depth is not specified, then use WebLength
        dCachedCEDepth = oAttrColl.CollectionOfAttributes("IJUAXSectionWeb").Item("WebLength").Value
    End If

    sCachedCECrossSectionType = sCECrossSectionType
    sCachedCECrossSectionSize = sCECrossSectionSize
    Err.Clear
  End If
  
  Set oXSection = Nothing
  Set oAttrColl = Nothing
  Set oRefDataService = Nothing
  
  Exit Sub
  
ErrorHandler:
    Err.Raise Err.Number

End Sub


Private Function CalcFilletOffset() As Double
  Dim dDepth As Double
    
  'The following Select is used because some cross sections do not
  'store the depth.  The depth would have been used otherwise.
  
  Select Case m_eCEMountingFace
    Case BottomSideOfInferiorFlange, TopSideOfSuperiorFlange
      dDepth = dCachedCEDepth
    
    Case LeftInteriorSideOfInferiorFlange, LeftInteriorSideOfSuperiorFlange, _
         LeftSideOfInferiorFlange, LeftSideOfSuperiorFlange, _
         RightInteriorSideOfInferiorFlange, RightInteriorSideOfSuperiorFlange, _
         RightSideOfInferiorFlange, RightSideOfSuperiorFlange
      dDepth = dCachedCEWidth
         
    Case Else
      CalcFilletOffset = 0# 'Unknown mounting face, therefore, no fillet offset
  End Select
  
  If dDepth = 0# Then         'If the depth = 0 Then
    CalcFilletOffset = 0#     '  unknown depth, therefore, no fillet offset
  Else                          'the fillet offset defaults to 10 mm
    CalcFilletOffset = 0.01   '  fillet offset is 10mm
  End If
    
End Function

Private Function CalcLapOffset(sLappingFace As String) As Double

  Dim dLap As Double
  Dim dThickness As Double
  Dim dMaxLap As Double
  Dim dTSArea As Double
  On Error GoTo ErrorHandler
  
  'Determine which thickness to use
  If UCase(sLappingFace) = "WEB" Then
    
    dMaxLap = dCachedCEDepth - 0.05
    If dMaxLap <= 0 Then
      Err.Raise SBR_E_INADEQUATE_LAP
    End If
    
  ElseIf UCase(sLappingFace) = "FLANGE" Then
    
    dMaxLap = dCachedCEWidth - 0.05
    If dMaxLap <= 0 Then
      Err.Raise SBR_E_INADEQUATE_LAP
    End If
    
  Else
    Err.Raise E_INVALIDARG
  End If
    
  If dCachedTSWebThickness > dCachedCEWebThickness Then
       dThickness = dCachedCEWebThickness
  Else
       dThickness = dCachedTSWebThickness
  End If
    
  dThickness = dThickness - 0.0015 'Computation uses thickness - 1.5mm
  
  'In this case, the area is approximated as the weblength × web thickness
  'For an accurate computation, the mounting face of the tripping stiffener
  'must be considered
  dTSArea = dCachedTSDepth * dCachedTSWebThickness
  
  dLap = (dTSArea) / (2 * 0.707106781186548 * dThickness) 'Compute Lap
   
  If dLap > dMaxLap Then 'If the lap is greater than the maximum lap, fail
    Err.Raise SBR_E_INADEQUATE_LAP
  End If
  
  CalcLapOffset = dLap
  
  Exit Function
  
ErrorHandler:
  Err.Raise Err.Number
End Function







