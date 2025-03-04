VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectBracket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  FILE:  ConnectRules.cls
'
'  DESCRIPTION:
'  Rule to compute the keypoint on the connecting profile that the tripping
'  bracket should connect to and the offsets from that keypoint
'
'  AUTHOR:  Daniel Pizzillo & Jonathan Stephenson
'
'  HISTORY
'  040210    (JLS) created from TrippingStiffenerRules
'-------------------------------------------------------------------------------

Option Explicit
Implements IJProfileConnectRule

Private dCachedCEArea As Double
Private dCachedCEWebThickness As Double
Private dCachedCEFlangeThickness As Double
Private dCachedCEDiameter As Double
Private dCachedCEWidth As Double
Private dCachedCEDepth As Double
Private sCachedCECrossSectionType As String
Private sCachedCECrossSectionSize As String
Private m_eCEMountingFace As StiffenerRuleLib.ProfileFaceName

Private Sub IJProfileConnectRule_ProfileConnectRule(ByVal dApproachAngle As Double, _
                                                    ByVal lConnectedMountingPoint As Long, _
                                                    ByVal sConnectedCrossSectionType As String, _
                                                    ByVal sConnectedCrossSectionSize As String, _
                                                    ByVal bConnectedIsStiffener As Boolean, _
                                                    ByVal eConnectedMountingFace As PlaneByElementsRuleLib.ProfileFaceName, _
                                                    ByVal eConnectType As PlaneByElementsRuleLib.PlaneDefinitionAttachmentMethod, _
                                                    ByVal pProfileSupport1 As Object, _
                                                    ByVal pProfileSupport2 As Object, _
                                                    ByVal pProfileSupport3 As Object, _
                                                    ByVal strSupportName As String, _
                                                    plAttachmentLoadPoint As Long, _
                                                    pdUOffset As Double, _
                                                    pdVOffset As Double)
Const METHOD = "IJProfileConnectRule_ProfileConnectOffsets"

Dim dRadius As Double

On Error GoTo ErrorHandler

'Initialize outputs
plAttachmentLoadPoint = 1
pdUOffset = 0#
pdVOffset = 0#

If dApproachAngle < 0# Or dApproachAngle >= 360# Then
  Err.Raise E_INVALIDARG
End If

m_eCEMountingFace = eConnectedMountingFace 'used by CalcFilletOffset

RefreshCache sConnectedCrossSectionType, sConnectedCrossSectionSize, eConnectedMountingFace

dRadius = 0.5 * dCachedCEDiameter  'get the radius for circular profiles

'''''''''''''''Attached Connection''''''''''''''''''''''''
If bConnectedIsStiffener And eConnectType = PDA_ATTACHED Then
  Select Case sConnectedCrossSectionType
    Case "B", "EA", "UA", "BUTL3"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            plAttachmentLoadPoint = 11
            pdUOffset = CalcFilletOffset()
          ElseIf dApproachAngle < 90# Then
            plAttachmentLoadPoint = 11
            pdVOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 270# Then
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If

        Case TopSideOfSuperiorFlange
          If dApproachAngle < 90# Or dApproachAngle >= 270# Then
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 270# Then
            plAttachmentLoadPoint = 3
          Else
            Err.Raise E_INVALIDARG
          End If

        Case LeftWeb
          If dApproachAngle < 180# Then
            plAttachmentLoadPoint = lConnectedMountingPoint
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
            plAttachmentLoadPoint = lConnectedMountingPoint
          Else
            Err.Raise E_INVALIDARG
          End If

        Case Else
          Err.Raise E_INVALIDARG
      End Select

    Case "BUT", "BUTL2", "TSType", "T_XType"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            Select Case lConnectedMountingPoint
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
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If

        Case TopSideOfSuperiorFlange
          If dApproachAngle < 90# Or dApproachAngle >= 270# Then
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 270# Then
            Select Case lConnectedMountingPoint
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
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 360# Then
            Select Case lConnectedMountingPoint
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
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            Select Case lConnectedMountingPoint
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
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If

        Case LeftSideOfSuperiorFlange
          If dApproachAngle < 180# Then
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 225# Then
            plAttachmentLoadPoint = 23
            pdUOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 270# Then
            Select Case lConnectedMountingPoint
              Case 11
                plAttachmentLoadPoint = 15
              Case 9
                plAttachmentLoadPoint = 17
              Case 25
                plAttachmentLoadPoint = 23
            End Select
          ElseIf dApproachAngle < 315# Then
            Select Case lConnectedMountingPoint
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
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            plAttachmentLoadPoint = 11
            pdUOffset = CalcFilletOffset()
          ElseIf dApproachAngle < 90# Then
            plAttachmentLoadPoint = 11
            pdVOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 270# Then
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 315# Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise E_INVALIDARG
          End If

        Case LeftWeb
          If dApproachAngle < 180# Then
            plAttachmentLoadPoint = lConnectedMountingPoint
          ElseIf dApproachAngle < 225# Then
            plAttachmentLoadPoint = 23
            pdUOffset = -1 * CalcFilletOffset()
          ElseIf dApproachAngle < 270# Then
            Select Case lConnectedMountingPoint
              Case 11
                plAttachmentLoadPoint = 15
              Case 9
                plAttachmentLoadPoint = 17
              Case 25
                plAttachmentLoadPoint = 23
            End Select
          ElseIf dApproachAngle < 315# Then
            Select Case lConnectedMountingPoint
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
            Select Case lConnectedMountingPoint
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
            plAttachmentLoadPoint = lConnectedMountingPoint
          Else
            Err.Raise E_INVALIDARG
          End If

        Case Else
          Err.Raise E_INVALIDARG
      End Select

    Case "FB"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 90# Or dApproachAngle >= 270# Then
            Select Case lConnectedMountingPoint
              Case 3
                plAttachmentLoadPoint = 11
              Case 1
                plAttachmentLoadPoint = 13
              Case 23
                plAttachmentLoadPoint = 15
            End Select
          Else
            plAttachmentLoadPoint = lConnectedMountingPoint
          End If
        Case LeftWeb
          If dApproachAngle < 180# Then
            plAttachmentLoadPoint = lConnectedMountingPoint
          Else
            Select Case lConnectedMountingPoint
              Case 11
                plAttachmentLoadPoint = 15
                pdVOffset = -1 * CalcFilletOffset()
              Case 25
                plAttachmentLoadPoint = 26
              Case 3
                plAttachmentLoadPoint = 23
                pdVOffset = CalcFilletOffset
            End Select
          End If

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
ElseIf eConnectType = PDA_LAPPED Then 'Lapped

  Select Case sConnectedCrossSectionType
    Case "B", "EA", "UA", "BUTL3"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If (dApproachAngle < 2# Or dApproachAngle > 358#) And lConnectedMountingPoint = 3 Then
            plAttachmentLoadPoint = 11
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case TopSideOfSuperiorFlange
          If dApproachAngle > 178# And dApproachAngle < 182# And lConnectedMountingPoint = 11 Then
            plAttachmentLoadPoint = 23
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case LeftWeb
          If dApproachAngle > 268# And dApproachAngle < 272# And lConnectedMountingPoint = 11 Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case RightSideOfSuperiorFlange
          If dApproachAngle > 88# And dApproachAngle < 92# And lConnectedMountingPoint = 15 _
            And Not sConnectedCrossSectionType = "BUTL3" Then
            plAttachmentLoadPoint = 11
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case Else
          Err.Raise E_INVALIDARG
      End Select

    Case "BUT", "BUTL2", "TSType", "T_XType"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          Err.Raise PBR_E_NO_SOLUTION
        Case TopSideOfSuperiorFlange
          If dApproachAngle > 178# And dApproachAngle < 182# And lConnectedMountingPoint = 12 Then
            plAttachmentLoadPoint = 3
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case LeftSideOfSuperiorFlange
          If dApproachAngle > 268# And dApproachAngle < 272# And lConnectedMountingPoint = 11 Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case Else
          Err.Raise E_INVALIDARG
      End Select

    Case "I", "ISType"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          Err.Raise PBR_E_NO_SOLUTION
        Case LeftSideOfSuperiorFlange
          If dApproachAngle > 268# And dApproachAngle < 272# And lConnectedMountingPoint = 11 Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case Else
          Err.Raise E_INVALIDARG
      End Select

    Case "CSType", "C_SS"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 2# Or dApproachAngle > 358# And lConnectedMountingPoint = 3 Then
            plAttachmentLoadPoint = 11
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case LeftWeb
          If dApproachAngle > 268# And dApproachAngle < 272# And lConnectedMountingPoint = 11 Then
            plAttachmentLoadPoint = 15
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case RightSideOfSuperiorFlange
          If dApproachAngle > 88# And dApproachAngle < 92# And lConnectedMountingPoint = 15 Then
            plAttachmentLoadPoint = 11
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case Else
          Err.Raise E_INVALIDARG
      End Select

    Case "FB"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If (dApproachAngle < 2# Or dApproachAngle > 358#) And lConnectedMountingPoint = 3 Then
            plAttachmentLoadPoint = 11
          Else
            Err.Raise PBR_E_NO_SOLUTION
          End If
        Case LeftWeb
          Err.Raise PBR_E_NO_SOLUTION
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
        (sCECrossSectionType As String, _
         sCECrossSectionSize As String, _
         eCEMountingFace As StiffenerRuleLib.ProfileFaceName)
            
  Dim oRefDataService As RefDataMiddleServices.RefdataSOMMiddleServices
  Dim oXSection As IJCrossSection
  Dim oAttrColl As IJDAttributes
  
  On Error GoTo ErrorHandler
  
  If Not sCECrossSectionType = sCachedCECrossSectionType Or Not sCECrossSectionSize = sCachedCECrossSectionSize Then
    
    Set oRefDataService = New RefDataMiddleServices.RefdataSOMMiddleServices
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
  ElseIf dDepth < 0.1 Then    'If the depth < 100mm Then
    CalcFilletOffset = 0.01   '  fillet offset is 10mm
  ElseIf dDepth <= 0.2 Then    'If the depth is between 100-200mm
    CalcFilletOffset = 0.015  '  fillet offset is 15mm
  ElseIf dDepth > 0.2 Then    'If the depth is greater than 200mm
    CalcFilletOffset = 0.02    '  fillet offset is 20mm
  End If
    
End Function
