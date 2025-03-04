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
Implements IJBracketConnectRule

Private dCachedCEArea As Double
Private dCachedCEWebThickness As Double
Private dCachedCEFlangeThickness As Double
Private dCachedCEDiameter As Double
Private dCachedCEWidth As Double
Private dCachedCEDepth As Double
Private sCachedCECrossSectionType As String
Private sCachedCECrossSectionSize As String
Private m_eCEMountingFace As StiffenerRuleLib.ProfileFaceName

Private Sub IJBracketConnectRule_MemberConnectRule(ByVal dApproachAngle As Double, _
                                                   ByVal lConnectedMountingPoint As Long, _
                                                   ByVal sConnectedCrossSectionStandard As String, _
                                                   ByVal sConnectedCrossSectionType As String, _
                                                   ByVal sConnectedCrossSectionName As String, _
                                                   ByVal bConnectedIsMember As Boolean, _
                                                   ByVal eConnectType As PlaneByElementsRuleLib.PlaneDefinitionAttachmentMethod, _
                                                   ByVal pSupport1 As Object, _
                                                   ByVal pSupport2 As Object, _
                                                   ByVal pSupport3 As Object, _
                                                   ByVal bstrSupportName As String, _
                                                   plAttachmentLoadPoint As Long, _
                                                   pdUOffset As Double, _
                                                   pdVOffset As Double)

Const METHOD = "IJBracketConnectRule_MemberConnectRule"

Dim dRadius As Double

On Error GoTo ErrorHandler

'Initialize outputs
plAttachmentLoadPoint = 1
pdUOffset = 0#
pdVOffset = 0#

If dApproachAngle < 0# Or dApproachAngle >= 360# Then
  Err.Raise E_INVALIDARG
End If

m_eCEMountingFace = 0 'used by CalcFilletOffset

RefreshCache sConnectedCrossSectionStandard, sConnectedCrossSectionType, sConnectedCrossSectionName, m_eCEMountingFace

'''''''''''''''Attached Connection''''''''''''''''''''''''
If bConnectedIsMember And eConnectType = PDA_ATTACHED Then
    Select Case sConnectedCrossSectionType
        Case "S", "M", "HP", "HSSR", "W", "C", "MC", "ST", "MT", "WT", "L"
            'approach angle is in between -90 to 90 degrees
            If dApproachAngle >= 270# Or dApproachAngle < 90# Then
                If lConnectedMountingPoint = 1 Or _
                   lConnectedMountingPoint = 4 Or _
                   lConnectedMountingPoint = 7 Or _
                   lConnectedMountingPoint = 12 Then
                   plAttachmentLoadPoint = 7
                ElseIf lConnectedMountingPoint = 2 Or _
                       lConnectedMountingPoint = 5 Or _
                       lConnectedMountingPoint = 8 Then
                       plAttachmentLoadPoint = 8
                ElseIf lConnectedMountingPoint = 3 Or _
                        lConnectedMountingPoint = 6 Or _
                        lConnectedMountingPoint = 9 Or _
                        lConnectedMountingPoint = 13 Then
                        plAttachmentLoadPoint = 9
                ElseIf lConnectedMountingPoint = 10 Or _
                        lConnectedMountingPoint = 11 Or _
                        lConnectedMountingPoint = 14 Then
                        plAttachmentLoadPoint = 14
                Else
                    plAttachmentLoadPoint = lConnectedMountingPoint
                End If
            'approach angle is in between 90 to 270 degrees
            ElseIf dApproachAngle >= 90# Or dApproachAngle < 270# Then
                If lConnectedMountingPoint = 1 Or _
                   lConnectedMountingPoint = 4 Or _
                   lConnectedMountingPoint = 7 Or _
                   lConnectedMountingPoint = 12 Then
                   plAttachmentLoadPoint = 1
                ElseIf lConnectedMountingPoint = 2 Or _
                       lConnectedMountingPoint = 5 Or _
                       lConnectedMountingPoint = 8 Then
                       plAttachmentLoadPoint = 2
                ElseIf lConnectedMountingPoint = 3 Or _
                        lConnectedMountingPoint = 6 Or _
                        lConnectedMountingPoint = 9 Or _
                        lConnectedMountingPoint = 13 Then
                        plAttachmentLoadPoint = 3
                ElseIf lConnectedMountingPoint = 10 Or _
                        lConnectedMountingPoint = 11 Or _
                        lConnectedMountingPoint = 14 Then
                        plAttachmentLoadPoint = 11
                Else
                    plAttachmentLoadPoint = lConnectedMountingPoint
                End If
            Else
                Err.Raise PBR_E_NO_SOLUTION
            End If
        Case "HSSC", "PIPE"
            'approach angle is in between -90 to 90 degrees
            If dApproachAngle >= 270# Or dApproachAngle < 90# Then
                plAttachmentLoadPoint = 8
            'approach angle is in between 90 to 270 degrees
            ElseIf dApproachAngle >= 90# Or dApproachAngle < 270# Then
                plAttachmentLoadPoint = 2
            Else
                Err.Raise PBR_E_NO_SOLUTION
            End If
    End Select
''''''''''Lapped Connection''''''''''''''''''''''''
ElseIf eConnectType = PDA_LAPPED Then 'Lapped
    Select Case sConnectedCrossSectionType
        Case "S", "M", "HP", "HSSR", "W", "C", "MC", "ST", "MT", "WT", "L", "HSSC", "PIPE"
            'approach angle is in between -90 to 90 degrees
            If dApproachAngle >= 270# Or dApproachAngle < 90# Then
                If lConnectedMountingPoint = 1 Or _
                   lConnectedMountingPoint = 4 Or _
                   lConnectedMountingPoint = 7 Or _
                   lConnectedMountingPoint = 12 Then
                   plAttachmentLoadPoint = 7
                ElseIf lConnectedMountingPoint = 2 Or _
                       lConnectedMountingPoint = 5 Or _
                       lConnectedMountingPoint = 8 Then
                       plAttachmentLoadPoint = 7
                ElseIf lConnectedMountingPoint = 3 Or _
                        lConnectedMountingPoint = 6 Or _
                        lConnectedMountingPoint = 9 Or _
                        lConnectedMountingPoint = 13 Then
                        plAttachmentLoadPoint = 9
                ElseIf lConnectedMountingPoint = 10 Or _
                        lConnectedMountingPoint = 11 Or _
                        lConnectedMountingPoint = 14 Then
                        plAttachmentLoadPoint = 7
                Else
                    plAttachmentLoadPoint = lConnectedMountingPoint
                End If
            'approach angle is in between 90 to 270 degrees
            ElseIf dApproachAngle >= 90# Or dApproachAngle < 270# Then
                If lConnectedMountingPoint = 1 Or _
                   lConnectedMountingPoint = 4 Or _
                   lConnectedMountingPoint = 7 Or _
                   lConnectedMountingPoint = 12 Then
                   plAttachmentLoadPoint = 1
                ElseIf lConnectedMountingPoint = 2 Or _
                       lConnectedMountingPoint = 5 Or _
                       lConnectedMountingPoint = 8 Then
                       plAttachmentLoadPoint = 1
                ElseIf lConnectedMountingPoint = 3 Or _
                        lConnectedMountingPoint = 6 Or _
                        lConnectedMountingPoint = 9 Or _
                        lConnectedMountingPoint = 13 Then
                        plAttachmentLoadPoint = 3
                ElseIf lConnectedMountingPoint = 10 Or _
                        lConnectedMountingPoint = 11 Or _
                        lConnectedMountingPoint = 14 Then
                        plAttachmentLoadPoint = 1
                Else
                    plAttachmentLoadPoint = lConnectedMountingPoint
                End If
            Else
                Err.Raise PBR_E_NO_SOLUTION
            End If
    End Select
End If

Exit Sub

ErrorHandler:
  Err.Raise Err.Number

End Sub

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

RefreshCache "ShipShapes", sConnectedCrossSectionType, sConnectedCrossSectionSize, eConnectedMountingFace

dRadius = 0.5 * dCachedCEDiameter  'get the radius for circular profiles

'''''''''''''''Attached Connection''''''''''''''''''''''''
If bConnectedIsStiffener And eConnectType = PDA_ATTACHED Then
  Select Case sConnectedCrossSectionType
    Case "B", "EA", "UA", "BUTL3", "B_TSN"
      Select Case eConnectedMountingFace
        Case BottomSideOfInferiorFlange
          If dApproachAngle < 45# Or dApproachAngle >= 315# Then
            plAttachmentLoadPoint = 11
            If pProfileSupport1 Is Nothing Or _
                 pProfileSupport2 Is Nothing Then
                pdUOffset = CalcFilletOffset() 'Logic prior to TR-180101 will be used
            ElseIf strSupportName = "Support1" And TypeOf pProfileSupport2 Is IJERSystem Or _
               strSupportName = "Support2" And TypeOf pProfileSupport1 Is IJERSystem Then
                   If lConnectedMountingPoint = 24 Then
                       plAttachmentLoadPoint = 14
                   ElseIf lConnectedMountingPoint = 1 Then
                       plAttachmentLoadPoint = 13
                   Else
                       plAttachmentLoadPoint = 11
                   End If
            Else
                'If and ElseIf are not True; so need following part of TR-180101 fix
                pdUOffset = CalcFilletOffset()
            End If
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
            'DM252888/ TR224673 similar to TR214913
            'DM331715 //TR310420  (ER toward plate case)//
            If dApproachAngle < 45# Or dApproachAngle >= 315# Then
                plAttachmentLoadPoint = 11
            ElseIf dApproachAngle >= 45# And dApproachAngle < 135 Then
                plAttachmentLoadPoint = lConnectedMountingPoint
            ElseIf dApproachAngle >= 135# And dApproachAngle < 225# Then
                plAttachmentLoadPoint = 3
            ElseIf dApproachAngle >= 225# And dApproachAngle < 315# Then
                plAttachmentLoadPoint = 19
            End If

'Original rule before 2011R1
'          If dApproachAngle < 180# Then
'            plAttachmentLoadPoint = lConnectedMountingPoint
'          Else
'            Select Case lConnectedMountingPoint
'              Case 11
'                plAttachmentLoadPoint = 15
'                pdVOffset = -1 * CalcFilletOffset()
'              Case 25
'                plAttachmentLoadPoint = 26
'              Case 3
'                plAttachmentLoadPoint = 23
'                pdVOffset = CalcFilletOffset
'            End Select
'          End If

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
    Case "B", "EA", "UA", "BUTL3", "B_TSN"
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

    Dim oRetER As IJERSystem
    Dim oRetStiffener As IJStiffener ''IJProfile
    Dim dLoadPointOffset As Double
    Dim dParentPlateThk As Double
    If CheckIfOneSuppIsERAndOtherIsStiffener(pProfileSupport1, pProfileSupport2, _
            oRetER, oRetStiffener) Then
        If oRetER.Position = ER_OnEdgeCentered Then
            dParentPlateThk = GetParentPlateSystemThickness(oRetER)
            dLoadPointOffset = -dParentPlateThk / 2
        End If
        '2S bracket case and 3S bracket case
        Select Case LCase(strSupportName)
        Case "support1"
            If TypeOf pProfileSupport1 Is IJStiffener And _
                        Not (pProfileSupport1 Is oRetER) Then
                If pProfileSupport2 Is oRetER Then
                    pdVOffset = 0
                End If
            End If
        Case "support2"
            If TypeOf pProfileSupport2 Is IJStiffener And _
                        Not (pProfileSupport2 Is oRetER) Then
                If pProfileSupport1 Is oRetER Then
                    pdVOffset = 0
                End If
            End If
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
        (sCECrossSectionStandard As String, _
         sCECrossSectionType As String, _
         sCECrossSectionSize As String, _
         eCEMountingFace As StiffenerRuleLib.ProfileFaceName)
            
  Dim oRefDataService As RefDataMiddleServices.RefdataSOMMiddleServices
  Dim oXSection As IJCrossSection
  Dim oAttrColl As IJDAttributes
  
  On Error GoTo ErrorHandler
  
  If Not sCECrossSectionType = sCachedCECrossSectionType Or Not sCECrossSectionSize = sCachedCECrossSectionSize Then
    
    Set oRefDataService = New RefDataMiddleServices.RefdataSOMMiddleServices
    Set oXSection = oRefDataService.GetCrossSection(sCECrossSectionStandard, sCECrossSectionSize)
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
  Else
    CalcFilletOffset = 0.01   '  fillet offset is 10mm
  End If
    
End Function

Private Function GetParentPlateSystemThickness(ByVal oERSystem As IJERSystem) As Double
    On Error GoTo ErrorHandler
    Dim oIJStiff As IJStiffener
    Set oIJStiff = oERSystem
    Dim oPlate As IJPlate
    Set oPlate = oIJStiff.PlateSystem
    GetParentPlateSystemThickness = oPlate.Thickness
    
    'Cleanup
    Set oIJStiff = Nothing
    Set oPlate = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function CheckIfOneSuppIsERAndOtherIsStiffener(ByVal pProfileSupport1 As Object, _
        ByVal pProfileSupport2 As Object, _
        oRetER As IJERSystem, _
        oRetStiffener As IJStiffener) As Boolean
    On Error GoTo ErrorHandler
    Dim blnReturnFlag As Boolean
    
    blnReturnFlag = False
If Not (pProfileSupport1 Is Nothing) And _
    Not (pProfileSupport2 Is Nothing) Then
    If TypeOf pProfileSupport1 Is IJERSystem Then
        If TypeOf pProfileSupport2 Is IJStiffener Then
            Set oRetER = pProfileSupport1
            Set oRetStiffener = pProfileSupport2
            blnReturnFlag = True
        End If
    ElseIf TypeOf pProfileSupport2 Is IJERSystem Then
        If TypeOf pProfileSupport1 Is IJStiffener Then
            Set oRetER = pProfileSupport2
            Set oRetStiffener = pProfileSupport1
            blnReturnFlag = True
        End If
    End If
End If
    CheckIfOneSuppIsERAndOtherIsStiffener = blnReturnFlag
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function

