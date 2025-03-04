VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtWeldXParm"
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
'
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "PhysConnRules.ButtWeldXParm"
Const m_ParameterRuleName As String = CUSTOMERID + "PhysConnRules.ButtWeldXParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "PhysConnRules\ButtWeldXParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
  
  pOH.SetOutput "RootGap"
  pOH.SetOutput "NRRootGap"
  pOH.SetOutput "Nose"                      'User code "N"
  pOH.SetOutput "NRNose"
  pOH.SetOutput "NoseOrientationAngle"          'User code "n"
  pOH.SetOutput "NRNoseOrientationAngle"        'User code "n" for second part
  pOH.SetOutput "RefSideFirstBevelDepth"        'User code "A"
  pOH.SetOutput "NRRefSideFirstBevelDepth"      'User code "A" for second part
  pOH.SetOutput "RefSideFirstBevelAngle"        'User code "a"
  pOH.SetOutput "NRRefSideFirstBevelAngle"      'User code "a" for second part
  pOH.SetOutput "AntiRefSideFirstBevelDepth"    'User code "B"
  pOH.SetOutput "NRAntiRefSideFirstBevelDepth"  'User code "B" for second part
  pOH.SetOutput "AntiRefSideFirstBevelAngle"    'User code "b"
  pOH.SetOutput "NRAntiRefSideFirstBevelAngle"  'User code "b" for second part
  pOH.SetOutput "RefSideSecondBevelDepth"       'User code "C"
  pOH.SetOutput "NRRefSideSecondBevelDepth"     'User code "C" for second part
  pOH.SetOutput "RefSideSecondBevelAngle"       'User code "c"
  pOH.SetOutput "NRRefSideSecondBevelAngle"     'User code "c" for second part
  pOH.SetOutput "AntiRefSideSecondBevelDepth"   'User code "D"
  pOH.SetOutput "NRAntiRefSideSecondBevelDepth" 'User code "D" for second part
  pOH.SetOutput "AntiRefSideSecondBevelAngle"   'User code "d"
  pOH.SetOutput "NRAntiRefSideSecondBevelAngle" 'User code "d" for second part
  pOH.SetOutput "ButtCategory"
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "RefPartName", imsARGUMENT_IS_BSTR
  pOH.SetOutput "NRReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "NonRefPartName", imsARGUMENT_IS_BSTR
  pOH.SetOutput "ReferencePart"
  
  ' call sub to declare the computed BUTT weld outputs for IJWeldSymbol
  AddWeldParmRuleOutputs BUTT_WELD_X, pOH
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler
  
    '*****************************************************
    'For SM Implementation
    'This is an X weld.  It has a 2/3 bevel on the molded side, and a 1/3 bevel on the antimolded side
    'the bevel angle on the molded side is 50; it is 60 on the antimolded side
    '*******************************************************
    
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Dim oPlate1 As New StructDetailObjects.PlatePart
    Dim oPlate2 As New StructDetailObjects.PlatePart
    Dim dThickness As Double
    
    Set oPhysConn.object = pPRL.SmartOccurrence
    
    'Variables to be used for getting the part names
    Dim sPartNameRef As String
    Dim sPartNameNR As String
    Dim oNamedItemRef As IJNamedItem
    Dim oNamedItemNR As IJNamedItem
    
    ' variable to indicate if plate normals are reversed
    Dim bNormalsReversed As Boolean
    bNormalsReversed = False
    
    'dim variables for bevel parameters
    Dim dBevelAngle As Double
    Dim dNose As Double
    Dim dNRNose As Double
    Dim dFirstBevel As Double
    Dim dAntiFirstBevel As Double
    Dim dNRFirstBevel As Double
    Dim dNRAntiFirstBevel As Double
    Dim dAntiBevelAngle As Double
    
    ' keep track of the reference part number and objects
    Dim iRefPartNum As Integer
    Dim oPartRef As Object
    Dim oPartNR As Object
   
    ' *************************************************************************************
    ' create some variables to hold the "addition" values to account for any thickness
    ' differences between the parts that is not accounted for by chamfers.  There are four
    ' addition values, one for each side of each plate at the connection, but at least two
    ' of them will always be 0.  The other two will be 0 if there is no thickness change
    ' at that point or greater than 0 if this face of this part is larger than the same
    ' face on the connected part.
    '
    ' This rule will compute the overlapping thickness between the two parts.  These addition
    ' values are used to grow the appropriate nose and/or dBevel length to account for
    ' thickness changes between the parts.
    '
    ' This is currently only computed for edge connections between two plate parts.  The
    ' values will be 0 for other connections.
    ' *************************************************************************************
    Dim dAdditionBaseRef As Double
    Dim dAdditionOffsetRef As Double
    Dim dAdditionBaseNR As Double
    Dim dAdditionOffsetNR As Double
    
    ' initialize them all to 0.
    dAdditionBaseRef = 0
    dAdditionOffsetRef = 0
    dAdditionBaseNR = 0
    dAdditionOffsetNR = 0
    
    'get answers from selector
    Dim sCategory As String
    GetSelectorAnswer pPRL, "Category", sCategory
    
    Dim sWorkcenter As String
    GetSelectorAnswer pPRL, "Workcenter", sWorkcenter
    
    Dim sChamferType As String
    GetSelectorAnswer pPRL, "ChamferType", sChamferType
    
    Dim sFirstWeldingSide As String
    GetSelectorAnswer pPRL, "FirstWeldingSide", sFirstWeldingSide
    
    ' ********************************************************************
    ' determine which part is the reference and which is the non-reference
    ' ********************************************************************
    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
        ' reference part is item 2
        iRefPartNum = 2
        Set oPartRef = oPhysConn.ConnectedObject2
        Set oPartNR = oPhysConn.ConnectedObject1
        Set oNamedItemRef = oPhysConn.ConnectedObject2
        Set oNamedItemNR = oPhysConn.ConnectedObject1
    Else
        ' reference part is item 1
        iRefPartNum = 1
        Set oPartRef = oPhysConn.ConnectedObject1
        Set oPartNR = oPhysConn.ConnectedObject2
        Set oNamedItemRef = oPhysConn.ConnectedObject1
        Set oNamedItemNR = oPhysConn.ConnectedObject2
    End If
    
    ' *************************************************
    ' set the part names for the reference and NR parts
    ' *************************************************
    sPartNameRef = oNamedItemRef.Name
    sPartNameNR = oNamedItemNR.Name
              
    ' ***************************************************************************************
    ' Get the overlapping thickness between the parts and the additions that would have to
    ' be made to either face of either part to make them align with the corresponding face
    ' of the other part.  If this is not a plate-edge to plate-edge connection, the thickness
    ' will be the thickness of the thinner part and the difference in thickness will be
    ' returned as an addition in dAdditionOffsetNR
    ' ***************************************************************************************
    GetButtWeldOverlappingThicknessAndAdditions pPRL, _
                                                oPartRef, _
                                                oPartNR, _
                                                sChamferType, _
                                                iRefPartNum, _
                                                dThickness, _
                                                dAdditionBaseRef, _
                                                dAdditionOffsetRef, _
                                                dAdditionBaseNR, _
                                                dAdditionOffsetNR, _
                                                bNormalsReversed

                    
    'Calculation of angles, check for knuckle
    Dim dButtMountingAngle As Double
    Dim dNOA As Double
    Dim dNRNOA As Double
    
            'Modification is done as per TR-171305 --- manish
    ' This used to use CornerButtMountingAngle, but this method causes some
    ' data loss.  Changing to use TeeMountingAngle (TR170710)
    'dButtMountingAngle = oPhysConn.CornerButtMountingAngle
           dButtMountingAngle = oPhysConn.TeeMountingAngle
    ' the code is expecting the angle to be less than 180 degrees (PI),
    ' but it could be greater than PI because of an issue in GetMountingAngle
    ' adjust it before using it  (TR170710) --
    ' **** this should be changed when DI 171304 is implemented ****
    If (dButtMountingAngle - TOL) >= PI Then
        dButtMountingAngle = (2 * PI) - dButtMountingAngle
    End If
    
    If (dButtMountingAngle - TOL) >= 1.570796 Then
        dButtMountingAngle = PI - dButtMountingAngle
    End If
    
    ' ************************************************************************************************
    ' get the reference sides.  This returns "Base" or "Offset", depending which is the "molded" side.
    ' The answer can be different for the two parts.
    ' ************************************************************************************************
    Dim sRefSideRef As String
    Dim sRefSideNR As String
    sRefSideRef = GetRefSide(oPhysConn.object, oPartRef)
    sRefSideNR = GetRefSide(oPhysConn.object, oPartNR)
    
          'get the reference side
    Dim sMoldedSide_ProfileRef As String
    Dim sMoldedSide_ProfileNR As String
    
    Dim bRefIsMolded As Boolean
    
    If sRefSideRef = "Base" Then
          pPRL.Add "ReferenceSide", "molded"
          bRefIsMolded = True
    ElseIf sRefSideRef = "Offset" Then
          pPRL.Add "ReferenceSide", "antimolded"
          bRefIsMolded = False
    Else
          sMoldedSide_ProfileRef = GetMoldedSide(oPhysConn.object, oPartRef)
          pPRL.Add "ReferenceSide", sRefSideRef
          bRefIsMolded = True
    End If
    
    If sRefSideNR = "Base" Then
          pPRL.Add "NRReferenceSide", "molded"
    ElseIf sRefSideNR = "Offset" Then
          pPRL.Add "NRReferenceSide", "antimolded"
    Else
          sMoldedSide_ProfileNR = GetMoldedSide(oPhysConn.object, oPartNR)
          pPRL.Add "NRReferenceSide", sRefSideNR
    End If


    dNose = 0#
     
     ' store rootgaps for use in storing the output and for computing IJWeldSymbol paramters
    Dim dRootGap As Double
    Dim dNRRootGap As Double
    dRootGap = 0#
    dNRRootGap = 0#
    pPRL.Add "RootGap", dRootGap
    pPRL.Add "NRRootGap", dNRRootGap
    
    If sFirstWeldingSide = "Molded" Then
        If sRefSideRef = "Base" Or sMoldedSide_ProfileRef = "WebLeft" Or sMoldedSide_ProfileRef = "TopFlangeTopFace" Or sMoldedSide_ProfileRef = "BottomFlangeBottomFace" Then
            dBevelAngle = 0.4363323 '25 degrees, to create an opening of 50 degrees
            dAntiBevelAngle = 0.5235988 '30 degrees, to create an opening of 60 degrees
        ElseIf sRefSideRef = "Offset" Or sMoldedSide_ProfileRef = "WebRight" Or _
            sMoldedSide_ProfileRef = "TopFlangeBottomFace" Or sMoldedSide_ProfileRef = "BottomFlangeTopFace" Then
            dBevelAngle = 0.5235988 '30 degrees, to create an opening of 60 degrees
            dAntiBevelAngle = 0.4363323 '25 degrees, to create an opening of 50 degrees
        Else 'use web left for a profile
            dBevelAngle = 0.4363323 '25 degrees, to create an opening of 50 degrees
            dAntiBevelAngle = 0.5235988 '30 degrees, to create an opening of 60 degrees
        End If
    Else 'the weld side is antimolded
        If sRefSideRef = "Offset" Or sMoldedSide_ProfileRef = "WebRight" Or _
            sMoldedSide_ProfileRef = "TopFlangeBottomFace" Or sMoldedSide_ProfileRef = "BottomFlangeTopFace" Then
            dBevelAngle = 0.4363323 '25 degrees, to create an opening of 50 degrees
            dAntiBevelAngle = 0.5235988 '30 degrees, to create an opening of 60 degrees
        ElseIf sRefSideRef = "Base" Or sMoldedSide_ProfileRef = "WebLeft" Or sMoldedSide_ProfileRef = "TopFlangeTopFace" Or sMoldedSide_ProfileRef = "BottomFlangeBottomFace" Then
            dBevelAngle = 0.5235988 '30 degrees, to create an opening of 60 degrees
            dAntiBevelAngle = 0.4363323 '25 degrees, to create an opening of 50 degrees
        Else 'use web right for a profile
            dBevelAngle = 0.5235988 '30 degrees, to create an opening of 60 degrees
            dAntiBevelAngle = 0.4363323 '25 degrees, to create an opening of 50 degrees
        End If
    End If
    
    'the following calculates the angles and nose value based on a knuckled case
    dNOA = 1.570796

    If oPhysConn.ConnectedObject1Type = SDOBJECT_PLATE And oPhysConn.ConnectedObject2Type = SDOBJECT_PLATE Then
    Select Case sFirstWeldingSide
        Case "Molded"
            If (dButtMountingAngle - TOL) >= -0.0001 And (dButtMountingAngle - TOL) <= 0.0001 Or _
                (dButtMountingAngle - TOL) >= (PI - 0.0001) And (dButtMountingAngle - TOL) <= (PI + 0.0001) Then 'no knuckle
                dNOA = 1.570796
            Else 'knuckle
                'use the nose value that applies for knuckles
                dNose = dNose * Cos(dButtMountingAngle / 2)
                If oPhysConn.InsideFaceOfKnuckle = "Base" Then
                    If sRefSideRef = "Base" Then 'the molded side and inside face of knuckle match
                        'reset the bevel angles based on the knuckle; same angle used for ref and non ref
                        dNOA = 1.570796 + dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - 1.570796 + dNOA
                        dAntiBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                    Else 'molded side is offset
                        dNOA = 1.570796 - dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                        dAntiBevelAngle = dBevelAngle - 1.570796 + dNOA
                    End If
                Else 'it is offset
                    If sRefSideRef = "Offset" Then 'the molded side and inside face of knuckle match
                        'reset the bevel angles based on the knuckle; same angle used for ref and non ref
                        dNOA = 1.570796 + dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                        dAntiBevelAngle = dBevelAngle - 1.570796 + dNOA
                    Else 'molded side is base
                        dNOA = 1.570796 - dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - 1.570796 + dNOA
                        dAntiBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                    End If
                End If
            End If
        Case "AntiMolded"
            If (dButtMountingAngle - TOL) >= -0.0001 And (dButtMountingAngle - TOL) <= 0.0001 Or _
                (dButtMountingAngle - TOL) >= (PI - 0.0001) And (dButtMountingAngle - TOL) <= (PI + 0.0001) Then 'no knuckle
                dNOA = 1.570796
            Else 'knuckle
                'use the nose value that applies for knuckles
                dNose = dNose * Cos(dButtMountingAngle / 2)
                If oPhysConn.InsideFaceOfKnuckle = "Base" Then
                    If sRefSideRef = "Base" Then 'the molded side and inside face of knuckle match
                        'reset the bevel angles based on the knuckle; same angle used for ref and non ref
                        dNOA = 1.570796 + dButtMountingAngle / 2
                        dAntiBevelAngle = dBevelAngle - 1.570796 + dNOA
                        dBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                    Else 'molded side is offset
                        dNOA = 1.570796 - dButtMountingAngle / 2
                        dAntiBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                        dBevelAngle = dBevelAngle - 1.570796 + dNOA
                    End If
                Else 'it is offset
                    If sRefSideRef = "Offset" Then 'the molded side and inside face of knuckle match
                        'reset the bevel angles based on the knuckle; same angle used for ref and non ref
                        dNOA = 1.570796 + dButtMountingAngle / 2
                        dAntiBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                        dBevelAngle = dBevelAngle - 1.570796 + dNOA
                    Else 'molded side is base
                        dNOA = 1.570796 - dButtMountingAngle / 2
                        dAntiBevelAngle = dBevelAngle - 1.570796 + dNOA
                        dBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                    End If
                End If
            End If
    End Select
    End If
    
    ' **********************************************************************************
    ' Set and adjust the final nose and bevel values to accout for thickness differences
    ' **********************************************************************************
    If dAdditionBaseRef < 0.00001 And _
       dAdditionOffsetRef < 0.00001 And _
       dAdditionBaseNR < 0.00001 And _
       dAdditionOffsetNR < 0.00001 Then
       'In this case, do nothing for ButtWeldX
        dFirstBevel = (dThickness * 2) / 3
        dAntiFirstBevel = dThickness / 3
        
        dNRFirstBevel = dFirstBevel
        dNRAntiFirstBevel = dAntiFirstBevel
    Else
        ' we have offsets to consider, start with the values equal and adjust them as needed
        ' at this point, we are correct if there are no adjustments to make due to thickness steps
        
        dFirstBevel = (dThickness * 2) / 3
        dAntiFirstBevel = dThickness / 3
        
        dNRFirstBevel = dFirstBevel
        dNRAntiFirstBevel = dAntiFirstBevel
        
        If sFirstWeldingSide = "Molded" Then
            If sRefSideRef = "Base" Then
                dFirstBevel = dFirstBevel + dAdditionBaseRef
                dAntiFirstBevel = dAntiFirstBevel + dAdditionOffsetRef
                
                dNRFirstBevel = dNRFirstBevel + dAdditionBaseNR
                dNRAntiFirstBevel = dNRAntiFirstBevel + dAdditionOffsetNR
            Else
                dFirstBevel = dFirstBevel + dAdditionOffsetRef
                dAntiFirstBevel = dAntiFirstBevel + dAdditionBaseRef
                
                dNRFirstBevel = dNRFirstBevel + dAdditionOffsetNR
                dNRAntiFirstBevel = dNRAntiFirstBevel + dAdditionBaseNR
            End If
        Else
            If sRefSideRef = "Offset" Then
                dFirstBevel = dFirstBevel + dAdditionBaseRef
                dAntiFirstBevel = dAntiFirstBevel + dAdditionOffsetRef
                
                dNRFirstBevel = dNRFirstBevel + dAdditionBaseNR
                dNRAntiFirstBevel = dNRAntiFirstBevel + dAdditionOffsetNR
            Else
                dFirstBevel = dFirstBevel + dAdditionOffsetRef
                dAntiFirstBevel = dAntiFirstBevel + dAdditionBaseRef
                
                dNRFirstBevel = dNRFirstBevel + dAdditionOffsetNR
                dNRAntiFirstBevel = dNRAntiFirstBevel + dAdditionBaseNR
            End If
        End If
    End If
    
    'Adjust the nose and bevel depths for both parts here
    ' also, determine if the bevel is on the base or offset side
    Dim bBevelRefSideRef As Boolean
    Dim bBevelRefSideNR As Boolean
    
    If sFirstWeldingSide = "Molded" Then
        If sRefSideRef = "Base" Or sRefSideRef = "Outer" Or sMoldedSide_ProfileRef = "WebLeft" Or sMoldedSide_ProfileRef = "TopFlangeTopFace" Or sMoldedSide_ProfileRef = "BottomFlangeBottomFace" Then
            bBevelRefSideRef = True
        ElseIf sRefSideRef = "Offset" Or sRefSideRef = "Top" Or sMoldedSide_ProfileRef = "WebRight" Or _
            sMoldedSide_ProfileRef = "TopFlangeBottomFace" Or sMoldedSide_ProfileRef = "BottomFlangeTopFace" Then
            bBevelRefSideRef = False
        Else 'profile, web left
            bBevelRefSideRef = True
        End If
    Else 'first welding side is antimolded
        If sRefSideRef = "Base" Or sRefSideRef = "Outer" Or sMoldedSide_ProfileRef = "WebLeft" Or sMoldedSide_ProfileRef = "TopFlangeTopFace" Or sMoldedSide_ProfileRef = "BottomFlangeBottomFace" Then
            bBevelRefSideRef = False
        ElseIf sRefSideRef = "Offset" Or sRefSideRef = "Top" Or sMoldedSide_ProfileRef = "WebRight" Or _
            sMoldedSide_ProfileRef = "TopFlangeBottomFace" Or sMoldedSide_ProfileRef = "BottomFlangeTopFace" Then
            bBevelRefSideRef = True
        Else 'profile, web left
            bBevelRefSideRef = True
        End If
    End If
    
    ' *********************************************************************************
    ' if the normals are the same between Reference and NonReference part, we put the
    ' bevels on the same side relative to base and offset.  Otherwise, we reverse them.
    ' *********************************************************************************
    If bNormalsReversed Then
        bBevelRefSideNR = Not bBevelRefSideRef
    Else
        bBevelRefSideNR = bBevelRefSideRef
    End If
    
    'Set Category
    pPRL.Add "ButtCategory", 65537
    
    pPRL.Add "NonRefPartName", sPartNameNR
    pPRL.Add "RefPartName", sPartNameRef
    
    ' **********************
    ' set the reference part
    ' **********************
    pPRL.Add "ReferencePart", iRefPartNum
                
    'set nose and ref part values
    pPRL.Add "Nose", dNose
    pPRL.Add "NRNose", dNose
    pPRL.Add "NoseOrientationAngle", 0#
    pPRL.Add "NRNoseOrientationAngle", 0#
    
     ' set the bevel depths and angles
    Dim dRefSideFirstBevelDepth As Double
    Dim dAntiRefSideFirstBevelDepth As Double
    Dim dRefSideFirstBevelAngle As Double
    Dim dAntiRefSideFirstBevelAngle As Double
    Dim dNRRefSideFirstBevelDepth As Double
    Dim dNRAntiRefSideFirstBevelDepth As Double
    Dim dNRRefSideFirstBevelAngle As Double
    Dim dNRAntiRefSideFirstBevelAngle As Double
       
    If bBevelRefSideRef = True Then
        dRefSideFirstBevelDepth = dFirstBevel
        dRefSideFirstBevelAngle = dBevelAngle
        dAntiRefSideFirstBevelDepth = dAntiFirstBevel
        dAntiRefSideFirstBevelAngle = dAntiBevelAngle
    Else
        dRefSideFirstBevelDepth = dAntiFirstBevel
        dRefSideFirstBevelAngle = dAntiBevelAngle
        dAntiRefSideFirstBevelDepth = dFirstBevel
        dAntiRefSideFirstBevelAngle = dBevelAngle
    End If
    
    If bBevelRefSideNR = True Then
        dNRRefSideFirstBevelDepth = dNRFirstBevel
        dNRRefSideFirstBevelAngle = dBevelAngle
        dNRAntiRefSideFirstBevelDepth = dNRAntiFirstBevel
        dNRAntiRefSideFirstBevelAngle = dAntiBevelAngle
    Else
        dNRRefSideFirstBevelDepth = dNRAntiFirstBevel
        dNRRefSideFirstBevelAngle = dAntiBevelAngle
        dNRAntiRefSideFirstBevelDepth = dNRFirstBevel
        dNRAntiRefSideFirstBevelAngle = dBevelAngle
    End If
    
    ' set the actual outputs
    pPRL.Add "RefSideFirstBevelDepth", dRefSideFirstBevelDepth
    pPRL.Add "RefSideFirstBevelAngle", dRefSideFirstBevelAngle
    pPRL.Add "AntiRefSideFirstBevelDepth", dAntiRefSideFirstBevelDepth
    pPRL.Add "AntiRefSideFirstBevelAngle", dAntiRefSideFirstBevelAngle
    pPRL.Add "NRRefSideFirstBevelDepth", dNRRefSideFirstBevelDepth
    pPRL.Add "NRRefSideFirstBevelAngle", dNRRefSideFirstBevelAngle
    pPRL.Add "NRAntiRefSideFirstBevelDepth", dNRAntiRefSideFirstBevelDepth
    pPRL.Add "NRAntiRefSideFirstBevelAngle", dNRAntiRefSideFirstBevelAngle
    
    pPRL.Add "RefSideSecondBevelDepth", 0
    pPRL.Add "RefSideSecondBevelAngle", 0
    pPRL.Add "AntiRefSideSecondBevelDepth", 0
    pPRL.Add "AntiRefSideSecondBevelAngle", 0
    pPRL.Add "NRRefSideSecondBevelDepth", 0
    pPRL.Add "NRRefSideSecondBevelAngle", 0
    pPRL.Add "NRAntiRefSideSecondBevelDepth", 0
    pPRL.Add "NRAntiRefSideSecondBevelAngle", 0
  
    ' set the IJWeldingSymbol properties
    SetCalculatedButtWeldParams pPRL, BUTT_WELD_X, bRefIsMolded, _
                                dRefSideFirstBevelDepth, dAntiRefSideFirstBevelDepth, _
                                dRefSideFirstBevelAngle, dAntiRefSideFirstBevelAngle, _
                                dNRRefSideFirstBevelDepth, dNRAntiRefSideFirstBevelDepth, _
                                dNRRefSideFirstBevelAngle, dNRAntiRefSideFirstBevelAngle, _
                                pcr_WG_V, dRootGap, dNRRootGap

  
Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub


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
  pIH.definition = pPR
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
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************






