VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtWeld3Parm"
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

Const m_ParameterRuleProgid As String = "PhysConnRules.ButtWeld3Parm"
Const m_ParameterRuleName As String = "PhysConnRules.ButtWeld3Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\PhysConnRules\ButtWeld3Parm.cls"

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
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler
  
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Dim oPlate1 As New StructDetailObjects.PlatePart
    Dim oPlate2 As New StructDetailObjects.PlatePart
    Dim dThickness As Double
    Dim dThickness1 As Double
    Dim dThickness2 As Double
    Dim dThicknessOffset1 As Double
    Dim dThicknessoffset2 As Double
      
    Set oPhysConn.object = pPRL.SmartOccurrence
    
    'Variables to be used for getting the part names
    Dim sPartName1 As String
    Dim sPartName2 As String
    Dim oNamedItem1 As IJNamedItem
    Dim oNamedItem2 As IJNamedItem
    Set oNamedItem1 = oPhysConn.ConnectedObject1
    Set oNamedItem2 = oPhysConn.ConnectedObject2
       
    'dim variables for bevel parameters
    Dim dBevel As Double
    Dim dBevelAngle As Double
    Dim dNRBevel As Double
    Dim dNose As Double
    Dim dNRNose As Double
    Dim dBaseDelta As Double
    Dim dOffsetDelta As Double
    Dim dFirstBevel As Double
    Dim dAntiFirstBevel As Double
    Dim dNRFirstBevel As Double
    Dim dNRAntiFirstBevel As Double
    Dim dAntiBevelAngle As Double

    'get answers from selector
    Dim sCategory As String
    sCategory = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "Category")
    Dim sWorkcenter As String
    sWorkcenter = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "Workcenter")
    Dim sUpside As String
    sUpside = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "UpSide")
    Dim sChamferType As String
    sChamferType = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "ChamferType")
    Dim sFirstWeldingSide As String
    sFirstWeldingSide = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "FirstWeldingSide")
    
    'find thicknesses, based on whether or not there is a chamfer
    Select Case sChamferType
        Case "None"
        'if there is no chamfer, use the thickness of each part
            If TypeOf oPhysConn.ConnectedObject1 Is IJPlate And _
                TypeOf oPhysConn.ConnectedObject2 Is IJPlate Then
                 Set oPlate1.object = oPhysConn.ConnectedObject1
                 Set oPlate2.object = oPhysConn.ConnectedObject2
                 dThicknessOffset1 = oPlate1.thicknessOffset
                 dThicknessoffset2 = oPlate2.thicknessOffset
                 If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                    dThickness = oPhysConn.Object2Thickness
                    dThickness2 = oPhysConn.Object1Thickness
                    sPartName1 = oNamedItem2.Name
                    sPartName2 = oNamedItem1.Name
                 Else
                    dThickness = oPhysConn.Object1Thickness
                    dThickness2 = oPhysConn.Object2Thickness
                    sPartName1 = oNamedItem1.Name
                    sPartName2 = oNamedItem2.Name
                 End If
                 
                'we also need the delta values, to be used in bevel calculations
                dBaseDelta = oPlate1.OffsetToBaseFace(oPlate2) - oPlate2.OffsetToBaseFace
                dOffsetDelta = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToOffsetFace
                
                If dBaseDelta < 0.0001 And dBaseDelta > -0.0001 Then
                    dBaseDelta = 0  'no chamfer on base side
                End If
                
                If dOffsetDelta < 0.0001 And dOffsetDelta > -0.0001 Then
                    dOffsetDelta = 0 ' no chamfer on offset side
                End If

            Else
                 dThicknessOffset1 = 0
                 dThicknessoffset2 = 0
                 If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                    dThickness = oPhysConn.Object2Thickness
                    dThickness2 = oPhysConn.Object1Thickness
                    sPartName1 = oNamedItem2.Name
                    sPartName2 = oNamedItem1.Name
                 Else
                    dThickness = oPhysConn.Object1Thickness
                    dThickness2 = oPhysConn.Object2Thickness
                    sPartName1 = oNamedItem1.Name
                    sPartName2 = oNamedItem2.Name
                 End If
                 
                 dBaseDelta = 0
                 dOffsetDelta = 0
                 
            End If
       
        Case "Obj1Base", "Obj2Base", "Obj1Offset", "Obj2Offset"
            'if there is a chamfer, only use the intersecting area of the two objects as the thickness
            Set oPlate1.object = oPhysConn.ConnectedObject1
            Set oPlate2.object = oPhysConn.ConnectedObject2
               
            'The base and offset delta are calculated here; this calculation is separate from
            'the chamfer
            dBaseDelta = oPlate1.OffsetToBaseFace(oPlate2) - oPlate2.OffsetToBaseFace
            dOffsetDelta = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToOffsetFace
            
            If dBaseDelta < 0.0001 And dBaseDelta > -0.0001 Then
                dBaseDelta = 0  'no chamfer on base side
            End If
            
            If dOffsetDelta < 0.0001 And dOffsetDelta > -0.0001 Then
                dOffsetDelta = 0 ' no chamfer on offset side
            End If
            
            Dim dDelta As Double
            If sChamferType = "Obj1Base" Or sChamferType = "Obj2Base" Then
                dDelta = dOffsetDelta
            ElseIf sChamferType = "Obj1Offset" Or sChamferType = "Obj2Offset" Then
                dDelta = dBaseDelta
            End If
                
            If oPlate1.OffsetToOffsetFace(oPlate2) <= oPlate2.OffsetToOffsetFace And _
                oPlate1.OffsetToBaseFace(oPlate2) <= oPlate2.OffsetToBaseFace Then
                If oPhysConn.Object1Thickness < oPhysConn.Object2Thickness Then
                    dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace
                    dThickness2 = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace + dDelta
                    sPartName1 = oNamedItem1.Name
                    sPartName2 = oNamedItem2.Name
                ElseIf oPhysConn.Object2Thickness < oPhysConn.Object1Thickness Then
                    dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace + dDelta
                    dThickness2 = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace
                    sPartName1 = oNamedItem2.Name
                    sPartName2 = oNamedItem1.Name
                Else
                    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                        dThickness = oPhysConn.Object2Thickness
                        dThickness2 = oPhysConn.Object1Thickness
                        sPartName1 = oNamedItem2.Name
                        sPartName2 = oNamedItem1.Name
                    Else
                        dThickness = oPhysConn.Object1Thickness
                        dThickness2 = oPhysConn.Object2Thickness
                        sPartName1 = oNamedItem1.Name
                        sPartName2 = oNamedItem2.Name
                    End If
                End If
            
            ElseIf oPlate1.OffsetToOffsetFace(oPlate2) <= oPlate2.OffsetToOffsetFace And _
                oPlate1.OffsetToBaseFace(oPlate2) >= oPlate2.OffsetToBaseFace Then
                If oPhysConn.Object1Thickness < oPhysConn.Object2Thickness Then
                    dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate1.OffsetToBaseFace(oPlate2)
                    dThickness2 = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate1.OffsetToBaseFace(oPlate2) + dDelta
                    sPartName1 = oNamedItem1.Name
                    sPartName2 = oNamedItem2.Name
                ElseIf oPhysConn.Object2Thickness < oPhysConn.Object1Thickness Then
                    dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate1.OffsetToBaseFace(oPlate2) + dDelta
                    dThickness2 = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate1.OffsetToBaseFace(oPlate2)
                    sPartName1 = oNamedItem2.Name
                    sPartName2 = oNamedItem1.Name
                Else
                    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                        dThickness = oPhysConn.Object2Thickness
                        dThickness2 = oPhysConn.Object1Thickness
                        sPartName1 = oNamedItem2.Name
                        sPartName2 = oNamedItem1.Name
                    Else
                        dThickness = oPhysConn.Object1Thickness
                        dThickness2 = oPhysConn.Object2Thickness
                        sPartName1 = oNamedItem1.Name
                        sPartName2 = oNamedItem2.Name
                    End If
                End If
            
            ElseIf oPlate1.OffsetToOffsetFace(oPlate2) >= oPlate2.OffsetToOffsetFace And _
                oPlate1.OffsetToBaseFace(oPlate2) >= oPlate2.OffsetToBaseFace Then
                If oPhysConn.Object1Thickness < oPhysConn.Object2Thickness Then
                    dThickness = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToBaseFace(oPlate2)
                    dThickness2 = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToBaseFace(oPlate2) + dDelta
                    sPartName1 = oNamedItem1.Name
                    sPartName2 = oNamedItem2.Name
                ElseIf oPhysConn.Object2Thickness < oPhysConn.Object1Thickness Then
                    dThickness = oPlate2.OffsetToOffsetFace - oPlate2.OffsetToBaseFace + dDelta
                    dThickness2 = oPlate2.OffsetToOffsetFace - oPlate2.OffsetToBaseFace
                    sPartName1 = oNamedItem2.Name
                    sPartName2 = oNamedItem1.Name
                Else
                    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                        dThickness = oPhysConn.Object2Thickness
                        dThickness2 = oPhysConn.Object1Thickness
                        sPartName1 = oNamedItem2.Name
                        sPartName2 = oNamedItem1.Name
                    Else
                        dThickness = oPhysConn.Object1Thickness
                        dThickness2 = oPhysConn.Object2Thickness
                        sPartName1 = oNamedItem1.Name
                        sPartName2 = oNamedItem2.Name
                    End If
                End If
            
            ElseIf oPlate1.OffsetToOffsetFace(oPlate2) >= oPlate2.OffsetToOffsetFace And _
                oPlate1.OffsetToBaseFace(oPlate2) <= oPlate2.OffsetToBaseFace Then
                If oPhysConn.Object1Thickness < oPhysConn.Object2Thickness Then
                    dThickness = oPlate2.OffsetToOffsetFace - oPlate2.OffsetToBaseFace
                    dThickness2 = oPlate2.OffsetToOffsetFace - oPlate2.OffsetToBaseFace + dDelta
                    sPartName1 = oNamedItem1.Name
                    sPartName2 = oNamedItem2.Name
                ElseIf oPhysConn.Object2Thickness < oPhysConn.Object1Thickness Then
                    dThickness = oPlate2.OffsetToOffsetFace - oPlate2.OffsetToBaseFace + dDelta
                    dThickness2 = oPlate2.OffsetToOffsetFace - oPlate2.OffsetToBaseFace
                    sPartName1 = oNamedItem2.Name
                    sPartName2 = oNamedItem1.Name
                Else
                    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                        dThickness = oPhysConn.Object2Thickness
                        dThickness2 = oPhysConn.Object1Thickness
                        sPartName1 = oNamedItem2.Name
                        sPartName2 = oNamedItem1.Name
                    Else
                        dThickness = oPhysConn.Object1Thickness
                        dThickness2 = oPhysConn.Object2Thickness
                        sPartName1 = oNamedItem1.Name
                        sPartName2 = oNamedItem2.Name
                    End If
                End If
            
            Else
                If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                    dThickness = oPhysConn.Object2Thickness
                    dThickness2 = oPhysConn.Object1Thickness
                    sPartName1 = oNamedItem2.Name
                    sPartName2 = oNamedItem1.Name
                Else
                    dThickness = oPhysConn.Object1Thickness
                    dThickness2 = oPhysConn.Object2Thickness
                    sPartName1 = oNamedItem1.Name
                    sPartName2 = oNamedItem2.Name
                End If
            End If
            
        Case "Obj1Double", "Obj2Double", "Obj1BaseObj2Offset", "Obj1OffsetObj2Base"
            Set oPlate1.object = oPhysConn.ConnectedObject1
            Set oPlate2.object = oPhysConn.ConnectedObject2
            If oPlate1.OffsetToBaseFace(oPlate2) < oPlate2.OffsetToBaseFace And _
                oPlate1.OffsetToOffsetFace(oPlate2) < oPlate2.OffsetToOffsetFace Then
                dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace
                dThickness2 = dThickness
                sPartName1 = oNamedItem1.Name
                sPartName2 = oNamedItem2.Name
            ElseIf oPlate1.OffsetToBaseFace(oPlate2) > oPlate2.OffsetToBaseFace And _
                oPlate1.OffsetToOffsetFace(oPlate2) > oPlate2.OffsetToOffsetFace Then
                dThickness = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToBaseFace(oPlate2)
                dThickness2 = dThickness
                sPartName1 = oNamedItem1.Name
                sPartName2 = oNamedItem2.Name
            Else
                If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                    dThickness = oPhysConn.Object2Thickness
                    dThickness2 = oPhysConn.Object1Thickness
                    sPartName1 = oNamedItem2.Name
                    sPartName2 = oNamedItem1.Name
                Else
                    dThickness = oPhysConn.Object1Thickness
                    dThickness2 = oPhysConn.Object2Thickness
                    sPartName1 = oNamedItem1.Name
                    sPartName2 = oNamedItem2.Name
                End If
            End If
    
    End Select
            
    dBevelAngle = 0.392699
    dAntiBevelAngle = 0.392699
    
    'Calculation of angles, check for knuckle
    Dim dButtMountingAngle As Double
    Dim dNOA As Double
    Dim dNRNOA As Double
    
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
  
    ' it is actually expecting it to be less than 90 degrees (PI/2).  I am not
    ' sure about this part of the check, but the bug above was making this value
    ' negative.
    If (dButtMountingAngle - TOL) >= 1.570796 Then
        dButtMountingAngle = PI - dButtMountingAngle
    End If

    If dThickness <= 0.02 Then
       Select Case sWorkcenter
          Case gsMachine1
                dNose = 0.005
                pPRL.Add "RootGap", -0.001
                pPRL.Add "NRRootGap", -0.001
          Case gsMachine2
                dNose = 0.006
                pPRL.Add "RootGap", -0.0011
                pPRL.Add "NRRootGap", -0.0011
       End Select
    
    ElseIf dThickness > 0.02 And dThickness <= 0.025 Then
       Select Case sWorkcenter
          Case gsMachine1
                dNose = 0.005
                pPRL.Add "RootGap", -0.001
                pPRL.Add "NRRootGap", -0.001
          Case gsMachine2
                dNose = 0.006
                pPRL.Add "RootGap", -0.0011
                pPRL.Add "NRRootGap", -0.0011
       End Select
       
    ElseIf dThickness > 0.025 Then
       Select Case sWorkcenter
          Case gsMachine1
                dNose = 0.003
                pPRL.Add "RootGap", -0.001
                pPRL.Add "NRRootGap", -0.001
          Case gsMachine2
                dNose = 0.004
                pPRL.Add "RootGap", -0.0011
                pPRL.Add "NRRootGap", -0.0011
       End Select
    End If
    
    Dim sRefSide As String
    sRefSide = GetRefSide(oPhysConn.ConnectedObject1)
    
    'the following calculates the angles and nose value based on a knuckled case
    Select Case sFirstWeldingSide
        Case "Molded"
            If (dButtMountingAngle - TOL) >= -0.0001 And (dButtMountingAngle - TOL) <= 0.0001 Or _
                (dButtMountingAngle - TOL) >= (PI - 0.0001) And (dButtMountingAngle - TOL) <= (PI + 0.0001) Then 'no knuckle
                dNOA = 1.570796
            Else 'knuckle
                'use the nose value that applies for knuckles
                dNose = dNose * Cos(dButtMountingAngle / 2)
                If oPhysConn.InsideFaceOfKnuckle = "Base" Then
                    If sRefSide = "Base" Then 'the molded side and inside face of knuckle match
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
                    If sRefSide = "Offset" Then 'the molded side and inside face of knuckle match
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
                    If sRefSide = "Base" Then 'the molded side and inside face of knuckle match
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
                    If sRefSide = "Offset" Then 'the molded side and inside face of knuckle match
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
    
    'set the bevel for both parts
    If dThickness < dThickness2 Then
        dFirstBevel = (dThickness - dNose) / 2
        dAntiFirstBevel = (dThickness - dNose) / 2
        dNRFirstBevel = (dThickness - dNose) / 2 + Abs(dThicknessoffset2 - dThicknessOffset1)
        dNRAntiFirstBevel = dThickness2 - dNose - dNRFirstBevel
    Else
        dFirstBevel = (dThickness2 - dNose) / 2
        dAntiFirstBevel = (dThickness2 - dNose) / 2
        dNRFirstBevel = (dThickness2 - dNose) / 2 + Abs(dThicknessoffset2 - dThicknessOffset1)
        dNRAntiFirstBevel = dThickness - dNose - dNRFirstBevel
    End If
                
    'get the reference side
'    Dim sRefSide As String
'    sRefSide = GetRefSide(oPhysConn.ConnectedObject1)
    If sRefSide = "Base" Then
          pPRL.Add "ReferenceSide", "molded"
          pPRL.Add "NRReferenceSide", "molded"
    ElseIf sRefSide = "Offset" Then
          pPRL.Add "ReferenceSide", "antimolded"
          pPRL.Add "NRReferenceSide", "antimolded"
    Else
          pPRL.Add "ReferenceSide", sRefSide
          pPRL.Add "NRReferenceSide", sRefSide
    End If
          
    If sCategory = gsOneSided Then
          pPRL.Add "ButtCategory", 65536
    Else
          pPRL.Add "ButtCategory", 65537
    End If
                
    pPRL.Add "NonRefPartName", sPartName2
    pPRL.Add "RefPartName", sPartName1
    
    'set the reference part
    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
        pPRL.Add "ReferencePart", 2
    Else
        pPRL.Add "ReferencePart", 1
    End If
                
    'set nose and ref part values
    pPRL.Add "Nose", dNose
    pPRL.Add "NRNose", dNose
    pPRL.Add "NoseOrientationAngle", dNOA
    pPRL.Add "NRNoseOrientationAngle", dNOA
                
    pPRL.Add "RefSideFirstBevelDepth", dFirstBevel
    pPRL.Add "RefSideFirstBevelAngle", dBevelAngle
    pPRL.Add "AntiRefSideFirstBevelDepth", dAntiFirstBevel
    pPRL.Add "AntiRefSideFirstBevelAngle", dAntiBevelAngle
    pPRL.Add "NRRefSideFirstBevelDepth", dNRFirstBevel
    pPRL.Add "NRRefSideFirstBevelAngle", dBevelAngle
    pPRL.Add "NRAntiRefSideFirstBevelDepth", dNRAntiFirstBevel
    pPRL.Add "NRAntiRefSideFirstBevelAngle", dAntiBevelAngle
    pPRL.Add "RefSideSecondBevelDepth", 0
    pPRL.Add "RefSideSecondBevelAngle", 0
    pPRL.Add "AntiRefSideSecondBevelDepth", 0
    pPRL.Add "AntiRefSideSecondBevelAngle", 0
    pPRL.Add "NRRefSideSecondBevelDepth", 0
    pPRL.Add "NRRefSideSecondBevelAngle", 0
    pPRL.Add "NRAntiRefSideSecondBevelDepth", 0
    pPRL.Add "NRAntiRefSideSecondBevelAngle", 0
  
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
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


