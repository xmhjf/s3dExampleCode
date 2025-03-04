VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtWeld2Parm"
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

Const m_ParameterRuleProgid As String = "PhysConnRules.ButtWeld2Parm"
Const m_ParameterRuleName As String = "PhysConnRules.ButtWeld2Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\PhysConnRules\ButtWeld2Parm.cls"

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
  pOH.SetOutput "NoseOrientationAngle"      'User code "n"
  pOH.SetOutput "NRNoseOrientationAngle"
  pOH.SetOutput "RefSideFirstBevelDepth"        'User code "A"
  pOH.SetOutput "NRRefSideFirstBevelDepth"      'User code "A" for second part
  pOH.SetOutput "RefSideFirstBevelAngle"        'User code "a"
  pOH.SetOutput "NRRefSideFirstBevelAngle"      'User code "a" for second part
  pOH.SetOutput "AntiRefSideFirstBevelDepth"    'User code "B"
  pOH.SetOutput "NRAntiRefSideFirstBevelDepth"  'User code "B" for second part
  pOH.SetOutput "AntiRefSideFirstBevelAngle"    'User code "b"
  pOH.SetOutput "NRAntiRefSideFirstBevelAngle"  'User code "b" for second part
  pOH.SetOutput "ButtCategory"
  pOH.SetOutput "RefPartName", imsARGUMENT_IS_BSTR
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR
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
    Dim dFirstBevel As Double
    Dim dNRFirstBevel As Double
    Dim dBaseDelta As Double
    Dim dOffsetDelta As Double
    Dim dBaseDelta2 As Double
    Dim dOffsetDelta1 As Double
    Dim dOffsetDelta2 As Double
    
    'get answers from selector
    Dim sCategory As String
    sCategory = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "Category")
    Dim sWorkcenter As String
    sWorkcenter = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "Workcenter")
    Dim sUpside As String
    sUpside = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "UpSide")
    Dim sFirstWeldingSide As String
    sFirstWeldingSide = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "FirstWeldingSide")
    Dim sChamferType As String
    sChamferType = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "ChamferType")
          
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
                
                dBaseDelta2 = oPlate2.OffsetToBaseFace - oPlate1.OffsetToBaseFace(oPlate2)
                dOffsetDelta1 = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToOffsetFace
                dOffsetDelta2 = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToOffsetFace(oPlate2)

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
                dBaseDelta2 = 0
                dOffsetDelta1 = 0
                dOffsetDelta2 = 0
                 
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
                    dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace
                    dThickness2 = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace + dDelta
                   
            ElseIf oPlate1.OffsetToOffsetFace(oPlate2) <= oPlate2.OffsetToOffsetFace And _
                oPlate1.OffsetToBaseFace(oPlate2) >= oPlate2.OffsetToBaseFace Then
                    dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate1.OffsetToBaseFace(oPlate2)
                    dThickness2 = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate1.OffsetToBaseFace(oPlate2) + dDelta
                   
            ElseIf oPlate1.OffsetToOffsetFace(oPlate2) >= oPlate2.OffsetToOffsetFace And _
                oPlate1.OffsetToBaseFace(oPlate2) >= oPlate2.OffsetToBaseFace Then
                    dThickness = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToBaseFace(oPlate2)
                    dThickness2 = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToBaseFace(oPlate2) + dDelta
                  
            ElseIf oPlate1.OffsetToOffsetFace(oPlate2) >= oPlate2.OffsetToOffsetFace And _
                oPlate1.OffsetToBaseFace(oPlate2) <= oPlate2.OffsetToBaseFace Then
                    dThickness = oPlate2.OffsetToOffsetFace(oPlate1) - oPlate2.OffsetToBaseFace(oPlate1)
                    dThickness2 = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate1.OffsetToBaseFace(oPlate2) + dDelta
                   
            Else
                If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                    dThickness = oPhysConn.Object2Thickness
                    dThickness2 = oPhysConn.Object1Thickness
                Else
                    dThickness = oPhysConn.Object1Thickness
                    dThickness2 = oPhysConn.Object2Thickness
                End If
            End If
            
            If oPhysConn.Object1Thickness < oPhysConn.Object2Thickness Then
                sPartName1 = oNamedItem1.Name
                sPartName2 = oNamedItem2.Name
            Else    '...If oPhysConn.Object2Thickness < oPhysConn.Object1Thickness Then
                sPartName1 = oNamedItem2.Name
                sPartName2 = oNamedItem1.Name
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
       
    Select Case sCategory
       Case gsOneSided
          dBevelAngle = 0.392699
          Select Case sWorkcenter
             Case gsMachine1
                If dThickness < 0.012 Then
                   dNose = 0.003
                   pPRL.Add "RootGap", -0.001
                   pPRL.Add "NRRootGap", -0.001
                   
                ElseIf dThickness > 0.012 And dThickness <= 0.022 Then
                   dNose = 0.003
                   pPRL.Add "RootGap", -0.001
                   pPRL.Add "NRRootGap", -0.001
                   
                ElseIf dThickness > 0.022 And dThickness <= 0.026 Then
                   dNose = 0.004
                   pPRL.Add "RootGap", -0.001
                   pPRL.Add "NRRootGap", -0.001
                   
                ElseIf dThickness > 0.026 And dThickness <= 0.03001 Then
                   dNose = 0.005
                   pPRL.Add "RootGap", -0.001
                   pPRL.Add "NRRootGap", -0.001
                
                Else
                   dNose = 0.005
                   pPRL.Add "RootGap", -0.001
                   pPRL.Add "NRRootGap", -0.001
                End If
                
             Case gsMachine2
                If dThickness < 0.012 Then
                   dNose = 0.0035
                   pPRL.Add "RootGap", -0.0012
                   pPRL.Add "NRRootGap", -0.0012
                   
                ElseIf dThickness > 0.012 And dThickness <= 0.022 Then
                   dNose = 0.0035
                   pPRL.Add "RootGap", -0.0012
                   pPRL.Add "NRRootGap", -0.0012
                   
                ElseIf dThickness > 0.022 And dThickness <= 0.026 Then
                   dNose = 0.0045
                   pPRL.Add "RootGap", -0.0012
                   pPRL.Add "NRRootGap", -0.0012
                   
                ElseIf dThickness > 0.026 And dThickness <= 0.03001 Then
                   dNose = 0.0055
                   pPRL.Add "RootGap", -0.0012
                   pPRL.Add "NRRootGap", -0.0012
                
                Else
                   dNose = 0.0055
                   pPRL.Add "RootGap", -0.0012
                   pPRL.Add "NRRootGap", -0.0012
                End If
          End Select
       
       Case gsTwoSided
          dBevelAngle = 0.6544985
          Select Case sWorkcenter
             Case gsMachine1
                   dNose = 0.007
                   pPRL.Add "RootGap", -0.001
                   pPRL.Add "NRRootGap", -0.001
             Case gsMachine2
                   dNose = 0.006
                   pPRL.Add "RootGap", -0.0011
                   pPRL.Add "NRRootGap", -0.0011
          End Select
    
    End Select
   
    'get the reference side
    Dim sRefSide As String
    sRefSide = GetRefSide(oPhysConn.ConnectedObject1)
   
    'Calculation of angles
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
       
    'section which calculates the bevels based on welding side and offsets
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
                    Else 'molded side is offset
                        dNOA = 1.570796 - dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                    End If
                Else 'the inside face of knuckle is offset
                    If sRefSide = "Offset" Then 'the molded side and inside face of knuckle match
                        'reset the bevel angles based on the knuckle; same angle used for ref and non ref
                        dNOA = 1.570796 + dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                    Else 'molded side is base
                        dNOA = 1.570796 - dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - 1.570796 + dNOA
                    End If
                End If
            End If

            Select Case sChamferType
                Case "Obj1Base", "Obj2Base"
                    'Set the bevel for both parts
                    dBevel = dThickness - dNose
                    dNRBevel = dBevel
                    dNRNose = dThickness2 - dNRBevel
                     
                Case "None"
                    'Set the bevel for both parts
                    dBevel = dThickness - dNose
                    If oPhysConn.Object1Thickness < oPhysConn.Object2Thickness Then
                        If dBaseDelta < 0 Then
                            dNRBevel = dBevel + dBaseDelta
                        Else
                            dNRBevel = dBevel + Abs(dBaseDelta)
                        End If
                    Else
                        If dBaseDelta2 < 0 Then
                            dNRBevel = dBevel + dBaseDelta2
                        Else
                            dNRBevel = dBevel + Abs(dBaseDelta2)
                        End If
                    End If
                    dNRNose = dThickness2 - dNRBevel
                
                Case "Obj1Offset", "Obj2Offset"
                    'If the offset is on the offset side, and the bevel is on the base side
                    'because the welding side is set to molded, find the FirstBevel, then calculate
                    'the NRFirstBevel
                    dBevel = dThickness - dNose
                    dNRBevel = dBevel + Abs(dBaseDelta)
                    dNRNose = dThickness2 - dNRBevel
                    
                Case "Obj1BaseObj2Offset", "Obj1OffsetObj2Base", "Obj1Double", "Obj2Double"
                    'in these cases, chamfers are applied so that both plates are reduced to the
                    'same thickness
                    dNRNose = dNose
                    dBevel = dThickness - dNose
                    dNRBevel = dBevel

            End Select

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
                        dBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                    Else 'molded side is offset
                        dNOA = 1.570796 - dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - 1.570796 + dNOA
                    End If
                Else 'it is offset
                    If sRefSide = "Offset" Then 'the molded side and inside face of knuckle match
                        'reset the bevel angles based on the knuckle; same angle used for ref and non ref
                        dNOA = 1.570796 + dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - 1.570796 + dNOA
                    Else 'molded side is base
                        dNOA = 1.570796 - dButtMountingAngle / 2
                        dBevelAngle = dBevelAngle - (1.570796 - (PI - dNOA))
                    End If
                End If
            End If

            Select Case sChamferType
                'If the offset is on the base side, and the bevel is on the antimolded side
                'because the welding side is set to antimolded, find the FirstBevel, then calculate
                'the NRFirstBevel

                Case "Obj1Base", "Obj2Base"
                    dBevel = dThickness - dNose
                    dNRBevel = dBevel + Abs(dOffsetDelta)
                    dNRNose = dThickness2 - dNRBevel
                    
                Case "None"
                    dBevel = dThickness - dNose
                    If oPhysConn.Object1Thickness < oPhysConn.Object2Thickness Then
                        If dOffsetDelta1 > 0 Then
                            dNRBevel = dBevel - dOffsetDelta1
                        Else
                            dNRBevel = dBevel + Abs(dOffsetDelta)
                        End If
                    Else
                        If dOffsetDelta2 > 0 Then
                            dNRBevel = dBevel - dOffsetDelta2
                        Else
                            dNRBevel = dBevel + Abs(dOffsetDelta2)
                        End If
                    End If
                    dNRNose = dThickness2 - dNRBevel
                    
                Case "Obj1BaseObj2Offset", "Obj1OffsetObj2Base", "Obj1Double", "Obj2Double"
                    'in these cases, chamfers are applied so that both plates are reduced to the
                    'same thickness
                    dNRNose = dNose
                    dBevel = dThickness - dNose
                    dNRBevel = dBevel

                'If the offset is on the offset side, and the bevel is on the antimolded side
                'because the welding side is set to antimolded, find the FirstBevel, then calculate
                'the NRFirstBevel
                Case "Obj1Offset", "Obj2Offset"
                     dBevel = dThickness - dNose
                     dNRBevel = dBevel
                     dNRNose = dNose + Abs(dBaseDelta)
            End Select
                        
    End Select
    
    If sCategory = gsOneSided Then
        pPRL.Add "ButtCategory", 65536
    Else
        pPRL.Add "ButtCategory", 65537
    End If


    'Set nose and reference part values
    pPRL.Add "Nose", dNose
    pPRL.Add "NRNose", dNRNose
    pPRL.Add "NoseOrientationAngle", dNOA
    pPRL.Add "NRNoseOrientationAngle", dNOA
    
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
    
    pPRL.Add "NonRefPartName", sPartName2
    pPRL.Add "RefPartName", sPartName1
    
    'set the reference part
    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
        pPRL.Add "ReferencePart", 2
    Else
        pPRL.Add "ReferencePart", 1
    End If
           
    Select Case sFirstWeldingSide
       Case "Molded"
              pPRL.Add "RefSideFirstBevelDepth", dBevel
              pPRL.Add "RefSideFirstBevelAngle", dBevelAngle
              pPRL.Add "AntiRefSideFirstBevelDepth", 0
              pPRL.Add "AntiRefSideFirstBevelAngle", 0
              pPRL.Add "NRRefSideFirstBevelDepth", dNRBevel
              pPRL.Add "NRRefSideFirstBevelAngle", dBevelAngle
              pPRL.Add "NRAntiRefSideFirstBevelDepth", 0
              pPRL.Add "NRAntiRefSideFirstBevelAngle", 0
       Case "AntiMolded"
              pPRL.Add "RefSideFirstBevelDepth", 0
              pPRL.Add "RefSideFirstBevelAngle", 0
              pPRL.Add "AntiRefSideFirstBevelDepth", dBevel
              pPRL.Add "AntiRefSideFirstBevelAngle", dBevelAngle
              pPRL.Add "NRRefSideFirstBevelDepth", 0
              pPRL.Add "NRRefSideFirstBevelAngle", 0
              pPRL.Add "NRAntiRefSideFirstBevelDepth", dNRBevel
              pPRL.Add "NRAntiRefSideFirstBevelAngle", dBevelAngle
       End Select

  
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

