VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:11 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          Regis Dechance
'   Creation Date:  Thursday, Oct 26 2005
'   Description:
'       Implementation of the Simple Physical representation of a door/window Symbol.  The symbol places a door/window frame, panel and outer curve
'       that is used to cut the door/window's opening.   This implementation on places the door/window on the outer surface of the supporting element
'       ie the wall or slab when the door/window is to be placed.
'
'      Door/window origin is in the center,  placement positions are calculated relative to this central point.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   26.Oct.2005     SymbolTeam(Paris)       Copyright Information, Header  is added.
'   20 Apr 2007     PL, PW, RD (Paris)      When thickness parameter is 0 set to very small value to allow mating to continue but casement to be
'                                           hidden.  Allow placement to be flush on wall.   Remove commented code and add comments.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Private Sub PlaceWindowPanesSymbol(ByVal pOutputCollection As IJDOutputCollection, ByVal sOutputName As String, coords() As Double, lRowPanes As Long, lColPanes As Long, pT4x4OfTranslation As DT4x4)
    Const METHOD = "PlaceWindowPanesSymbol:"
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim defColl As IJDDefinitionCollection
    Dim WindowPanesDef As IJDSymbolDefinition
    Dim definitionParams As Variant
    Dim pObjectOfOutputCollection As IJDObject
    Dim newEnumArg2 As IJDEnumArgument
    Dim IJDInputsArg2 As IMSSymbolEntities.IJDInputsArg
    Dim IJEditJDArg2 As IJDEditJDArgument
    Dim PC2 As IJDParameterContent
    Dim argument2 As IJDArgument
    Dim ii As Integer
    
    Set pObjectOfOutputCollection = pOutputCollection
    Set defColl = oSymbolFactory.DefinitionCollection(pObjectOfOutputCollection.ResourceManager)
    
    definitionParams = ""
    Set WindowPanesDef = defColl.GetDefinitionByProgId(True, "SimpleDoorAsm.WindowPanes", vbNullString, definitionParams)
    
    Dim WindowPanesSymbol As IMSSymbolEntities.DSymbol
    Set WindowPanesSymbol = oSymbolFactory.PlaceSymbol(WindowPanesDef, pObjectOfOutputCollection.ResourceManager)
    
    Set newEnumArg2 = New DEnumArgument
    Set IJEditJDArg2 = newEnumArg2.IJDEditJDArgument
    
    For ii = 1 To 9
        Set PC2 = New DParameterContent
        Set argument2 = New DArgument
    
        PC2.uomValue = coords(ii)
        PC2.Type = igValue
        PC2.uomType = 1
        ' Feed the Argument
        argument2.Index = ii
        argument2.Entity = PC2
        ' Add the argument to the arg collection
        IJEditJDArg2.SetArg argument2
        Set PC2 = Nothing
        Set argument2 = Nothing
    Next
    
    'nRowPanes argument
    Set PC2 = New DParameterContent
    Set argument2 = New DArgument
    PC2.uomValue = lRowPanes
    PC2.Type = igValue
    PC2.uomType = 1
    argument2.Index = 10
    argument2.Entity = PC2
    IJEditJDArg2.SetArg argument2
    Set PC2 = Nothing
    Set argument2 = Nothing
    
    'nColPanes argument
    Set PC2 = New DParameterContent
    Set argument2 = New DArgument
    PC2.uomValue = lColPanes
    PC2.Type = igValue
    PC2.uomType = 1
    argument2.Index = 11
    argument2.Entity = PC2
    IJEditJDArg2.SetArg argument2
    Set PC2 = Nothing
    Set argument2 = Nothing
    
    
    WindowPanesSymbol.IJDValuesArg.SetValues newEnumArg2
    Set IJDInputsArg2 = WindowPanesSymbol
    IJDInputsArg2.Update
    Set IJDInputsArg2 = Nothing
    Set IJEditJDArg2 = Nothing
    Set newEnumArg2 = Nothing
    Set WindowPanesDef = Nothing
    
    Call pOutputCollection.AddOutput(sOutputName, Geometry_Transform(WindowPanesSymbol, pT4x4OfTranslation))

Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Sub


Sub Run(ByVal pOutputCollection As IJDOutputCollection, ByRef arrayOfInputs(), arrayOfOutputs() As String)
        
    Const METHOD = "Run:"
    On Error GoTo ErrorHandler
    
    Dim pObjectOfOutputCollection As IJDObject
    Set pObjectOfOutputCollection = pOutputCollection
    
    Dim pGeometryFactory As GeometryFactory
    Set pGeometryFactory = New GeometryFactory
    With pGeometryFactory
    
'input taken from catalog data

    Dim Kinematics As Long
    Dim lPush As Long
    
    Dim dDoorXPosition As Double
    Dim dDoorYPosition As Double
    Dim dDoorZPosition As Double
    
    Dim dOpeningRatio As Double
    Dim dWidth As Double
    Dim dHeight As Double
    
    Dim TopFrameDepth As Double
    Dim TopFrameWidth As Double
    Dim CTEdge As Double
    
    Dim LeftFrameDepth As Double
    Dim LeftFrameWidth As Double
    Dim CLEdge As Double
    
    Dim RightFrameDepth As Double
    Dim RightFrameWidth As Double
    Dim CREdge As Double
    
    Dim LowerFrameDepth As Double
    Dim LowerFrameWidth As Double
    Dim CBEdge As Double
    
    Dim lRowPanes As Long
    Dim lColPanes As Long
    Dim coords(1 To 9) As Double
    
   
    'Door construction
    dOpeningRatio = arrayOfInputs(3) ' 0
    Kinematics = arrayOfInputs(4) ' 0
    lPush = arrayOfInputs(5) ' 0
    dHeight = arrayOfInputs(6) ' 800
    dWidth = arrayOfInputs(7) '400
    
    TopFrameDepth = arrayOfInputs(9) '150
    TopFrameWidth = arrayOfInputs(10) '150
    
    LowerFrameDepth = arrayOfInputs(12) ' 50
    LowerFrameWidth = arrayOfInputs(13) ' 10
    
    LeftFrameDepth = arrayOfInputs(15) ' 50
    LeftFrameWidth = arrayOfInputs(16) '10
    
    RightFrameDepth = arrayOfInputs(18) ' 50
    RightFrameWidth = arrayOfInputs(19) '10
    
    lRowPanes = arrayOfInputs(26)
    lColPanes = arrayOfInputs(27)

    
    'Handle 0 values for items that need at least a minimum size (if dim =0, set to 1e-4 not to fail in math)
    If Abs(TopFrameWidth) < 0.00001 Then
        If TopFrameWidth < 0 Then TopFrameWidth = -0.0001 Else TopFrameWidth = 0.0001
    End If
    If Abs(LowerFrameWidth) < 0.00001 Then
        If LowerFrameWidth < 0 Then LowerFrameWidth = -0.0001 Else LowerFrameWidth = 0.0001
    End If
    If Abs(LeftFrameWidth) < 0.00001 Then
        If LeftFrameWidth < 0 Then LeftFrameWidth = -0.0001 Else LeftFrameWidth = 0.0001
    End If
    If Abs(RightFrameWidth) < 0.00001 Then
        If RightFrameWidth < 0 Then RightFrameWidth = -0.0001 Else RightFrameWidth = 0.0001
    End If
    If Abs(TopFrameDepth) < 0.00001 Then
        If TopFrameDepth < 0 Then TopFrameDepth = -0.0001 Else TopFrameDepth = 0.0001
    End If
    If Abs(LowerFrameDepth) < 0.00001 Then
        If LowerFrameDepth < 0 Then LowerFrameDepth = -0.0001 Else LowerFrameDepth = 0.0001
    End If
    If Abs(LeftFrameDepth) < 0.00001 Then
        If LeftFrameDepth < 0 Then LeftFrameDepth = -0.0001 Else LeftFrameDepth = 0.0001
    End If
    If Abs(RightFrameDepth) < 0.00001 Then
        If RightFrameDepth < 0 Then RightFrameDepth = -0.0001 Else RightFrameDepth = 0.0001
    End If
    
    
    Dim dPanelThickness As Double
    
    dPanelThickness = arrayOfInputs(20) ' 2
    
    If Abs(dPanelThickness) < 0.00001 Then
        If dPanelThickness < 0 Then dPanelThickness = -0.0001 Else dPanelThickness = 0.0001
    End If
    
    dDoorXPosition = arrayOfInputs(21) '0
    dDoorYPosition = arrayOfInputs(23) 'dHeight / 2
    dDoorZPosition = -arrayOfInputs(22)
''''    MsgBox " x y z " & Str(dDoorXPosition) & "," & Str(dDoorYPosition) & "," & Str(dDoorZPosition)
    'ugly kludge
    If dOpeningRatio = 1 Then
        dOpeningRatio = 0.999
    Else
    End If
    
    Dim pT4x4OfTranslation As New DT4x4
    ' make the left Lower corner of the mating surface coincident with the origin of the coordinate system
    If True Then
        Dim pVectorOfTranslation As New DVector
''        Call pVectorOfTranslation.Set(-(0 - dWidth / 2 + LeftFrameWidth ), -(0 - dHeight / 2 + LowerFrameWidth ), -(0))
        Call pVectorOfTranslation.Set(-(0 - dWidth / 2), -(0 - dHeight / 2), -(0))
        Call Vector_Permute(pVectorOfTranslation)
        Call pT4x4OfTranslation.LoadIdentity
        Call pT4x4OfTranslation.Translate(pVectorOfTranslation)
''''    MsgBox " vector x y z " & Str(pVectorOfTranslation.x) & "," & Str(pVectorOfTranslation.y) & "," & Str(pVectorOfTranslation.z)
    End If
    
' Set the outputs
    Dim iOutput As Integer
    iOutput = 0
    
    Dim pVectorOfProjection As New DVector
    Call pVectorOfProjection.Set(0, 1, 0)
    Call Vector_Permute(pVectorOfProjection)
    
    Dim dLengthOfProjection As Double
    dLengthOfProjection = dHeight
    ' Place Frames (door/window frame) by 4 lines that are projected to form the side of a Frame)
    ' Door position is in the center of the door, Frame positions are calculated relative to this central point.
    ' Can have an offset between Frame and door edge, usually set by value named CxEdge
    
    '
    ' Port Faces Of Left/Right Frame Strip 1st parallel to door face, 2nd orthogonal away from door, 3rd parallel at Frame width, 4th
    ' orthogonal back towards door.
    '
    Dim pElementsOfCurves As IJElements
    Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition + LeftFrameDepth)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition + LeftFrameDepth, _
                        dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition + LeftFrameDepth)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition + LeftFrameDepth, _
                        dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)))
    iOutput = 1 '"LeftFrameFace_"
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(1), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
  
    'Left Frame Strip Solid
    Dim pComplexString3d As ComplexString3d
    Set pComplexString3d = .ComplexStrings3d.CreateByCurves( _
                Nothing, pElementsOfCurves)

    Dim pProjection3d As Projection3d
    Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
             pComplexString3d, _
             pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, dLengthOfProjection, True)

    iOutput = 2 '"LeftFrameStrip"
    Call pOutputCollection.AddOutput(arrayOfOutputs(2), Geometry_Transform(pProjection3d, pT4x4OfTranslation))

    ' Port Faces of Right Frame Strip
    Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                       dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                       dDoorXPosition + dWidth / 2 - RightFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                       dDoorXPosition + dWidth / 2 - RightFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                       dDoorXPosition + dWidth / 2 - RightFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition + RightFrameDepth)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                       dDoorXPosition + dWidth / 2 - RightFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition + RightFrameDepth, _
                       dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition + RightFrameDepth)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                       dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition + RightFrameDepth, _
                       dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)))
    
    iOutput = 3 '"RightFrameFace_"
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(3), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
    
    ' Right Frame Strip Solid
     Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
                
     Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                      pComplexString3d, _
                                                      pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                      dLengthOfProjection, _
                                                      True)
    
    iOutput = 4 '"RightFrameStrip"
    Call pOutputCollection.AddOutput(arrayOfOutputs(4), Geometry_Transform(pProjection3d, pT4x4OfTranslation))
    
    
    '
    ' Port Faces Of Top/Lower Frame Strip: 1st orthogonal away from door face, 2nd parallel to door face at distance Frame width Cnwidth,
    ' 3rd orthogonal towards door face, 4th parallel to door no offset.
    '
    Call pVectorOfProjection.Set(1, 0, 0)
    Call Vector_Permute(pVectorOfProjection)

    dLengthOfProjection = dWidth - LeftFrameWidth - RightFrameWidth
    
    ' Top Frame strip
    Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition + dHeight / 2 - TopFrameWidth, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition + dHeight / 2 - TopFrameWidth, dDoorZPosition + TopFrameDepth)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition + dHeight / 2 - TopFrameWidth, dDoorZPosition + TopFrameDepth, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition + dHeight / 2, dDoorZPosition + TopFrameDepth)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition + dHeight / 2, dDoorZPosition + TopFrameDepth, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition + dHeight / 2, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition + dHeight / 2 - TopFrameWidth, dDoorZPosition)))
    
    iOutput = 5 '"TopFrameFace_"
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(5), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
    
    ' Top Frame Strip Solid
    Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
                
    Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                      pComplexString3d, _
                                                      pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                      dLengthOfProjection, _
                                                      True)
    
    iOutput = 6 '"TopFrameStrip"
    Call pOutputCollection.AddOutput(arrayOfOutputs(6), Geometry_Transform(pProjection3d, pT4x4OfTranslation))
   
   '  Port faces of Lower Frame strip
    Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition + LowerFrameDepth)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition + LowerFrameDepth, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2 + LowerFrameWidth, dDoorZPosition + LowerFrameDepth)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2 + LowerFrameWidth, dDoorZPosition + LowerFrameDepth, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2 + LowerFrameWidth, dDoorZPosition)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2 + LowerFrameWidth, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 + LeftFrameWidth, dDoorYPosition - dHeight / 2, dDoorZPosition)))
                        
    
    iOutput = 7 '"LowerFrameFace_"
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(7), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
    
    ' Lower Frame Strip Solid
    Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
                
    Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                      pComplexString3d, _
                                                      pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                      dLengthOfProjection, _
                                                      True)
    
    iOutput = 8 '"LowerFrameStrip"
    Call pOutputCollection.AddOutput(arrayOfOutputs(8), Geometry_Transform(pProjection3d, pT4x4OfTranslation))
    
     
    'Faces of first Panel solid
    
    Const PI = 3.14159

    ' Manage type of kinematics swing, slide, rollup
    Dim lNumberOfPanel As Long
    Dim dKinematics As Double
    Dim dSlide As Double
    Dim dRoll As Double
    Dim dEps As Double
    Dim pVectorOfTranslation1 As New DVector
    Dim pVectorOfRotation As New DVector
    Dim pVectorOfTranslation2 As New DVector
    
    Select Case Kinematics:
        Case 0: 'Slide on left side
            lNumberOfPanel = 1
            dKinematics = 0
            dSlide = dOpeningRatio
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, 1, 0)
            Call pVectorOfTranslation2.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
        Case 1: 'Slide on right side
            lNumberOfPanel = 1
            dKinematics = 0
            dSlide = dOpeningRatio
            dRoll = 1
            dEps = -1
            Call pVectorOfTranslation1.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, 1, 0)
            Call pVectorOfTranslation2.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
        Case 2: 'Swing along vertical left axis
            lNumberOfPanel = 1
            dKinematics = 1
            dSlide = 0
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, -1, 0)
            Call pVectorOfTranslation2.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
        Case 3: 'Swing along vertical right axis
            lNumberOfPanel = 1
            dKinematics = 1
            dSlide = 0
            dRoll = 1
            dEps = -1
            Call pVectorOfTranslation1.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, 1, 0)
            Call pVectorOfTranslation2.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
        Case 4: 'Slide up
            lNumberOfPanel = 1
            dKinematics = 0
            dSlide = dOpeningRatio
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1.Set(0, dEps * dRoll * dHeight / 2, 0)
            Call pVectorOfRotation.Set(1, 0, 0)
            Call pVectorOfTranslation2.Set(0, dEps * ((-dHeight / 2) + dHeight * dSlide), 0)
        Case 5: 'Swing up along horizontal axis
            lNumberOfPanel = 1
            dKinematics = 1
            dSlide = 0
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1.Set(0, dEps * dRoll * dHeight / 2, 0)
            Call pVectorOfRotation.Set(-1, 0, 0)
            Call pVectorOfTranslation2.Set(0, dEps * ((-dHeight / 2) + dHeight * dSlide), 0)
        Case 6: 'Roll Up
            lNumberOfPanel = 1
            dKinematics = 0
            dSlide = dOpeningRatio
            dRoll = 1 - dOpeningRatio
            dEps = 1
            Call pVectorOfTranslation1.Set(0, dEps * dRoll * dHeight / 2, 0)
            Call pVectorOfRotation.Set(1, 0, 0)
            Call pVectorOfTranslation2.Set(0, dEps * ((-dHeight / 2) + dHeight * dSlide), 0)
        Case 7: 'Double swing
            
            Dim pVectorOfTranslation1a As New DVector
            Dim pVectorOfTranslation2a As New DVector
            Dim pVectorOfTranslation1b As New DVector
            Dim pVectorOfTranslation2b As New DVector
            
            lNumberOfPanel = 2
            dKinematics = 1
            dSlide = 0
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1a.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, 1, 0)
            Call pVectorOfTranslation2a.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
            dEps = -1
            Call pVectorOfTranslation1b.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, 1, 0)
            Call pVectorOfTranslation2b.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
    End Select
    
    Call Vector_Permute(pVectorOfTranslation1)
    Call Vector_Permute(pVectorOfRotation)
    Call Vector_Permute(pVectorOfTranslation2)
    Call Vector_Permute(pVectorOfTranslation1a)
    Call Vector_Permute(pVectorOfTranslation2a)
    Call Vector_Permute(pVectorOfTranslation1b)
    Call Vector_Permute(pVectorOfTranslation2b)
    
    If lNumberOfPanel = 1 Then
        Dim pT4x4 As New DT4x4
        Call pT4x4.LoadIdentity
        'First translation since origin is on the middle of the door panel
        Call pT4x4.Translate(pVectorOfTranslation1)
        'Rotate along Y of local coordinate system on left side of the door
        Call pT4x4.Rotate((dEps * dKinematics * lPush * -PI / 2) * dOpeningRatio, pVectorOfRotation)
        'Second translate along the x
        Call pT4x4.Translate(pVectorOfTranslation2)
    Else
        Dim pT4x4a As New DT4x4
        Call pT4x4a.LoadIdentity
        Call pT4x4a.Translate(pVectorOfTranslation1a)
        Call pT4x4a.Rotate((dEps * dKinematics * lPush * -PI / 2) * dOpeningRatio, pVectorOfRotation)
        Call pT4x4a.Translate(pVectorOfTranslation2a)
        
        Dim pT4x4b As New DT4x4
        Call pT4x4b.LoadIdentity
        Call pT4x4b.Translate(pVectorOfTranslation1b)
        Call pT4x4b.Rotate((dEps * dKinematics * lPush * PI / 2) * dOpeningRatio, pVectorOfRotation)
        Call pT4x4b.Translate(pVectorOfTranslation2b)
        
    End If
    'compute Panel  (the moving part of a door/window) shape

    If lNumberOfPanel = 1 Then
        Dim pPositionLocalOfPoint(4) As New DPosition
        Call pPositionLocalOfPoint(1).Set(dDoorXPosition - (dWidth / 2) + LeftFrameWidth, dDoorYPosition - dRoll * dHeight / 2 + LowerFrameWidth, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint(1))
        Call pPositionLocalOfPoint(2).Set(dDoorXPosition + (dWidth / 2) - RightFrameWidth, dDoorYPosition - dRoll * dHeight / 2 + LowerFrameWidth, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint(2))
        Call pPositionLocalOfPoint(3).Set(dDoorXPosition + (dWidth / 2) - RightFrameWidth, dDoorYPosition + dRoll * dHeight / 2 - TopFrameWidth, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint(3))
        Call pPositionLocalOfPoint(4).Set(dDoorXPosition - (dWidth / 2) + LeftFrameWidth, dDoorYPosition + dRoll * dHeight / 2 - TopFrameWidth, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint(4))
    Else
        Dim pPositionLocalOfPoint1(8) As New DPosition
        Call pPositionLocalOfPoint1(1).Set(dDoorXPosition - (dWidth / 2) + LeftFrameWidth, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(1))
        Call pPositionLocalOfPoint1(2).Set(dDoorXPosition, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(2))
        Call pPositionLocalOfPoint1(3).Set(dDoorXPosition, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(3))
        Call pPositionLocalOfPoint1(4).Set(dDoorXPosition - (dWidth / 2) + LeftFrameWidth, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(4))
        
        Call pPositionLocalOfPoint1(5).Set(dDoorXPosition, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(5))
        Call pPositionLocalOfPoint1(6).Set(dDoorXPosition + (dWidth / 2) - RightFrameWidth, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(6))
        Call pPositionLocalOfPoint1(7).Set(dDoorXPosition + (dWidth / 2) - RightFrameWidth, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(7))
        Call pPositionLocalOfPoint1(8).Set(dDoorXPosition, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(8))
    End If
    
    'move and rotate Panel points
    Dim i As Integer
    If lNumberOfPanel = 1 Then
        Dim pPositionGlobalOfPoint(4) As IJDPosition
        For i = 1 To 4
            Set pPositionGlobalOfPoint(i) = pT4x4.TransformPosition(pPositionLocalOfPoint(i))
        Next
    Else
        Dim pPositionGlobalOfPoint1(8) As IJDPosition
        For i = 1 To 4
            Set pPositionGlobalOfPoint1(i) = pT4x4a.TransformPosition(pPositionLocalOfPoint1(i))
        Next
        For i = 5 To 8
            Set pPositionGlobalOfPoint1(i) = pT4x4b.TransformPosition(pPositionLocalOfPoint1(i))
        Next
    End If
    
'    Debug.Print "Position 1rst point"
'    Debug.Print Str(pPositionLocalOfPoint(1).x) + " : ";
'    Debug.Print Str(pPositionLocalOfPoint(1).y) + " : ";
'    Debug.Print Str(pPositionLocalOfPoint(1).z)
'    Debug.Print
'    Debug.Print "Position 2nd point"
'    Debug.Print Str(pPositionLocalOfPoint(2).x) + " : ";
'    Debug.Print Str(pPositionLocalOfPoint(2).y) + " : ";
'    Debug.Print Str(pPositionLocalOfPoint(2).z)
'    Debug.Print
'    Debug.Print "Position 3rd point"
'    Debug.Print Str(pPositionLocalOfPoint(3).x) + " : ";
'    Debug.Print Str(pPositionLocalOfPoint(3).y) + " : ";
'    Debug.Print Str(pPositionLocalOfPoint(3).z)
'    Debug.Print
'    Debug.Print "Position 4th point"
'    Debug.Print Str(pPositionLocalOfPoint(4).x) + " : ";
'    Debug.Print Str(pPositionLocalOfPoint(4).y) + " : ";
'    Debug.Print Str(pPositionLocalOfPoint(4).z)
'    Debug.Print


'-------> shape creation
    ' create Panel curves in position from Panel points
    Call pVectorOfProjection.Set(0, 0, 1)
    Call Vector_Permute(pVectorOfProjection)
    dLengthOfProjection = dPanelThickness
    
    If lNumberOfPanel = 1 Then
         'place window panel panes only if lRowPanes > 1 or lColPanes > 1
        If lRowPanes > 1 Or lColPanes > 1 Then
        
            coords(1) = pPositionGlobalOfPoint(1).x
            coords(2) = pPositionGlobalOfPoint(1).y
            coords(3) = pPositionGlobalOfPoint(1).z
            coords(4) = pPositionGlobalOfPoint(2).x
            coords(5) = pPositionGlobalOfPoint(2).y
            coords(6) = pPositionGlobalOfPoint(2).z
            coords(7) = pPositionGlobalOfPoint(4).x
            coords(8) = pPositionGlobalOfPoint(4).y
            coords(9) = pPositionGlobalOfPoint(4).z
            
            Call PlaceWindowPanesSymbol(pOutputCollection, "WindowPanes", coords, lRowPanes, lColPanes, pT4x4OfTranslation)
        Else
            'if panes, don't display panel
            Set pElementsOfCurves = CurvesAsElements( _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint(1).x, pPositionGlobalOfPoint(1).y, pPositionGlobalOfPoint(1).z, _
                            pPositionGlobalOfPoint(2).x, pPositionGlobalOfPoint(2).y, pPositionGlobalOfPoint(2).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint(2).x, pPositionGlobalOfPoint(2).y, pPositionGlobalOfPoint(2).z, _
                            pPositionGlobalOfPoint(3).x, pPositionGlobalOfPoint(3).y, pPositionGlobalOfPoint(3).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint(3).x, pPositionGlobalOfPoint(3).y, pPositionGlobalOfPoint(3).z, _
                            pPositionGlobalOfPoint(4).x, pPositionGlobalOfPoint(4).y, pPositionGlobalOfPoint(4).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint(4).x, pPositionGlobalOfPoint(4).y, pPositionGlobalOfPoint(4).z, _
                            pPositionGlobalOfPoint(1).x, pPositionGlobalOfPoint(1).y, pPositionGlobalOfPoint(1).z))
        
    
            iOutput = 9 '"PanelSolidFace_"
            Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(9), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
        End If
   Else
          'place window panel panes only if lRowPanes > 1 or lColPanes > 1
        If lRowPanes > 1 Or lColPanes > 1 Then
                    
            coords(1) = pPositionGlobalOfPoint1(1).x
            coords(2) = pPositionGlobalOfPoint1(1).y
            coords(3) = pPositionGlobalOfPoint1(1).z
            coords(4) = pPositionGlobalOfPoint1(2).x
            coords(5) = pPositionGlobalOfPoint1(2).y
            coords(6) = pPositionGlobalOfPoint1(2).z
            coords(7) = pPositionGlobalOfPoint1(4).x
            coords(8) = pPositionGlobalOfPoint1(4).y
            coords(9) = pPositionGlobalOfPoint1(4).z
            
            Call PlaceWindowPanesSymbol(pOutputCollection, "WindowPanes1", coords, lRowPanes, lColPanes, pT4x4OfTranslation)
            
            coords(1) = pPositionGlobalOfPoint1(5).x
            coords(2) = pPositionGlobalOfPoint1(5).y
            coords(3) = pPositionGlobalOfPoint1(5).z
            coords(4) = pPositionGlobalOfPoint1(6).x
            coords(5) = pPositionGlobalOfPoint1(6).y
            coords(6) = pPositionGlobalOfPoint1(6).z
            coords(7) = pPositionGlobalOfPoint1(8).x
            coords(8) = pPositionGlobalOfPoint1(8).y
            coords(9) = pPositionGlobalOfPoint1(8).z
            
            Call PlaceWindowPanesSymbol(pOutputCollection, "WindowPanes2", coords, lRowPanes, lColPanes, pT4x4OfTranslation)
        Else
            'if panes, don't display panel
              Dim pElementsOfCurves1 As IJElements
               Set pElementsOfCurves1 = CurvesAsElements( _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint1(1).x, pPositionGlobalOfPoint1(1).y, pPositionGlobalOfPoint1(1).z, _
                            pPositionGlobalOfPoint1(2).x, pPositionGlobalOfPoint1(2).y, pPositionGlobalOfPoint1(2).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint1(2).x, pPositionGlobalOfPoint1(2).y, pPositionGlobalOfPoint1(2).z, _
                            pPositionGlobalOfPoint1(3).x, pPositionGlobalOfPoint1(3).y, pPositionGlobalOfPoint1(3).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint1(3).x, pPositionGlobalOfPoint1(3).y, pPositionGlobalOfPoint1(3).z, _
                            pPositionGlobalOfPoint1(4).x, pPositionGlobalOfPoint1(4).y, pPositionGlobalOfPoint1(4).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint1(4).x, pPositionGlobalOfPoint1(4).y, pPositionGlobalOfPoint1(4).z, _
                            pPositionGlobalOfPoint1(1).x, pPositionGlobalOfPoint1(1).y, pPositionGlobalOfPoint1(1).z))
                            
                Dim pElementsOfCurves2 As IJElements
                Set pElementsOfCurves2 = CurvesAsElements( _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint1(5).x, pPositionGlobalOfPoint1(5).y, pPositionGlobalOfPoint1(5).z, _
                            pPositionGlobalOfPoint1(6).x, pPositionGlobalOfPoint1(6).y, pPositionGlobalOfPoint1(6).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint1(6).x, pPositionGlobalOfPoint1(6).y, pPositionGlobalOfPoint1(6).z, _
                            pPositionGlobalOfPoint1(7).x, pPositionGlobalOfPoint1(7).y, pPositionGlobalOfPoint1(7).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint1(7).x, pPositionGlobalOfPoint1(7).y, pPositionGlobalOfPoint1(7).z, _
                            pPositionGlobalOfPoint1(8).x, pPositionGlobalOfPoint1(8).y, pPositionGlobalOfPoint1(8).z), _
                         .Lines3d.CreateBy2Points(Nothing, _
                            pPositionGlobalOfPoint1(8).x, pPositionGlobalOfPoint1(8).y, pPositionGlobalOfPoint1(8).z, _
                            pPositionGlobalOfPoint1(5).x, pPositionGlobalOfPoint1(5).y, pPositionGlobalOfPoint1(5).z))
    
            iOutput = 9 '"PanelSolidFace_"
            Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(9), pElementsOfCurves1, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
            Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(9), pElementsOfCurves2, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
        End If
    End If
   
   
    ' Pannel Solid (only if no panes)
    If lRowPanes <= 1 And lColPanes <= 1 Then
        If lNumberOfPanel = 1 Then
            Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
                        
            Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                              pComplexString3d, _
                                                              pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                              dLengthOfProjection, _
                                                              True)
            
            
            iOutput = 10 '"PanelSolid_"
            Call pOutputCollection.AddOutput(arrayOfOutputs(10), Geometry_Transform(pProjection3d, pT4x4OfTranslation))
        Else
            Dim pComplexString3d1 As ComplexString3d
            Dim pComplexString3d2 As ComplexString3d
            
            Set pComplexString3d1 = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves1)
            Set pComplexString3d2 = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves2)
            
            Dim pProjection3d1 As Projection3d
            Dim pProjection3d2 As Projection3d
            Set pProjection3d1 = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                              pComplexString3d1, _
                                                              pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                              dLengthOfProjection, _
                                                              True)
                                                              
            Set pProjection3d2 = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                              pComplexString3d2, _
                                                              pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                              dLengthOfProjection, _
                                                              True)
            
            iOutput = 10 '"PanelSolid_"
            Call pOutputCollection.AddOutput(arrayOfOutputs(10), Geometry_Transform(pProjection3d1, pT4x4OfTranslation))
            Call pOutputCollection.AddOutput(arrayOfOutputs(10), Geometry_Transform(pProjection3d2, pT4x4OfTranslation))
        End If
   End If
    
    ' Mandatory mating surface required by equipment in order to be placed on something.  This is calculated by drawing 2 curves around the
    ' outer edge of the door one offset from the other by the absolute value  of the thickness parameter.  This is supposed to be hidden visually
    ' by the frame.
    Dim AbsTopFrameWidth As Double
    AbsTopFrameWidth = Abs(TopFrameWidth)
    Dim AbsLowerFrameWidth As Double
    AbsLowerFrameWidth = Abs(LowerFrameWidth)
    Dim AbsLeftFrameWidth As Double
    AbsLeftFrameWidth = Abs(LeftFrameWidth)
    Dim AbsRightFrameWidth As Double
    AbsRightFrameWidth = Abs(RightFrameWidth)
  

        ' reverse the order of the points 8-7, 6-5, 4-3, 2-1
'''        Set pElementsOfCurves = CurvesAsElements( _
'''            Line_Permute( _
'''                .Lines3d.CreateBy2Points(Nothing, _
'''                    dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
'''                    dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)), _
'''            Line_Permute( _
'''                .Lines3d.CreateBy2Points(Nothing, _
'''                    dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
'''                    dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
'''            Line_Permute( _
'''                .Lines3d.CreateBy2Points(Nothing, _
'''                    dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
'''                    dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
'''            Line_Permute( _
'''                .Lines3d.CreateBy2Points(Nothing, _
'''                    dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
'''                    dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)))
        Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                .Lines3d.CreateBy2Points(Nothing, _
                    dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                    dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                .Lines3d.CreateBy2Points(Nothing, _
                    dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
                    dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                .Lines3d.CreateBy2Points(Nothing, _
                    dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
                    dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                .Lines3d.CreateBy2Points(Nothing, _
                    dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                    dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)))

        Set pElementsOfCurves1 = CurvesAsElements( _
                        Line_Permute( _
                            .Lines3d.CreateBy2Points(Nothing, _
                                dDoorXPosition - (dWidth / 2) + AbsLeftFrameWidth, dDoorYPosition - (dHeight / 2) + AbsLowerFrameWidth, dDoorZPosition, _
                                dDoorXPosition + (dWidth / 2) - AbsRightFrameWidth, dDoorYPosition - (dHeight / 2) + AbsLowerFrameWidth, dDoorZPosition)), _
                        Line_Permute( _
                            .Lines3d.CreateBy2Points(Nothing, _
                                dDoorXPosition + (dWidth / 2) - AbsRightFrameWidth, dDoorYPosition - (dHeight / 2) + AbsLowerFrameWidth, dDoorZPosition, _
                                dDoorXPosition + (dWidth / 2) - AbsRightFrameWidth, dDoorYPosition + (dHeight / 2) - AbsTopFrameWidth, dDoorZPosition)), _
                        Line_Permute( _
                            .Lines3d.CreateBy2Points(Nothing, _
                                dDoorXPosition + (dWidth / 2) - AbsRightFrameWidth, dDoorYPosition + (dHeight / 2) - AbsTopFrameWidth, dDoorZPosition, _
                                dDoorXPosition - (dWidth / 2) + AbsLeftFrameWidth, dDoorYPosition + (dHeight / 2) - AbsTopFrameWidth, dDoorZPosition)), _
                        Line_Permute( _
                            .Lines3d.CreateBy2Points(Nothing, _
                                dDoorXPosition - (dWidth / 2) + AbsLeftFrameWidth, dDoorYPosition + (dHeight / 2) - AbsTopFrameWidth, dDoorZPosition, _
                                dDoorXPosition - (dWidth / 2) + AbsLeftFrameWidth, dDoorYPosition - (dHeight / 2) + AbsLowerFrameWidth, dDoorZPosition)))

        Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
        
        Dim pMatingPlane As IJPlane
        Set pMatingPlane = .Planes3d.CreateByOuterBdry(pOutputCollection.ResourceManager, pComplexString3d)
        
        Dim pInternalBoundary As ComplexString3d
        Set pInternalBoundary = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves1)
        
        Call pMatingPlane.AddHole(pInternalBoundary)
                
        iOutput = 11 '"DefaultSurface"
        Call pOutputCollection.AddOutput(arrayOfOutputs(11), Geometry_Transform(pMatingPlane, pT4x4OfTranslation))
    
    
    End With
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
