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
    
    Dim CTwidth As Double
    Dim CTThickness As Double
    Dim CTEdge As Double
    
    Dim CLwidth As Double
    Dim CLThickness As Double
    Dim CLEdge As Double
    
    Dim CRwidth As Double
    Dim CRThickness As Double
    Dim CREdge As Double
    
    Dim CBwidth As Double
    Dim CBThickness As Double
    Dim CBEdge As Double
    
    
    'Door construction
    dOpeningRatio = arrayOfInputs(3) ' 0
    Kinematics = arrayOfInputs(4) ' 0
    lPush = arrayOfInputs(5) ' 0
    dHeight = arrayOfInputs(6) ' 800
    dWidth = arrayOfInputs(7) '400
    
    CTwidth = arrayOfInputs(9) '150
    CTThickness = arrayOfInputs(10) '150
    CTEdge = arrayOfInputs(11)
    
    CBwidth = arrayOfInputs(13) ' 50
    CBThickness = arrayOfInputs(14) ' 10
    CBEdge = arrayOfInputs(15)
    
    CLwidth = arrayOfInputs(17) ' 50
    CLThickness = arrayOfInputs(18) '10
    CLEdge = arrayOfInputs(19)
    
    CRwidth = arrayOfInputs(21) ' 50
    CRThickness = arrayOfInputs(22) '10
    CREdge = arrayOfInputs(23)
    
    'Handle Null casements (if dim =0, set to 1e-4 not to fail in math)
    If Abs(CTThickness) < 0.00001 Then
        If CTThickness < 0 Then CTThickness = -0.0001 Else CTThickness = 0.0001
    End If
    If Abs(CBThickness) < 0.00001 Then
        If CBThickness < 0 Then CBThickness = -0.0001 Else CBThickness = 0.0001
    End If
    If Abs(CLThickness) < 0.00001 Then
        If CLThickness < 0 Then CLThickness = -0.0001 Else CLThickness = 0.0001
    End If
    If Abs(CRThickness) < 0.00001 Then
        If CRThickness < 0 Then CRThickness = -0.0001 Else CRThickness = 0.0001
    End If
    If Abs(CTwidth) < 0.00001 Then
        If CTwidth < 0 Then CTwidth = -0.0001 Else CTwidth = 0.0001
    End If
    If Abs(CBwidth) < 0.00001 Then
        If CBwidth < 0 Then CBwidth = -0.0001 Else CBwidth = 0.0001
    End If
    If Abs(CLwidth) < 0.00001 Then
        If CLwidth < 0 Then CLwidth = -0.0001 Else CLwidth = 0.0001
    End If
    If Abs(CRwidth) < 0.00001 Then
        If CRwidth < 0 Then CRwidth = -0.0001 Else CRwidth = 0.0001
    End If
    
    
    Dim dPannelThickness As Double
    
    dPannelThickness = arrayOfInputs(24) ' 2
    
    dDoorXPosition = arrayOfInputs(25) '0
    dDoorYPosition = arrayOfInputs(27) 'dHeight / 2
    dDoorZPosition = -arrayOfInputs(26)
    
    'ugly kludge
    If dOpeningRatio = 1 Then
        dOpeningRatio = 0.999
    Else
    End If
    
    Dim pT4x4OfTranslation As New DT4x4
    ' make the left bottom corner of the mating surface coincident with the origin of the coordinate system
    If True Then
        Dim pVectorOfTranslation As New DVector
        Call pVectorOfTranslation.Set(-(0 - dWidth / 2 - CLThickness - CLEdge), -(0 - dHeight / 2 - CBThickness - CBEdge), -(0))
        Call Vector_Permute(pVectorOfTranslation)
        Call pT4x4OfTranslation.LoadIdentity
        Call pT4x4OfTranslation.Translate(pVectorOfTranslation)
    End If
    
' Set the outputs
    Dim iOutput As Integer
    iOutput = 0
    
    Dim pVectorOfProjection As New DVector
    Call pVectorOfProjection.Set(0, 1, 0)
    Call Vector_Permute(pVectorOfProjection)
    
    Dim dLengthOfProjection As Double
    dLengthOfProjection = dHeight + CTEdge + CBEdge
    ' Place casements (door/window frame) by 4 lines that are projected to form the side of a casement)
    ' Door position is in the center of the door, casement positions are calculated relative to this central point.
    ' Can have an offset between casement and door edge, usually set by value named CxEdge
    
    '
    ' Port Faces Of Left/Right Casement Strip 1st parallel to door face, 2nd orthogonal away from door, 3rd parallel at casement width, 4th
    ' orthogonal back towards door.
    '
    Dim pElementsOfCurves As IJElements
    Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CLwidth)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CLwidth, _
                        dDoorXPosition - dWidth / 2 - CLEdge, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CLwidth)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CLwidth, _
                        dDoorXPosition - dWidth / 2 - CLEdge, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition)))
    iOutput = iOutput + 1
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(iOutput), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
  
    'Left Casement Strip Solid
    Dim pComplexString3d As ComplexString3d
    Set pComplexString3d = .ComplexStrings3d.CreateByCurves( _
                Nothing, pElementsOfCurves)

    Dim pProjection3d As Projection3d
    Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
             pComplexString3d, _
             pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, dLengthOfProjection, True)

    iOutput = iOutput + 1
    Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(pProjection3d, pT4x4OfTranslation))

    ' Port Faces of Right Casement Strip
    Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                       dDoorXPosition + dWidth / 2 + CREdge, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition, _
                       dDoorXPosition + dWidth / 2 + CREdge + CRThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                       dDoorXPosition + dWidth / 2 + CREdge + CRThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition, _
                       dDoorXPosition + dWidth / 2 + CREdge + CRThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CRwidth)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                       dDoorXPosition + dWidth / 2 + CREdge + CRThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CRwidth, _
                       dDoorXPosition + dWidth / 2 + CREdge, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CRwidth)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                       dDoorXPosition + dWidth / 2 + CREdge, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CRwidth, _
                       dDoorXPosition + dWidth / 2 + CREdge, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition)))
    
    iOutput = iOutput + 1
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(iOutput), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
    
    ' Right Casement Strip Solid
     Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
                
     Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                      pComplexString3d, _
                                                      pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                      dLengthOfProjection, _
                                                      True)
    
    iOutput = iOutput + 1
    Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(pProjection3d, pT4x4OfTranslation))
    
    
    '
    ' Port Faces Of Top/bottom Casement Strip: 1st orthogonal away from door face, 2nd parallel to door face at distance casement width Cnwidth,
    ' 3rd orthogonal towards door face, 4th parallel to door no offset.
    '
   '  Port faces of Bottom Casement strip
    Call pVectorOfProjection.Set(1, 0, 0)
    Call Vector_Permute(pVectorOfProjection)

    dLengthOfProjection = dWidth + CLEdge + CREdge + CLThickness + CRThickness
    
    Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CBwidth)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition + CBwidth, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge - CBThickness, dDoorZPosition + CBwidth)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge - CBThickness, dDoorZPosition + CBwidth, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge - CBThickness, dDoorZPosition)), _
            Line_Permute( _
                    .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge - CBThickness, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition - dHeight / 2 - CBEdge, dDoorZPosition)))
                        
    
    iOutput = iOutput + 1
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(iOutput), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
    
    ' Bottom Casement Strip Solid
    Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
                
    Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                      pComplexString3d, _
                                                      pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                      dLengthOfProjection, _
                                                      True)
    
    iOutput = iOutput + 1
    Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(pProjection3d, pT4x4OfTranslation))
    
    ' Top Casement strip
    Set pElementsOfCurves = CurvesAsElements( _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition + dHeight / 2 + CTEdge + CTThickness, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition + dHeight / 2 + CTEdge + CTThickness, dDoorZPosition + CTwidth)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition + dHeight / 2 + CTEdge + CTThickness, dDoorZPosition + CTwidth, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition + dHeight / 2 + CTEdge, dDoorZPosition + CTwidth)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition + dHeight / 2 + CTEdge, dDoorZPosition + CTwidth, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition + dHeight / 2 + CTEdge, dDoorZPosition)), _
            Line_Permute( _
                     .Lines3d.CreateBy2Points(Nothing, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition + dHeight / 2 + CTEdge, dDoorZPosition, _
                        dDoorXPosition - dWidth / 2 - CLEdge - CLThickness, dDoorYPosition + dHeight / 2 + CTEdge + CTThickness, dDoorZPosition)))
    
    iOutput = iOutput + 1
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(iOutput), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
    
    ' Top Casement Strip Solid
    Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
                
    Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                      pComplexString3d, _
                                                      pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                      dLengthOfProjection, _
                                                      True)
    
    iOutput = iOutput + 1
    Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(pProjection3d, pT4x4OfTranslation))
     
    'Faces of first pannel solid
    
    Const PI = 3.14159

    ' Manage type of kinematics swing, slide, rollup
    Dim lNumberOfPannel As Long
    Dim dKinematics As Double
    Dim dSlide As Double
    Dim dRoll As Double
    Dim dEps As Double
    Dim pVectorOfTranslation1 As New DVector
    Dim pVectorOfRotation As New DVector
    Dim pVectorOfTranslation2 As New DVector
    
    Select Case Kinematics:
        Case 0: 'Slide on left side
            lNumberOfPannel = 1
            dKinematics = 0
            dSlide = dOpeningRatio
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, 1, 0)
            Call pVectorOfTranslation2.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
        Case 1: 'Slide on right side
            lNumberOfPannel = 1
            dKinematics = 0
            dSlide = dOpeningRatio
            dRoll = 1
            dEps = -1
            Call pVectorOfTranslation1.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, 1, 0)
            Call pVectorOfTranslation2.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
        Case 2: 'Swing along vertical left axis
            lNumberOfPannel = 1
            dKinematics = 1
            dSlide = 0
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, -1, 0)
            Call pVectorOfTranslation2.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
        Case 3: 'Swing along vertical right axis
            lNumberOfPannel = 1
            dKinematics = 1
            dSlide = 0
            dRoll = 1
            dEps = -1
            Call pVectorOfTranslation1.Set(dEps * -dWidth / 2, 0, 0)
            Call pVectorOfRotation.Set(0, 1, 0)
            Call pVectorOfTranslation2.Set(dEps * ((dWidth / 2) - dWidth * dSlide), 0, 0)
        Case 4: 'Slide up
            lNumberOfPannel = 1
            dKinematics = 0
            dSlide = dOpeningRatio
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1.Set(0, dEps * dRoll * dHeight / 2, 0)
            Call pVectorOfRotation.Set(1, 0, 0)
            Call pVectorOfTranslation2.Set(0, dEps * ((-dHeight / 2) + dHeight * dSlide), 0)
        Case 5: 'Swing up along horizontal axis
            lNumberOfPannel = 1
            dKinematics = 1
            dSlide = 0
            dRoll = 1
            dEps = 1
            Call pVectorOfTranslation1.Set(0, dEps * dRoll * dHeight / 2, 0)
            Call pVectorOfRotation.Set(-1, 0, 0)
            Call pVectorOfTranslation2.Set(0, dEps * ((-dHeight / 2) + dHeight * dSlide), 0)
        Case 6: 'Roll Up
            lNumberOfPannel = 1
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
            
            lNumberOfPannel = 2
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
    
    If lNumberOfPannel = 1 Then
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
    'compute pannel  (the moving part of a door/window) shape

    If lNumberOfPannel = 1 Then
        Dim pPositionLocalOfPoint(4) As New DPosition
        Call pPositionLocalOfPoint(1).Set(dDoorXPosition - dWidth / 2, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint(1))
        Call pPositionLocalOfPoint(2).Set(dDoorXPosition + dWidth / 2, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint(2))
        Call pPositionLocalOfPoint(3).Set(dDoorXPosition + dWidth / 2, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint(3))
        Call pPositionLocalOfPoint(4).Set(dDoorXPosition - dWidth / 2, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint(4))
    Else
        Dim pPositionLocalOfPoint1(8) As New DPosition
        Call pPositionLocalOfPoint1(1).Set(dDoorXPosition - dWidth / 2, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(1))
        Call pPositionLocalOfPoint1(2).Set(dDoorXPosition, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(2))
        Call pPositionLocalOfPoint1(3).Set(dDoorXPosition, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(3))
        Call pPositionLocalOfPoint1(4).Set(dDoorXPosition - dWidth / 2, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(4))
        
        Call pPositionLocalOfPoint1(5).Set(dDoorXPosition, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(5))
        Call pPositionLocalOfPoint1(6).Set(dDoorXPosition + dWidth / 2, dDoorYPosition - dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(6))
        Call pPositionLocalOfPoint1(7).Set(dDoorXPosition + dWidth / 2, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(7))
        Call pPositionLocalOfPoint1(8).Set(dDoorXPosition, dDoorYPosition + dRoll * dHeight / 2, dDoorZPosition)
        Call Position_Permute(pPositionLocalOfPoint1(8))
    End If
    
    'move and rotate pannel points
    Dim i As Integer
    If lNumberOfPannel = 1 Then
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
    ' create pannel curves in position from pannel points
    Call pVectorOfProjection.Set(0, 0, 1)
    Call Vector_Permute(pVectorOfProjection)
    dLengthOfProjection = dPannelThickness
    
    If lNumberOfPannel = 1 Then
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
    

    iOutput = iOutput + 1
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(iOutput), pElementsOfCurves, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
   
   Else
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

    iOutput = iOutput + 1
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(iOutput), pElementsOfCurves1, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
    Call ProjectSplitComplexString1(pOutputCollection, arrayOfOutputs(iOutput), pElementsOfCurves2, pVectorOfProjection, dLengthOfProjection, pT4x4OfTranslation)
    End If
   
   
    ' Pannel Solid
    If lNumberOfPannel = 1 Then
        Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
                    
        Set pProjection3d = .Projections3d.CreateByCurve(pOutputCollection.ResourceManager, _
                                                          pComplexString3d, _
                                                          pVectorOfProjection.x, pVectorOfProjection.y, pVectorOfProjection.z, _
                                                          dLengthOfProjection, _
                                                          True)
        
        
        iOutput = iOutput + 1
        Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(pProjection3d, pT4x4OfTranslation))
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
        
        iOutput = iOutput + 1
        Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(pProjection3d1, pT4x4OfTranslation))
        Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(pProjection3d2, pT4x4OfTranslation))
    End If
       
    
    ' Mandatory mating surface required by equipment in order to be placed on something.  This is calculated by drawing 2 curves around the
    ' outer edge of the door one offset from the other by the absolute value  of the thickness parameter.
    Dim AbsCTThickness As Double
    AbsCTThickness = Abs(CTThickness)
    Dim AbsCBThickness As Double
    AbsCBThickness = Abs(CBThickness)
    Dim AbsCLThickness As Double
    AbsCLThickness = Abs(CLThickness)
    Dim AbsCRThickness As Double
    AbsCRThickness = Abs(CRThickness)
  

        ' reverse the order of the points 8-7, 6-5, 4-3, 2-1
        Set pElementsOfCurves = CurvesAsElements( _
                        Line_Permute( _
                            .Lines3d.CreateBy2Points(Nothing, _
                                dDoorXPosition - (dWidth / 2) - AbsCLThickness, dDoorYPosition - (dHeight / 2) - AbsCBThickness, dDoorZPosition, _
                                dDoorXPosition - (dWidth / 2) - AbsCLThickness, dDoorYPosition + (dHeight / 2) + AbsCTThickness, dDoorZPosition)), _
                        Line_Permute( _
                            .Lines3d.CreateBy2Points(Nothing, _
                                dDoorXPosition - (dWidth / 2) - AbsCLThickness, dDoorYPosition + (dHeight / 2) + AbsCTThickness, dDoorZPosition, _
                                dDoorXPosition + (dWidth / 2) + AbsCRThickness, dDoorYPosition + (dHeight / 2) + AbsCTThickness, dDoorZPosition)), _
                        Line_Permute( _
                            .Lines3d.CreateBy2Points(Nothing, _
                                dDoorXPosition + (dWidth / 2) + AbsCRThickness, dDoorYPosition + (dHeight / 2) + AbsCTThickness, dDoorZPosition, _
                                dDoorXPosition + (dWidth / 2) + AbsCRThickness, dDoorYPosition - (dHeight / 2) - AbsCBThickness, dDoorZPosition)), _
                        Line_Permute( _
                            .Lines3d.CreateBy2Points(Nothing, _
                                dDoorXPosition + (dWidth / 2) + AbsCRThickness, dDoorYPosition - (dHeight / 2) - AbsCBThickness, dDoorZPosition, _
                                dDoorXPosition - (dWidth / 2) - AbsCLThickness, dDoorYPosition - (dHeight / 2) - AbsCBThickness, dDoorZPosition)))

        Set pComplexString3d = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves)
        
        Dim pMatingPlane As IJPlane
        Set pMatingPlane = .Planes3d.CreateByOuterBdry(pOutputCollection.ResourceManager, pComplexString3d)
        
        Set pElementsOfCurves1 = CurvesAsElements( _
            Line_Permute( _
                .Lines3d.CreateBy2Points(Nothing, _
                    dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                    dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                .Lines3d.CreateBy2Points(Nothing, _
                    dDoorXPosition + dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition, _
                    dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                .Lines3d.CreateBy2Points(Nothing, _
                    dDoorXPosition + dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
                    dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition)), _
            Line_Permute( _
                .Lines3d.CreateBy2Points(Nothing, _
                    dDoorXPosition - dWidth / 2, dDoorYPosition + dHeight / 2, dDoorZPosition, _
                    dDoorXPosition - dWidth / 2, dDoorYPosition - dHeight / 2, dDoorZPosition)))
        
        Dim pInternalBoundary As ComplexString3d
        Set pInternalBoundary = .ComplexStrings3d.CreateByCurves(Nothing, pElementsOfCurves1)
        
        Call pMatingPlane.AddHole(pInternalBoundary)
                
        iOutput = iOutput + 1
        Call pOutputCollection.AddOutput(arrayOfOutputs(iOutput), Geometry_Transform(pMatingPlane, pT4x4OfTranslation))
    
    
    End With
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

