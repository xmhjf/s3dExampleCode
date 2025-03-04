Attribute VB_Name = "basShapesPlace3D"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   FileName:       basShapesPlace3D.bas
'   ProgID:         basShapesPlace3D
'   Author:         3D Config Team
'   Compiler:       HL
'   Creation Date:  Thursday, Jan 23, 2003
'   Description:    This module contains lower level primitive create
'                   functions and procedures to create primitive objects.
'
'                   placeCircularTorus
'                   placeConicalTankHead
'                   placeEllipticalDish
'                   placeKnuckleRadiusTankHead
'                   placeMiterCutCircTorus
'                   placeMiterCutRectTorus
'                   placePlane
'                   placePrismaticNEdgeTorus
'                   placeProjectedPolygonFromVertices
'                   PlaceSlopedBottomCylinder
'                   placeSnout
'                   placeSphericalDish
'                   placeTransformableCone
'                   placeTransformableCylinder
'                   placeTransformableSphere
'                   placeTruncNEdgePrism
'
'   Change History:
'   dd.mmm.yyyy   who           change description
'   -----------   ---           ------------------
'   23.Jan.2003   HL      Collected code from 3D Config Team into this module
'   28.Jan.2003   HL      Added function placeCircularTorus
'   29.Jan.2003   HL      Added Conditional Compilation to have debug info
'   04.Feb.2003   HL      Changed all return type to IMSSymbolEntities.DSymbol
'   04.Feb.2003   HL      Combined placeProjectedPolygonFromVertices and
'                               placeTransformableProjectedPolygonFromVertices to one.
'   04.Feb.2003   Doug Hempel   Added two new functions placeConicalTankHead and
'                               placeKnuckleRadiusTankHead
'   06.Feb.2003   HL      Changed return type of all functions to their specific types.
'                               Make transformation matrix to be an optional parameter.
'                               Change all Dposition/Dvector/Ori to pass by ref.
'   13.Feb.2003   KVanWinkle    Added sample of Place Function Calls.
'   24.Mar.2003   JG    Added placeMiterCutCircTorus function.
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is Updated.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Const MODULE = "basShapesPlace3D"
'Set debug mode to true (-1), or false (0)
#Const INDEBUG = 0

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placeCircularTorus
'   Author:     HL
'   Inputs:
'               Output collection object
'               Cross Section center
'               Cross Section Normal vector
'               Revolving vector
'               Minor Radius
'               Sweep Angle
'               Is Capped?
'               Transformation Matrix
'
'   Outputs:
'               Circular Torus
'
'   Description:
'               This function places a circular torus.  Originally,
'               Tori3d.CreateByAxisMajorMinorRadiusSweep was used.  However,
'               it does not have capped functionality.  Therefore, a circle
'               and its revolution is used to create the circular torus.
'
'   Example of Call:
'
'               Set CreateCircularTorus = placeCircularTorus(objOutputColl, posStart, yVector, vecRev, dblMinorR, dblSweepAngle, True, tmxMatrix)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   28.Jan.2003     HL        Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeCircularTorus(ByVal objOutputColl As Object, _
                                   posStart As IJDPosition, _
                                   vecNormal As IJDVector, _
                                   vecRev As IJDVector, _
                                   ByVal dblRadius As Double, _
                                   ByVal dblSweepAngle As Double, _
                                   ByVal blnIsCapped As Boolean, _
                                   Optional tmxMatrix As IngrGeom3D.IJDT4x4) _
                                   As IngrGeom3D.Revolution3d

    Dim CenterPos As IJDPosition
    Dim objGeomFactory As IngrGeom3D.GeometryFactory
    Dim objCircle As IngrGeom3D.Circle3d
    Dim objRevolution As IngrGeom3D.Revolution3d
        
    Set CenterPos = New DPosition
    CenterPos.Set 0, 0, 0
    
    Set objGeomFactory = New IngrGeom3D.GeometryFactory
    
'    #If INDEBUG Then
'        MsgBox "Sweep Angle is " & dblSweepAngle
'    #End If
    
    'Use a circle and revolve it
    Set objCircle = objGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
        posStart.x, posStart.y, posStart.z, vecNormal.x, vecNormal.y, vecNormal.z, _
        dblRadius)
        
    'objCircle.Transform tmxMatrix
    
    Set objRevolution = objGeomFactory.Revolutions3d.CreateByCurve(objOutputColl.ResourceManager, _
        objCircle, vecRev.x, vecRev.y, vecRev.z, CenterPos.x, CenterPos.y, CenterPos.z, dblSweepAngle, _
        blnIsCapped)

    If Not (tmxMatrix Is Nothing) Then
        objRevolution.Transform tmxMatrix
    End If

    Set placeCircularTorus = objRevolution
    
    Set objRevolution = Nothing
    Set objCircle = Nothing
    Set objGeomFactory = Nothing
    Set CenterPos = Nothing
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Function:   placeEllipticalDish
'   Author:     Doug Hempel
'   Inputs:
'               m_OutputColl          => The collection object
'               dblDishDiameter       => Dish diameter.
'               dblMinorMajorRatio    => Ratio of minor to major axis of ellipse.
'               blnIsCapped           => Boolean indicating whether to cap (close) the open end of
'                                        the elliptical dish.  Currently capping functionality is
'                                        not implemented.
'               tmxTMatrix            => Transformation matrix used to transform the dish.
'
'   Outputs:
'               Object - Returns a revolution object that is a spherical dish
'
'   Description:
'               Builds an elliptical dish from the given dish diameter and dish
'               height and ratio. This function is a low level dish builder and
'               assumes that the diameter and ratio value have already been defined.
'               This function is not normally called directly by symbol programmers,
'               but rather is called by the higher level dish create.
'
'   Example of Call:
'
'               Set CreateDish = placeEllipticalDish(objOutputColl, dblDiameter, dblHeight / (dblDiameter / 2), blnIsCapped, xform)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   31.Jan.2003     Doug Hempel     Implemented dish capping capability
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeEllipticalDish _
( _
   ByVal m_outputColl As Object, _
   ByVal dblDishDiameter As Double, _
   ByVal dblMinorMajorRatio As Double, _
   ByVal blnIsCapped As Boolean, _
   Optional ByRef tmxTMatrix As IngrGeom3D.IJDT4x4 _
) As IngrGeom3D.Revolution3d

   Dim objGeomFactory          As IngrGeom3D.GeometryFactory
   Dim objRevolution           As IngrGeom3D.Revolution3d
   Dim objEllipticalArc        As IngrGeom3D.EllipticalArc3d
   Dim objCappingLine          As IngrGeom3D.Line3d
   Dim objComplexString        As IngrGeom3D.ComplexString3d
   Dim posStartofComplexString As IJDPosition
   Dim posCappingLineStart     As IJDPosition
   Dim posCappingLineEnd       As IJDPosition
   Dim posDishBuildOrigin      As IJDPosition
   Dim objTmpCollection        As Collection
   
   Set objGeomFactory = New IngrGeom3D.GeometryFactory
   Set objTmpCollection = New Collection
   Set posStartofComplexString = New DPosition
   Set posCappingLineStart = New DPosition
   Set posCappingLineEnd = New DPosition
   Set posDishBuildOrigin = New DPosition


'  --- build the dish in standard position and orientation ---
   posDishBuildOrigin.Set 0, 0, 0

'  --- build an ellipse ---
   Set objEllipticalArc = objGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, posDishBuildOrigin.x, posDishBuildOrigin.y, posDishBuildOrigin.z, 0, 0, 1, 0, dblDishDiameter / 2, 0, dblMinorMajorRatio, degreeToRadian(270), degreeToRadian(90))
   
   If (blnIsCapped = True) Then
'     --- add dome ellipse to temporary collection ---
      objTmpCollection.Add objEllipticalArc
      
'     --- build the capping line ---
      posCappingLineStart.Set posDishBuildOrigin.x, posDishBuildOrigin.y, posDishBuildOrigin.z
      posCappingLineEnd.Set posDishBuildOrigin.x, posDishBuildOrigin.y - (dblDishDiameter / 2), posDishBuildOrigin.z
      Set objCappingLine = objGeomFactory.Lines3d.CreateBy2Points(Nothing, posCappingLineStart.x, posCappingLineStart.y, posCappingLineStart.z, posCappingLineEnd.x, posCappingLineEnd.y, posCappingLineEnd.z)
      objTmpCollection.Add objCappingLine

'     --- create complex string from temporary collection of capping line and dome arc ---
      posStartofComplexString.Set posDishBuildOrigin.x + dblMinorMajorRatio * dblDishDiameter / 2, posDishBuildOrigin.y, posDishBuildOrigin.z
      Set objComplexString = PlaceTrCString(posStartofComplexString, objTmpCollection)

'     --- revolve the ellipse about x axis to form a elliptical dish ---
      Set objRevolution = objGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, objComplexString, 1, 0, 0, posDishBuildOrigin.x, posDishBuildOrigin.y, posDishBuildOrigin.z, degreeToRadian(360), True)
   
   Else
'     --- revolve the ellipse about x axis to form a elliptical dish ---
      Set objRevolution = objGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, objEllipticalArc, 1, 0, 0, posDishBuildOrigin.x, posDishBuildOrigin.y, posDishBuildOrigin.z, degreeToRadian(360), True)
   End If
   If Not (tmxTMatrix Is Nothing) Then
        '  --- transform the object to right position and orientation ---
        objRevolution.Transform tmxTMatrix
   End If
   
'  --- set function return value ---
   Set placeEllipticalDish = objRevolution
       
'  --- cleanup ---
   Set objEllipticalArc = Nothing
   Set objRevolution = Nothing
   Set objGeomFactory = Nothing
   Set posDishBuildOrigin = Nothing
   If (blnIsCapped = True) Then
      Set objComplexString = Nothing
      Set objCappingLine = Nothing
   End If
   Set objTmpCollection = Nothing
   Set posCappingLineStart = Nothing
   Set posCappingLineEnd = Nothing
   Set posStartofComplexString = Nothing
   
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Function:   placeMiterCutRectTorus
'   Author:     JG
'   Inputs:
'               Output Collection object
'               Sweep Angle
'               array of points
'               Number of segments
'               Transformation Matrix
'
'   Outputs:
'
'
'   Description:
'               This function places
'
'   Example of Call:
'
'               Set CreateRectangularTorus = placeMiterCutRectTorus(objOutputColl, revAngle, arrayPoints, intSegments, tmxMatrix)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   07.Jan.2003     JG      Function creation
'
'   26.Mar.2003     JG      Added ability choose whether
'                                   to cap the output or not.
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeMiterCutRectTorus(ByVal objOutputColl As Object, _
                                        dblSweepAngle As Double, _
                                        ByRef arrTest() As IJDPosition, _
                                        dblNumOfSeg As Integer, _
                                        tmxMatrix As IJDT4x4, _
                                        blnIsCapped As Boolean) _
                                        As IMSSymbolEntities.DSymbol
    
    ' Get or create the definition
    Dim defColl             As IJDDefinitionCollection
    Dim boxDef              As IJDSymbolDefinition
    Dim definitionParams    As Variant
    Dim oEnv                As IMSSymbolEntities.DSymbol
    Dim newEnumArg          As IJDEnumArgument
    Dim IJEditJDArg         As IJDEditJDArgument
    Dim PC                  As IJDParameterContent
    Dim argument            As IJDArgument
    Dim iCount              As Integer
    Dim oSymbolFactory      As IMSSymbolEntities.DSymbolEntitiesFactory

    Set oSymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set defColl = oSymbolFactory.DefinitionCollection(objOutputColl.ResourceManager)
    definitionParams = "5"
    Set boxDef = defColl.GetDefinitionByProgId(True, "MiterCutPrsmTorus.MiterCutPrsmServices", vbNullString, definitionParams)

    Set oEnv = oSymbolFactory.PlaceSymbol(boxDef, objOutputColl.ResourceManager)
    Set oSymbolFactory = Nothing
    Set boxDef = Nothing
    Set defColl = Nothing

    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument
    
    Dim i As Integer
    
'Start array entries

    'Sweep angle *****************************
    iCount = 1

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = dblSweepAngle
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing

    'nseg *************************************
    iCount = 2

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = dblNumOfSeg
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'matrix ***********************************
    iCount = 3

    'Pass the Transformation Matrix
    For i = 0 To 15
        
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = tmxMatrix.IndexValue(i)
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
        iCount = iCount + 1
            
    Next i

    'Points ***********************************
    'iCount = 16

    For i = 0 To 3
    
        'Set X
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = arrTest(i).x
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
        iCount = iCount + 1
        
        'Set Y
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = arrTest(i).y
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
        iCount = iCount + 1
        
        'Set Z
        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = arrTest(i).z
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
        iCount = iCount + 1
        
    Next i
    
    Dim intIsCapped As Integer
    If blnIsCapped = True Then
        intIsCapped = 1
    Else
        intIsCapped = 0
    End If
    
    'Cap torus  ******************************
    iCount = 31

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = intIsCapped
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    
'End array entries
    
    oEnv.IJDValuesArg.SetValues newEnumArg
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = oEnv
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    Set placeMiterCutRectTorus = oEnv
    Set oEnv = Nothing
    Set oSymbolFactory = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Function:   placeMiterCutCircTorus
'   Author:     JG
'   Inputs:
'               Output Collection object
'               Sweep Angle
'               array of points
'               Number of segments
'               Transformation Matrix
'
'   Outputs:
'
'
'   Description:
'               This function places
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   17.Feb.2003     JG      Function creation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeMiterCutCircTorus(ByVal objOutputColl As Object, _
                                        dblSweepAngle As Double, _
                                        dblRIns As Double, _
                                        dblROut As Double, _
                                        dblNumOfSeg As Integer, _
                                        dblScalingFactor As Double, _
                                        Optional tmxMatrix As IJDT4x4, _
                                        Optional bolIsCapped As Boolean) _
                                       As IMSSymbolEntities.DSymbol
    
    ' Get or create the definition
    Dim defColl             As IJDDefinitionCollection
    Dim boxDef              As IJDSymbolDefinition
    Dim definitionParams    As Variant
    Dim oEnv                As IMSSymbolEntities.DSymbol
    Dim newEnumArg          As IJDEnumArgument
    Dim IJEditJDArg         As IJDEditJDArgument
    Dim PC                  As IJDParameterContent
    Dim argument            As IJDArgument
    Dim iCount              As Integer
    Dim oSymbolFactory      As IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set oSymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set defColl = oSymbolFactory.DefinitionCollection(objOutputColl.ResourceManager)
    definitionParams = "5"
    Set boxDef = defColl.GetDefinitionByProgId(True, "MiterCutCircTorus.MiterCutCircServices", vbNullString, definitionParams)

    Set oEnv = oSymbolFactory.PlaceSymbol(boxDef, objOutputColl.ResourceManager)
    Set oSymbolFactory = Nothing
    Set boxDef = Nothing
    Set defColl = Nothing

    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument
    
    Dim i As Integer
    If tmxMatrix Is Nothing Then
        Set tmxMatrix = New DT4x4
        tmxMatrix.LoadIdentity
    End If
    
'Start array entries

    'Sweep angle *****************************
    iCount = 1

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = dblSweepAngle
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing

    'nseg *************************************
    iCount = 2

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = dblNumOfSeg
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'RINS *********************************
    iCount = 3

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = dblRIns
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
    
    'ROUT *********************************
    iCount = 4

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = dblROut
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Reduction Ratio **********************
    iCount = 5

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = dblScalingFactor
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Reduction Ratio **********************
    
    'matrix ***********************************
    iCount = 6
    'Pass the Transformation Matrix
    For i = 0 To 15

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = tmxMatrix.IndexValue(i)
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing

        iCount = iCount + 1

    Next i

        Set PC = New DParameterContent
        Set argument = New DArgument
        If bolIsCapped = True Then
            PC.uomValue = 1
        Else
            PC.uomValue = 0
        End If
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing

    
'End array entries
    
    oEnv.IJDValuesArg.SetValues newEnumArg
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = oEnv
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    Set placeMiterCutCircTorus = oEnv
    Set oEnv = Nothing
    Set oSymbolFactory = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placePrismaticNEdgeTorus
'   Author:     KV and HL
'   Inputs:
'               Output collection object
'               Number of vertices =>  This parameter is used so that we could
'                                      re-use an array of points partially.
'               Array of points Dposition defines the top surface
'               Vector to revolve
'               Center to revolve
'               Sweep Angle
'               Transformation Matrix
'
'   Outputs:
'               Irregular prismatic Torus
'
'   Description:
'               This function places a prismatic torus.  The points must be
'               defined in order (clockwise or counter clockwise)
'
'   Example of Call:
'
'               Set CreatePrismaticTorus = placePrismaticNedgeTorus(objOutputColl, intVertices, arrPts, vecRevVector, dblAngle, tmxForm)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   19.Dec.2002     KV & Hong    Created this function.
'   14.Jan.2003     JG      Added Transformation Matrix
'   14.Jan.2003     HL        Added checking existence of the matrix
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function placePrismaticNedgeTorus(ByVal objOutputColl As Object, _
                        ByVal nVertices As Integer, _
                        ByRef arrPrismPoints() As IJDPosition, _
                        ByRef vecRevVector As IJDVector, _
                        ByVal swAngle As Double, _
                        Optional ByRef tmxMatrix As IJDT4x4, _
                        Optional ByVal blnIsCapped As Boolean) As IngrGeom3D.Revolution3d

    Dim CenterPos As IJDPosition
    
    Dim i As Integer
    Dim objGeomFactory As IngrGeom3D.GeometryFactory
    Dim oLinestr As IngrGeom3D.LineString3d
    Dim oRevolution As IngrGeom3D.Revolution3d
    Dim arrayPoints() As Double
    
    ReDim arrayPoints(3 * nVertices + 2) As Double
    
    Set CenterPos = New DPosition
    CenterPos.Set 0, 0, 0
    
    For i = 0 To nVertices - 1
        arrayPoints(3 * i) = arrPrismPoints(i).x
        arrayPoints(3 * i + 1) = arrPrismPoints(i).y
        arrayPoints(3 * i + 2) = arrPrismPoints(i).z
    Next i
    
    arrayPoints(3 * nVertices) = arrPrismPoints(0).x
    arrayPoints(3 * nVertices + 1) = arrPrismPoints(0).y
    arrayPoints(3 * nVertices + 2) = arrPrismPoints(0).z
        
    Set objGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLinestr = New IngrGeom3D.LineString3d
    
    Set oLinestr = objGeomFactory.LineStrings3d.CreateByPoints(Nothing, nVertices + 1, arrayPoints)

    Set oRevolution = objGeomFactory.Revolutions3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    oLinestr, _
                                                    vecRevVector.x, vecRevVector.y, vecRevVector.z, _
                                                    CenterPos.x, CenterPos.y, CenterPos.z, _
                                                    swAngle, blnIsCapped)

    If Not (tmxMatrix Is Nothing) Then
        oRevolution.Transform tmxMatrix
    End If

    Set placePrismaticNedgeTorus = oRevolution
    Set oLinestr = Nothing
    Set objGeomFactory = Nothing
    Set CenterPos = Nothing
    Set oRevolution = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placeTruncNEdgePrism
'   Author:     HL
'   Inputs:
'               Output collection object
'               Number of vertices  => This parameter is used so that we could
'                                      re-use an array of points partially.
'               Array of points Dposition defines the top surface
'               Array of points Dposition defines the bottom surface
'               Transformation Matrix
'
'   Outputs:
'               Irregular prism based on the inputs
'
'   Description:
'               This function places a truncated N Edge prism.  The number of
'               vertices must match in both surfaces.  The function will put
'               in side surfaces.  It is suggested that point0 of top surface
'               should align with the point0 of the bottom surface, and in
'               the same order (clockwise/counter clockwise) of the points
'               defined in sequence.
'
'   Example of Call:
'
'               Set CreateBasePrism = placeTruncNEdgePrism(objOutputColl, intVertices, arrTopPts, arrBottomPts, tmxForm)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   18.Dec.2002     HL        Created this function.
'   14.Jan.2002     HL        Added Transformation Matrix
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeTruncNEdgePrism(ByVal objOutputColl As Object, _
                        ByVal nVertices As Integer, _
                        ByRef topSurfacePoints() As IJDPosition, _
                        ByRef bottomSurfacePoints() As IJDPosition, _
                        Optional ByRef tmxMatrix As IJDT4x4, _
                        Optional ByVal blnIsCapped As Boolean) As IngrGeom3D.RuledSurface3d
    
    Dim i As Integer
    Dim arrayTopPoints() As Double
    Dim arrayBottomPoints() As Double
    
    Dim objGeomFactory As IngrGeom3D.GeometryFactory
    Dim oLinestrTop As IngrGeom3D.LineString3d
    Dim oLinestrBottom As IngrGeom3D.LineString3d
            
    Dim oNEdgePrism As IngrGeom3D.RuledSurface3d
        
    ReDim arrayTopPoints(3 * nVertices + 2) As Double
    ReDim arrayBottomPoints(3 * nVertices + 2) As Double
    
    ''These are to obtain the coordinates of various vertices.
    For i = 0 To nVertices - 1
        arrayTopPoints(3 * i) = topSurfacePoints(i).x
        arrayTopPoints(3 * i + 1) = topSurfacePoints(i).y
        arrayTopPoints(3 * i + 2) = topSurfacePoints(i).z
        
        arrayBottomPoints(3 * i) = bottomSurfacePoints(i).x
        arrayBottomPoints(3 * i + 1) = bottomSurfacePoints(i).y
        arrayBottomPoints(3 * i + 2) = bottomSurfacePoints(i).z
    Next i
    
    arrayTopPoints(3 * nVertices) = topSurfacePoints(0).x
    arrayTopPoints(3 * nVertices + 1) = topSurfacePoints(0).y
    arrayTopPoints(3 * nVertices + 2) = topSurfacePoints(0).z
    
    arrayBottomPoints(3 * nVertices) = bottomSurfacePoints(0).x
    arrayBottomPoints(3 * nVertices + 1) = bottomSurfacePoints(0).y
    arrayBottomPoints(3 * nVertices + 2) = bottomSurfacePoints(0).z

'    For i = 0 To 3 * nVertices + 2
'        logError (arrayTopPoints(i))
'    Next i
'
'    For i = 0 To 3 * nVertices + 2
'        logError (arrayBottomPoints(i))
'    Next i

    Set objGeomFactory = New IngrGeom3D.GeometryFactory

    Set oLinestrTop = objGeomFactory.LineStrings3d.CreateByPoints(Nothing, nVertices + 1, arrayTopPoints)
    Set oLinestrBottom = objGeomFactory.LineStrings3d.CreateByPoints(Nothing, nVertices + 1, arrayBottomPoints)

    Set oNEdgePrism = objGeomFactory.RuledSurfaces3d.CreateByCurves(objOutputColl.ResourceManager, oLinestrBottom, oLinestrTop, blnIsCapped)

    If Not (tmxMatrix Is Nothing) Then
        oNEdgePrism.Transform tmxMatrix
    End If

    Set placeTruncNEdgePrism = oNEdgePrism

    Dim objLineString As IJDObject
    Set objLineString = oLinestrTop
    Set oLinestrTop = Nothing
    objLineString.Remove

    Set objLineString = oLinestrBottom
    Set oLinestrBottom = Nothing
    objLineString.Remove
    
    Set oNEdgePrism = Nothing
    Set objGeomFactory = Nothing
     
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placeTransformableSphere
'   Author:     CYW
'   Inputs:
'               Output collection object
'               Radius as Double
'               transformation matrix
'
'   Outputs:
'               Sphere object
'
'   Description:
'               This function places a persistent sphere at 0,0,0.  Must use
'               transformation matrix to move it.
'
'   Example of call:
'
'               Set CreateSphere = placeTransformableSphere(objOutputColl, dblDiameter / 2, tmxForm)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   07.Jan.2003     CYW  Function creation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeTransformableSphere(ByVal objOutputColl As Object, _
                        radius As Double, _
                        Optional ByRef tmxMatrix As IJDT4x4) _
                        As IngrGeom3D.Sphere3d
        
    Dim objSphere   As IngrGeom3D.Sphere3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
        
    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set objSphere = geomFactory.Spheres3d.CreateByCenterRadius( _
                                                    objOutputColl.ResourceManager, _
                                                    0, 0, 0, _
                                                    radius, _
                                                    True)
    If Not (tmxMatrix Is Nothing) Then
        objSphere.Transform tmxMatrix
    End If

    Set placeTransformableSphere = objSphere
    Set objSphere = Nothing
    Set geomFactory = Nothing
     
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placeSphericalDish
'   Author:     Doug Hempel
'   Inputs:
'               m_OutputColl          => The collection object
'               posDishArcRadiusPnt   => Position of the radius point of the arc that is revolved
'                                        to build the spherical dish.
'               posDishArcStartPnt    => Position of the start point of the arc that is revolved
'                                        to build the spherical dish.
'               posDishArcEndPnt      => Position of the end point of the arc that is revolved
'                                        to build the spherical dish.
'               blnIsCapped           => Boolean indicating whether to cap (close) the open end
'                                        the spherical dish.  Currently capping functionality is
'                                        not implemented.
'               tmxTMatrix           <=> Transformation matrix used to transform the dish.
'
'   Outputs:
'               Object - Returns a revolution object that is a spherical dish
'
'   Description:
'               Builds a spherical dish from positions of an arc that is revolved
'               to create the dish.  This function is a low level dish builder and
'               assumes that the arc position values have already been defined.
'               This function is not normally called directly by symbol programmers,
'               but rather is called by the higher level dish create.
'
'   Example of Call:
'
'               Set CreateDish = placeSphericalDish(objOutputColl, posDishArcRadiusPnt, posDishArcStartPnt, posDishArcEndPnt, blnIsCapped, xform)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   31.Jan.2003     Doug Hempel     Implemented dish capping capability
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeSphericalDish _
( _
   ByVal m_outputColl As Object, _
   posDishArcRadiusPnt As IJDPosition, _
   posDishArcStartPnt As IJDPosition, _
   posDishArcEndPnt As IJDPosition, _
   ByVal blnIsCapped As Boolean, _
   Optional ByRef tmxTMatrix As IngrGeom3D.IJDT4x4 _
) As IngrGeom3D.Revolution3d
   
   Dim objGeomFactory          As IngrGeom3D.GeometryFactory
   Dim objRevolution           As IngrGeom3D.Revolution3d
   Dim objComplexString        As IngrGeom3D.ComplexString3d
   Dim objArc3d                As IngrGeom3D.Arc3d
   Dim objCappingLine          As IngrGeom3D.Line3d
   Dim objTmpCollection        As Collection
   Dim posStartofComplexString As IJDPosition
   Dim posDishBuildOrigin      As IJDPosition
   Dim posCappingLineStart     As IJDPosition
   Dim posCappingLineEnd       As IJDPosition
   
   Set objGeomFactory = New IngrGeom3D.GeometryFactory
   Set objTmpCollection = New Collection
   Set posStartofComplexString = New DPosition
   Set posCappingLineStart = New DPosition
   Set posCappingLineEnd = New DPosition
   Set posDishBuildOrigin = New DPosition
   
'  --- build the dish in standard position and orientation ---
   posDishBuildOrigin.Set 0, 0, 0
   
'  --- if dish is to be capped, construct capping line ---
   If (blnIsCapped = True) Then
      posCappingLineStart.Set posDishBuildOrigin.x, posDishBuildOrigin.y, posDishBuildOrigin.z
      posCappingLineEnd.Set posDishArcStartPnt.x, posDishArcStartPnt.y, posDishArcStartPnt.z
      Set objCappingLine = objGeomFactory.Lines3d.CreateBy2Points(Nothing, posCappingLineStart.x, posCappingLineStart.y, posCappingLineStart.z, posCappingLineEnd.x, posCappingLineEnd.y, posCappingLineEnd.z)
      objTmpCollection.Add objCappingLine
   End If
   
'  --- build the dome arc ---
   Set objArc3d = PlaceTrArcByCenter(posDishArcStartPnt, posDishArcEndPnt, posDishArcRadiusPnt)
   
   If (blnIsCapped = True) Then
'     --- add dome arc to temporary collection
      objTmpCollection.Add objArc3d

'     --- create a complex string from temporary collection of capping line and dome arc ---
      posStartofComplexString.Set posCappingLineStart.x, posCappingLineStart.y, posCappingLineStart.z
      Set objComplexString = PlaceTrCString(posStartofComplexString, objTmpCollection)
   
'     --- revolve the complex string about x axis to form a spherical dish ---
      Set objRevolution = objGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, objComplexString, 1, 0, 0, posDishBuildOrigin.x, posDishBuildOrigin.y, posDishBuildOrigin.z, degreeToRadian(360), True)
   
   Else
'     --- revolve the arc about x axis to form a spherical dish ---
      Set objRevolution = objGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, objArc3d, 1, 0, 0, posDishBuildOrigin.x, posDishBuildOrigin.y, posDishBuildOrigin.z, degreeToRadian(360), True)
   End If
   
   If Not (tmxTMatrix Is Nothing) Then
      '  --- transform the object to right position and orientation ---
      objRevolution.Transform tmxTMatrix
   End If
   
'  --- set function return value ---
   Set placeSphericalDish = objRevolution
   
'  --- cleanup ---
   Set objArc3d = Nothing
   Set objRevolution = Nothing  'dev note: check to see if this should be done another way
   Set objGeomFactory = Nothing
   Set posDishBuildOrigin = Nothing
   If (blnIsCapped = True) Then
      Set objComplexString = Nothing
      Set objCappingLine = Nothing
   End If
   Set objTmpCollection = Nothing
   Set posCappingLineStart = Nothing
   Set posCappingLineEnd = Nothing
   Set posStartofComplexString = Nothing
   
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placeSnout
'   Author:     Doug Hempel
'   Inputs:
'               m_OutputColl     => The collection object
'               posBottomCenter  => Bottom (left) position of snout.
'               posTopCenter     => Top (right) position of snout.
'               dblBottomCenter  => Bottom (left) diameter of snout.
'               dblTopCenter     => Top (right) diameter of snout.
'               blnIsCapped      => Boolean indicating whether to cap (close) the open ends
'                                   the snout.  Currently capping functionality is
'                                   not implemented.
'               tmxTMatrix       => Transformation matrix used to transform the snout.
'
'   Outputs:
'               Object - Returns a ruled surface object that is a snout.
'
'   Description:
'               Builds a snout from the given top and bottom face positions and
'               the diameters of the faces.  This function is a low level snout
'               builder and assumes that the low level parameter values have
'               been defined at a higher level.  This function is not normally
'               called directly by symbol programmers, but rather is called by
'               the higher level snout create.
'
'   Example of Call:
'
'               Set CreateSnout = placeSnout(objOutputColl, posBottomCenter, posTopCenter, dblBottomDiameter, dblTopDiameter, True, xform)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeSnout _
( _
   ByVal m_outputColl As Object, _
   posBottomCenter As IJDPosition, _
   posTopCenter As IJDPosition, _
   ByVal dblBottomDia As Double, _
   ByVal dblTopDia As Double, _
   ByVal isCapped As Boolean, _
   Optional ByRef tmxTMatrix As IngrGeom3D.IJDT4x4 _
) As IngrGeom3D.RuledSurface3d
   
   Dim objGeomFactory As IngrGeom3D.GeometryFactory
   Dim objRuledSurface As IngrGeom3D.RuledSurface3d
   Dim objTopCircle As IngrGeom3D.Circle3d
   Dim objBottomCircle As IngrGeom3D.Circle3d
   
   Dim i As Integer
   
   Set objGeomFactory = New IngrGeom3D.GeometryFactory

'  --- create bottom circle ---
   Set objBottomCircle = objGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, posBottomCenter.x, posBottomCenter.y, posBottomCenter.z, -1, 0, 0, dblBottomDia / 2)

'  --- create top circle ---
   Set objTopCircle = objGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, posTopCenter.x, posTopCenter.y, posTopCenter.z, 1, 0, 0, dblTopDia / 2)
   
'  --- create a ruled surface between top and bottom circle ---
   Set objRuledSurface = objGeomFactory.RuledSurfaces3d.CreateByCurves(m_outputColl.ResourceManager, objTopCircle, objBottomCircle, isCapped)

   If Not (tmxTMatrix Is Nothing) Then
    '  --- transform the object to desired position and orientation ---
       objRuledSurface.Transform tmxTMatrix
   End If
   
'  --- set function value ---
   Set placeSnout = objRuledSurface
   
'  --- function cleanup ---
   Set objGeomFactory = Nothing
   Set objTopCircle = Nothing
   Set objBottomCircle = Nothing
   Set objRuledSurface = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placeProjectedPolygonFromVertices
'   Author:     Doug Hempel
'   Inputs:
'               output collection
'               array of vertices
'               direction
'               length
'               is capped
'           Optional
'               transformation matrix
'
'   Outputs:
'               a polygon
'
'   Description:
'               This function creates persistent projcetion of polygon
'
'   Example of call:
'
'               Set CreateBox = placeProjectedPolygonFromVertices(objOutputColl, posBoxBottomFaceVertices(), vecProjDir, dblZLength, blnIsCapped, xform)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   03.Feb.2003     HL        Changed nVertices from a local var to a parameter.  This
'                                   will allow us to re-use array of dposition partially.
'   04.Feb.2003     HL        Added Optional parameter transformation matrix
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeProjectedPolygonFromVertices _
( _
   ByVal m_outputColl As Object, _
   ByRef posPolygonVertices() As IJDPosition, _
   vecProjDir As IJDVector, _
   ByVal dblProjLength As Double, _
   ByVal isCapped As Boolean, _
   Optional ByRef tmxTMatrix As IngrGeom3D.IJDT4x4, _
   Optional ByVal nVertices As Integer _
) As IngrGeom3D.Projection3d
   
   Dim objGeomFactory As IngrGeom3D.GeometryFactory
   Dim objProjection As IngrGeom3D.Projection3d
   Dim oLinestr As IngrGeom3D.LineString3d
      
   Dim i As Integer
   Dim dblVertPnts() As Double
   
   Set objGeomFactory = New IngrGeom3D.GeometryFactory
   Set oLinestr = New IngrGeom3D.LineString3d

'  --- obtain the number of vertices of the 0 based input position array ---
   If nVertices = 0 Then
      nVertices = UBound(posPolygonVertices) + 1
   End If
   If (nVertices < 3) Then
'      MsgBox "Programmer Error: Need at least 3 Vertices"
      objGeomFactory = Nothing
      oLinestr = Nothing
      Exit Function
   End If

'  --- load dynamic double array with repeating xyz values from input points ---
   ReDim dblVertPnts(0 To 3 * nVertices + 2)
   For i = 0 To nVertices - 1
      dblVertPnts(3 * i) = posPolygonVertices(i).x
      dblVertPnts(3 * i + 1) = posPolygonVertices(i).y
      dblVertPnts(3 * i + 2) = posPolygonVertices(i).z
   Next i
   
'  --- add one extra set of xyz values (force last point = first point) ---
   dblVertPnts(3 * nVertices) = posPolygonVertices(0).x
   dblVertPnts(3 * nVertices + 1) = posPolygonVertices(0).y
   dblVertPnts(3 * nVertices + 2) = posPolygonVertices(0).z
    
'  --- create temporary linestring to use for projection function ---
   Set oLinestr = objGeomFactory.LineStrings3d.CreateByPoints(Nothing, nVertices + 1, dblVertPnts)
   
'  --- project the linestring ---
   Set objProjection = PlaceProjection(m_outputColl, oLinestr, vecProjDir, dblProjLength, isCapped)
   
'  --- transform the object ---
   If Not (tmxTMatrix Is Nothing) Then
        objProjection.Transform tmxTMatrix
   End If
   
'  --- release the outstanding objects ---
   Set objGeomFactory = Nothing
   
   Dim objLineString As IJDObject
   Set objLineString = oLinestr
   Set oLinestr = Nothing
   objLineString.Remove

'  --- set function value ---
   Set placeProjectedPolygonFromVertices = objProjection
   Set objProjection = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placeTransformableCone
'   Author:     HL
'   Inputs:
'               Output collection object
'               Base Center Position
'               Top Center Position
'               Base Radius
'               Top Tadius
'           Optional
'               Is Capped
'               Transformation Matrix
'
'   Outputs:
'               Cone
'
'   Description:
'               This function places a cone.
'
'   Example of call:
'               Dim stPoint   As IJDPosition
'               Dim enPoint   As IJDPosition
'               Dim objCone  As object
'
'               Set stPoint = new DPosition
'               Set enPoint = new DPosition
'               stPoint.set 0, 0, 0
'               enPoint.set 0, 0, 1
'               set objCone = PlaceTransformableCone(m_OutputColl, stPoint, enPoint, 2, 1)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
'               Set objCone = Nothing

'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   04.Feb.2003     HL        Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeTransformableCone(ByVal objOutputColl As Object, _
                        centerBase As IJDPosition, _
                        centerTop As IJDPosition, _
                        ByVal radiusBase As Double, _
                        ByVal radiusTop As Double, _
                        Optional ByVal isCapped As Boolean = True, _
                        Optional tmxMatrix As IngrGeom3D.IJDT4x4) As IngrGeom3D.Cone3d
                                
    Dim startBase   As IJDPosition
    Dim startTop    As IJDPosition
    Dim vecNorm     As IJDVector
    Dim objCone     As IngrGeom3D.Cone3d
    Dim geomFactory As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set startBase = New DPosition
    Set startTop = New DPosition
    Set vecNorm = New DVector
    
    Set vecNorm = getOrthonormal(centerTop.x - centerBase.x, _
                   centerTop.y - centerBase.y, _
                   centerTop.z - centerBase.z)
    vecNorm.Length = radiusBase
    Set startBase = centerBase.Offset(vecNorm)
    vecNorm.Length = radiusTop
    Set startTop = centerTop.Offset(vecNorm)
    Set objCone = geomFactory.Cones3d.CreateBy4Pts(objOutputColl.ResourceManager, _
                            centerBase.x, centerBase.y, centerBase.z, _
                            centerTop.x, centerTop.y, centerTop.z, _
                            startBase.x, startBase.y, startBase.z, _
                            startTop.x, startTop.y, startTop.z, _
                            isCapped)
                            
    If Not (tmxMatrix Is Nothing) Then
        objCone.Transform tmxMatrix
    End If
    
    Set placeTransformableCone = objCone
    
    Set objCone = Nothing
    Set geomFactory = Nothing
    Set startBase = Nothing
    Set startTop = Nothing
    Set vecNorm = Nothing
            
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   PlaceTransformableCylinder
'   Author:     JG
'   Inputs:
'               Start Point
'               End Point
'               Diameter
'               isCapped
'               Transformation Matrix
'
'   Outputs:
'               Cylinder
'
'   Description:
'               This function creates persistent projetion of circle
'               based on two points (axis of cylinder) and diameter
'
'   Example of call:
'               Dim stPoint   As new AutoMath.DPosition
'               Dim enPoint   As new AutoMath.DPosition
'               Dim ldiam     as Double
'               Dim objCylinder  As object
'               stPoint.set 0, 0, 0
'               enPoint.set 0, 0, 1
'               lDiam = 1.5
'               set objCylinder = PlaceTransformableCylinder(m_OutputColl, stPoint, enPoint, lDiam, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'               Set objCylinder = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   07.Jan.2003     JG      Function creation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function PlaceTransformableCylinder(ByVal objOutputColl As Object, _
                                lStartPoint As AutoMath.DPosition, _
                                lEndPoint As AutoMath.DPosition, _
                                ByVal lDiameter As Double, _
                                ByVal isCapped As Boolean, Optional tMatrix As IngrGeom3D.IJDT4x4) _
                                As IngrGeom3D.Projection3d
    
    Dim circleCenter    As AutoMath.DPosition
    Dim circleNormal    As AutoMath.DVector
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim dblCylWidth     As Double
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory

    Set circleCenter = New AutoMath.DPosition
    circleCenter.Set lStartPoint.x, lStartPoint.y, lStartPoint.z
    Set circleNormal = New AutoMath.DVector
    circleNormal.Set lEndPoint.x - lStartPoint.x, _
                     lEndPoint.y - lStartPoint.y, _
                     lEndPoint.z - lStartPoint.z
    dblCylWidth = circleNormal.Length
    circleNormal.Length = 1
    
' Construct a circle that will be used to project the disc
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        circleCenter.x, circleCenter.y, circleCenter.z, _
                        circleNormal.x, circleNormal.y, circleNormal.z, _
                        lDiameter / 2)
    
' Project the disc of body
    Set objProjection = geomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                        objCircle, _
                                                        circleNormal.x, circleNormal.y, circleNormal.z, _
                                                        dblCylWidth, isCapped)
                                                    
    If Not (tMatrix Is Nothing) Then
        objProjection.Transform tMatrix
    End If
    
    Set objCircle = Nothing
    
    Set PlaceTransformableCylinder = objProjection
    Set objProjection = Nothing
    Set geomFactory = Nothing
        
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Function:   PlaceSlopedBottomCylinder
'   Author:     JG
'   Inputs:
'               Output Collection
'               start point
'               start direction
'               end point
'               end direction
'               diameter
'               is capped
'               transformation matrix
'
'   Outputs:
'               Sloped Bottom Cylinder
'
'   Description:
'               This function uses a Cone3D object and replaces the boundaries with 2 elipses
'               based based on two points, 2 vectors (axis of end elipses), diameter, and
'               transfomation matrix to move and properly orientate the final object
'
'   Example of call:
'               Dim posStartPoint   As new AutoMath.DPosition
'               Dim vecStartVect    As new Automath.Dvector
'               Dim posEndPoint   As new AutoMath.DPosition
'               Dim vectEndVect    As New AutoMath.DVector
'               Dim tmxMatrix As IngrGeom3D.IJDT4x4
'               Set tmxMatrix = New DT4x4
'               Dim dblDiam     as Double
'               Dim objCylinder  As object
'               stPoint.set 0, 0, 0
'               enPoint.set 0, 0, 1
'               stVect.Set -1, 0, -1
'               enVect.Set 1, 0, -1
'               lDiam = 1.5
'               set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, lDiam, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'               Set objCylinder = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   07.Jan.2003     JG      Function creation
'
'   13.Mar.2003     JG      Changed the way the ends are built from
'                                   useing circles to elipses.  Made cylinder
'                                   acurate to proper diminsions.
'
'   17.Mar.2003     JG      Updated documentation.
'
'   19.Mar.2003     JG      Converted function to use a cone vs a ruledsurface
'                                   also converted to a nested symbol to use
'                                   planes to cap the ends.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function PlaceSlopedBottomCylinder(ByVal objOutputColl As Object, _
                                            posStartPoint As AutoMath.DPosition, _
                                            vecStartVector As AutoMath.DVector, _
                                            posEndPoint As AutoMath.DPosition, _
                                            vecEndVector As AutoMath.DVector, _
                                            dblDiameter As Double, _
                                            bolIsCapped As Boolean, _
                                            tmxMatrix As IngrGeom3D.IJDT4x4) As Object
    
    ' Get or create the definition
    Dim defColl             As IJDDefinitionCollection
    Dim boxDef              As IJDSymbolDefinition
    Dim definitionParams    As Variant
    Dim oEnv                As IMSSymbolEntities.DSymbol
    Dim newEnumArg          As IJDEnumArgument
    Dim IJEditJDArg         As IJDEditJDArgument
    Dim PC                  As IJDParameterContent
    Dim argument            As IJDArgument
    Dim iCount              As Integer
    Dim oSymbolFactory      As IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set oSymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set defColl = oSymbolFactory.DefinitionCollection(objOutputColl.ResourceManager)
    definitionParams = ""
    Set boxDef = defColl.GetDefinitionByProgId(True, "SlopedBotCylinder.SlopBotCylServices", vbNullString, definitionParams)

    Set oEnv = oSymbolFactory.PlaceSymbol(boxDef, objOutputColl.ResourceManager)
    Set oSymbolFactory = Nothing
    Set boxDef = Nothing
    Set defColl = Nothing

    Set newEnumArg = New DEnumArgument
    Set IJEditJDArg = newEnumArg.IJDEditJDArgument

    Dim i As Integer
    If tmxMatrix Is Nothing Then
        Set tmxMatrix = New DT4x4
        tmxMatrix.LoadIdentity
    End If
    
'Start array entries

    'Start point *****************************
    'X
    iCount = 1

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = posStartPoint.x
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Y
    iCount = 2

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = posStartPoint.y
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Z
    iCount = 3

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = posStartPoint.z
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Start vector  ***************************
    'X
    iCount = 4

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = vecStartVector.x
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Y
    iCount = 5

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = vecStartVector.y
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Z
    iCount = 6

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = vecStartVector.z
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'End Point *******************************
    'X
    iCount = 7

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = posEndPoint.x
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Y
    iCount = 8

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = posEndPoint.y
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Z
    iCount = 9

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = posEndPoint.z
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'End Vector  ****************************
    'X
    iCount = 10

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = vecEndVector.x
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Y
    iCount = 11

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = vecEndVector.y
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Z
    iCount = 12

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = vecEndVector.z
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'Diameter ***************************
    
        iCount = 13

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = dblDiameter
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
    'isCapped  **************************
    Dim intCapped As Integer
    If bolIsCapped = True Then
        intCapped = 1
    Else
        intCapped = 0
    End If
    
        iCount = 14

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = intCapped
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing
        
        'matrix ***********************************

        iCount = 15

    'Pass the Transformation Matrix
    'This contains 16 elements
    For i = 0 To 15

        Set PC = New DParameterContent
        Set argument = New DArgument

        PC.uomValue = tmxMatrix.IndexValue(i)
        PC.Type = igValue
        PC.uomType = 1
    ' Feed the Argument
        argument.Index = iCount
        argument.Entity = PC
    ' Add the argument to the arg collection
        IJEditJDArg.SetArg argument
        Set PC = Nothing
        Set argument = Nothing

        iCount = iCount + 1

    Next i
        
'End array entries
    
    oEnv.IJDValuesArg.SetValues newEnumArg
    Dim IJDInputsArg As IMSSymbolEntities.IJDInputsArg
    Set IJDInputsArg = oEnv
    IJDInputsArg.Update
    Set IJDInputsArg = Nothing
    Set IJEditJDArg = Nothing
    Set newEnumArg = Nothing

    Set PlaceSlopedBottomCylinder = oEnv
    Set oEnv = Nothing
    Set oSymbolFactory = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   placeKnuckleRadiusTankHead
'   Author:     Doug Hempel
'   Inputs:
'               m_OutputColl           => The collection object
'               posKnuckleArcRadiusPnt => Position of knuckle arc radius point.
'               posKnuckleArcStartPnt  => Position of knuckle arc start point.
'               posKnuckleArcEndPnt    => Position of knuckle arc end point.
'               posDomeArcRadiusPnt    => Position of dome arc radius point.
'               posDomeArcStartPnt     => Position of dome arc start point.
'               posDomeArcEndPnt       => Position of dome arc end point.
'               posCappingLineStartPnt => Position of capping line start point.
'               posCappingLineEndPnt   => Position of capping line end point.
'               blnIsCapped            => True/False, is tank head open or closed.
'               tmxTMatrix            <=> Transformation matrix used to transform the tank head.
'
'   Outputs:
'               Object - Returns a revolution object that is an elliptical tank head.
'
'   Description:
'               Builds a knuckle radius tank head from the given arc and line
'               positions.  This function is a low level tank head builder.
'               This function is not normally called directly by symbol
'               programmers, but rather is called by the higher level create
'               function.
'
'   Example of call:
'
'               Set CreateKLRatioTankHead = placeKnuckleRadiusTankHead(objOutputColl, posKnuckleArcRadiusPnt,
'                   posKnuckleArcStartPnt, posKnuckleArcEndPnt,
'                   posDomeArcRadiusPnt, posDomeArcStartPnt, posDomeArcEndPnt,
'                   posCappingLineStartPnt, posCappingLineEndPnt,
'                   blnIsCapped, xform)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeKnuckleRadiusTankHead _
( _
   ByVal m_outputColl As Object, _
   posKnuckleArcRadiusPnt As IJDPosition, _
   posKnuckleArcStartPnt As IJDPosition, _
   posKnuckleArcEndPnt As IJDPosition, _
   posDomeArcRadiusPnt As IJDPosition, _
   posDomeArcStartPnt As IJDPosition, _
   posDomeArcEndPnt As IJDPosition, _
   posCappingLineStartPnt As IJDPosition, _
   posCappingLineEndPnt As IJDPosition, _
   ByVal blnIsCapped As Boolean, _
   Optional ByRef tmxTMatrix As IngrGeom3D.IJDT4x4 _
) As IngrGeom3D.Revolution3d


'  --- items that will be set rather than "newed" in this function scope ---
   Dim objRevolution     As IngrGeom3D.Revolution3d
   Dim objComplexString  As IngrGeom3D.ComplexString3d
   Dim objDomeArc        As IngrGeom3D.Arc3d
   Dim objKnuckleArc     As IngrGeom3D.Arc3d
   Dim objCappingLine    As IngrGeom3D.Line3d
   
'  --- items that will have to be "newed" in this function scope ---
   Dim objGeomFactory          As IngrGeom3D.GeometryFactory
   Dim posDomeBuildOrigin      As IJDPosition
   Dim posStartofComplexString As IJDPosition
   Dim objTemporaryCollection  As Collection
   

'  --- other variables ---
   Dim i As Integer

'  --- create new objects ---
   Set objGeomFactory = New IngrGeom3D.GeometryFactory
   Set objTemporaryCollection = New Collection
   Set posDomeBuildOrigin = New DPosition
   Set posStartofComplexString = New DPosition

   
   posDomeBuildOrigin.Set 0, 0, 0
   
'  --- if tankhead is to be capped, construct capping line ---
   If (blnIsCapped = True) Then
      Set objCappingLine = objGeomFactory.Lines3d.CreateBy2Points(Nothing, posCappingLineStartPnt.x, posCappingLineStartPnt.y, posCappingLineStartPnt.z, posCappingLineEndPnt.x, posCappingLineEndPnt.y, posCappingLineEndPnt.z)
      objTemporaryCollection.Add objCappingLine
   End If

'  --- construct knuckle arc ---
   Set objKnuckleArc = objGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, posKnuckleArcRadiusPnt.x, posKnuckleArcRadiusPnt.y, posKnuckleArcRadiusPnt.z, posKnuckleArcStartPnt.x, posKnuckleArcStartPnt.y, posKnuckleArcStartPnt.z, posKnuckleArcEndPnt.x, posKnuckleArcEndPnt.y, posKnuckleArcEndPnt.z)
   objTemporaryCollection.Add objKnuckleArc

'  --- construct dome arc ---
   Set objDomeArc = objGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, posDomeArcRadiusPnt.x, posDomeArcRadiusPnt.y, posDomeArcRadiusPnt.z, posDomeArcStartPnt.x, posDomeArcStartPnt.y, posDomeArcStartPnt.z, posDomeArcEndPnt.x, posDomeArcEndPnt.y, posDomeArcEndPnt.z)
   objTemporaryCollection.Add objDomeArc

'  --- set starting point (starting point along the path of complex string)
   If (blnIsCapped = True) Then
      posStartofComplexString.Set posCappingLineStartPnt.x, posCappingLineStartPnt.y, posCappingLineStartPnt.z
   Else
      posStartofComplexString.Set posKnuckleArcStartPnt.x, posKnuckleArcStartPnt.y, posKnuckleArcStartPnt.z
   End If

'  --- create a complex string from the temporary collection ---
   Set objComplexString = PlaceTrCString(posStartofComplexString, objTemporaryCollection)

'  --- revolve the complex object about x axis to form a dish ---
   Set objRevolution = objGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, objComplexString, 1, 0, 0, posDomeBuildOrigin.x, posDomeBuildOrigin.y, posDomeBuildOrigin.z, degreeToRadian(360), True)
   
   If Not (tmxTMatrix Is Nothing) Then
       '  --- transform the object to right position and orientation ---
       objRevolution.Transform tmxTMatrix
   End If
'  --- set function return value ---
   Set placeKnuckleRadiusTankHead = objRevolution
       
'  --- cleanup ---
   If (blnIsCapped = True) Then
      Set objCappingLine = Nothing
   End If
   For i = 1 To objTemporaryCollection.Count
      objTemporaryCollection.Remove 1
   Next i
   Set objKnuckleArc = Nothing
   Set objDomeArc = Nothing
   Set objGeomFactory = Nothing
   Set objRevolution = Nothing
   Set posDomeBuildOrigin = Nothing
   Set posStartofComplexString = Nothing
   Set objComplexString = Nothing
   
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Function:   placeConicalTankHead
'   Author:     Doug Hempel
'   Inputs:
'               m_OutputColl           => The collection object
'               posConeLineStartPnt    => Position of cone line start point.
'               posConeLineEndPnt      => Position of cone line end point.
'               posConeTopStartPnt     => Position of cone top start point.
'               posConeTopEndPnt       => Position of cone top end point.
'               posCappingLineStartPnt => Position of capping line start point.
'               posCappingLineEndPnt   => Position of capping line end point.
'               blnIsCapped            => True/False, is tank head open or closed.
'               tmxTMatrix            <=> Transformation matrix used to transform the tank head.
'
'   Outputs:
'               Object - Returns a revolution object that is an elliptical tank head.
'
'   Description:
'               Builds a conical tank head from the given arc and line
'               positions.  This function is a low level tank head builder.
'               This function is not normally called directly by symbol
'               programmers, but rather is called by the higher level create
'               function.
'
'   Example of call:
'
'               Set CreateConicalTankHead = placeConicalTankHead(objOutputColl, posConeLineStartPnt, posConeLineEndPnt, posConeTopStartPnt, posConeTopEndPnt, posCappingLineStartPnt, posCappingLineEndPnt, blnIsCapped, xform)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function placeConicalTankHead _
( _
   ByVal m_outputColl As Object, _
   posConeLineStartPnt As IJDPosition, _
   posConeLineEndPnt As IJDPosition, _
   posConeTopStartPnt As IJDPosition, _
   posConeTopEndPnt As IJDPosition, _
   posCappingLineStartPnt As IJDPosition, _
   posCappingLineEndPnt As IJDPosition, _
   ByVal blnIsCapped As Boolean, _
   Optional ByRef tmxTMatrix As IngrGeom3D.IJDT4x4 _
) As IngrGeom3D.Revolution3d

'  --- items that will be set rather than "newed" in this function scope ---
   Dim objRevolution    As IngrGeom3D.Revolution3d
   Dim objComplexString As IngrGeom3D.ComplexString3d
   Dim objConeLine      As IngrGeom3D.Line3d
   Dim objConeTop       As IngrGeom3D.Line3d
   Dim objCappingLine   As IngrGeom3D.Line3d

'  --- items that will have to be "newed" in this function scope ---
   Dim objGeomFactory          As IngrGeom3D.GeometryFactory
   Dim posTankHeadBuildOrigin      As IJDPosition
   Dim posStartofComplexString As IJDPosition
   Dim objTemporaryCollection  As Collection


'  --- other variables ---
   Dim i As Integer

'  --- create new objects ---
   Set objGeomFactory = New IngrGeom3D.GeometryFactory
   Set objTemporaryCollection = New Collection
   Set posTankHeadBuildOrigin = New DPosition
   Set posStartofComplexString = New DPosition


   posTankHeadBuildOrigin.Set 0, 0, 0

'  --- if tankhead is to be capped, construct capping line ---
   If (blnIsCapped = True) Then
      Set objCappingLine = objGeomFactory.Lines3d.CreateBy2Points(Nothing, posCappingLineStartPnt.x, posCappingLineStartPnt.y, posCappingLineStartPnt.z, posCappingLineEndPnt.x, posCappingLineEndPnt.y, posCappingLineEndPnt.z)
      objTemporaryCollection.Add objCappingLine
   End If

'  --- construct cone line ---
   Set objConeLine = objGeomFactory.Lines3d.CreateBy2Points(Nothing, posConeLineStartPnt.x, posConeLineStartPnt.y, posConeLineStartPnt.z, posConeLineEndPnt.x, posConeLineEndPnt.y, posConeLineEndPnt.z)
   objTemporaryCollection.Add objConeLine

'  --- construct cone top ---
   Set objConeTop = objGeomFactory.Lines3d.CreateBy2Points(Nothing, posConeTopStartPnt.x, posConeTopStartPnt.y, posConeTopStartPnt.z, posConeTopEndPnt.x, posConeTopEndPnt.y, posConeTopEndPnt.z)
   objTemporaryCollection.Add objConeTop

'  --- set starting point (starting point along the path of complex string)
   If (blnIsCapped = True) Then
      posStartofComplexString.Set posCappingLineStartPnt.x, posCappingLineStartPnt.y, posCappingLineStartPnt.z
   Else
      posStartofComplexString.Set posConeLineStartPnt.x, posConeLineStartPnt.y, posConeLineStartPnt.z
   End If

'  --- create a complex string from the temporary collection ---
   Set objComplexString = PlaceTrCString(posStartofComplexString, objTemporaryCollection)

'  --- revolve the complex object about x axis to form a dish ---
   Set objRevolution = objGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, objComplexString, 1, 0, 0, posTankHeadBuildOrigin.x, posTankHeadBuildOrigin.y, posTankHeadBuildOrigin.z, degreeToRadian(360), True)

    If Not (tmxTMatrix Is Nothing) Then
        '  --- transform the object to right position and orientation ---
       objRevolution.Transform tmxTMatrix
    End If
'  --- set function return value ---
   Set placeConicalTankHead = objRevolution

'  --- cleanup ---
   If (blnIsCapped = True) Then
      Set objCappingLine = Nothing
   End If
   For i = 1 To objTemporaryCollection.Count
      objTemporaryCollection.Remove 1
   Next i
   Set objConeLine = Nothing
   Set objConeTop = Nothing
   Set objGeomFactory = Nothing
   Set objRevolution = Nothing
   Set posTankHeadBuildOrigin = Nothing
   Set posStartofComplexString = Nothing
   Set objComplexString = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   PlacePlane
'   Author:     HL
'   Inputs:
'               Output collection object
'               array of points
'       Optional
'               Number of points
'               transformation matrix
'
'   Outputs:
'               A Plane based on input
'
'   Description:
'               This function places a plane.
'
'   Example of Call:
'
'               Set oPlane = PlacePlane(objOutputColl, posVerts, 4, tmxForm)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   14.Feb.2003     HL        Created this function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function PlacePlane(ByVal objOutputColl As Object, _
                            ByRef posPolygonVertices() As IJDPosition, _
                            Optional ByVal nVertices As Integer, _
                            Optional ByRef tmxTMatrix As IngrGeom3D.IJDT4x4) As IngrGeom3D.Plane3d
   
    Dim objGeomFactory As IngrGeom3D.GeometryFactory
    Dim objPlane As IngrGeom3D.Plane3d
      
    Dim i As Integer
    Dim dblVerts() As Double
    
    If nVertices = 0 Then
        nVertices = UBound(posPolygonVertices) + 1
    End If
    
    If nVertices < 3 Then
        Err.Raise vbObjectError + 50000, "PlacePlane", "Need at least 3 Vertices"
        Exit Function
    End If
    
    ReDim dblVerts(0 To 3 * nVertices - 1)
    For i = 0 To nVertices - 1
        dblVerts(3 * i) = posPolygonVertices(i).x
        dblVerts(3 * i + 1) = posPolygonVertices(i).y
        dblVerts(3 * i + 2) = posPolygonVertices(i).z
    Next i
    
    Set objGeomFactory = New IngrGeom3D.GeometryFactory
    If objOutputColl Is Nothing Then
        Set objPlane = objGeomFactory.Planes3d.CreateByPoints(Nothing, _
                            nVertices, dblVerts)
    Else
        Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                            nVertices, dblVerts)
    End If
    
    If Not (tmxTMatrix Is Nothing) Then
        objPlane.Transform tmxTMatrix
    End If
    
    Set PlacePlane = objPlane
    Set objPlane = Nothing
    Set objGeomFactory = Nothing
    
End Function
