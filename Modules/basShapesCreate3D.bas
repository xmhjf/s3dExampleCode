Attribute VB_Name = "basShapesCreate3D"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   FileName:       basShapesCreate3D.bas
'   ProgID:         basShapesCreate3D
'   Author:         3D Config Team
'   Compiler:       HL
'   Creation Date:  Thursday, Jan 23, 2003
'   Description:    This module contains high level primitive create
'                   functions and procedures to create primitive objects.
'
'       Box:        CreateBox
'
'       Cone:       CreateCone
'
'       Cylinder:   CreateCylinder
'                   CreateSlopeBottomCylinder
'
'       Dish:       Create2to1EllipticalDish
'                   CreateDish
'
'                   Create2to1EllipticalTankHead
'                   CreateAnyKLRatioTankHead
'                   CreateASMETypeFDTankHead
'                   CreateConicalTankHead
'                   CreateHemisphereTankHead
'
'       Plane:      CreateDefaultPlane
'       Point:      CreatePoint
'
'       Prism:      CreateBasePrism
'                   CreatePrism
'                   CreateSlopeBottomPrism
'                   CreateSlopedPrism
'
'       Pyramid:    CreatePyramid
'
'       Torus:      CreateCircularTorus
'                   CreatePrismaticTorus
'                   CreateRectangularTorus
'
'       Snout:      CreateSnout
'
'       Sphere:     CreateSphere
'
'   Change History:
'   dd.mmm.yyyy   who           change description
'   -----------   ---           ------------------
'   23.Jan.2003   HL      Collected code from 3D Config Team into this module
'   27.Jan.2003   JG    Modified CreateCylinder function
'   29.Jan.2003   HL      Added Conditional Compilation to have debug info
'   03.Feb.2003   HL      Added function CreatePoint
'   04.Feb.2003   HL      Modified CreateCone function to use transformation matrix
'   04.Feb.2003   Doug Hempel   Added 5 new Tank Head create functions.
'   04.Feb.2003   HL      Changed all return type to IMSSymbolEntities.DSymbol
'   06.Feb.2003   HL      Changed all return type to their specific type.
'                               Change all Dposition/Dvector/Ori to pass by ref.
'   14.Feb.2003   HL      Added function CreateDefaultPlane.
'   20.Mar.2003   JG    Added CreateReducingMiterCutCircularTorus
'   25.Mar.2003   HL      Added isCapped to functions.
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is Updated.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'   To Do List:
'       1.  PTOR - NSeg
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Const MODULE = "basShapesCreate3D"
'Set debug mode to true (-1), or false (0)
#Const INDEBUG = 0

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   Create2to1EllipticalDish
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl   => The collection object
'               strOrigin       => The coordinate location of dish origin expressed
'                                  as a string.
'               strDirection    => A text description of the direction of the dish.
'               dblDiameter     => Diameter of dish.
'               blnIsCapped     => Boolean indicating whether to cap the open end of
'                                  the dish.
'
'   Outputs:
'               Object - Returns a generic object that is a dish with a fixed
'               height to diameter ratio.
'
'   Description:
'               Returns a dish as a generic object.  The dish can be given a
'               position and orientation in 3d space.  This dish is a special
'               case of the CreateDish that creates an elliptical dish with a
'               height that is 1/2 of the capping radius or a height to capping
'               diameter ratio of 0.25.  It calls the CreateDish function with
'               fixed parameter values to apply the above mentioned ratios.
'
'   Example of Call:
'               Dim objDish as Object
'               Set objDish = Create2to1EllipticalDish(m_OutputColl, "E 0 N 0 U 0", "E 90 U 0", 0.75, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDish
'               Set objDish = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function Create2to1EllipticalDish _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByVal strDirection As String, _
   ByVal dblDiameter As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.Revolution3d

   Dim dblHeight As Double
   
   dblHeight = 0.25 * dblDiameter
   
   Set Create2to1EllipticalDish = CreateDish(objOutputColl, strOrigin, strDirection, dblDiameter, dblHeight, 1, blnIsCapped)

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   Create2to1EllipticalTankHead
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl   => The collection object
'               strOrigin       => The coordinate location of dish origin expressed
'                                  as a string.
'               strDirection    => A text description of direction of the tank head.
'               dblDiameter     => Capping diameter of tank head.
'               blnIsCapped     => Boolean indicating whether to cap the open end of
'                                  the tank head.
'   Outputs:
'               Object - Returns a generic object that is a knuckle radius tank head with
'               a fixed height to diameter ratio.
'
'   Description:
'               Returns a knuckle radius tank head as a generic object.
'               The tank head can be given a position and orientation in 3d space.
'               This tank head is a special case of the CreateKnuckleRadiusTankHead
'               that creates a tank head with a height that is 1/2 of the capping
'               radius or a height to capping diameter ratio of 0.25.  It calls the
'               CreateKLRatioTankHead function with fixed parameter values to
'               apply the above mentioned ratios.
'
'   Example of Call:
'               Dim objTankHead as Object
'               Set objTankHead = Create2to1EllipticalTankHead(m_OutputColl, "E 0 N 0 U 0", "E 90 U 0", 0.75, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTankHead
'               Set objTankHead = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function Create2to1EllipticalTankHead _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByVal strDirection As String, _
   ByVal dblDiameter As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.Revolution3d

   Set Create2to1EllipticalTankHead = CreateKLRatioTankHead(objOutputColl, strOrigin, strDirection, dblDiameter, 0.9045, 0.17275, blnIsCapped)

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateASMETypeFDTankHead
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl     => The collection object
'               strOrigin         => The coordinate location of dish origin expressed
'                                    as a string.
'               strDirection      => A text description of the direction of the tank head.
'               dblDiameter       => Capping diameter of tank head.
'               blnIsCapped       => Boolean indicating whether to cap the open end of
'                                    the tank head.
'
'   Outputs:
'               Object - Returns a generic object that is a knuckle radius tank head
'               with a fixed height to diameter ratio.
'
'   Description:
'               Returns an ASME FD knuckle radius tank head as a generic object.
'               The tank head can be given a position and orientation in 3d space.
'               This tank head is a special case of the CreateKLRatioTankHead
'               using, K = 1.0 and L = 0.06. It calls the CreateKLRadiusTankHead
'               function with fixed parameter values.
'
'   Example of Call:
'               Dim objTankHead as Object
'               Set objTankHead = CreateASMETypeFDTankHead(m_OutputColl, "E 0 N 0 U 0", "E 90 U 0", 0.75, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTankHead
'               Set objTankHead = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateASMETypeFDTankHead _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByVal strDirection As String, _
   ByVal dblDiameter As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.Revolution3d

   Set CreateASMETypeFDTankHead = CreateKLRatioTankHead(objOutputColl, strOrigin, strDirection, dblDiameter, 1#, 0.06, blnIsCapped)

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateBasePrism
'   Author:     HL
'   Inputs:
'               Output collection object
'               Origin
'               Orientation
'               Length
'               Vertices in X and Y Coordinates
'               Optional X Direction
'               Optional Y Direction
'               Optional X Offset
'               Optional Y Offset
'               Optional X Scaling factor
'               Optional Y Scaling factor
'               Optional blnIsCapped
'
'   Outputs:
'               Prism based on input
'
'   Description:
'               This function places a prism.  It is the base function for all
'               other create prism functions.  The prism is built vertically.
'               X direction is the vector of the top surface.
'               Y direction is the vector of the bottom surface.
'
'   Example of Call:
'               Dim objPrism as Object
'               Set objPrism=CreateBasePrism(m_OutputColl, "E 0 N 0 U 0", Nothing, 1, "1,0,0,1,-1,0,0,-1")
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTankHead
'               Set objPrism = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   13.Jan.2003     HL        Created this function
'   13.Feb.2003     HL        Checked length of strOrigin
'   13.Mar.2003     HL        Use new intergraph's string convert function
'                                   X and Y dir don't have to start with U or D now.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateBasePrism(ByVal objOutputColl As Object, _
                            ByVal strOrigin As String, _
                            ByRef oriOrientation As Orientation, _
                            ByVal dblLength As Double, _
                            ByVal strVertices As String, _
                            Optional ByVal strXDirection As String, _
                            Optional ByVal strYDirection As String, _
                            Optional ByVal dblXOffset As Double, _
                            Optional ByVal dblYOffset As Double, _
                            Optional ByVal dblXScalingFactor As Double, _
                            Optional ByVal dblYScalingFactor As Double, _
                            Optional ByVal blnIsCapped As Boolean = True) _
                            As IngrGeom3D.RuledSurface3d

    Dim intVertices As Integer
    Dim strVerts() As String
    Dim arrTopPts() As IJDPosition
    Dim arrBottomPts() As IJDPosition
    Dim Xdir As IJDVector
    Dim Ydir As IJDVector
    Dim i As Integer
    Dim x, y, z As Double
    
    Dim tmxForm As IJDT4x4
    Dim vecPosition As IJDVector
    Dim vecDir As IJDVector
    
    strVerts = Split(strVertices, ",")
    intVertices = UBound(strVerts) + 1

    'Check to see if we have both x and y inputs
    If (intVertices Mod 2) <> 0 Then
        Err.Raise vbObjectError + 53000, "CreateBasePrism", "Must input both x and y"
        Exit Function
    End If
    intVertices = intVertices / 2

    ReDim arrTopPts(intVertices - 1) As IJDPosition
    ReDim arrBottomPts(intVertices - 1) As IJDPosition
    
    Set tmxForm = New DT4x4
    Set vecPosition = New DVector
    Set vecDir = New DVector

    'Check to see if X and Y direction has a value
    If Len(strXDirection) > 0 Then
        Set Xdir = convertDirectionToUnitVector(strXDirection)
'        #If INDEBUG Then
'            MsgBox "CreateBasePrism: xdir is " & Xdir.x & " " & Xdir.y & " " & Xdir.z
'        #End If
    End If
    
    If Len(strYDirection) > 0 Then
        Set Ydir = convertDirectionToUnitVector(strYDirection)
'        #If INDEBUG Then
'            MsgBox "CreateBasePrism: ydir is " & Ydir.x & " " & Ydir.y & " " & Ydir.z
'        #End If
    End If
    
    'Initialize X and Y Scaling Factor if not there.
    If dblXScalingFactor = 0 Then
        dblXScalingFactor = 1
    End If
    
    If dblYScalingFactor = 0 Then
        dblYScalingFactor = 1
    End If

    For i = 0 To intVertices - 1
        'Set points for the top surface
        Set arrTopPts(i) = New DPosition
        x = CDbl(strVerts(i * 2)) * dblXScalingFactor + dblXOffset / 2
        y = CDbl(strVerts(i * 2 + 1)) * dblYScalingFactor + dblYOffset / 2
        z = dblLength / 2
        If Not (Xdir Is Nothing) Then
            If Xdir.z <> 0 Then
                z = z + (-x * Xdir.x - y * Xdir.y) / Xdir.z
            End If
        End If
        arrTopPts(i).Set x, y, z
        
        'Set points for the bottom surface
        Set arrBottomPts(i) = New DPosition
        x = CDbl(strVerts(i * 2)) - dblXOffset / 2
        y = CDbl(strVerts(i * 2 + 1)) - dblYOffset / 2
        z = -dblLength / 2
        If Not (Ydir Is Nothing) Then
            If Ydir.z <> 0 Then
                z = z - (-x * Ydir.x - y * Ydir.y) / Ydir.z
            End If
        End If
        arrBottomPts(i).Set x, y, z
    Next i

    'initialize the transformation matrix to identify
    tmxForm.LoadIdentity

    If Len(strOrigin) > 0 Then
        'load the transformation matrix with a translate from origin to placement position
        Set vecPosition = convertPositionStringToDVector(strOrigin)
        tmxForm.Translate vecPosition
    End If
    
    If Not (oriOrientation Is Nothing) Then
        Call loadOriIntoTransformationMatrix(tmxForm, oriOrientation)
    End If

    Set CreateBasePrism = placeTruncNEdgePrism(objOutputColl, intVertices, arrTopPts, arrBottomPts, tmxForm, blnIsCapped)

    For i = 0 To intVertices - 1
        Set arrTopPts(i) = Nothing
        Set arrBottomPts(i) = Nothing
    Next i
    Set tmxForm = Nothing
    Set vecPosition = Nothing
    Set vecDir = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateBox
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl         => The collection object
'               strOrigin             => The coordinate location of box origin expressed
'                                        as a string. Box origin is in center of box.
'               oriBoxOrientation     => An orientation object used to effect the orientation
'                                        of the box.
'               dblXLength            => Length along the X axis direction of the box when box
'                                        is in standard orientation.
'               dblYLength            => Length along the Y axis direction of the box when box
'                                        is in standard orientation.
'               dblZLength            => Length along the Z axis direction of the box when box
'                                        is in standard orientation.
'               blnIsCapped           => boolean indicating whether to cap (close) the box on the
'                                        top and bottom.
'                                        False: box is open on top and bottom when in
'                                        standard orientation.
'
'   Outputs:
'               Object - Returns a generic object that is a box
'
'   Description:
'               Returns a box as a generic object.  The box can be given a
'               position and orientation in 3d space.
'
'   Example of Call:
'               Dim objBox as Object
'               Dim oriPartOri As Orientation
'               Set oriPartOri = New Orientation
'               oriPartOri.ResetDefaultAxis      'Reset to No Rotations
'               oriPartOri.RotationAboutZ = -90  'Set the amount of rotation about Z Axis to -90 degrees
'               oriPartOri.RotationAboutX = 55   'Set the amount of rotation about X Axis to 55 degrees
'               oriPartOri.ApplyRotations        'Rotates the oriention object by the indicated rotation amounts
'               Set objBox = CreateBox(m_OutputColl, "E 0 N 0 U 0", oriPartOri, 0.75, 0.5, 1.0, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
'               Set objBox = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   12.Feb.2003     HL        Check existence of strOrigin and oriBoxOrientation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateBox _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByRef oriBoxOrientation As Orientation, _
   ByVal dblXLength As Double, _
   ByVal dblYLength As Double, _
   ByVal dblZLength As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.Projection3d

   
   Dim xform As IngrGeom3D.IJDT4x4
   Dim posBoxBottomFaceVertices(0 To 3) As IJDPosition
   Dim vecProjDir As IJDVector
   Dim vecTmp As IJDVector
   
   Set xform = New DT4x4
   Dim i As Integer
   For i = 0 To 3
      Set posBoxBottomFaceVertices(i) = New DPosition
   Next i
   Set vecProjDir = New DVector
   Set vecTmp = New DVector
     
'  --- make vertices of box bottom face ---
   posBoxBottomFaceVertices(0).Set -dblXLength / 2, -dblYLength / 2, -dblZLength / 2
   posBoxBottomFaceVertices(1).Set dblXLength / 2, -dblYLength / 2, -dblZLength / 2
   posBoxBottomFaceVertices(2).Set dblXLength / 2, dblYLength / 2, -dblZLength / 2
   posBoxBottomFaceVertices(3).Set -dblXLength / 2, dblYLength / 2, -dblZLength / 2

'  --- bottom face of box will be projected upward ---
   vecProjDir.Set 0, 0, 1
   
'  --- initialize the transform matrix to identity ---
   xform.LoadIdentity

'  --- load the transform matrix with a translate from origin to placement position ---
   If Len(strOrigin) > 0 Then
        Set vecTmp = convertPositionStringToDVector(strOrigin)
        xform.Translate vecTmp
    End If
   
'  --- put ori values into xform matrix ---
   If Not (oriBoxOrientation Is Nothing) Then
      Call loadOriIntoTransformationMatrix(xform, oriBoxOrientation)
   End If

'  --- call place function to create the box ---
   Set CreateBox = placeProjectedPolygonFromVertices(objOutputColl, posBoxBottomFaceVertices(), vecProjDir, dblZLength, blnIsCapped, xform)
   
'  --- cleanup ---
   Set xform = Nothing
   For i = 0 To 3
      Set posBoxBottomFaceVertices(i) = Nothing
   Next i
   Set vecProjDir = Nothing
   Set vecTmp = Nothing
   
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateCircularTorus
'   Author:     Dan Fielder
'   Inputs:
'               Output Collection
'               Origin
'               Inside radius
'               Outside radius
'               X Direction
'               Y Direction
'               Sweep angle
'               orientation
'               Optional blnIsCapped
'
'   Outputs:
'               circular torus symbol
'
'   Description:
'               This function creates a circular torus symbol based upon
'               "engineering" type inputs.  They are two ways to input:
'               1.  By giving strXDirection and strYDirection, it will compute
'               the angle, the angle passed in as a parameter will be ignored.
'               It will always use the smaller than 180 angle.
'               2.  By just inputing an angle, it will compute xdir and ydir
'               on xy-plane.  The user must use ori to orient it.
'
'               If all 4 inputs are given, it will use the first method.
'
'   Example of Call:
'               Dim objCTor as Object
'               Set objCTor = CreateCircularTorus(m_outputColl, "E 0 N 0 U 0", 0.1016, 0.1524, 0, "ES 120 U 10", "NW 330 U 10")
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCTor
'               Set objCTor = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   19.Dec.2002     Dan Fielder     created
'   28.Jan.2003     HL        Seperate higher and lower level functions.
'   07.Feb.2003     HL        Put in the code to compute xdir and ydir
'                                   based on angle.
'   13.Feb.2003     HL        Added functionality for xdir and ydir.
'   24.Feb.2003     HL        Instead of computing new position of the circle
'                                   center point, use transformation matrix.
'   03.Mar.2003     JG      Added branching for segmented circular torus.
'   20.Mar.2003     JG      Updated call to placeMiterCutCircTorus to
'                                   reflect changes in the parameters.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateCircularTorus(ByVal objOutputColl As Object, _
                                    ByVal strOrigin As String, _
                                    ByVal dblInsideRadius As Double, _
                                    ByVal dblOutsideRadius As Double, _
                                    Optional ByVal intNumberOfSegments As Integer, _
                                    Optional ByVal strXDirection As String, _
                                    Optional ByVal strYDirection As String, _
                                    Optional ByVal dblSweepAngle As Double, _
                                    Optional ByRef oriDir As Orientation, _
                                    Optional ByVal blnIsCapped As Boolean = True) As Object
                                    
    Dim vecCenter As IJDVector
    Dim vecStart As IJDVector
    Dim posStart As IJDPosition
    Dim xVector As IJDVector
    Dim yVector As IJDVector
    Dim vecRev As IJDVector
    Dim dblMajorR As Double
    Dim dblMinorR As Double
    
    Set vecRev = New DVector
    Set vecCenter = New DVector
    Set vecStart = New DVector
    Set posStart = New DPosition
    
    dblMinorR = (dblOutsideRadius - dblInsideRadius) / 2
    dblMajorR = dblInsideRadius + dblMinorR
    
    'Used to move and orientate the object after construction.
    Dim tmxMatrix As IJDT4x4
    Set tmxMatrix = New DT4x4
        
    tmxMatrix.LoadIdentity
    If Len(strOrigin) > 0 Then
        ' convert the origin string to a center position vector
        Set vecCenter = convertPositionStringToDVector(strOrigin)
'        #If INDEBUG Then
'            MsgBox "Center is " & vecCenter.x & " " & vecCenter.y & " " & vecCenter.z
'        #End If
        
        'Move object into position
        tmxMatrix.Translate vecCenter
    End If

    If Len(strXDirection) > 0 Then
        Set xVector = convertDirectionToUnitVector(strXDirection)
    Else
        Set xVector = New DVector
        xVector.Set 1, 0, 0
    End If
    
    If Len(strYDirection) > 0 Then
        Set yVector = convertDirectionToUnitVector(strYDirection)
    Else
        Set yVector = New DVector
        yVector.Set 0, 1, 0
    End If
        
    If Len(strXDirection) = 0 And Len(strYDirection) = 0 Then
        ' if doesn't have xdir and ydir, use the angle
        ' convert the sweep angle to radians from degrees
        dblSweepAngle = degreeToRadian(dblSweepAngle)
        setXYDirFromAngle xVector, yVector, dblSweepAngle
    Else
        'if xVector and yVector are pointing to the opposite direction.
        If xVector.x = -yVector.x And xVector.y = -yVector.y And xVector.z = -yVector.z Then
            Err.Raise vbObjectError + 53002, "CreateCircularTorus", "Can not have the directions that are opposite."
            Exit Function
        Else
            'The angle will not be used.  It will be calculated from
            'xvector and yvector.
            Set oriDir = getOriFromXYDir(xVector, yVector, dblSweepAngle)
            'reset xdir and ydir to flat dirs so that a circle can be draw correctly.
            setXYDirFromAngle xVector, yVector, dblSweepAngle
        End If
    End If
    
    If dblSweepAngle = 0 Then
        Err.Raise vbObjectError + 53001, "CreateCircularTorus", "Can not have a sweep angle of 0"
        Exit Function
    End If
            
    'Only being used with angle given only.  Will not use when both xdir and ydir are passed.
    If Not (oriDir Is Nothing) Then
        loadOriIntoTransformationMatrix tmxMatrix, oriDir
    End If
    
    'Get the start point
    Set vecStart = getVectorFromPointWithAngle(-dblSweepAngle / 2, degreeToRadian(225))
    posStart.Set vecStart.x * dblMajorR, vecStart.y * dblMajorR, vecStart.z * dblMajorR
    vecRev.Set 0, 0, 1

'    #If INDEBUG Then
'        MsgBox "X Dir is " & xVector.x & " " & xVector.y & " " & xVector.z
'        MsgBox "Y Dir is " & yVector.x & " " & yVector.y & " " & yVector.z
'        MsgBox "Sweep Angle is " & radianToDegree(dblSweepAngle)
'        MsgBox "Ori X is " & oriDir.XAxis.x & " " & oriDir.XAxis.y & " " & oriDir.XAxis.z
'        MsgBox "Ori Y is " & oriDir.YAxis.x & " " & oriDir.YAxis.y & " " & oriDir.YAxis.z
'        MsgBox "Ori Z is " & oriDir.ZAxis.x & " " & oriDir.ZAxis.y & " " & oriDir.ZAxis.z
'        MsgBox "Revolving Vector is " & vecRev.x & " " & vecRev.y & " " & vecRev.z
'        MsgBox "Start Position is " & posStart.x & " " & posStart.y & " " & posStart.z
'    #End If
        
    If (intNumberOfSegments = 0) Then
        Set CreateCircularTorus = placeCircularTorus(objOutputColl, posStart, yVector, vecRev, _
                                                    dblMinorR, dblSweepAngle, blnIsCapped, tmxMatrix)
    Else
    
        Set CreateCircularTorus = placeMiterCutCircTorus(objOutputColl, dblSweepAngle, degreeToRadian(dblInsideRadius), _
                                                    degreeToRadian(dblOutsideRadius), intNumberOfSegments, 0, tmxMatrix, blnIsCapped)
    End If
    
    Set vecCenter = Nothing
    Set vecStart = Nothing
    Set posStart = Nothing
    Set xVector = Nothing
    Set yVector = Nothing
    Set vecRev = Nothing
    Set tmxMatrix = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateCone
'   Author:     Dan Fielder
'   Inputs:
'               Output Collection
'               Origin - "X, Y, Z" format
'               Direction - "X, Y, Z" format
'               Diameter
'               Height
'               Optional blnIsCapped
'
'   Outputs:
'               cone symbol
'
'   Description:
'               This function creates a cone symbol based upon "engineering"
'               type inputs.
'
'   Example of Call:
'               Dim objCone as Object
'               Set objCone = CreateCone(m_outputColl, "E 0 N 0 U 0", "N 0 U 0", 3, 2)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
'               Set objCone = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   19.Dec.2002     Dan Fielder     Created
'   04.Feb.2003     HL        Change AutoMath.Dvector to IJDVector
'   04.Feb.2003     HL        Seperate high and lower functions.
'                                   Use transformation matrix to reposition cone
'                                   instead of calculating the points based on position.
'                                   Add input validation.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateCone(ByVal objOutputColl As Object, _
                            ByVal strOrigin As String, _
                            ByVal strDirection As String, _
                            ByVal dblDiameter As Double, _
                            ByVal dblHeight As Double, _
                            Optional ByVal blnIsCapped As Boolean = True) _
                            As IngrGeom3D.Cone3d

    Dim vecBuildAxis As IJDVector
    Dim directionVector As IJDVector
    Dim vecCenter As IJDVector
    Dim centerBase As IJDPosition
    Dim centerTop As IJDPosition
    Dim radiusBase As Double
    Dim radiusTop As Double
    Dim xform As IngrGeom3D.IJDT4x4
    
    Set vecBuildAxis = New DVector
    Set directionVector = New DVector
    Set centerBase = New DPosition
    Set centerTop = New DPosition
    Set vecCenter = New DVector
    Set xform = New DT4x4
        
    'Set default build Axis to E
    vecBuildAxis.Set 1, 0, 0
    
    ' initialize the transform matrix to identity ---
    xform.LoadIdentity

    ' load the transform matrix with a translate from origin to placement position ---
    If Len(strOrigin) > 0 Then
        Set vecCenter = convertPositionStringToDVector(strOrigin)
        xform.Translate vecCenter
    End If
    
    If Len(strDirection) > 0 Then
        Set directionVector = convertDirectionToUnitVector(strDirection)
        Call addDirVectorComponentRotationsToMatrix(xform, directionVector, vecBuildAxis, 0)
    End If

    ' set the center base relative to (0,0,0)
    centerBase.Set -dblHeight / 2, 0, 0
                    
    ' set the center top relative to (0,0,0)
    centerTop.Set dblHeight / 2, 0, 0
                    
    radiusBase = dblDiameter / 2
    radiusTop = 0
    
    Set CreateCone = placeTransformableCone(objOutputColl, centerBase, centerTop, radiusBase, radiusTop, blnIsCapped, xform)

    ' clean up
    Set directionVector = Nothing
    Set centerBase = Nothing
    Set centerTop = Nothing
    Set vecCenter = Nothing
    Set xform = Nothing
        
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateConicalTankHead
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl     => The collection object
'               strOrigin         => The coordinate location of tank head origin
'                                    expressed as a string.
'               strDirection      => A text description of the direction of the tank head.
'               dblConeDiameter   => Capping Diameter of tank head.
'               dblEndDiameter    => Diameter of the cone top of the tank head.
'               dblConeHeight     => Height of the conical tank head.
'               blnIsCapped       => Boolean indicating whether to cap the open
'                                    end of the tank head.
'
'   Outputs:
'               Object - Returns a generic object that is a conical tank head.
'
'   Description:
'               Returns a conical tank head as a generic object.  The tank head
'               can be given a position and orientation in 3d space.
'
'   Example of Call:
'               Dim objTankHead as Object
'               Set objTankHead = CreateConicalTankHead(m_OutputColl, "E 0 N 0 U 0", "E 90 U 0", 0.75, 0.20, 1.0, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTankHead
'               Set objTankHead = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   13.Feb.2003     HL        Checked length of strOrigin
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateConicalTankHead _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByVal strDirection As String, _
   ByVal dblConeDiameter As Double, _
   ByVal dblEndDiameter As Double, _
   ByVal dblConeHeight As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.Revolution3d

   Dim xform       As IngrGeom3D.IJDT4x4
   Dim vecTmp      As IJDVector
   Dim vecDir      As IJDVector
   
   Dim posConeLineStartPnt As IJDPosition
   Dim posConeLineEndPnt As IJDPosition
   
   Dim posConeTopStartPnt As IJDPosition
   Dim posConeTopEndPnt As IJDPosition
   
   Dim posCappingLineStartPnt As IJDPosition
   Dim posCappingLineEndPnt As IJDPosition
   
   Dim posTankHeadBuildOrigin As IJDPosition
   
   Dim PI As Double

'  - - - - - - - - - - - - - - - - - - - - - - -
   
   Set vecTmp = New DVector
   Set vecDir = New DVector
   Set xform = New DT4x4
   
   Set posConeLineStartPnt = New DPosition
   Set posConeLineEndPnt = New DPosition
   
   Set posConeTopStartPnt = New DPosition
   Set posConeTopEndPnt = New DPosition
   
   Set posCappingLineStartPnt = New DPosition
   Set posCappingLineEndPnt = New DPosition
   
   Set posTankHeadBuildOrigin = New DPosition
   
'  - - - - - - - - - - - - - - - - - - - - - - -
      
   PI = 4 * Atn(1)

'  --- build the dish at position 0,0,0 ---
   posTankHeadBuildOrigin.Set 0, 0, 0

'  --- initialize the transform matrix to identity ---
   xform.LoadIdentity

   If Len(strOrigin) > 0 Then
        '  --- load the transform matrix with a translate from origin to placement position ---
        Set vecTmp = convertPositionStringToDVector(strOrigin)
        xform.Translate vecTmp
   End If
'  --- get the direction vector of the dish ---
   Set vecDir = convertDirectionToUnitVector(strDirection)
   
'  --- set primary build axis indicator (using a vector) in this case the X axis ---
   vecTmp.Set 1, 0, 0

'  --- concatenate component rotations of dish direction vector onto xform matrix ---
   Call addDirVectorComponentRotationsToMatrix(xform, vecDir, vecTmp, 0)
   
'  --- compute cone line start and end points ---
   posConeLineStartPnt.Set posTankHeadBuildOrigin.x, posTankHeadBuildOrigin.y - (dblConeDiameter / 2), posTankHeadBuildOrigin.z
   posConeLineEndPnt.Set posTankHeadBuildOrigin.x + dblConeHeight, posTankHeadBuildOrigin.y - (dblEndDiameter / 2), posTankHeadBuildOrigin.z

'  --- compute cone top start and end points ---
   posConeTopStartPnt.Set posConeLineEndPnt.x, posConeLineEndPnt.y, posConeLineEndPnt.z
   posConeTopEndPnt.Set posTankHeadBuildOrigin.x + dblConeHeight, posTankHeadBuildOrigin.y, posTankHeadBuildOrigin.z

'  --- compute capping line start and end points ---
   posCappingLineStartPnt.Set posTankHeadBuildOrigin.x, posTankHeadBuildOrigin.y, posTankHeadBuildOrigin.z
   posCappingLineEndPnt.Set posTankHeadBuildOrigin.x, posTankHeadBuildOrigin.y - (dblConeDiameter / 2), posTankHeadBuildOrigin.z


'  --- call low level function to get object ---
   Set CreateConicalTankHead = placeConicalTankHead(objOutputColl, _
      posConeLineStartPnt, posConeLineEndPnt, _
      posConeTopStartPnt, posConeTopEndPnt, _
      posCappingLineStartPnt, posCappingLineEndPnt, _
      blnIsCapped, xform)


'  --- cleanup ---
   Set vecTmp = Nothing
   Set vecDir = Nothing
   Set xform = Nothing
   Set posConeLineStartPnt = Nothing
   Set posConeLineEndPnt = Nothing
   Set posConeTopStartPnt = Nothing
   Set posConeTopEndPnt = Nothing
   Set posCappingLineStartPnt = Nothing
   Set posCappingLineEndPnt = Nothing
   Set posTankHeadBuildOrigin = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateCylinder
'   Author:     JG
'   Inputs:
'               output collection
'               origin - "X, Y, Z" format
'               diameter
'               length
'           Optional
'               direction
'               orientation
'               blnIsCapped
'
'   Outputs:
'               cylinder
'
'   Description:
'               This function creates a cylinder symbol based upon "engineering"
'               type inputs.
'
'   Example of Call:
'               Dim position   As String
'               Dim direction  As String
'               Dim diameter   As Double
'               Dim length     As Double
'               Dim objCylinder  As object
'           Optional Choose one of the following.
'               position = "N 5 W 5 D 2"
'               direction = "N 10 U 3 E"  reads north 10 degrees up 3 degrees east
'           Optional
'               Rotate to "N 10 U 3 E"
'               Dim myOri As Orientation
'               Set myOri = New Orientation
'               myOri.RotationAboutY = 10
'               myOri.RotationAboutZ = 87
'           End Optional
'               length = 10
'               diamter = 3
'           With Direction
'               set objCylinder = createCylinder(m_OutputColl, position, length, diameter, direction)
'           With Orientation
'               set objCylinder = createCylinder(m_OutputColl, position, length, diameter,, myOri)
'
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'               Set objCylinder = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   19.Dec.2002     JG      created
'   27.Jan.2003     JG      Replaced with newer version of code.
'                                   and fixed call to
'                                   addDirVectorComponentRotationsToMatrix
'   13.Feb.2003     HL        checked existence of strOrigin
'
'   To Do List:
'       change position.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateCylinder(ByVal objOutputColl As Object, _
                           ByVal strOrigin As String, _
                           ByVal dblDiameter As Double, _
                           ByVal dblLength As Double, _
                           Optional ByVal strDirection As String, _
                           Optional ByRef oriPlacement As Orientation, _
                           Optional ByVal blnIsCapped As Boolean = True) _
                           As IngrGeom3D.Projection3d
                                        
    Dim posCenter As IJDVector
    Set posCenter = New DVector
    Dim posStart As IJDPosition
    Set posStart = New DPosition
    Dim posEnd As IJDPosition
    Set posEnd = New DPosition
    Dim vecDirection As IJDVector
    Set vecDirection = New DVector
    Dim vecBuildAxis As IJDVector
    Set vecBuildAxis = New DVector
    
    vecBuildAxis.Set 1, 0, 0
    
    'Used to move and orientate the object after construction.
    Dim tmxMatrix As IJDT4x4
    Set tmxMatrix = New DT4x4
    
    tmxMatrix.LoadIdentity
    
    If Len(strOrigin) > 0 Then
        Set posCenter = convertPositionStringToDVector(strOrigin)
        'Move object into position
        tmxMatrix.Translate posCenter
    End If
    
    'Add the necessary offset for creation.
    posEnd.Set -dblLength / 2, 0, 0
    posStart.Set dblLength / 2, 0, 0
    
    'Direction takes precidence over an orientation.
    If Len(strDirection) > 0 Then
        Set vecDirection = convertDirectionToUnitVector(strDirection)
        Call addDirVectorComponentRotationsToMatrix(tmxMatrix, vecDirection, vecBuildAxis, 0)
    Else
        If Not (oriPlacement Is Nothing) Then
            Call loadOriIntoTransformationMatrix(tmxMatrix, oriPlacement)
        End If
    End If
    
    'Create the cylinder
    Set CreateCylinder = PlaceTransformableCylinder(objOutputColl, posStart, posEnd, dblDiameter, blnIsCapped, tmxMatrix)
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateDefaultPlane
'   Author:     HL
'   Inputs:
'               Output collection object
'               Origin
'               Plane X Length
'               Plane Y Length
'               Direction
'               Transformation Matrix
'
'   Outputs:
'               A Plane based on input.
'
'   Description:
'               This function places a plane.  It uses right-hand rule to determine
'               the direction of the plane.
'
'   Example of Call:
'               Dim objPlane as Object
'               Set objPlane = CreateDefaultPlane(m_outputColl, "E 0 N 0 U 0", 0.2, 0.1, "D")
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
'               Set objPlane = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   14.Feb.2003     HL        Created this function
'   04.Mar.2003     HL        Added Transformation Matrix
'   06.Mar.2003     HL        Added direction.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateDefaultPlane(ByVal objOutputColl As Object, _
                                    ByVal strOrigin As String, _
                                    ByVal dblXLength As Double, _
                                    ByVal dblYLength As Double, _
                                    ByVal strPlaneNormal As String, _
                                    Optional ByRef oriDir As Orientation) As IngrGeom3D.Plane3d
                                    
    Dim tmxForm As IJDT4x4
    Dim vecCenter As IJDVector
    Dim posVerts(3) As IJDPosition
    Dim oPlane As IngrGeom3D.Plane3d
    Dim i As Integer
    
    Set tmxForm = New DT4x4
    tmxForm.LoadIdentity
    
    If Len(strOrigin) > 0 Then
        Set vecCenter = convertPositionStringToDVector(strOrigin)
        tmxForm.Translate vecCenter
    End If
    
    If Not (oriDir Is Nothing) Then
        loadOriIntoTransformationMatrix tmxForm, oriDir
    End If
    
    For i = 0 To 3
        Set posVerts(i) = New DPosition
    Next i
    
    Select Case UCase(strPlaneNormal)
        Case "U"
            posVerts(0).Set -dblXLength / 2, -dblYLength / 2, 0
            posVerts(1).Set dblXLength / 2, -dblYLength / 2, 0
            posVerts(2).Set dblXLength / 2, dblYLength / 2, 0
            posVerts(3).Set -dblXLength / 2, dblYLength / 2, 0
        Case "D"
            posVerts(3).Set -dblXLength / 2, -dblYLength / 2, 0
            posVerts(2).Set dblXLength / 2, -dblYLength / 2, 0
            posVerts(1).Set dblXLength / 2, dblYLength / 2, 0
            posVerts(0).Set -dblXLength / 2, dblYLength / 2, 0
        Case Else
            Err.Raise vbObjectError + 53003, "CreateDefaultPlane", "Invalid string for Normal to plane, must be U or D"
            Exit Function
    End Select
    
    Set oPlane = PlacePlane(objOutputColl, posVerts, 4, tmxForm)
        
    Set CreateDefaultPlane = oPlane
    Set oPlane = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateDish
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl      => The collection object
'               strOrigin          => The coordinate location of dish origin
'                                     expressed as a string.
'               strDirection       => A text description of the direction of the dish.
'               dblDiameter        => Diameter of dish.
'               dblHeight          => Height of dish (also known as depth of dish)
'               dblKnuckleRadius   => 0 produces a spherical dish. Any nonzero produces
'                                     an elliptical dish with desired diameter and height
'                                     that approximates closely a dish with a knuckle radius.
'               blnIsCapped        => Boolean indicating whether to cap the open end of dish.
'
'   Outputs:
'               Object - Returns a generic object that is a dish.
'
'   Description:
'               Returns a dish as a generic object.  The dish can be given a
'               position and orientation in 3d space.
'
'   Example of Call:
'               Dim objDish as Object
'               Set objDish = CreateDish(m_OutputColl, "E 0 N 0 U 0", "E 90 U 0", 0.75, 0.25, 1, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDish
'               Set objDish = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   31.Jan.2003     Doug Hempel     Implemented dish capping capability
'   13.Feb.2003     HL        Checked existence of strOrigin and strDirection
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateDish _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByVal strDirection As String, _
   ByVal dblDiameter As Double, _
   ByVal dblHeight As Double, _
   ByVal dblKnuckleRadius As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.Revolution3d

   Dim xform       As IngrGeom3D.IJDT4x4
   Dim vecTmp      As IJDVector
   Dim vecDir      As IJDVector
   Dim posDishArcRadiusPnt As IJDPosition
   Dim posDishArcStartPnt As IJDPosition
   Dim posDishArcEndPnt As IJDPosition
   Dim posDishBuildOrigin As IJDPosition
   
   Set vecTmp = New DVector
   Set vecDir = New DVector
   Set xform = New DT4x4
   Set posDishArcRadiusPnt = New DPosition
   Set posDishArcStartPnt = New DPosition
   Set posDishArcEndPnt = New DPosition
   Set posDishBuildOrigin = New DPosition
   
   Dim dblDishArcRadius As Double
   
   Dim PI As Double

   PI = 4 * Atn(1)

'  --- initialize the transform matrix to identity ---
   xform.LoadIdentity

   If Len(strOrigin) > 0 Then
        '  --- load the transform matrix with a translate from origin to placement position ---
        Set vecTmp = convertPositionStringToDVector(strOrigin)
        xform.Translate vecTmp
   End If
      
'  --- set primary build axis indicator (using a vector) in this case the X axis ---
   vecTmp.Set 1, 0, 0

   If Len(strDirection) > 0 Then
        '  --- get the direction vector of the dish ---
           Set vecDir = convertDirectionToUnitVector(strDirection)
        
        '  --- concatenate component rotations of dish direction vector onto xform matrix ---
           Call addDirVectorComponentRotationsToMatrix(xform, vecDir, vecTmp, 0)
   End If
   
'  --- build the dish at position 0,0,0 ---
   posDishBuildOrigin.Set 0, 0, 0

'  --- compute arc radius of dish from height and diameter ---
   dblDishArcRadius = ((((dblDiameter / 2) ^ 2) / dblHeight) + dblHeight) / 2
   
'  --- compute dish arc center, start and end points ---
   posDishArcRadiusPnt.Set posDishBuildOrigin.x + dblHeight - dblDishArcRadius, posDishBuildOrigin.y, posDishBuildOrigin.z
   posDishArcStartPnt.Set posDishBuildOrigin.z, posDishBuildOrigin.y - dblDiameter / 2, posDishBuildOrigin.z
   posDishArcEndPnt.Set posDishBuildOrigin.x + dblHeight, posDishBuildOrigin.y, posDishBuildOrigin.z

'  --- use knuckle radius as switch only to determine which kind of dish ---
   If dblKnuckleRadius = 0 Then
      Set CreateDish = placeSphericalDish(objOutputColl, posDishArcRadiusPnt, posDishArcStartPnt, posDishArcEndPnt, blnIsCapped, xform)
   Else
      Set CreateDish = placeEllipticalDish(objOutputColl, dblDiameter, dblHeight / (dblDiameter / 2), blnIsCapped, xform)
   End If

'  --- cleanup ---
   Set vecTmp = Nothing
   Set vecDir = Nothing
   Set xform = Nothing
   Set posDishArcRadiusPnt = Nothing
   Set posDishArcStartPnt = Nothing
   Set posDishArcEndPnt = Nothing
   Set posDishBuildOrigin = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateHemisphereTankHead
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl    => The collection object
'               strOrigin        => The coordinate location of dish origin expressed
'                                   as a string.
'               strDirection     => A text description of the direction of the tank head.
'               dblDiameter      => Capping diameter of tank head.
'               blnIsCapped      => Boolean indicating whether to cap the open end of
'                                   the tank head.
'
'   Outputs:
'               Object - Returns a generic object that is a tank head with a fixed
'               height to diameter ratio is appropriate for creating a hemisphere.
'
'   Description:
'               Returns a hemispherical tank head as a generic object.
'               The tank head can be given a position and orientation in 3d space.
'               It calls the CreateDish function with height = diameter/2.
'
'   Example of Call:
'               Dim objTankHead as Object
'               Set objTankHead = CreateHemisphereTankHead(m_OutputColl, "E 0 N 0 U 0", "E 90 U 0", 0.75, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTankHead
'               Set objTankHead = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateHemisphereTankHead _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByVal strDirection As String, _
   ByVal dblDiameter As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.Revolution3d

   Set CreateHemisphereTankHead = CreateDish(objOutputColl, strOrigin, strDirection, dblDiameter, dblDiameter / 2, 0, blnIsCapped)

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateKLRatioTankHead
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl     => The collection object
'               strOrigin         => The coordinate location of tank head origin
'                                    expressed as a string.
'               strDirection      => A text description of the direction of the tank head.
'               dblDomeDiameter   => Diameter of tank head dome.
'               dblKValue         => Constant K in formula ACOS((0.5-L)/(K-L)) = KnuckleRadius Angle
'               dblLValue         => Constant L in formula ACOS((0.5-L)/(K-L)) = KnuckleRadius Angle.
'               blnIsCapped       => Boolean indicating whether to cap the open end
'                                    of the tank head.
'
'   Outputs:
'               Object - Returns a generic object that is a tank head.
'
'   Description:
'               Returns a knuckle radius tank head as a generic object.  The tank
'               head can be given a position and orientation in 3d space.
'
'   Example of Call:
'               Dim objTankHead as Object
'               Set objTankHead = CreateKLRatioTankHead(m_OutputColl, "E 0 N 0 U 0", "E 90 U 0", 0.75, 1.0, 0.06, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTankHead
'               Set objTankHead = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   13.Feb.2003     HL        Check existence of strOrigin and strDirection.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateKLRatioTankHead _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByVal strDirection As String, _
   ByVal dblDomeDiameter As Double, _
   ByVal dblKValue As Double, _
   ByVal dblLValue As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.Revolution3d

   Dim xform       As IngrGeom3D.IJDT4x4
   Dim vecTmp      As IJDVector
   Dim vecDir      As IJDVector
   
   Dim posDomeArcRadiusPnt As IJDPosition
   Dim posDomeArcStartPnt As IJDPosition
   Dim posDomeArcEndPnt As IJDPosition
   
   Dim posKnuckleArcRadiusPnt As IJDPosition
   Dim posKnuckleArcStartPnt As IJDPosition
   Dim posKnuckleArcEndPnt As IJDPosition
   
   Dim posCappingLineStartPnt As IJDPosition
   Dim posCappingLineEndPnt As IJDPosition
   
   Dim posTankHeadBuildOrigin As IJDPosition
   
   Dim dblDomeRadius As Double
   Dim dblKnuckleRadius As Double
   Dim dblKnuckleAngle As Double
   Dim dblDomeHalfAngle As Double
   Dim dblDomeHeight As Double
   
   Dim PI As Double

'  - - - - - - - - - - - - - - - - - - - - - - -
   
   Set vecTmp = New DVector
   Set vecDir = New DVector
   Set xform = New DT4x4
   
   Set posDomeArcRadiusPnt = New DPosition
   Set posDomeArcStartPnt = New DPosition
   Set posDomeArcEndPnt = New DPosition
   
   Set posKnuckleArcRadiusPnt = New DPosition
   Set posKnuckleArcStartPnt = New DPosition
   Set posKnuckleArcEndPnt = New DPosition
   
   Set posCappingLineStartPnt = New DPosition
   Set posCappingLineEndPnt = New DPosition
   
   Set posTankHeadBuildOrigin = New DPosition
   
      
   PI = 4 * Atn(1)

'  --- build the tank head at position 0,0,0 ---
   posTankHeadBuildOrigin.Set 0, 0, 0

'  --- initialize the transform matrix to identity ---
   xform.LoadIdentity

   If Len(strOrigin) > 0 Then
        '  --- load the transform matrix with a translate from origin to placement position ---
           Set vecTmp = convertPositionStringToDVector(strOrigin)
           xform.Translate vecTmp
   End If
   
'  --- set primary build axis indicator (using a vector) in this case the X axis ---
   vecTmp.Set 1, 0, 0

   If Len(strDirection) > 0 Then
        '  --- get the direction vector of the dish ---
           Set vecDir = convertDirectionToUnitVector(strDirection)
        
        '  --- concatenate component rotations of dish direction vector onto xform matrix ---
           Call addDirVectorComponentRotationsToMatrix(xform, vecDir, vecTmp, 0)
   End If
   
'  --- compute dome arc radius ---
   dblDomeRadius = dblKValue * dblDomeDiameter

'  --- compute knuckle arc radius ---
   dblKnuckleRadius = dblLValue * dblDomeDiameter

'  --- compute knuckle arc angle value ---
   dblKnuckleAngle = ACos((0.5 - dblLValue) / (dblKValue - dblLValue))

'  --- compute 1/2 dome arc angle value ---
   dblDomeHalfAngle = (degreeToRadian(90) - dblKnuckleAngle)
   
'  --- compute dome height value ---
   dblDomeHeight = dblDomeRadius - (Cos(dblDomeHalfAngle) * (dblDomeRadius - dblKnuckleRadius))

'  --- compute knuckle arc center, start and end points ---
   posKnuckleArcRadiusPnt.Set posTankHeadBuildOrigin.x, posTankHeadBuildOrigin.y - ((dblDomeDiameter / 2) - dblKnuckleRadius), posTankHeadBuildOrigin.z
   posKnuckleArcStartPnt.Set posTankHeadBuildOrigin.x, posTankHeadBuildOrigin.y - (dblDomeDiameter / 2), posTankHeadBuildOrigin.z
   posKnuckleArcEndPnt.Set posTankHeadBuildOrigin.x + (Sin(dblKnuckleAngle) * dblKnuckleRadius), posTankHeadBuildOrigin.y - ((Cos(dblKnuckleAngle) * dblKnuckleRadius) + (dblDomeDiameter / 2 - dblKnuckleRadius)), posTankHeadBuildOrigin.z

'  --- compute dome arc center, start and end points ---
   posDomeArcRadiusPnt.Set posTankHeadBuildOrigin.x - (dblDomeRadius - dblDomeHeight), posTankHeadBuildOrigin.y, posTankHeadBuildOrigin.z
   posDomeArcStartPnt.Set posKnuckleArcEndPnt.x, posKnuckleArcEndPnt.y, posKnuckleArcEndPnt.z
   posDomeArcEndPnt.Set posTankHeadBuildOrigin.x + dblDomeHeight, posTankHeadBuildOrigin.y, posTankHeadBuildOrigin.z

'  --- compute capping line start and end points ---
   posCappingLineStartPnt.Set posTankHeadBuildOrigin.x, posTankHeadBuildOrigin.y, posTankHeadBuildOrigin.z
   posCappingLineEndPnt.Set posTankHeadBuildOrigin.x, posTankHeadBuildOrigin.y - (dblDomeDiameter / 2), posTankHeadBuildOrigin.z


'  --- call low level function to get object ---
   Set CreateKLRatioTankHead = placeKnuckleRadiusTankHead(objOutputColl, _
      posKnuckleArcRadiusPnt, posKnuckleArcStartPnt, posKnuckleArcEndPnt, _
      posDomeArcRadiusPnt, posDomeArcStartPnt, posDomeArcEndPnt, _
      posCappingLineStartPnt, posCappingLineEndPnt, _
      blnIsCapped, xform)


'  --- cleanup ---
   Set vecTmp = Nothing
   Set vecDir = Nothing
   Set xform = Nothing
   Set posDomeArcRadiusPnt = Nothing
   Set posDomeArcStartPnt = Nothing
   Set posDomeArcEndPnt = Nothing
   Set posKnuckleArcRadiusPnt = Nothing
   Set posKnuckleArcStartPnt = Nothing
   Set posKnuckleArcEndPnt = Nothing
   Set posCappingLineStartPnt = Nothing
   Set posCappingLineEndPnt = Nothing
   Set posTankHeadBuildOrigin = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreatePoint
'   Author:     HL
'   Inputs:
'               Output collection object
'               Point
'               transformation matrix
'
'   Outputs:
'               Point
'
'   Description:
'               This function outputs a point.
'
'   Example of Call:
'               Dim objPoint As Object
'               Set objPoint = CreatePoint(m_outputColl, CenterPos)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPoint
'               Set objPoint = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   03.Feb.2003     HL        Created this function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreatePoint(ByVal objOutputColl As Object, _
                            ByRef posPoint As IJDPosition, _
                            Optional ByRef tmxMatrix As IJDT4x4) As IngrGeom3D.Point3d
    
    Dim objGeom As IngrGeom3D.GeometryFactory
    Dim objPoint As IngrGeom3D.Point3d
    
    Set objGeom = New IngrGeom3D.GeometryFactory
    Set objPoint = objGeom.Points3d.CreateByPoint(objOutputColl.ResourceManager, posPoint.x, posPoint.y, posPoint.z)
    
    If Not (tmxMatrix Is Nothing) Then
        objPoint.Transform tmxMatrix
    End If
    
    Set CreatePoint = objPoint
    Set objPoint = Nothing
    Set objGeom = Nothing
    
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreatePrism
'   Author:     HL
'   Inputs:
'               Output collection object
'               Origin
'               Orientation
'               Length
'               Vertices in X and Y Coordinates
'               Optional blnIsCapped
'
'   Outputs:
'               Prism based on input
'
'   Description:
'               This function outputs a prism.  It uses the CreateBasePrism
'               function.
'
'   Example of Call:
'               Dim oriOrientation As Orientation
'               Dim objPyramid As Object
'               Set oriOrientation = New Orientation
'               oriOrientation.ResetDefaultAxis
'               strOrigin = "E 0 N 0 U 0"
'               Set objPrism = CreatePrism(m_OutputColl, "E 0 N 0 U 0", oriOrientation, 1, "0,-1,2,0,1,1,-1,1,-2,0")
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPrism
'               Set objPrism = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   13.Jan.2003     HL        Created this function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreatePrism(ByVal objOutputColl As Object, _
                            ByVal strOrigin As String, _
                            ByRef oriOrientation As Orientation, _
                            ByVal dblLength As Double, _
                            ByVal strVertices As String, _
                            Optional ByVal blnIsCapped As Boolean = True) _
                            As IngrGeom3D.RuledSurface3d

    Set CreatePrism = CreateBasePrism(objOutputColl, strOrigin, oriOrientation, _
                      dblLength, strVertices, , , , , , , blnIsCapped)
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Function:   CreatePrismaticTorus
'   Author:     KV and HL
'   Inputs:
'               Output collection object
'               Origin
'               Radius
'               Number of Segments
'               Vertices in X and Y coordinates
'               X Direction
'               Y Direction
'               Angle
'               Orientation
'               Optional blnIsCapped
'
'   Outputs:
'               Prismatic Torus based on input
'
'   Description:
'               This function creates a prismatic torus symbol based upon
'               "engineering" type inputs.  They are two ways to input:
'               1.  By giving strXDirection and strYDirection, it will compute
'               the angle, the angle passed in as a parameter will be ignored.
'               It will always use the smaller than 180 angle.
'               2.  By just inputing an angle, it will compute xdir and ydir
'               on xy-plane.  The user must use ori to orient it.
'
'               If all 4 inputs are given, it will use the first method.
'
'   Example of Call:
'               Dim objPTOR as Object
'               Set objPTOR = CreatePrismaticTorus(m_outputColl, "E 0 N 0 U 0", 5, 0, "0,-1,2,0,1,1,-1,1,-2,0", "", "", 350)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPTOR
'               Set objPTOR = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   10.Jan.2003     KV and Hong  Created this function
'   19.Feb.2003     DH and HL       Added xdir and ydir functionality.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreatePrismaticTorus(ByVal objOutputColl As Object, _
                            ByVal strOrigin As String, _
                            ByVal dblRadius As Double, _
                            ByVal intSegments As Integer, _
                            ByVal strVertices As String, _
                            Optional ByVal strXDirection As String, _
                            Optional ByVal strYDirection As String, _
                            Optional ByVal dblAngle As Double, _
                            Optional ByRef oriOrientation As Orientation, _
                            Optional ByVal blnIsCapped As Boolean = True) _
                            As IngrGeom3D.Revolution3d

    Dim intVertices As Integer
    Dim strVerts() As String
    Dim arrPts() As IJDPosition
    Dim Xdir As IJDVector
    Dim Ydir As IJDVector
    Dim i As Integer
    Dim x, y, z As Double
    
    Dim tmxForm As IJDT4x4
    Dim vecRevVector As IJDVector
    Dim vecStart As IJDVector
    Dim posStart As IJDPosition
    Dim vecPosition As IJDVector
    
    Dim tmxTemp As IJDT4x4
    
    'Get x and y coordinates into an array
    strVerts = Split(strVertices, ",")
    intVertices = UBound(strVerts) + 1

    'Check to see if we have both x and y inputs
    If (intVertices Mod 2) <> 0 Then
        Err.Raise vbObjectError + 53000, "CreatePrismaticTorus", "Must input both x and y"
        Exit Function
    End If
    intVertices = intVertices / 2

    ReDim arrPts(intVertices - 1) As IJDPosition
    
    Set vecStart = New DVector
    Set posStart = New DPosition
    Set vecRevVector = New DVector
    Set tmxForm = New DT4x4
    'initialize the transformation matrix to identify
    tmxForm.LoadIdentity

    If Len(strOrigin) > 0 Then
        'load the transformation matrix with a translate from origin to placement position
        Set vecPosition = convertPositionStringToDVector(strOrigin)
        tmxForm.Translate vecPosition
    End If
    
    'Check to see if X and Y direction has a value
    If Len(strXDirection) > 0 Then
        Set Xdir = convertDirectionToUnitVector(strXDirection)
    Else
        Set Xdir = New DVector
        Xdir.Set 1, 0, 0
    End If
    
    If Len(strYDirection) > 0 Then
        Set Ydir = convertDirectionToUnitVector(strYDirection)
    Else
        Set Ydir = New DVector
        Ydir.Set 0, 1, 0
    End If
    
    If Len(strXDirection) = 0 And Len(strYDirection) = 0 Then
        'if doesn't have xdir and ydir, use the angle
        'convert the sweep angle to radians from degrees
        dblAngle = degreeToRadian(dblAngle)
        setXYDirFromAngle Xdir, Ydir, dblAngle
        vecRevVector.Set 0, 0, 1
    Else
        'if xdir and ydir are pointing to the opposite direction.
        If Xdir.x = -Ydir.x And Xdir.y = -Ydir.y And Xdir.z = -Ydir.z Then
            Err.Raise vbObjectError + 53002, "CreatePrismaticTorus", "Can not have the directions that are opposite."
            Exit Function
        Else
            'The angle will not be used.  It will be calculated from xdir and ydir
            dblAngle = degreeToRadian(180) - getAngleBetween2VectorsIn3D(Xdir, Ydir)
            Set vecRevVector = Xdir.Cross(Ydir)
            vecRevVector.Length = 1
        End If
    End If
    
    If dblAngle = 0 Then
        Err.Raise vbObjectError + 53001, "CreatePrismaticTorus", "Can not have a sweep angle of 0"
        Exit Function
    End If

'    #If INDEBUG Then
'        MsgBox "X Dir is " & Xdir.x & " " & Xdir.y & " " & Xdir.z
'        MsgBox "Y Dir is " & Ydir.x & " " & Ydir.y & " " & Ydir.z
'        MsgBox "Sweep Angle is " & radianToDegree(dblAngle)
'        MsgBox "Revolving Vector is " & vecRevVector.x & " " & vecRevVector.y & " " & vecRevVector.z
'    #End If

    'Only being used with angle given only.  Will not use when both xdir and ydir are passed.
    If Len(strXDirection) = 0 And Len(strYDirection) = 0 Then
        If Not (oriOrientation Is Nothing) Then
            loadOriIntoTransformationMatrix tmxForm, oriOrientation
        End If
        'Get the center point of the polygon
        Set vecStart = getVectorFromPointWithAngle(dblAngle / 2, degreeToRadian(225))
        posStart.Set vecStart.x * dblRadius, vecStart.y * dblRadius, vecStart.z * dblRadius
    Else
        'Get the center point of the polygon
        Set posStart = getDPosOnSphereFrom2Vectors(Xdir, Ydir, dblRadius, False)
    End If
    
'    #If INDEBUG Then
'        MsgBox "Translate Vector is " & vecStart.x & " " & vecStart.y & " " & vecStart.z
'    #End If
    
    'Set a seperate ori and transformation matrix for transform polygon vertices.
    Set tmxTemp = New DT4x4
    tmxTemp.LoadIdentity
    
    vecStart.Set posStart.x, posStart.y, posStart.z
    tmxTemp.Translate vecStart
    
    'Build Temporary Matrix to transform polygon points
    Set vecStart = vecRevVector.Cross(Xdir)
    vecStart.Length = 1
    tmxTemp.IndexValue(0) = vecStart.x
    tmxTemp.IndexValue(1) = vecStart.y
    tmxTemp.IndexValue(2) = vecStart.z
    tmxTemp.IndexValue(4) = vecRevVector.x
    tmxTemp.IndexValue(5) = vecRevVector.y
    tmxTemp.IndexValue(6) = vecRevVector.z
    tmxTemp.IndexValue(8) = Xdir.x
    tmxTemp.IndexValue(9) = Xdir.y
    tmxTemp.IndexValue(10) = Xdir.z
    
'    For i = 0 To 15
'        logError tmxTemp.IndexValue(i)
'    Next i
    
    For i = 0 To intVertices - 1
        Set arrPts(i) = New DPosition
        x = CDbl(strVerts(i * 2))
        y = CDbl(strVerts(i * 2 + 1))
        z = 0
        posStart.Set x, y, z
        Set arrPts(i) = getTransformedDPos(posStart, tmxTemp)
    Next i

    vecRevVector.Set -vecRevVector.x, -vecRevVector.y, -vecRevVector.z
    
    Set CreatePrismaticTorus = placePrismaticNedgeTorus(objOutputColl, intVertices, arrPts, vecRevVector, dblAngle, tmxForm, blnIsCapped)
    
    For i = 0 To intVertices - 1
        Set arrPts(i) = Nothing
    Next i
    Set tmxForm = Nothing
    Set tmxTemp = Nothing
    
    Set vecRevVector = Nothing
    Set vecPosition = Nothing
    Set Xdir = Nothing
    Set Ydir = Nothing
    Set vecStart = Nothing
    Set posStart = Nothing
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreatePyramid
'   Author:     HL
'   Inputs:
'               Output collection object
'               Origin
'               Orientation
'               XBottom
'               YBottom
'               XTop
'               YTop
'               XOffset
'               YOffset
'               Height
'               Optional blnIsCapped
'
'   Outputs:
'               Pyramid
'
'   Description:
'               This function places a pyramid.
'
'   Example of Call:
'               Dim strOrigin As String
'               Dim objPyramid As Object
'               strOrigin = "E 0 N 0 U 0"
'               Set objPyramid = CreatePyramid(m_OutputColl, strOrigin, "", 1, 1.5, 0.7, 1, 0, 0, 1)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPyramid
'               Set objPyramid = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   10.Jan.2003     HL        Created this function.
'   13.Feb.2003     HL        Checked existence of strOrigin
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreatePyramid(ByVal objOutputColl As Object, _
                          ByVal strOrigin As String, _
                          ByRef oriOrientation As Orientation, _
                          ByVal dblXBottom As Double, _
                          ByVal dblYBottom As Double, _
                          ByVal dblXTop As Double, _
                          ByVal dblYTop As Double, _
                          ByVal dblXOffset As Double, _
                          ByVal dblYOffset As Double, _
                          ByVal dblHeight As Double, _
                          Optional ByVal blnIsCapped As Boolean = True) _
                          As IngrGeom3D.RuledSurface3d
        
    Dim tmxForm As IJDT4x4
    Dim vecPosition As IJDVector
    Dim vecDir As IJDVector
    
    Dim arrTopPts(3) As IJDPosition
    Dim arrBottomPts(3) As IJDPosition
        
    Dim i As Integer
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    
    Set tmxForm = New DT4x4
    Set vecPosition = New DVector
    Set vecDir = New DVector
    
    For i = 0 To 3
        Set arrTopPts(i) = New DPosition
        Set arrBottomPts(i) = New DPosition
    Next i
            
    'initialize the transformation matrix to identify
    tmxForm.LoadIdentity

    If Len(strOrigin) > 0 Then
        'load the transformation matrix with a translate from origin to placement position
        Set vecPosition = convertPositionStringToDVector(strOrigin)
        tmxForm.Translate vecPosition
    End If
    
    If Not (oriOrientation Is Nothing) Then
        Call loadOriIntoTransformationMatrix(tmxForm, oriOrientation)
    End If
    
    'get point positions from length and height
    x1 = dblXBottom / 2
    y1 = dblYBottom / 2
    z1 = dblHeight / 2 * -1
    
    x2 = dblXTop / 2
    y2 = dblYTop / 2
    z2 = dblHeight / 2
    
    If dblXTop = 0 Or dblXBottom = 0 Then
        arrTopPts(0).Set -x1 - dblXOffset / 2, -y1 - dblYOffset / 2, z1
        arrTopPts(1).Set x1 - dblXOffset / 2, -y1 - dblYOffset / 2, z1
        arrTopPts(2).Set x2 + dblXOffset / 2, -y2 + dblYOffset / 2, z2
        arrTopPts(3).Set -x2 + dblXOffset / 2, -y2 + dblYOffset / 2, z2
        
        arrBottomPts(0).Set -x1 - dblXOffset / 2, y1 - dblYOffset / 2, z1
        arrBottomPts(1).Set x1 - dblXOffset / 2, y1 - dblYOffset / 2, z1
        arrBottomPts(2).Set x2 + dblXOffset / 2, y2 + dblYOffset / 2, z2
        arrBottomPts(3).Set -x2 + dblXOffset / 2, y2 + dblYOffset / 2, z2
    Else
        If dblYTop = 0 Or dblYBottom = 0 Then
            arrTopPts(0).Set x1 - dblXOffset / 2, -y1 - dblYOffset / 2, z1
            arrTopPts(1).Set x1 - dblXOffset / 2, y1 - dblYOffset / 2, z1
            arrTopPts(2).Set x2 + dblXOffset / 2, y2 + dblYOffset / 2, z2
            arrTopPts(3).Set x2 + dblXOffset / 2, -y2 + dblYOffset / 2, z2
            
            arrBottomPts(0).Set -x1 - dblXOffset / 2, -y1 - dblYOffset / 2, z1
            arrBottomPts(1).Set -x1 - dblXOffset / 2, y1 - dblYOffset / 2, z1
            arrBottomPts(2).Set -x2 + dblXOffset / 2, y2 + dblYOffset / 2, z2
            arrBottomPts(3).Set -x2 + dblXOffset / 2, -y2 + dblYOffset / 2, z2
        Else
            arrTopPts(0).Set -x1 - dblXOffset / 2, -y1 - dblYOffset / 2, z1
            arrTopPts(1).Set x1 - dblXOffset / 2, -y1 - dblYOffset / 2, z1
            arrTopPts(2).Set x1 - dblXOffset / 2, y1 - dblYOffset / 2, z1
            arrTopPts(3).Set -x1 - dblXOffset / 2, y1 - dblYOffset / 2, z1
            
            arrBottomPts(0).Set -x2 + dblXOffset / 2, -y2 + dblYOffset / 2, z2
            arrBottomPts(1).Set x2 + dblXOffset / 2, -y2 + dblYOffset / 2, z2
            arrBottomPts(2).Set x2 + dblXOffset / 2, y2 + dblYOffset / 2, z2
            arrBottomPts(3).Set -x2 + dblXOffset / 2, y2 + dblYOffset / 2, z2
        End If
    End If
        
'    For i = 0 To 3
'        logError arrTopPts(i).x & " " & arrTopPts(i).y & " " & arrTopPts(i).z
'    Next i
'
'    For i = 0 To 3
'        logError arrBottomPts(i).x & " " & arrBottomPts(i).y & " " & arrBottomPts(i).z
'    Next i

    Set CreatePyramid = placeTruncNEdgePrism(objOutputColl, 4, arrTopPts, arrBottomPts, tmxForm, blnIsCapped)

    For i = 0 To 3
        Set arrTopPts(i) = Nothing
        Set arrBottomPts(i) = Nothing
    Next i

    Set tmxForm = Nothing
    Set vecPosition = Nothing
    Set vecDir = Nothing
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateRectangularTorus
'   Author:     JG
'   Inputs:
'               Output Collection
'               Origin - "X, Y, Z" format
'               X Direction
'               Y Direction
'               Inside Radius
'               Outside Radius
'               Angle
'               Number of Segments
'               Height
'               Orientation object
'               Optional blnIsCapped
'
'   Outputs:
'               cylinder
'
'   Description:
'               This function creates a rectangular torus symbol based upon
'               "engineering" type inputs.
'
'   Example of Call:
'               Dim objRTOR as Object
'               Set objRTOR = CreateRectangularTorus(m_outputColl, "e 0 n 0 u 0", 0.5, 1.7, 90, 0, 1)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCTOR
'               Set objRTOR = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   19.Dec.2002     JG      created
'   14.Feb.2003     JG      Changed the default orientation of the output
'   17.Feb.2003     JG      Changed call to placeMiter to placeMiterCutRectTorus
'   26.Mar.2003     JG      Added the ability to choose whether to cap
'                                   the output or not for the miter cut torus.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateRectangularTorus( _
                                    ByVal objOutputColl As Object, _
                                    ByVal strOrigin As String, _
                                    ByVal dblInsideRadius As Double, _
                                    ByVal dblOutsideRadius As Double, _
                                    ByVal dblAngle As Double, _
                                    ByVal intSegments As Integer, _
                                    ByVal dblHeight As Double, _
                                    Optional ByVal strXDirection As String, _
                                    Optional ByVal strYDirection As String, _
                                    Optional ByRef oriNewOri As Orientation, _
                                    Optional ByVal blnIsCapped As Boolean = True) As Object
                                        
'These are used to build the item at the origin
    Dim vecNormal As AutoMath.DVector
    Set vecNormal = New DVector
    vecNormal.Set 0, 0, 1
    Dim posOrigin As AutoMath.DPosition
    Set posOrigin = New DPosition
    posOrigin.Set 0, 0, 0
    
'Get the Angle of Rotation
    Dim revAngle As Double
    revAngle = degreeToRadian(dblAngle)

'Get the center position of the torus
    Dim posCenter As AutoMath.DVector
    Set posCenter = New DVector
    Set posCenter = convertPositionStringToDVector(strOrigin)
 
'Init the point array
    Dim arrayPoints(0 To 3) As IJDPosition
    Set arrayPoints(0) = New DPosition
    Set arrayPoints(1) = New DPosition
    Set arrayPoints(2) = New DPosition
    Set arrayPoints(3) = New DPosition

'Figure the inside points of the original shape
    arrayPoints(0).Set 0, -1 * dblInsideRadius, dblHeight / 2
    arrayPoints(1).Set 0, -1 * dblInsideRadius, -dblHeight / 2

'Figure the outside points of the original shape
    arrayPoints(2).Set 0, -1 * dblOutsideRadius, -dblHeight / 2
    arrayPoints(3).Set 0, -1 * dblOutsideRadius, dblHeight / 2
    
'Set up the transformation matrix to be passed the the place function
    Dim tmxMatrix As IngrGeom3D.IJDT4x4
    Set tmxMatrix = New DT4x4
    tmxMatrix.LoadIdentity
    tmxMatrix.Translate posCenter  'Load matrix with the move
    
    If (strXDirection <> "") Then
        Dim vecXDirection As AutoMath.DVector
        Set vecXDirection = New DVector
        Set vecXDirection = convertDirectionToUnitVector(strXDirection)
        
        Dim vecYDirection As AutoMath.DVector
        Set vecYDirection = New DVector
        Set vecYDirection = convertDirectionToUnitVector(strYDirection)
    
        Dim myOri As Orientation
        Set myOri = New Orientation
        Set myOri.XAxis = vecXDirection
        Set myOri.YAxis = vecYDirection
        myOri.FindRemainingUnsetAxis
        Call loadOriIntoTransformationMatrix(tmxMatrix, myOri) 'Load it with orientation
    ElseIf (Not oriNewOri Is Nothing) Then
        Call loadOriIntoTransformationMatrix(tmxMatrix, oriNewOri)
    End If
       
'Call the lower level function to create the object.
    If (intSegments = 0) Then
        Set CreateRectangularTorus = placePrismaticNedgeTorus(objOutputColl, 4, arrayPoints, vecNormal, revAngle, tmxMatrix, blnIsCapped)
    Else
        Set CreateRectangularTorus = placeMiterCutRectTorus(objOutputColl, revAngle, arrayPoints, intSegments, tmxMatrix, blnIsCapped)
    End If

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateReducingMiterCutCircularTorus
'   Author:     JG
'   Inputs:
'               Output Collection
'               Origin
'               Start Diameter
'               End Diameter
'               Center Line Radius
'               Number of segments
'               X Direction
'               Y Direction
'               Sweep angle
'               Orientation
'               Is Capped
'
'   Outputs:
'               Reducing segmented circular torus symbol
'
'   Description:
'               This function creates a circular torus symbol based upon
'               "engineering" type inputs.  They are two ways to input:
'               1.  By giving strXDirection and strYDirection, it will compute
'               the angle, the angle passed in as a parameter will be ignored.
'               It will always use the smaller than 180 angle.
'               2.  By just inputing an angle, it will compute xdir and ydir
'               on xy-plane.  The user must use ori to orient it.
'
'               If all 4 inputs are given, it will use the first method.
'
'   Example of Call:
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   20.Mar.2003     JG      Copied CreateCircularTorus and adapted
'                                   it for this special case torus.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateReducingMiterCutCircularTorus( _
                                    ByVal objOutputColl As Object, _
                                    ByVal strOrigin As String, _
                                    ByVal dblStartDiameter As Double, _
                                    ByVal dblEndDiameter As Double, _
                                    ByVal dblCenterLineRadius As Double, _
                                    Optional ByVal intNumberOfSegments As Integer, _
                                    Optional ByVal strXDirection As String, _
                                    Optional ByVal strYDirection As String, _
                                    Optional ByVal dblSweepAngle As Double, _
                                    Optional ByRef oriDir As Orientation, _
                                    Optional ByVal blnIsCapped As Boolean = True) As Object

    Dim vecCenter As IJDVector
    Dim vecStart As IJDVector
    Dim posStart As IJDPosition
    Dim xVector As IJDVector
    Dim yVector As IJDVector
    Dim vecRev As IJDVector
    Dim dblMajorR As Double
    Dim dblMinorR As Double
    
    Set vecRev = New DVector
    Set vecCenter = New DVector
    Set vecStart = New DVector
    Set posStart = New DPosition
    
    Dim dlbScaleingFactor As Double
    Dim dblInsideRadius As Double
    Dim dblOutsideRadius As Double
    
    dlbScaleingFactor = dblEndDiameter / dblStartDiameter
    
    If dblStartDiameter > dblEndDiameter Then
        dblInsideRadius = dblCenterLineRadius - (dblStartDiameter / 2)
        dblOutsideRadius = dblCenterLineRadius + (dblStartDiameter / 2)
    Else
        dblInsideRadius = dblCenterLineRadius - (dblEndDiameter / 2)
        dblOutsideRadius = dblCenterLineRadius + (dblEndDiameter / 2)
    End If
    
    dblMinorR = (dblOutsideRadius - dblInsideRadius) / 2
    dblMajorR = dblInsideRadius + dblMinorR

    'Used to move and orientate the object after construction.
    Dim tmxMatrix As IJDT4x4
    Set tmxMatrix = New DT4x4
        
    tmxMatrix.LoadIdentity
    If Len(strOrigin) > 0 Then
        ' convert the origin string to a center position vector
        Set vecCenter = convertPositionStringToDVector(strOrigin)
'        #If INDEBUG Then
'            MsgBox "Center is " & vecCenter.x & " " & vecCenter.y & " " & vecCenter.z
'        #End If
        
        'Move object into position
        tmxMatrix.Translate vecCenter
    End If

    If Len(strXDirection) > 0 Then
        Set xVector = convertDirectionToUnitVector(strXDirection)
    Else
        Set xVector = New DVector
        xVector.Set 1, 0, 0
    End If
    
    If Len(strYDirection) > 0 Then
        Set yVector = convertDirectionToUnitVector(strYDirection)
    Else
        Set yVector = New DVector
        yVector.Set 0, 1, 0
    End If
        
    If Len(strXDirection) = 0 And Len(strYDirection) = 0 Then
        ' if doesn't have xdir and ydir, use the angle
        ' convert the sweep angle to radians from degrees
        dblSweepAngle = degreeToRadian(dblSweepAngle)
        setXYDirFromAngle xVector, yVector, dblSweepAngle
    Else
        'if xVector and yVector are pointing to the opposite direction.
        If xVector.x = -yVector.x And xVector.y = -yVector.y And xVector.z = -yVector.z Then
            Err.Raise vbObjectError + 53002, "CreateCircularTorus", "Can not have the directions that are opposite."
            Exit Function
        Else
            'The angle will not be used.  It will be calculated from
            'xvector and yvector.
            Set oriDir = getOriFromXYDir(xVector, yVector, dblSweepAngle)
            'reset xdir and ydir to flat dirs so that a circle can be draw correctly.
            setXYDirFromAngle xVector, yVector, dblSweepAngle
        End If
    End If
    
    If dblSweepAngle = 0 Then
        Err.Raise vbObjectError + 53001, "CreateCircularTorus", "Can not have a sweep angle of 0"
        Exit Function
    End If
            
    'Only being used with angle given only.  Will not use when both xdir and ydir are passed.
    If Not (oriDir Is Nothing) Then
        loadOriIntoTransformationMatrix tmxMatrix, oriDir
    End If
    
    'Get the start point
    Set vecStart = getVectorFromPointWithAngle(-dblSweepAngle / 2, degreeToRadian(225))
    posStart.Set vecStart.x * dblMajorR, vecStart.y * dblMajorR, vecStart.z * dblMajorR
    vecRev.Set 0, 0, 1

    Set CreateReducingMiterCutCircularTorus = placeMiterCutCircTorus(objOutputColl, dblSweepAngle, degreeToRadian(dblInsideRadius), _
                                                    degreeToRadian(dblOutsideRadius), intNumberOfSegments, _
                                                    dlbScaleingFactor, tmxMatrix, blnIsCapped)

    Set vecCenter = Nothing
    Set vecStart = Nothing
    Set posStart = Nothing
    Set xVector = Nothing
    Set yVector = Nothing
    Set vecRev = Nothing
    Set tmxMatrix = Nothing


End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateSlopeBottomCylinder
'   Author:     JG
'   Inputs:
'               Output Collection
'               Origin
'               Direction
'               X Direction
'               Y Direction
'               Diameter
'               Length
'               Degree of Rotation
'               Optional IsCapped
'
'   Outputs:
'               Slope-Bottom Cylinder
'
'   Description:
'               This function creates a Slope-Bottom Cylinder symbol based upon
'               "engineering" type inputs.
'
'   Example of call:
'               Dim position   As String
'               Dim direction  As String
'               Dim xdir       As string
'               Dim ydir       As string
'               Dim diameter   As Double
'               Dim length     As Double
'               Dim
'               Dim objCylinder  As object
'               position = "N 5 W 5 D 2"
'               direction = "N 10 U 3 E"  reads north 10 degrees up 3 degrees east
'               xdir = "E 45 D"
'               ydir = "Y 45 D"
'               length = 10
'               diamter = 3
'               set objCylinder = createCylinder(m_OutputColl, position, length, diameter, direction)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'               Set objCylinder = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   19.Dec.2002     JG      created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateSlopeBottomCylinder(ByVal objOutputColl As Object, _
                                        ByVal strOrigin As String, _
                                        ByVal strDirection As String, _
                                        ByVal strXdir As String, _
                                        ByVal strYdir As String, _
                                        ByVal dblDiameter As Double, _
                                        ByVal dblLength As Double, _
                                        Optional ByVal dblDegreeOfRotation As Double, _
                                        Optional ByVal blnIsCapped As Boolean = True) As Object
                                                                
    'Dim all vars
    'Used to place the object after construction. Center point
    Dim centerVect As IJDVector
    If Len(strOrigin) > 0 Then
        Set centerVect = convertPositionStringToDVector(strOrigin)
    End If
    
    'Used to orientate the object after construction.   Direction of tube
    Dim vDirection As IJDVector
    If Len(strDirection) > 0 Then
        Set vDirection = convertDirectionToUnitVector(strDirection)
    End If
    
    'Used to orientate the object defines the axis that it is build about
    Dim vecMajorAxis As IJDVector
    Set vecMajorAxis = New DVector
    vecMajorAxis.Set 0, 0, 1
    
    'Orientates slope of the west end of the cylinder
    Dim xVect As IJDVector
    If Len(strXdir) > 0 Then
        Set xVect = convertDirectionToUnitVector(strXdir)
    End If
    
    'Orientates slope of the east end of the cylinder
    Dim yVect As IJDVector
    If Len(strYdir) > 0 Then
        Set yVect = convertDirectionToUnitVector(strYdir)
    End If
    
    'Used in the construction of the circles.  Defines the pos of the westerly circle
    Dim startPos As IJDPosition
    Set startPos = New DPosition
    startPos.Set -1 * (dblLength / 2), 0, 0
    
    'Used in the construction of the circles.  Defines the pos of the easterly circle
    Dim endPos As IJDPosition
    Set endPos = New DPosition
    endPos.Set (dblLength / 2), 0, 0
    
    'Used for the transformation of the final object.
    'Create an empty matrix to pass to the moveAndOrientate function
    Dim tmxMatrix As IngrGeom3D.IJDT4x4
    Set tmxMatrix = New DT4x4
    Dim dblRotRadian As Double
    dblRotRadian = degreeToRadian(dblDegreeOfRotation)
    'End Declaration
    
    'Fill out the transformation matrix to orientate the final object
    'Call moveAndOrientate(vecMajorAxis, centerVect, vDirection, tMatrix, dblRotRadian)
    tmxMatrix.LoadIdentity 'Initialize the matrix
    If Not (centerVect Is Nothing) Then
        tmxMatrix.Translate centerVect  'Move the object to the new location
    End If
    vecMajorAxis.Set 1, 0, 0
    Call addDirVectorComponentRotationsToMatrix(tmxMatrix, vDirection, vecMajorAxis, dblRotRadian) 'Get rotations
           
    'Get Sloped bottom cylinder
    Set CreateSlopeBottomCylinder = PlaceSlopedBottomCylinder(objOutputColl, startPos, xVect, endPos, yVect, dblDiameter, blnIsCapped, tmxMatrix)
      
    Set centerVect = Nothing
    Set vDirection = Nothing
    Set vecMajorAxis = Nothing
    Set xVect = Nothing
    Set yVect = Nothing
    Set startPos = Nothing
    Set endPos = Nothing
    
    Exit Function
            
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateSlopeBottomPrism
'   Author:     HL
'   Inputs:
'               Output collection object
'               Origin
'               Orientation
'               X Direction
'               Y Direction
'               Length
'               Vertices in X and Y Coordinates
'               Optional blnIsCapped
'
'   Outputs:
'               Prism based on input
'
'   Description:
'               This function outputs a Sloped Bottom Prism.  It uses function
'               createbaseprism
'
'   Example of Call:
'               Dim objPrism as Object
'               Set objPrism = CreateSlopeBottomPrism(m_outputColl, "E 0 N 0 U 0", oriOrientation, "NE 80 U 70", "W 270 D -90", 1, "0,-1,2,0,1,1,-1,1,-2,0")
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPrism
'               Set objPrism = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'   13.Jan.2003     HL    Created this function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateSlopeBottomPrism(ByVal objOutputColl As Object, _
                            ByVal strOrigin As String, _
                            ByRef oriOrientation As Orientation, _
                            ByVal strXDirection As String, _
                            ByVal strYDirection As String, _
                            ByVal dblLength As Double, _
                            ByVal strVertices As String, _
                            Optional ByVal blnIsCapped As Boolean = True) _
                            As IngrGeom3D.RuledSurface3d

    Set CreateSlopeBottomPrism = CreateBasePrism(objOutputColl, strOrigin, oriOrientation, dblLength, strVertices, strXDirection, strYDirection, , , , , blnIsCapped)
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateSlopedPrism
'   Author:     HL
'   Inputs:
'               Output collection object
'               Origin
'               Orientation
'               X Offset
'               Y Offset
'               X Scaling Factor
'               Y Scaling Factor
'               Length
'               Vertices in X and Y Coordinates
'               Optional blnIsCapped
'
'   Outputs:
'               Prism based on input
'
'   Description:
'               This function outputs a Sloped Bottom Prism.  It uses function
'               createbaseprism
'
'   Example of Call:
'               Dim objPrism as Object
'               Dim oriOrientation As Orientation
'               Set oriOrientation = New Orientation
'               oriOrientation.ResetDefaultAxis
'               Set objPrism = CreateSlopedPrism(m_outputColl, "E 0 N 0 U 0", oriOrientation, 0, 0, 0.5, 0.5, 1, "0,-1,2,0,1,1,-1,1,-2,0")
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPrism
'               Set objPrism = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'   13.Jan.2003     HL    Created this function
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateSlopedPrism(ByVal objOutputColl As Object, _
                            ByVal strOrigin As String, _
                            ByRef oriOrientation As Orientation, _
                            ByVal dblXOffset As Double, _
                            ByVal dblYOffset As Double, _
                            ByVal dblXScalingFactor As Double, _
                            ByVal dblYScalingFactor As Double, _
                            ByVal dblLength As Double, _
                            ByVal strVertices As String, _
                            Optional ByVal blnIsCapped As Boolean = True) As IngrGeom3D.RuledSurface3d

    Set CreateSlopedPrism = CreateBasePrism(objOutputColl, strOrigin, oriOrientation, dblLength, strVertices, "", "", dblXOffset, dblYOffset, dblXScalingFactor, dblYScalingFactor, blnIsCapped)
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateSnout
'   Author:     Doug Hempel
'   Inputs:
'               objOutputColl         => The collection object
'               strOrigin             => The coordinate location of snout origin expressed
'                                        as a string. Snout origin is in center of snout.
'               oriSnoutOrientation   => An orientation object used to effect the orientation
'                                        of the snout.
'               dblZOffset            => The amount of offset of the center of the top face from
'                                        the center of the bottom face in the Z direction when
'                                        the snout is in the standard orientation.
'               dblTopDiameter        => The diameter of the right most face (top) when the snout
'                                        is in standard orientation.
'               dblBottomDiameter     => The diameter of the left most face (bottom) when the snout
'                                        is in standard orientation.
'               dblLength             => The length of the snout along the X axis when the snout
'                                        is in standard orientation.
'               blnIsCapped           => boolean indicating whether to cap (close) the ends
'                                        the snout.
'
'   Outputs:
'               Object - Returns a generic object that is a snout
'
'   Description:
'               Returns a snout as a generic object.  The snout can be given a
'               position and orientation in 3d space.
'
'   Example of Call:
'               Dim objSnout as Object
'               Dim oriPartOri As Orientation
'               Set oriPartOri = New Orientation
'               oriPartOri.ResetDefaultAxis      'Reset to No Rotations
'               oriPartOri.RotationAboutZ = -90  'Set the amount of rotation about Z Axis to -90 degrees
'               oriPartOri.RotationAboutX = 55   'Set the amount of rotation about X Axis to 55 degrees
'               oriPartOri.ApplyRotations        'Rotates the oriention object by the indicated rotation amounts
'               Set objSnout = CreateSnout(m_OutputColl, "E 0 N 0 U 0", oriPartOri, 0.2, 0.5, 1.0, 1.5, True)
'               m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSnout
'               Set objSnout = Nothing
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     Doug Hempel     Created
'   13.Feb.2003     HL        Added checking of strOrigin and oriSnoutOrientation
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateSnout _
( _
   ByVal objOutputColl As Object, _
   ByVal strOrigin As String, _
   ByVal oriSnoutOrientation As Orientation, _
   ByVal dblZOffset As Double, _
   ByVal dblTopDiameter As Double, _
   ByVal dblBottomDiameter As Double, _
   ByVal dblLength As Double, _
   Optional ByVal blnIsCapped As Boolean = True _
) As IngrGeom3D.RuledSurface3d

   Dim xform           As IngrGeom3D.IJDT4x4
   Dim vecTmp          As IJDVector
   Dim vecDir          As IJDVector
   Dim vecZDir         As IJDVector
   Dim posBuildOrigin  As IJDPosition
   Dim posBottomCenter As IJDPosition
   Dim posTopCenter    As IJDPosition
      
   Set vecTmp = New DVector
   Set vecDir = New DVector
   Set vecZDir = New DVector
   Set xform = New DT4x4
   Set posBuildOrigin = New DPosition
   Set posBottomCenter = New DPosition
   Set posTopCenter = New DPosition
   
   
   Dim PI As Double
   
   PI = 4 * Atn(1)


'  --- build the snout at position 0,0,0 ---
'   posBuildOrigin.Set 0, 0, 0

'  --- handle z offset ---
   If dblZOffset <> 0 Then   'z has positive or negative offset
      posBottomCenter.x = -(dblLength / 2)
      posBottomCenter.y = 0
      posBottomCenter.z = -dblZOffset / 2
   
      posTopCenter.x = dblLength / 2
      posTopCenter.y = 0
      posTopCenter.z = dblZOffset / 2
 
   Else    'z has no offset
      posBottomCenter.x = -(dblLength / 2)
      posBottomCenter.y = 0
      posBottomCenter.z = 0
   
      posTopCenter.x = dblLength / 2
      posTopCenter.y = 0
      posTopCenter.z = 0
   
   End If
   
'  --- initialize the transform matrix to identity ---
   xform.LoadIdentity

   If Len(strOrigin) > 0 Then
        '  --- load the transform matrix with a translate from origin to placement position ---
           Set vecTmp = convertPositionStringToDVector(strOrigin)
           xform.Translate vecTmp
   End If
   
'  --- put ori values into xform matrix ---
   If Not (oriSnoutOrientation Is Nothing) Then
        Call loadOriIntoTransformationMatrix(xform, oriSnoutOrientation)
   End If
   
   Set CreateSnout = placeSnout(objOutputColl, posBottomCenter, posTopCenter, dblBottomDiameter, dblTopDiameter, True, xform)
   
'  --- cleanup ---
   Set vecTmp = Nothing
   Set vecDir = Nothing
   Set vecZDir = Nothing
   Set xform = Nothing
   Set posBuildOrigin = Nothing

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   CreateSphere
'   Author:     HL
'   Inputs:
'               Output collection object
'               Origin
'               Diameter
'
'   Outputs:
'               Sphere
'
'   Description:
'               This function outputs a Sphere.
'
'   Example of Call:
'               Dim objSphere as Object
'               Set objSphere = CreateSphere(m_OutputColl, "E 0 N 0 U 0", 1.5)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   23.Jan.2003     HL        Created this function
'   13.Feb.2003     HL        Added checking of existence of strOrigin
'   13.Mar.2003     CYW  Added comments for Example of call
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function CreateSphere(ByVal objOutputColl As Object, _
                            ByVal strOrigin As String, _
                            ByVal dblDiameter As Double) As IngrGeom3D.Sphere3d

    Dim tmxForm As IJDT4x4
    Dim vecPosition As IJDVector
    Dim vecDir As IJDVector

    Set tmxForm = New DT4x4
    Set vecPosition = New DVector
    Set vecDir = New DVector

    'initialize the transformation matrix to identify
    tmxForm.LoadIdentity

    If Len(strOrigin) > 0 Then
        'load the transformation matrix with a translate from origin to placement position
        Set vecPosition = convertPositionStringToDVector(strOrigin)
        tmxForm.Translate vecPosition
    End If
    
    Set CreateSphere = placeTransformableSphere(objOutputColl, dblDiameter / 2, tmxForm)
    
    Set tmxForm = Nothing
    Set vecPosition = Nothing
    Set vecDir = Nothing
End Function

