Attribute VB_Name = "basUtility3D"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   FileName:       basUtility3D.bas
'   ProgID:         basUtility3D
'   Author:         3D Config Team
'   Compiler:       HL
'   Creation Date:  Thursday, Jan 23 2003
'   Description:    This module contains helper functions:
'
'       Math:       ACos
'                   ASin
'                   radianToDegree
'                   degreeToRadian
'
'       Convert:    convertDirectionToUnitVector
'                   convertDirToVectorNE (may not needed)
'                   convertOriStrToVectors
'                   convertPositionStringToDPos
'                   convertPositionStringToDVector
'
'       Formula:    getAngleBetween2VectorsIn3D
'                   getNormalDirVectorFrom3PlanePnts
'                   getPerpendicularVectorIn2D
'                   getUnitDirVectorfrom2Points
'                   getVectorFromPointWithAngle
'                   setXYDirFromAngle
'                   getNewPosition
'                   getTransformedDPos
'                   getOriFromXYDir
'
'       Matrix:     addDirVectorComponentRotationsToMatrix
'                   loadOriIntoTransformationMatrix
'
'       Other:      logError
'
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'   23.Jan.2003     HL    Collected code from 3D Config Team into this module
'   28.Jan.2003     DH    Modified Acos function to handle input of 0
'   28.Jan.2003     HL    Added Function getAngleBetween2VectorsIn3D
'   29.Jan.2003     HL    Added Conditional Compilation to have debug info
'   30.Jan.2003     HL    Added Function getVectorFromPointWithAngle
'   05.Feb.2003     HL    Added Function getPerpendicularVectorIn2D
'   05.Feb.2003     HL    Added Function setXYDirFromAngle
'   13.Feb.2003     HL    Added Function getNewPosition
'   19.Feb.2003     HL    Added Function getNewPositionByTransformation
'   24.Feb.2003     HL    Added Function GetOriFromXYDir
'   24.Feb.2003     JG  Modified Function convertDirectionToUnitVector
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is Updated.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Const MODULE = "basUtility3D"
Const LOGNAME = "c:\dow\error.log"
'Set debug mode to true (-1), or false (0)
#Const INDEBUG = 0

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Sub:        addDirVectorComponentRotationsToMatrix
'   Author:     DH
'   Inputs:
'               tmxTMatrix                   <=> transformation matrix that receives
'                                                discerned rotation angle values.
'               vecDir                        => direction vector, point the matrix this way.
'               vecPrimaryBuildAxisIndicator  => a vector that indicates which is the
'                                                primary build axis by a 1 on the chosen
'                                                axis, 0's for other two axis.
'               dblSpinAngleAboutBuildAxis    => a spin angle in radians about the
'                                                primary build axis. 0 for no spin.
'
'   Outputs:
'               This subroutine does not return anything.  However, by passing
'               tmxTMatrix by reference, it is changed to have new values.
'
'   Description:
'               Discerns rotations from the given vector and installs the
'               the correct rotation angle values into the given transformation
'               matrix.  The installed rotations will point the primary
'               build axis (via the xform) to the direction of the input vector.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------        ---             ------------------
'   01.Jan.2003     DH             Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub addDirVectorComponentRotationsToMatrix _
( _
  ByRef tmxTMatrix As IngrGeom3D.IJDT4x4, _
  vecDir As IngrGeom3D.IJDVector, _
  vecPrimaryBuildAxisIndicator As IngrGeom3D.IJDVector, _
  ByVal dblSpinAngleAboutBuildAxis As Double _
)

   Dim vecAxisSelector As IJDVector
   Dim dblAboutXRotAngle As Double
   Dim dblAboutYRotAngle As Double
   Dim dblAboutZRotAngle As Double


'  --- allocate a new vector to use as an axis selector ---
   Set vecAxisSelector = New DVector

'  --- make sure given vector is a unit vector ---
   vecDir.Length = 1#

'  --- develope specific rotations to point direction vector ---
'  --- handle when primary build axis is X ---
   If (vecPrimaryBuildAxisIndicator.x <> 0 And vecPrimaryBuildAxisIndicator.y = 0 And vecPrimaryBuildAxisIndicator.z = 0) Then
'     --- get rotation angle about X axis in yz plane ---
      dblAboutXRotAngle = dblSpinAngleAboutBuildAxis
'     MsgBox "X Rot Angle:" & " " & Str(DegofRad(dblAboutXRotAngle))

'     --- compute rotation angle about Y axis in xz plane ---
      dblAboutYRotAngle = -(ASin(vecDir.z / vecDir.Length))
'     MsgBox "Y Rot Angle:" & " " & Str(DegofRad(dblAboutYRotAngle))

'     --- compute rotation angle about Z axis in xy plane ---
'     --- handle division by zero ---
      If (vecDir.x = 0 And vecDir.y = 0) Then
         dblAboutZRotAngle = 0
      ElseIf (vecDir.x = 0 And vecDir.y > 0) Then
         dblAboutZRotAngle = degreeToRadian(90)
      ElseIf (vecDir.x = 0 And vecDir.y < 0) Then
         dblAboutZRotAngle = degreeToRadian(270)
      Else  '--- handle quadrants ---
         If (vecDir.x >= 0 And vecDir.y >= 0) Then 'Q1
            dblAboutZRotAngle = Atn(vecDir.y / vecDir.x)
         ElseIf (vecDir.x < 0) Then 'Q2 & Q3
            dblAboutZRotAngle = degreeToRadian(180) + Atn(vecDir.y / vecDir.x)
         ElseIf (vecDir.x >= 0 And vecDir.y < 0) Then 'Q4
            dblAboutZRotAngle = degreeToRadian(360) + Atn(vecDir.y / vecDir.x)
         End If
      End If
'     MsgBox "Z Rot Angle:" & " " & Str(DegofRad(dblAboutZRotAngle))

'     --- concatenate the rotations onto the given transformation matrix in the required order ---
      vecAxisSelector.Set 0, 0, 1
      tmxTMatrix.Rotate dblAboutZRotAngle, vecAxisSelector

      vecAxisSelector.Set 0, 1, 0
      tmxTMatrix.Rotate dblAboutYRotAngle, vecAxisSelector

      vecAxisSelector.Set 1, 0, 0
      tmxTMatrix.Rotate dblAboutXRotAngle, vecAxisSelector



'  --- handle when Y is primary build axis ---
   ElseIf (vecPrimaryBuildAxisIndicator.x = 0 And vecPrimaryBuildAxisIndicator.y <> 0 And vecPrimaryBuildAxisIndicator.z = 0) Then

'     might need some code here in the future,
'     so far configuration team has no need. 01/22/03






'  --- handle when Z is primary build axis ---
   ElseIf (vecPrimaryBuildAxisIndicator.x = 0 And vecPrimaryBuildAxisIndicator.y = 0 And vecPrimaryBuildAxisIndicator.z <> 0) Then

'     might need some code here in the future,
'     so far configuration team has no need. 01/22/03





   Else
'      MsgBox "Programmer Error"
   End If


'  --- release objects ---
   Set vecAxisSelector = Nothing

End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Sub:        loadOriIntoTransformationMatrix
'   Author:     DH
'   Inputs:
'               tmxTMatrix       <=> transformation matrix that receives
'                                    orientation.
'               oriOrientation    => orientation object that defines the
'                                    orientation for the matrix.
'
'   Outputs:
'               This subroutine does not return anything.  However, by passing
'               tmxTMatrix by reference, it is changed to have new values.
'
'   Description:
'               Loads orientation values into the given transformation matrix.
'               The orientation is passed into the subroutine via the orientation
'               object.
'
'   Sub:        loadOriIntoTransformationMatrix
'   Author:     DH
'   Inputs:
'               tmxTMatrix       <=> transformation matrix that receives
'                                    orientation.
'               oriOrientation    => orientation object that defines the
'                                    orientation for the matrix.
'
'   Outputs:
'               This subroutine does not return anything.  However, by passing
'               tmxTMatrix by reference, it is changed to have new values.
'
'   Description:
'               Loads orientation values into the given transformation matrix.
'               The orientation is passed into the subroutine via the orientation
'               object.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     DH                  Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub loadOriIntoTransformationMatrix _
( _
  ByRef tmxTMatrix As IngrGeom3D.IJDT4x4, _
  ByRef oriOrientation As Orientation _
)

   tmxTMatrix.IndexValue(0) = oriOrientation.XAxis.x
   tmxTMatrix.IndexValue(1) = oriOrientation.XAxis.y
   tmxTMatrix.IndexValue(2) = oriOrientation.XAxis.z

   tmxTMatrix.IndexValue(4) = oriOrientation.YAxis.x
   tmxTMatrix.IndexValue(5) = oriOrientation.YAxis.y
   tmxTMatrix.IndexValue(6) = oriOrientation.YAxis.z

   tmxTMatrix.IndexValue(8) = oriOrientation.ZAxis.x
   tmxTMatrix.IndexValue(9) = oriOrientation.ZAxis.y
   tmxTMatrix.IndexValue(10) = oriOrientation.ZAxis.z

End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   getUnitDirVectorfrom2Points
'   Author:     DH
'   Inputs:
'               posFromPosition  => the from position
'               posToPosition    => the to position
'
'   Outputs:
'               IJDVector - A unit vector pointing from -> to
'
'   Description:
'               Computes a unit vector pointing in the direction of
'               from position -> to position.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     DH     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function getUnitDirVectorfrom2Points _
( _
   posFromPosition As IJDPosition, _
   posToPosition As IJDPosition _
) As IJDVector

   Dim vecUnitDirectionVector As IJDVector
   Set vecUnitDirectionVector = New DVector

'  --- move vector between the coordinates to origin 0,0,0 ---
   vecUnitDirectionVector.Set posToPosition.x - posFromPosition.x, posToPosition.y - posFromPosition.y, posToPosition.z - posFromPosition.z

'  --- reset vector length to be 1.0 to create unit direction vector ---
   vecUnitDirectionVector.Length = 1#

'  --- set function return value ---
   Set getUnitDirVectorfrom2Points = vecUnitDirectionVector
   
'  --- clean up ---
   Set vecUnitDirectionVector = Nothing
   
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   getNormalDirVectorFrom3PlanePnts
'   Author:     DH
'   Inputs:
'               posPnt0   => position point 0 on the plane
'               posPnt1   => position point 1 on the plane
'               posPnt2   => position point 2 on the plane
'
'   Outputs:
'               IJDVector - A unit vector perpendicular to plane containing the
'               three given position points.
'
'   Description:
'               Computes a unit vector pointing perpendicular to the plane
'               containing the three position points. CW or CCW point order
'               changes returned vector from one side of plane to the other.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     DH                   Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function getNormalDirVectorFrom3PlanePnts _
( _
   posPnt0 As IJDPosition, _
   posPnt1 As IJDPosition, _
   posPnt2 As IJDPosition _
) As IJDVector


   Dim vecA As IJDVector  ' vector from P0 to P1
   Dim vecC As IJDVector  ' vector from P0 to P2
   Dim vecD As IJDVector  ' vector C cross A

   Set vecA = New DVector
   Set vecC = New DVector
   Set vecD = New DVector
   
'  --- move vector tails to 0,0,0 ---
   vecA.Set posPnt1.x - posPnt0.x, posPnt1.y - posPnt0.y, posPnt1.z - posPnt0.z
   vecC.Set posPnt2.x - posPnt0.x, posPnt2.y - posPnt0.y, posPnt2.z - posPnt0.z

'  --- compute vector cross product to get a normal vector to plane ---
   Set vecD = vecC.Cross(vecA)
   
'  --- convert normal vector to a unit direction vector ---
   vecD.Length = 1#

'  --- set function value ---
   Set getNormalDirVectorFrom3PlanePnts = vecD

'  --- clean up ---
   Set vecA = Nothing
   Set vecC = Nothing
   Set vecD = Nothing
   
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   aSin
'   Author:     DH
'   Inputs:
'               dblInputValue  => value for which arcsine is computed.
'
'   Outputs:
'               Double - The arcsine of the input value. It is an angle in radians.
'
'   Description:
'               Computes the arcsine of the input value. It will handle properly
'               the input values of 0, 1 or -1.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------         ---             ------------------
'   01.Jan.2003        DH              Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function ASin(ByVal dblInputValue As Double) As Double
  
    Select Case dblInputValue
        Case 0
            ASin = 0
        Case 1
            ASin = degreeToRadian(90)
        Case -1
            ASin = degreeToRadian(-90)
        Case Else
            ASin = Atn(dblInputValue / ((-dblInputValue * dblInputValue + 1) ^ 0.5))
    End Select

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   aCos
'   Author:     DH
'   Inputs:
'               dblInputValue  => value for which arccosine is computed.
'
'   Outputs:
'               Double - The arccosine of the input value. It is an angle in radians.
'
'   Description:
'               Computes the arccosine of the input value.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     DH    Created
'   28.Jan.2003     DH     Revised the function to handle input of 0 correctly.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function ACos(ByVal dblInputValue As Double) As Double
  
    Select Case dblInputValue
        Case 0
            ACos = degreeToRadian(90)
        Case -1
            ACos = degreeToRadian(180)
        Case 1
            ACos = 0
        Case Else
            ACos = Atn(-dblInputValue / ((-dblInputValue * dblInputValue + 1) ^ 0.5)) + 2 * Atn(1)
    End Select

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   radianToDegree
'   Author:     DH
'   Inputs:
'               Radians   => Input value in radians.
'
'   Outputs:
'               Double - The equivalent degree value of the radian input value.
'
'   Description:
'               Converts the given radian value into degrees.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   01.Jan.2003     DH     Created
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function radianToDegree(ByVal Radians As Double) As Double
   Dim PI As Double
   
   PI = 4 * Atn(1)
   radianToDegree = (180 / PI) * Radians
   
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Sub:        logError
'   Author:     HL
'   Inputs:
'               Error Message
'
'   Outputs:
'               None
'
'   Description:
'               This subroutine puts error message in a file defined in
'               LOGNAME constant.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   20.Nov.2002     HL        Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub logError(strMsg)
    Dim FileNumber As Integer
    
    FileNumber = FreeFile()
    
    Open LOGNAME For Append As #FileNumber
    Print #FileNumber, strMsg
    Close #FileNumber
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Function:   convertDirToVectorNE
'   Author:     HL
'   Inputs:
'               strDir => Direction string
'
'   Outputs:
'               IJDVector
'
'   Description:
'               This function converts a direction string to a vector.  It is
'               being used only in Truncated N Edge Prism.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   19.Dec.2002     HL        Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function convertDirToVectorNE(ByVal strDir As String) As IJDVector
    Dim arrStr() As String
    Dim i As Integer
    Dim iTotal As Integer
    Dim theItem As String
    Dim alpha As Double
    Dim beta As Double
    Dim vecDir As IJDVector
    
    Set vecDir = New DVector
    
    If Len(strDir) = 0 Then
        Exit Function
    End If

    arrStr = Split(UCase(strDir), " ")
    iTotal = UBound(arrStr)

    theItem = arrStr(0)
    If iTotal = 0 Then
        Select Case theItem
            Case "U"
                vecDir.Set 0, 0, 1
            Case "D"
                vecDir.Set 0, 0, -1
            Case Else
'                MsgBox "Direction must start either U or D"
                Exit Function
        End Select
        Set convertDirToVectorNE = vecDir
        Exit Function
    End If
    
    If ((iTotal + 1) Mod 2) <> 1 Then
        Exit Function
    End If
    
    For i = 2 To iTotal Step 2
        theItem = arrStr(i)
        Select Case theItem
            Case "E"
                alpha = CDbl(arrStr(i - 1))
            Case "N"
                beta = CDbl(arrStr(i - 1))
            Case "W"
                alpha = -CDbl(arrStr(i - 1))
            Case "S"
                beta = -CDbl(arrStr(i - 1))
        End Select
    Next i
    alpha = degreeToRadian(alpha)
    beta = degreeToRadian(beta)

    vecDir.Set Sin(alpha) * Cos(beta), Sin(beta), Cos(alpha) * Cos(beta)
    Set convertDirToVectorNE = vecDir
    'MsgBox vecDir.x & " " & vecDir.y & " " & vecDir.z
    
    Set vecDir = Nothing
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   degreeToRadian
'   Author:     JG
'   Inputs:
'               Degree
'
'   Outputs:
'               Radian
'
'   Description:
'               This function converts degrees into radians
'
'   Example of call:
'               Dim dblDegree As Double
'               Dim dblRadian As Double
'               dblRadian = degreeToRadian(dblDegree)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   20.Dec.2002     JG      Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function degreeToRadian(degree As Double) As Double
    Dim PI As Double
    
    PI = 4 * Atn(1)
    degreeToRadian = (degree * PI) / 180
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   convertPositionStringToDVector
'   Author:     JG
'   Inputs:
'               string containing position vector
'
'   Outputs:
'               DVection
'
'   Description:
'               This function converts a string containing a real world
'               vector to a DVector
'
'   Example of call:
'
'               Dim newVector as IJDVector
'               Set newVector = new DVector
'               Dim Vector   As String
'               Vector = "N 5 W 5 D 2"
'               set newVector = convertStringToDVector(Vector)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   20.Dec.2002     JG      Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function convertPositionStringToDVector(strVect As String) As AutoMath.DVector
    Dim arrVectString() As String
    Dim tempVector As IJDVector 'Used to create the position to be passed back.
    Set tempVector = New DVector
    Dim intCounter As Integer
    intCounter = 0
    
    'Split up the string
    strVect = UCase(strVect)
    arrVectString = Split(strVect, " ", , vbTextCompare)
    
    'Convert string directions to +- values
    Do While (intCounter <= UBound(arrVectString))
        If (arrVectString(intCounter) = "E") Then
            tempVector.x = CDbl(arrVectString(intCounter + 1))
        ElseIf (arrVectString(intCounter) = "W") Then
            tempVector.x = -1 * CDbl(arrVectString(intCounter + 1))
        ElseIf (arrVectString(intCounter) = "N") Then
            tempVector.y = CDbl(arrVectString(intCounter + 1))
        ElseIf (arrVectString(intCounter) = "S") Then
            tempVector.y = -1 * CDbl(arrVectString(intCounter + 1))
        ElseIf (arrVectString(intCounter) = "U") Then
            tempVector.z = CDbl(arrVectString(intCounter + 1))
        ElseIf (arrVectString(intCounter) = "D") Then
            tempVector.z = -1 * CDbl(arrVectString(intCounter + 1))
        End If
        intCounter = intCounter + 2  'Move to the next direction indicator.
    Loop
                        
    'Pass back object.
    Set convertPositionStringToDVector = tempVector
    
    Set tempVector = Nothing
    
    Exit Function
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   convertPositionStringToDPos
'   Author:     JG
'   Inputs:
'               position string
'
'   Outputs:
'               DPosition
'
'   Description:
'               This function converts a string containing a real world
'               position to a DPosition
'
'   Example of call:
'
'               Dim newPosition as IJDPosition
'               Set newPosition = new DPosition
'               Dim position   As String
'               position = "N 5 W 5 D 2"
'               set newPosition = convertStringToDPos(position)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   20.Dec.2002     JG      Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function convertPositionStringToDPos(strPos As String) As AutoMath.DPosition
             
    Dim arrPosString() As String
    Dim tempPosition As IJDPosition 'Used to create the position to be passed back.
    Set tempPosition = New DPosition
    Dim intCounter As Integer
    intCounter = 0
    
    'Split up the string
    strPos = UCase(strPos)
    arrPosString = Split(strPos, " ", , vbTextCompare)
    
    'Convert string directions to +- values
    Do While (intCounter <= UBound(arrPosString))
         If (arrPosString(intCounter) = "E") Then
            tempPosition.x = CDbl(arrPosString(intCounter + 1))
        ElseIf (arrPosString(intCounter) = "W") Then
            tempPosition.x = -1 * CDbl(arrPosString(intCounter + 1))
        ElseIf (arrPosString(intCounter) = "N") Then
            tempPosition.y = CDbl(arrPosString(intCounter + 1))
        ElseIf (arrPosString(intCounter) = "S") Then
            tempPosition.y = -1 * CDbl(arrPosString(intCounter + 1))
        ElseIf (arrPosString(intCounter) = "U") Then
            tempPosition.z = CDbl(arrPosString(intCounter + 1))
        ElseIf (arrPosString(intCounter) = "D") Then
            tempPosition.z = -1 * CDbl(arrPosString(intCounter + 1))
        End If
        intCounter = intCounter + 2  'Move to the next direction indicator.
    Loop
                        
    'Pass back object.
    Set convertPositionStringToDPos = tempPosition
    
    Set tempPosition = Nothing
        
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   convertOriStrToVectors
'   Author:     JG
'   Inputs:
'               Orientation string
'
'   Outputs:
'               Orientation object
'
'   Description:
'               This function converts a string containing a real
'               world orientation and converts it to a type
'               containing 3 unit vectors
'
'   Example of call:
'
'               Dim typOri as OrientationType
'               set typOri = convertOriStrToVectors("X IS E AND Y IS N")
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   20.Dec.2002     JG      Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function convertOriStrToVectors(strOri As String) As Orientation
    
    Dim strOriArray() As String
    Dim strFirstOri() As String
    Dim strSecondOri() As String
    Const intOriginalDir = 0
    Const intNewDir = 1
    Dim vecTemp As AutoMath.DVector
    Set vecTemp = New DVector
   
    Dim myOri As Orientation
    Set myOri = New Orientation
    
    myOri.Tolerance = 0.001
 
    strOri = UCase(strOri)
    strOriArray = Split(strOri, " AND ", , vbTextCompare)
    strFirstOri = Split(strOriArray(0), " IS ", , vbTextCompare)
    strSecondOri = Split(strOriArray(1), " IS ", , vbTextCompare)

    'Determine the first original direction and set the new direction.
    If (strFirstOri(intOriginalDir) = "X") Then
        Set vecTemp = convertDirectionToUnitVector(strFirstOri(intNewDir))
        Set myOri.XAxis = vecTemp
    ElseIf (strFirstOri(intOriginalDir) = "Y") Then
        Set vecTemp = convertDirectionToUnitVector(strFirstOri(intNewDir))
        Set myOri.YAxis = vecTemp
    ElseIf (strFirstOri(intOriginalDir) = "Z") Then
        Set vecTemp = convertDirectionToUnitVector(strFirstOri(intNewDir))
        Set myOri.ZAxis = vecTemp
    End If
    
    'Determine the second original direction and set the new direction.
    If (strSecondOri(intOriginalDir) = "X") Then
        Set vecTemp = convertDirectionToUnitVector(strSecondOri(intNewDir))
        Set myOri.XAxis = vecTemp
    ElseIf (strSecondOri(intOriginalDir) = "Y") Then
        Set vecTemp = convertDirectionToUnitVector(strSecondOri(intNewDir))
        Set myOri.YAxis = vecTemp
    ElseIf (strSecondOri(intOriginalDir) = "Z") Then
        Set vecTemp = convertDirectionToUnitVector(strSecondOri(intNewDir))
        Set myOri.ZAxis = vecTemp
    End If
    
    'Determine the third original direction and find the new direction
    'which would be the perpendicular to the other two.
    myOri.FindRemainingUnsetAxis
    
    Set convertOriStrToVectors = myOri
 
    Set myOri = Nothing
    Set vecTemp = Nothing
        
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   convertDirectionToUnitVector
'   Author:     JG
'   Inputs:
'               Orientation string
'
'   Outputs:
'               IJDVector
'
'   Description:
'               This function converts a string containing a real
'               world polar direction and converts it to a unit vector
'
'   Example of call:
'
'               Dim dirVector as IJDVector
'               Set dirVector = New DVector
'               Dim direction As String
'               direction = "NE 10 U 20"
'               Set dirVector = convertDirectionToUnitVector(direction)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   14.Feb.2003     JG      Created this function.
'
'   17.Feb.2003     JG      Added error checking to the function.
'                                   Checks to see if the angle and the quadrent
'                                   passed are equal.
'
'   24.Feb.2003     JG      Renamed function to overwrite original
'                                   function so that others may use this
'                                   methode of defining headings versus the
'                                   previous.
'
'   25.Feb.2003     JG      Fixed some error checing bugs and added
'                                   more descriptive debugging statements.
'
'   26.Mar.2003     HL        Added error checking for passing a string of elements\
'                                   not equal to 4
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function convertDirectionToUnitVector(direction As String) As IJDVector
    
    'Split the string up into an usable array
    Dim dblAngle As Double
    Dim dblElevation As Double
    Dim dblXPnt As Double
    Dim dblYPnt As Double
    Dim dblZPnt As Double
    Dim dblXYHyp As Double
    Dim vecReturn As IJDVector
    Set vecReturn = New DVector
    vecReturn.Set 0, 1, 0
    
    'Helper vars
    Dim intWholeNumber As Integer
    Dim strErrMsg As String
    
    'Parse the string
    Dim dirString() As String
    direction = Trim(UCase(direction))
    dirString = Split(direction, " ", , vbTextCompare)
        
    Dim bolZeroY As Boolean
    bolZeroY = False
    Dim bolZeroX As Boolean
    bolZeroX = False
    
    'Check the string
    Dim bolValid As Boolean
    bolValid = True
    
    If UBound(dirString) <> 3 Then
        strErrMsg = "Invalid direction string: " & Chr(34) & direction & Chr(34) & vbCrLf & _
                    "           "
        Err.Raise vbObjectError + 56000, "convertDirectionToUnitVector", strErrMsg
        Exit Function
    End If
    
    'Normalize the angle
    'If it is greater then 360 degrees then remove any extra
    If ((Val(dirString(1))) >= 360) Then
        intWholeNumber = (Val(dirString(1)) / 360)
        dirString(1) = Val(dirString(1)) - (360 * intWholeNumber)
    End If
    
    'Check to make sure that the angle falls into the proper quadrent.
    If (dirString(0) <> "N" And dirString(0) <> "S" And dirString(0) <> "E" And dirString(0) <> "W" And _
        dirString(0) <> "NE" And dirString(0) <> "NW" And dirString(0) <> "SE" And dirString(0) <> "SW") Then
        bolValid = False
    ElseIf (dirString(0) = "N") Then
        bolZeroX = True
        If (Val(dirString(1)) <> 0) Then
            bolValid = False
        End If
    ElseIf (dirString(0) = "E") Then
        bolZeroY = True
        If (Val(dirString(1)) <> 90) Then
            bolValid = False
        End If
    ElseIf (dirString(0) = "S") Then
        bolZeroX = True
        If (Val(dirString(1)) <> 180) Then
            bolValid = False
        End If
    ElseIf (dirString(0) = "W") Then
        bolZeroY = True
        If (Val(dirString(1)) <> 270) Then
            bolValid = False
        End If
    ElseIf (dirString(0) = "NE") Then
        If (Not (Val(dirString(1)) > 0) Or Not (Val(dirString(1)) < 90)) Then
            bolValid = False
        End If
    ElseIf (dirString(0) = "SE") Then
        If (Not (Val(dirString(1)) > 90) Or Not (Val(dirString(1)) < 180)) Then
            bolValid = False
        End If
    ElseIf (dirString(0) = "SW") Then
        If (Not (Val(dirString(1)) > 180) Or Not (Val(dirString(1)) < 270)) Then
            bolValid = False
        End If
    ElseIf (dirString(0) = "NW") Then
        If (Not (Val(dirString(1)) > 270) Or Not (Val(dirString(1)) < 360)) Then
            bolValid = False
        End If
    End If
    
    If bolValid = False Then
        Err.Raise vbObjectError + 56001, "convertDirectionToUnitVector", strErrMsg
        Exit Function
    End If
    
    'Check the elevation.
    'Elevation must be between -90 and 90 degrees.
    '-90 = down
    '0   = horizontal
    '90  = up
    If (dirString(2) <> "U" And dirString(2) <> "D") Then
        bolValid = False
    ElseIf (dirString(2) = "U") Then
        If (Val(dirString(3)) < 0 Or Val(dirString(3)) > 90) Then
            bolValid = False
            strErrMsg = strErrMsg & "Invalid Elevation passed " & dirString(2) & " " & dirString(3)
        End If
    ElseIf (dirString(2) = "D") Then
        If (Val(dirString(3)) > 0 Or Val(dirString(3)) < -90) Then
            bolValid = False
            strErrMsg = strErrMsg & "Invalid Elevation passed " & dirString(2) & " " & dirString(3)
        End If
    End If
    
    If bolValid = False Then
        Err.Raise vbObjectError + 56002, "convertDirectionToUnitVector", strErrMsg
        Exit Function
    End If
    
    'Solve for the vector.
    If (bolValid = True) Then
        'Get the angles and convert then to radians
        dblAngle = degreeToRadian(Val(dirString(1)))
        dblElevation = degreeToRadian(Val(dirString(3)))
        
        'Get the z distance and new hypotenous to use later
        'Hypotenous of this triangle is 1
        dblZPnt = Sin(dblElevation)
        dblXYHyp = Cos(dblElevation)
        
        'Find XY points
        'If the vector is not just up or down
        'If it is then skip this section.
        If (Abs(dblZPnt) <> 1) Then
            If bolZeroX = False Then
                dblXPnt = Sin(dblAngle) * dblXYHyp
            End If
            If bolZeroY = False Then
                dblYPnt = Cos(dblAngle) * dblXYHyp
            End If
        End If
        
        vecReturn.Set dblXPnt, dblYPnt, dblZPnt
    End If
    
    Set convertDirectionToUnitVector = vecReturn
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   convertDirectionToUnitVectorPV
'   Author:     JG
'   Inputs:
'               Orientation string
'
'   Outputs:
'               Orientation object
'
'   Description:
'               This function converts a string containing a real
'               world direction and converts it to a unit vector
'
'   Example of call:
'
'               Dim dirVector as IJDVector
'               Set dirVector = New DVector
'               Dim direction As String
'               direction = "N 5 W 5 D 2"
'               Set dirVector = convertDirection(direction)
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   20.Dec.2002     JG      Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function convertDirectionToUnitVectorPV(direction As String) As AutoMath.DVector
        
    'Dim Vars
    Dim unitVect As IJDVector
    Set unitVect = New DVector
    unitVect.Set 0, 0, 0
    Dim dirString() As String
    Dim dblRadian As Double
    Dim dblXYHyp As Double
    dblXYHyp = 1 'Initialize this to 1
    Dim intFirstDir As Integer
    Dim intSecondDir As Integer
    intFirstDir = 0  ''Initialize this to the normal 0 position
    intSecondDir = 2  'Initialize this to the normal 2 position
        
    'Split the string up into an usable array
    direction = UCase(direction)
    dirString = Split(direction, " ", , vbTextCompare)

    'Runs only for 3D deminsions
    'If this is a 3D direction then we need to
    'Determine which posistion contains the elevation
    'and calculate the z dir value and the hypotenus
    'that will be used to calculate the rest of the
    'values
    If (UBound(dirString) = 4) Then
        If (dirString(4) = "U" Or dirString(4) = "D") Then
            intFirstDir = 0
            intSecondDir = 2
            dblRadian = degreeToRadian(Val(dirString(3)))
            dblXYHyp = Cos(dblRadian)
            If (dirString(4) = "D") Then
                unitVect.z = -1 * Sin(dblRadian)
            Else
                unitVect.z = 1 * Sin(dblRadian)
            End If
        ElseIf (dirString(2) = "U" Or dirString(2) = "D") Then
            intFirstDir = 0
            intSecondDir = 4
            dblRadian = degreeToRadian(Val(dirString(1)))
            dblXYHyp = Cos(dblRadian)
            If (dirString(2) = "D") Then
                unitVect.z = -1 * Sin(dblRadian)
            Else
                 unitVect.z = 1 * Sin(dblRadian)
            End If
        ElseIf (dirString(0) = "U" Or dirString(0) = "D") Then
            intFirstDir = 2
            intSecondDir = 4
            dblRadian = degreeToRadian(Val(dirString(1)))
            dblXYHyp = Sin(dblRadian)
             If (dirString(0) = "D") Then
                unitVect.z = -1 * Cos(dblRadian)
            Else
                unitVect.z = 1 * Cos(dblRadian)
            End If
        End If
    End If
    
    'Runs for both 2D and 3D deminsions
    'CASE 3D:  The U and D will not be used due to we allready
    '           found and calculated these valuse
    '           It will calculate the value for the either the
    '           2nd or 4th pos which is in the XY plane.
    '           Of the 2 remaining it is the farthest right.
    '           EX.  U 10 N 10 W  it uses W
    'CASE 2D:   Calculates the value that is the farthest right.
    '           EX.  N 10 E  it uses E
    If (UBound(dirString) > 1) Then
        dblRadian = degreeToRadian(Val(dirString(intSecondDir - 1)))
        If dirString(intSecondDir) = "U" Then
            unitVect.z = Sin(dblRadian) * dblXYHyp
        ElseIf dirString(intSecondDir) = "D" Then
            unitVect.z = -1 * Sin(dblRadian) * dblXYHyp
        ElseIf dirString(intSecondDir) = "N" Then
            unitVect.y = Sin(dblRadian) * dblXYHyp
        ElseIf dirString(intSecondDir) = "S" Then
            unitVect.y = -1 * Sin(dblRadian) * dblXYHyp
        ElseIf dirString(intSecondDir) = "E" Then
            unitVect.x = Sin(dblRadian) * dblXYHyp
        ElseIf dirString(intSecondDir) = "W" Then
            unitVect.x = -1 * Sin(dblRadian) * dblXYHyp
        End If
    End If

    'Runs for all types of directions
    'Get the first position.  If only one direction then cosine = 1
    'For 2D and 3D values it calculates the values farthest left.
    '       Ex.  N 10 E it uses N
    '            N 10 U 5 E it uses N
    If (UBound(dirString) > 0) Then
        dblRadian = degreeToRadian(Val(dirString(intFirstDir + 1)))
    End If
        
    If dirString(intFirstDir) = "E" Then
        unitVect.x = Cos(dblRadian) * dblXYHyp
    ElseIf dirString(intFirstDir) = "W" Then
        unitVect.x = -1 * Cos(dblRadian) * dblXYHyp
    ElseIf dirString(intFirstDir) = "N" Then
        unitVect.y = Cos(dblRadian) * dblXYHyp
    ElseIf dirString(intFirstDir) = "S" Then
        unitVect.y = -1 * Cos(dblRadian) * dblXYHyp
    ElseIf dirString(intFirstDir) = "U" Then
        unitVect.z = Cos(dblRadian) * dblXYHyp
    ElseIf dirString(intFirstDir) = "D" Then
        unitVect.z = -1 * Cos(dblRadian) * dblXYHyp
    End If
        
    'Set for for return
    Set convertDirectionToUnitVectorPV = unitVect
    
    Set unitVect = Nothing
        
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   getAngleBetween2VectorsIn3D
'   Author:     HL
'   Inputs:
'               Cos value of an angle
'
'   Outputs:
'               An angle in radians.
'
'   Description:
'               This function will compute an angle between 2 vectors.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   28.Jan.2003     HL        Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function getAngleBetween2VectorsIn3D(V1 As IJDVector, V2 As IJDVector)
'    Dim cosValue As Double
'    Dim dblLengthV1 As Double
'    Dim dblLengthV2 As Double
'
'    dblLengthV1 = (V1.x * V1.x + V1.y * V1.y + V1.z * V1.z) ^ 0.5
'    dblLengthV2 = (V2.x * V2.x + V2.y * V2.y + V2.z * V2.z) ^ 0.5
'
'    cosValue = (V1.x * V2.x + V1.y * V2.y + V1.z * V2.z) / (dblLengthV1 * dblLengthV2)
    
    getAngleBetween2VectorsIn3D = ACos(V1.Dot(V2))
    
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   getVectorFromPointWithAngle
'   Author:     HL
'   Inputs:
'               Angle - to offset, in radian
'               StartAngle - To measure from in radian
'
'   Outputs:
'               New Position vector based on the inputs
'
'   Description:
'               This function will compute position of a point from an angle from
'               start position's angle on xy plane.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   30.Jan.2003     HL        Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function getVectorFromPointWithAngle(ByVal dblAngle As Double, ByVal dblStartAngle As Double) As IJDVector
    Dim x As Double, y As Double
    Dim vecEnd As IJDVector
    Set vecEnd = New DVector
    
    dblAngle = (dblStartAngle + dblAngle)
    
    'The select statement is used because VB doesn't return 0, instead it is a very small number.
    Select Case radianToDegree(dblAngle)
        Case 90, 270:
            x = 0
        Case Else:
            x = Cos(dblAngle)
    End Select
    
    Select Case radianToDegree(dblAngle)
        Case 0, 180:
            y = 0
        Case Else:
            y = Sin(dblAngle)
    End Select
    vecEnd.Set x, y, 0
    
'    #If INDEBUG Then
'        MsgBox "getVectorFromPointWithAngle: new angle is in degree " & radianToDegree(dblAngle)
'        MsgBox "getVectorFromPointWithAngle: New vector position is " & vecEnd.x & " " & vecEnd.y & " " & vecEnd.z
'    #End If
    
    Set getVectorFromPointWithAngle = vecEnd
    Set vecEnd = Nothing
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Sub:        setXYDirFromAngle
'   Author:     HL
'   Inputs:
'               xVector
'               yVector
'               Angle
'
'   Outputs:
'               None.  However, by passing xVector and yVector by reference,
'               they will be set to have the new values
'
'   Description:
'               This function will compute xdir and ydir based on an angle
'               in x-y plane.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   05.Feb.2003     HL        Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub setXYDirFromAngle(ByRef xVector As IJDVector, ByRef yVector As IJDVector, ByVal dblAngle As Double)
    
    Set xVector = getVectorFromPointWithAngle(dblAngle / 2, degreeToRadian(225))
    Set yVector = getVectorFromPointWithAngle(-dblAngle / 2, degreeToRadian(225))
    
'    #If INDEBUG Then
'        MsgBox "X Dir is from angle " & radianToDegree(dblAngle / 2) & ": " & xVector.x & " " & xVector.y & " " & xVector.z
'        MsgBox "Y Dir is from angle " & radianToDegree(-dblAngle / 2) & ": " & yVector.x & " " & yVector.y & " " & yVector.z
'    #End If
    
    Set xVector = getPerpendicularVectorIn2D(xVector, False)
    Set yVector = getPerpendicularVectorIn2D(yVector, True)
    
'    #If INDEBUG Then
'        MsgBox "Normal X Dir is : " & xVector.x & " " & xVector.y & " " & xVector.z
'        MsgBox "Normal Y Dir is : " & yVector.x & " " & yVector.y & " " & yVector.z
'    #End If
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   getPerpendicularVectorIn2D
'   Author:     HL
'   Inputs:
'               Vector
'
'   Outputs:
'               Vector perpendicular to the given vector in xy plane.
'
'   Description:
'               This function will compute a parpendicular vector to the given
'               vector in x-y plane.
'               One Method:
'                   Suppose input vector is (a,b,0), the new vector is (x,y,0)
'                   ax + by = 0    (from dot product)
'                   x*x + y*y = 1  (A unit vector)
'                   It will give two solutions, one is clock wise, the other
'                   counterclock wise.
'                   x = sqrt(b*b /(a*a+b*b))
'                   y = sqrt(a*a /(a*a+b*b))
'
'               Another Method:
'                   Use Cross Product of two vectors.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   05.Feb.2003     HL        Created this function.
'   06.Feb.2003     HL        Use Second method to determine the signs of
'                                   x and y programmatically.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function getPerpendicularVectorIn2D(vecDir As IJDVector, _
        ByVal blnClockwise As Boolean) As IJDVector
    Dim x As Double, y As Double
    Dim vecNew As IJDVector
    Set vecNew = New DVector
    
    Dim zVec As IJDVector
    Set zVec = New DVector
    
'   Method #1: This will require handle signs of x and y individually.
'    x = vecDir.y * vecDir.y / (vecDir.x * vecDir.x + vecDir.y * vecDir.y)
'    x = x ^ 0.5
'    y = vecDir.x * vecDir.x / (vecDir.x * vecDir.x + vecDir.y * vecDir.y)
'    y = y ^ 0.5
'
'   Method #2: Use Cross Product of two vectors.
    
    zVec.Set 0, 0, 1
    If blnClockwise Then
        Set vecNew = vecDir.Cross(zVec)
    Else
        Set vecNew = zVec.Cross(vecDir)
    End If
    
'    #If INDEBUG Then
'        MsgBox "The new vector is " & vecNew.x & " " & vecNew.y & " " & vecNew.z
'    #End If
    
    Set getPerpendicularVectorIn2D = vecNew
    Set vecNew = Nothing
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   getDPosOnSphereFrom2Vectors
'   Author:     DH and HL
'   Inputs:
'               X Vector
'               Y Vector
'               Radius
'
'   Outputs:
'               Position
'
'   Description:
'               This function will compute the new position on a sphere with radius
'               given two vectors.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   13.Feb.2003     DH and HL       Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function getDPosOnSphereFrom2Vectors(Xdir As IJDVector, Ydir As IJDVector, ByVal dblRadius As Double, ByVal blnClockwise As Boolean) As IJDPosition
    Dim Zdir As IJDVector
    Dim vecDir As IJDVector
    Dim posNew As IJDPosition
    
    Set posNew = New DPosition
    
'    #If INDEBUG Then
'        MsgBox "X vector is " & Xdir.x & " " & Xdir.y & " " & Xdir.z
'        MsgBox "Y vector is " & Ydir.x & " " & Ydir.y & " " & Ydir.z
'    #End If

    Set Zdir = Xdir.Cross(Ydir)
    Zdir.Length = 1
'    #If INDEBUG Then
'        MsgBox "Z vector is " & Zdir.x & " " & Zdir.y & " " & Zdir.z
'    #End If
    
    If blnClockwise Then
        Set vecDir = Zdir.Cross(Ydir)
    Else
        Set vecDir = Xdir.Cross(Zdir)
    End If
    vecDir.Length = dblRadius
'    #If INDEBUG Then
'        MsgBox "New vector is " & vecDir.x & " " & vecDir.y & " " & vecDir.z
'    #End If
    
    posNew.Set vecDir.x, vecDir.y, vecDir.z
    Set getDPosOnSphereFrom2Vectors = posNew
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   getTransformedDPos
'   Author:     DH and HL
'   Inputs:
'               Old Position
'               Transformation Matrix
'
'   Outputs:
'               New Position
'
'   Description:
'               This function will compute the new position based on the transformation
'               matrix.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   19.Feb.2003     DH and HL       Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function getTransformedDPos(posOld As IJDPosition, tmxMatrix As IJDT4x4) As IJDPosition
    Dim posNew As IJDPosition
    Dim arrayPoints(3) As Double
    Dim i As Integer
    
    Set posNew = New DPosition
    For i = 0 To 3
        arrayPoints(i) = tmxMatrix.IndexValue(i) * posOld.x + tmxMatrix.IndexValue(i + 4) * posOld.y + tmxMatrix.IndexValue(i + 8) * posOld.z + tmxMatrix.IndexValue(i + 12) * 1
    Next i
    posNew.Set arrayPoints(0) / arrayPoints(3), arrayPoints(1) / arrayPoints(3), arrayPoints(2) / arrayPoints(3)
'    #If INDEBUG Then
'        MsgBox "New Point is " & posNew.x & " " & posNew.y & " " & posNew.z
'    #End If
    Set getTransformedDPos = posNew
    Set posNew = Nothing
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   getOriFromXYDir
'   Author:     DH and HL
'   Inputs:
'               XDir
'               YDir
'               SweepAngle
'
'   Outputs:
'               Ori.
'               Also, by passing sweep angle by reference, it will also be set.
'
'   Description:
'               This function will compute an ori based on Xdir and Ydir.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'   24.Feb.2003     DH and HL       Created this function.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function getOriFromXYDir(ByRef Xdir As IJDVector, ByRef Ydir As IJDVector, ByRef dblSweepAngle As Double) As Orientation
    Dim oriNew As Orientation
    Dim Zdir As IJDVector
    Dim dblAngle As Double
    Dim vecX As IJDVector
    Dim dblDistance As Double
    Dim x As Double, y As Double, z As Double
    
    Set Zdir = Xdir.Cross(Ydir)
    Zdir.Length = 1
    
'    #If INDEBUG Then
'        MsgBox "GetOriFromXYDir:Z vector is " & Zdir.x & " " & Zdir.y & " " & Zdir.z
'    #End If

    Set oriNew = New Orientation
    'Get angle between the two vectors.
    dblSweepAngle = degreeToRadian(180) - getAngleBetween2VectorsIn3D(Xdir, Ydir)
    
    Set vecX = Xdir.Cross(Zdir)
    vecX.Length = 1
    
    'If less or equal to 90 then
    If dblSweepAngle <= degreeToRadian(90) Then
        dblAngle = degreeToRadian(45) - dblSweepAngle / 2
        dblDistance = vecX.Length * Tan(dblAngle)
    Else
        dblAngle = dblSweepAngle / 2 - degreeToRadian(45)
        dblDistance = -vecX.Length * Tan(dblAngle)
    End If
    x = Xdir.x * dblDistance + vecX.x
    y = Xdir.y * dblDistance + vecX.y
    z = Xdir.z * dblDistance + vecX.z
    
    oriNew.YAxis.Set -x, -y, -z
    oriNew.YAxis.Length = 1
    Set oriNew.ZAxis = Zdir
    Set oriNew.XAxis = oriNew.YAxis.Cross(oriNew.ZAxis)
    oriNew.XAxis.Length = 1
    
    Set getOriFromXYDir = oriNew
    Set oriNew = Nothing
    Set vecX = Nothing
    
End Function
