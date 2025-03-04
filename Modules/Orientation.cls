VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Orientation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   Orientation
'   ProgID:         Orientation.Orientation
'   Author:         JG
'   Creation Date:  Tuesday, Jan 07 2003
'   Description:
'       Gives SP3D VB symbol components the concept of an orientation.
'       You may manually set each vector, set two and have it find the third
'       or rotate the major axis about x,y, and z acquiring your new axis.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   07.Jan.2003     JLG     Created initial routines
'   13.Jan.2003     JLG     Added resetDefaultAxis sub to allow the client to
'                           reset the axis back to the default values.
'   14.Jan.2003     JLG     Changed the use of 3 arrays to store rotations to
'                           useing only one array to store all rotations.  Then
'                           step through the array backwards to apply the
'                           rotations.  Aslo changed Class_Initialize to call
'                           resetDefaultAxis instead of having duplicate code.
'   14.May.2003     JLG     Added ability to set the angle input type to radians.
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

'Dim private vars used in the class
Private mNewX As IJDVector
Private mNewY As IJDVector
Private mNewZ As IJDVector
Private mBoolXSet As Boolean
Private mBoolYSet As Boolean
Private mBoolZSet As Boolean
Private mTolerance As Double
Private mCount As Integer
Private mXYZRotate(0 To 5) As String
Private mBoolRadianInput As Boolean

Private Const MODULE = "Orientation:" 'Used for error messages

Private Sub Class_Initialize()

''' Initializes all paramaters for use in
''' the class.  Uses the resetDefaultAxis
''' sub routine.

    Set mNewX = New DVector
    Set mNewY = New DVector
    Set mNewZ = New DVector
    mTolerance = 0.0000001  'Initialize the tolerance to 7 deci
    mBoolRadianInput = False
    
    Call ResetDefaultAxis 'This sets all values to thier default setting.

End Sub

Public Sub ResetDefaultAxis()

''' This sub resets the axis and
''' all module variables
''' Example of call:
''' myOri.resetDefaultAxis

    mCount = 0
    mBoolXSet = False   'Initialize all to false due to
    mBoolYSet = False   'the user has not set them
    mBoolZSet = False
    
    mNewX.Set 1, 0, 0   'Initialize the orientation
    mNewY.Set 0, 1, 0   'to the major axis x,y,z
    mNewZ.Set 0, 0, 1
    
    Dim i As Integer    'Clear out any stored rotations.
    For i = 0 To 5      'by initializing them all to 0
        mXYZRotate(i) = 0
    Next i
    
End Sub

Public Property Let Tolerance(ByVal dblTol As Double)

''' This property sets the tolerance
''' used in testing
''' This will be based on the accuracy of
''' your inputs.  It gets initialized to
''' 0.0000001
''' Example of call:
''' myOri.tolerance = 0.001

    mTolerance = dblTol
    
End Property

Public Property Get Tolerance() As Double

''' This property retrieves the
''' tolerance that was set
''' either by default or by the user
''' Example of call:
''' Dim dblTol As Double
''' dblTol = myOri.tolerance

    Tolerance = mTolerance
    
End Property

Public Property Let RadianInput(ByVal bolRadInput As Boolean)

'''  This property sets a flag to
'''  state that the angle input is in
'''  radians
'''  Example of call:
'''  myOri.RadianInput = true

    mBoolRadianInput = bolRadInput
    
End Property

Public Property Get RadianInput() As Boolean

'''  This property retrienves a flag
'''  that states whether the angle input is in
'''  radians or not
'''  Example of call:
'''  dim test as boolean
'''  test = myOri.RadianInput

    RadianInput = mBoolRadianInput
    
End Property

Public Property Set XAxis(ByRef vecPassed As AutoMath.DVector)
    
''' This property sets the
''' new X direction for
''' our orientation
''' Example of call:
''' myOri.SetX 1,3,5

    Const METHOD = "x:"
    On Error GoTo ErrorHandler
    
    Dim isPerp As Boolean
    Set mNewX = vecPassed
    mBoolXSet = True
            
    If (mBoolYSet = True) Then
    'Then we need to check if x and y are perp
        isPerp = CheckPerpendicular(mNewY, mNewX)
        If (Not isPerp) Then
            Err.Description = "Invalid Ori"
            Err.Raise -2147467259
        End If
    End If
    
    If (mBoolZSet = True) Then
    'Then we need to check if x and z are perp
        isPerp = CheckPerpendicular(mNewZ, mNewX)
        If (Not isPerp) Then
            Err.Description = "Invalid Ori"
            Err.Raise -2147467259
        End If
    End If
    
   Exit Property
   
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    
End Property

Public Property Set YAxis(ByRef vecPassed As AutoMath.DVector)

''' This property sets the
''' new Y direction for
''' our orientation
''' Example of call:
''' myOri.SetY 1,3,5

    Const METHOD = "y:"
    On Error GoTo ErrorHandler

    Dim isPerp As Boolean
    Set mNewY = vecPassed
    mBoolYSet = True

    If mBoolXSet Then
    'Then we need to check if x and y are perp
        isPerp = CheckPerpendicular(mNewX, mNewY)
        If (Not isPerp) Then
            Err.Description = "Invalid Ori"
            Err.Raise -2147467259
        End If
    End If
    
    If mBoolZSet Then
    'Then we need to check if x and y are perp
        isPerp = CheckPerpendicular(mNewZ, mNewY)
        If (Not isPerp) Then
            Err.Description = "Invalid Ori"
            Err.Raise -2147467259
        End If
    End If
    Exit Property
   
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Property

Public Property Set ZAxis(ByRef vecPassed As AutoMath.DVector)

''' This property sets the
''' new Z direction for
''' our orientation
''' Example of call:
''' myOri.SetZ 1,3,5

    Const METHOD = "y:"
    On Error GoTo ErrorHandler

    Dim isPerp As Boolean
    Set mNewZ = vecPassed
    mBoolZSet = True

    If mBoolXSet Then
    'Then we need to check if x and y are perp
        isPerp = CheckPerpendicular(mNewZ, mNewX)
        If (Not isPerp) Then
            Err.Description = "Invalid Ori"
            Err.Raise -2147467259
        End If
    End If
    
    If mBoolYSet Then
    'Then we need to check if x and y are perp
        isPerp = CheckPerpendicular(mNewZ, mNewY)
        If (Not isPerp) Then
            Err.Description = "Invalid Ori"
            Err.Raise -2147467259
        End If
    End If
    Exit Property
   
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Property


Public Property Get XAxis() As AutoMath.DVector

''' This property retrieves the
''' new X direction and passes it back
''' Example of call:
''' Dim myOri As New Orientation
''' Dim vecXOri as Automath.DVector
''' Set vecXOri = myOri.getX

    Set XAxis = mNewX

End Property

Public Property Get YAxis() As AutoMath.DVector

''' This property retrieves the
''' new Y direction and passes it back
''' Example of call:
''' Dim myOri As New Orientation
''' Dim vecYOri as Automath.DVector
''' Set vecYOri = myOri.getY

    Set YAxis = mNewY

End Property

Public Property Get ZAxis() As AutoMath.DVector

''' This property retrieves the
''' new Z direction and passes it back
''' Dim myOri As New Orientation
''' Example of call:
''' Dim vecZOri as Automath.DVector
''' Set vecZOri = myOri.getZ

    Set ZAxis = mNewZ

End Property

Public Property Get IsXAxisSet() As Boolean

''' This function informs the client
''' if the x vector has been set
''' Example of call:
''' Dim xset as Boolean
''' xset = myOri.isXSet

    IsXAxisSet = mBoolXSet

End Property

Public Property Get IsYAxisSet() As Boolean

''' This function informs the client
''' if the y vector has been set
''' Example of call:
''' Dim yset as Boolean
''' yset = myOri.isYSet

    IsYAxisSet = mBoolYSet

End Property

Public Property Get IsZAxisSet() As Boolean

''' This function informs the client
''' if the z vector has been set
''' Example of call:
''' Dim yset as Boolean
''' yset = myOri.isYSet

    IsZAxisSet = mBoolZSet

End Property


Public Sub FindRemainingUnsetAxis()

''' This function determins the
''' last vector if two have been set
''' Example of call:
''' myOri.findLastVec

    If (mBoolXSet = True And mBoolYSet = True) Then
        Set mNewZ = GetNormalDirVectorFrom2Vectors(mNewY, mNewX)
    ElseIf (mBoolXSet = True And mBoolZSet = True) Then
        Set mNewY = GetNormalDirVectorFrom2Vectors(mNewX, mNewZ)
    ElseIf (mBoolYSet = True And mBoolZSet = True) Then
        Set mNewX = GetNormalDirVectorFrom2Vectors(mNewZ, mNewY)
    End If

End Sub


Private Function GetNormalDirVectorFrom2Vectors(ByVal vecA As IJDVector, _
                                                ByVal vecC As IJDVector _
                                               ) As IJDVector

''' This function computes
''' the normal vector of two passed.
''' This is accomplished by computing
''' the cross product of the two vectors
''' vectors
''' Example of call:
''' Dim vecOne As IJDVector
''' Dim vecTwo As IJDVector
''' Dim vecResult As IJDVector
''' vecOne.Set 1, 0, 0
''' vecTwo.Set 0, 1, 0
''' Set vecResult = getNormalDirVectorFrom2Vectors(vecOne, vecTwo)
''' the result is vecResult = 0,0,1

   Dim vecD As IJDVector  ' vector C cross A
   Set vecD = New DVector

'  --- compute vector cross product to get a normal vector to plane ---
   Set vecD = vecC.Cross(vecA)

'  --- convert normal vector to a unit direction vector ---
   vecD.length = 1#

'  --- set function value ---
   Set GetNormalDirVectorFrom2Vectors = vecD

'  --- clean up ---
   Set vecA = Nothing
   Set vecC = Nothing
   Set vecD = Nothing

End Function

Private Function CheckPerpendicular(ByVal vecOne As AutoMath.DVector, _
                                    ByVal vecTwo As AutoMath.DVector _
                                    ) As Boolean
                            
''' This function is to determins
''' whether 2 vectors are perpendicular
''' Example of call:
''' Dim perpTest As Boolean
''' perpTest = CheckPerpendicular(firstVec, SecondVec)
                            
    Dim dblAngle As Double

    'The dot product should be = 0 or w/in tolerance
    dblAngle = vecOne.Dot(vecTwo)
    
    If (Abs(dblAngle) < mTolerance) Then
    'Case:  They are 90 degres
        CheckPerpendicular = True
    Else
    'Case:  They are not 90 degres
        CheckPerpendicular = False
    End If
    
End Function

Public Property Let RotationAboutX(ByVal dblAngle As Double)
    
''' This property sets the
''' amount of rotation about X
''' Example of call:
''' myOri.setRotationAboutX = 45

    'Convert degrees to radians if
    'necessary else radians were passed
    Dim dblRadian As Double
    If (mBoolRadianInput = False) Then
        dblRadian = (dblAngle * (4 * Atn(1))) / 180
    Else
        dblRadian = dblAngle
    End If
    
    'Set the order and magnitude of the rotation
    'ex.  If they request this as the first rotation
    '     of 5 degress about x it will be the
    '     last one loaded into the matrix
    mXYZRotate(mCount) = "X"
    mCount = mCount + 1
    mXYZRotate(mCount) = dblRadian
    mCount = mCount + 1

End Property

Public Property Let RotationAboutY(ByVal dblAngle As Double)
    
''' This property sets the
''' amount of rotation about Y
''' Example of call:
''' myOri.setRotationAboutY = 45
    
    'Convert degrees to radians if
    'necessary else radians were passed
    Dim dblRadian As Double
    If (mBoolRadianInput = False) Then
        dblRadian = (dblAngle * (4 * Atn(1))) / 180
    Else
        dblRadian = dblAngle
    End If

    mXYZRotate(mCount) = "Y"
    mCount = mCount + 1
    mXYZRotate(mCount) = dblRadian
    mCount = mCount + 1

End Property

Public Property Let RotationAboutZ(ByVal dblAngle As Double)

''' This property sets the
''' amount of rotation about Y
''' Example of call:
''' myOri.setRotationAboutZ = 45

    'Convert degrees to radians if
    'necessary else radians were passed
    Dim dblRadian As Double
    If (mBoolRadianInput = False) Then
        dblRadian = (dblAngle * (4 * Atn(1))) / 180
    Else
        dblRadian = dblAngle
    End If

    mXYZRotate(mCount) = "Z"
    mCount = mCount + 1
    mXYZRotate(mCount) = dblRadian
    mCount = mCount + 1
    
End Property

Public Sub ApplyRotations()
    
''' This property applies the set rotations.
''' They are loaded into the matrix in
''' the reverse order then they were set.
''' Example of call:
''' myOri.applyRotations
    
    Dim tmxResult As IngrGeom3D.IJDT4x4
    Set tmxResult = New DT4x4
    tmxResult.LoadIdentity
    Dim vecAboutAxis As IJDVector
    Set vecAboutAxis = New DVector
    
    'Apply the rotations in revers order
    'from they way they were set by the client
    Dim i As Integer
    For i = 5 To 0 Step -2
        If mXYZRotate(i - 1) = "X" Then
            vecAboutAxis.Set 1, 0, 0
            tmxResult.Rotate Val(mXYZRotate(i)), vecAboutAxis
        ElseIf mXYZRotate(i - 1) = "Y" Then
            vecAboutAxis.Set 0, 1, 0
            tmxResult.Rotate Val(mXYZRotate(i)), vecAboutAxis
        ElseIf mXYZRotate(i - 1) = "Z" Then
            vecAboutAxis.Set 0, 0, 1
            tmxResult.Rotate Val(mXYZRotate(i)), vecAboutAxis
        End If
    Next i
        
    'Get the new axis vectors from the resultant vector.
    mNewX.Set tmxResult.IndexValue(0), tmxResult.IndexValue(1), tmxResult.IndexValue(2)
    mNewY.Set tmxResult.IndexValue(4), tmxResult.IndexValue(5), tmxResult.IndexValue(6)
    mNewZ.Set tmxResult.IndexValue(8), tmxResult.IndexValue(9), tmxResult.IndexValue(10)
End Sub
