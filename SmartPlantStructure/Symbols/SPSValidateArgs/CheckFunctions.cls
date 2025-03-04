VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CheckFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'   CheckFunctions.cls
'
'Author
'   Jayadev Pulaparty      10/22/2003
'
'Description
'   The functions in this file are custom check methods of the symbol machinery
'   They validate the attribute value set on the symbol instance
'
'   Example:
'
'   These functions will be set on the symbol defintion as shown here for the case
'   of the height attribute validation set on the handrail symbol definition
'   SPSHandrailMacros.TypeATopMounted -
'
'       oInput.Name = "Height"
'       oInput.Description = "Height of the Handrail from path to top of Top Rail"
'       oInput.Properties = igINPUT_IS_A_PARAMETER
'       oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
'
'   This check will ensure that the "Height" attribute for the symbol instance
'   can never be zero or negative
'
'Notes
'
'History:
'   3-Mar-06        SR          TR94637: GTZero function is mdofied to check for Minimum
'                               tolerance as most of length/width attributes are used for
'                               computational geometry.
'
'   15-Sep-09       WR          TR-163700.  Added SmallDISTTOL to the method BT90And75Check
'                               to allow inclusive entries of 75 and 90 degrees.
'
'*******************************************************************
Const PI = 3.14159265358979
Const SmallDISTTOL = 0.000001        '10 e-06
Private m_oLocalizer As IJLocalizer

'*************************************************************************
'
'Function
'
'   <BT90And75Check>
'
'Abstract
'
'   This function sets the boolean flag to true if the angle attribute values
'   to be between 75-90 degrees and false otherwise
'
'Arguments
'
'   pDef        -   symbol definition instance (input)
'   pobj        -   value object (input)
'   pbool       -   boolean flag (input/output)
'   ErrMessage  -   error message(input/output - optional)
'
'Return
'
'   None
'
'Exceptions
'
'   None
'
'***************************************************************************

Public Sub BT90And75Check(ByVal pDef As Object, ByVal pobj As Object, ByRef pbool As Boolean, ErrMessage As String)
On Error GoTo ErrHandler
    Dim dvalue As Double
    
    dvalue = pobj
    If dvalue <= 1.5707963267949 + SmallDISTTOL And dvalue > (75 * 1.5707963267949 / 90) - SmallDISTTOL Then
        pbool = True
    Else
        ErrMessage = m_oLocalizer.GetString(IDS_VALIDATE_VALUE_7590_DEGREE, "Value should be b/w 75-90 degree")
        pbool = False
    End If

Exit Sub
ErrHandler:
pbool = False

End Sub

'*************************************************************************
'
'Function
'
'   <GTZero>
'
'Abstract
'
'   This function sets the boolean flag to true if the attribute value is
'   greater than zero and false otherwise
'
'Arguments
'
'   pDef        -   symbol definition instance (input)
'   pobj        -   value object (input)
'   pbool       -   boolean flag (input/output)
'   ErrMessage  -   error message(input/output - optional)
'
'Return
'
'   None
'
'Exceptions
'
'   None
'
'***************************************************************************

Public Sub GTZero(ByVal pDef As Object, ByVal pobj As Object, ByRef pbool As Boolean, ErrMessage As String)
    On Error GoTo ErrHandler
    Dim dvalue As Double
    
    dvalue = pobj
    
    If dvalue > SmallDISTTOL Then
        pbool = True
    Else
        ErrMessage = m_oLocalizer.GetString(IDS_VALIDATE_VALUE_GREATERTHAN_ZERO, "Value should be greater than Zero")
        pbool = False
    End If
    Exit Sub
ErrHandler:
pbool = False
    
End Sub

'*************************************************************************
'
'Function
'
'   <NegativeNum>
'
'Abstract
'
'   This function sets the boolean flag to true if the attribute value is
'   greater than or equal to zero and false otherwise
'
'Arguments
'
'   pDef        -   symbol definition instance (input)
'   pobj        -   value object (input)
'   pbool       -   boolean flag (input/output)
'   ErrMessage  -   error message(input/output - optional)
'
'Return
'
'   None
'
'Exceptions
'
'   None
'
'***************************************************************************

Public Sub NegativeNum(ByVal pDef As Object, ByVal pobj As Object, ByRef pbool As Boolean, ErrMessage As String)
    On Error GoTo ErrHandler
   
    Dim dvalue As Double
    dvalue = pobj
    
    If dvalue >= 0 Then
        pbool = True
    Else
        ErrMessage = m_oLocalizer.GetString(IDS_VALIDATE_VALUE_POSITIVE, "Value should be zero or positive")
        pbool = False
    End If
    Exit Sub
ErrHandler:
pbool = False
    
End Sub

'*************************************************************************
'
'Function
'
'   <LT90Deg>
'
'Abstract
'
'   This function sets the boolean flag to true if the angle attribute values
'   to be between 65-25 degrees and false otherwise
'
'Arguments
'
'   pDef        -   symbol definition instance (input)
'   pobj        -   value object (input)
'   pbool       -   boolean flag (input/output)
'   ErrMessage  -   error message(input/output - optional)
'
'Return
'
'   None
'
'Exceptions
'
'   None
'
'***************************************************************************

Public Sub LT90Deg(ByVal pDef As Object, ByVal pobj As Object, ByRef pbool As Boolean, ErrMessage As String)
On Error GoTo ErrHandler
    Dim dvalue As Double
    dvalue = pobj
    If dvalue < 1.1344 And dvalue > 0.436 Then ' 65 deg
        pbool = True
    Else
        ErrMessage = m_oLocalizer.GetString(IDS_VALIDATE_VALUE_2565_DEGREE, "Value should be b/w 25-65 degree")
        pbool = False
    End If
    Exit Sub
ErrHandler:
pbool = False
End Sub

'*************************************************************************
'
'Function
'
'   <IN360Deg>
'
'Abstract
'
'   This function sets the boolean flag to true if the angle attribute values
'   to be between 0-360 degrees and false otherwise
'
'Arguments
'
'   pDef        -   symbol definition instance (input)
'   pobj        -   value object (input)
'   pbool       -   boolean flag (input/output)
'   ErrMessage  -   error message(input/output - optional)
'
'Return
'
'   None
'
'Exceptions
'
'   None
'
'***************************************************************************

Public Sub IN360Deg(ByVal pDef As Object, ByVal pobj As Object, ByRef pbool As Boolean, ErrMessage As String)
On Error GoTo ErrHandler
   
    Dim dvalue As Double
    dvalue = pobj
    
'    If dvalue <= PI And dvalue >= -PI Then
'   *************************************************************
    '   Modified : Murali Mohan Kaveti [May 15, 2003]
    '   (TR no. JTS-TR-CP33119)
    '   Code to suit for all values of TopRailsectionAngle, MiddleRailSectionAngle,
    '   PostSectionAngle and ToePlateSectionAngle. Modified code to consider
    '   angle from 0 to 360 deg. rather than -180 to 180 deg.
'   *************************************************************

    If dvalue > (2 * PI) Then
        dvalue = (dvalue) Mod (2 * PI)
    ElseIf dvalue < 0 Then
        Dim temp As Double
        dvalue = -(dvalue)
        temp = (dvalue) Mod (2 * PI)
        dvalue = (2 * PI) - temp
    End If

    If dvalue <= (2 * PI) And dvalue >= 0 Then
        pbool = True
    Else
        ErrMessage = m_oLocalizer.GetString(IDS_VALIDATE_VALUE_0360_DEGREE, "Value should be b/w 0-360 degree")
        pbool = False
    End If
    
    Exit Sub
    
ErrHandler:
    pbool = False
End Sub

'*************************************************************************
'
'Function
'
'   <IsCrossSectionName>
'
'Abstract
'
'   This function sets the boolean flag to true if the supplied crossSectionName
'   is a valid one and false otherwise
'
'   This function is returning true by default as of now
'   The user can plug-in his own custom criterion for this check to pass
'
'Arguments
'
'   pInput      -   symbol definition instance (input)
'   pObject     -   value object (input)
'   pIsValid    -   boolean flag (input/output)
'   errMsg      -   error message(input/output - optional)
'
'Return
'
'   None
'
'Exceptions
'
'   None
'
'***************************************************************************

Public Sub IsCrossSectionName(pInput As Object, pObject As Object, ByRef pIsValid As Boolean, ByRef errMsg As String)
    pIsValid = True
    Exit Sub
End Sub

'*************************************************************************
'
'Function
'
'   <CheckForSketch3d>
'
'Abstract
'
'   This function sets the boolean flag to true if the supplied object
'   is a IJDSketch3d object and false otherwise
'
'Arguments
'
'   pInput      -   symbol definition instance (input)
'   pObject     -   value object (input)
'   pIsValid    -   boolean flag (input/output)
'   errMsg      -   error message(input/output - optional)
'
'Return
'
'   None
'
'Exceptions
'
'   None
'
'***************************************************************************

Public Sub CheckForSketch3d(pInput As Object, pObject As Object, ByRef pIsValid As Boolean, ByRef errMsg As String)
    
    pIsValid = False
    On Error GoTo ErrHandler
    
    Dim pSketch3d As IJDSketch3d
    Set pSketch3d = pObject
    
    pIsValid = True
    Set pSketch3d = Nothing
    
    Exit Sub
    
ErrHandler:
    errMsg = m_oLocalizer.GetString(IDS_VALIDATE_INPUT_SKETCH3D, "Input has to be sketch3d ")
End Sub

'*************************************************************************
'
'Function
'
'   <IsMaterailGrade>
'
'Abstract
'
'   This function sets the boolean flag to true if the supplied material grade
'   is a valid one and false otherwise
'
'   This function is returning true by default as of now
'   The user can plug-in his own custom criterion for this check to pass
'
'Arguments
'
'   pInput      -   symbol definition instance (input)
'   pObject     -   value object (input)
'   pIsValid    -   boolean flag (input/output)
'   errMsg      -   error message(input/output - optional)
'
'Return
'
'   None
'
'Exceptions
'
'   None
'
'***************************************************************************

Public Sub IsMaterailGrade(pInput As Object, pObject As Object, ByRef pIsValid As Boolean, ByRef errMsg As String)
    pIsValid = True
    Exit Sub
End Sub

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
