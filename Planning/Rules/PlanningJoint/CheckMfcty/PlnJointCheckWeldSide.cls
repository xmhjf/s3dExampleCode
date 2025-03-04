VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlnJointCheckWeldSide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ****************************************************************************
' Copyright (c) 2002, Intergraph Corporation.  All rights reserved.
'
' File
'
'   PlnJointCheckWeldSide
'
' Author
'
'   Devi Kishore Adiraju
'
' Abstract
'
'   Planning Joint Weld Side Manufacturability Check
'
' Description
'
'   Checks if the weld side property of planning joints is correct
'
' History
'
'   7th Aug 2006   Kishore    Creation
'                             DI-CP·121259  Create CheckMfcty program to check Planning Joint has invalid RefPartName
'   16th Sept 2008 Srinivas   TR-CP-150370  Application error while running "CheckMfcty" command in Planning application
' ***************************************************************************

Option Explicit

Private Const Module = "PlnCheckMfcty.PlnJointCheckWeldSide"
Private Const PROG_ID = "PlnCheckMfcty.PlnJointCheckWeldSide"

Private Enum EDotProd
    Zero = 0
    Positive = 1
    Negative = 2
End Enum


Implements IJCheckMfcty

'********************************************************************
'Previously the determination of Weld Side property of Planning Joints
'relied on the property "RefPartName" which is exposed on occurrence
'interfaces of the physical connection objects. But there was a chance
'that the value of this property is not up to date and hence incorrect.
'This was causing incorrect value of the Weld Side property. This issue
'was resolved but, planning joints which are migrated can have incorrect
'Weld Side. This check uses the correct method to calculate weldside and
'compares it with the existing value of Weld Side of the joints to find out
'errant joints.
'********************************************************************
Private Sub IJCheckMfcty_ExecuteCheck _
       (ByVal oCollection As GSCADGlobalCheck.IJElements, _
        ByVal bStopOnError As Boolean, _
        pOptionCodes() As Variant, _
        ByVal pCallingObj As Object)

    Const Method = "IJCheckMfcty_ExecuteCheck"
    On Error GoTo ErrorHandler

    Dim oCallBack               As IJCheckMfctyCallback
    Dim oJoints                 As IJElements
    Dim oPlnJoint               As IJPlnJoint
    Dim lWeldSide               As JointWeldSide
    Dim lCorrectWeldSide        As JointWeldSide
    Dim strCorrectWeldSide      As String
    Dim strWrongWeldSide        As String
    Dim strErrMsg               As String
    Dim lngDummy                As Long
    Dim lngCancel               As Long
    
    ' Setup the Check Manufacturability call back mechanism.
    Set oCallBack = pCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    ' if the collection is empty, do not continue
    If oCollection.Count = 0 Then GoTo Cleanup
    
    ' get all the planning joints from the collection
    GatherPlanningJoints oCollection, oJoints
    
    For Each oPlnJoint In oJoints
        Dim oPlnJointProps As IJPlnJointProps
        
        Set oPlnJointProps = oPlnJoint
        lWeldSide = oPlnJointProps.WeldSide
        lCorrectWeldSide = CalculateWeldSide(oPlnJoint)
        
        If lCorrectWeldSide <> lWeldSide Then
                            
            strErrMsg = "The Weld Side of the Planning Joint is incorrect" & vbNewLine _
                    & "Current Weld Side Value: " + WeldSideString(lWeldSide) & vbNewLine _
                    & "Correct Weld Side Value: " + WeldSideString(lCorrectWeldSide)

            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oPlnJoint              the object with error/warning
            '   PROG_ID                prog id of this rule
            '   ESeverity.siError      severity of the error (101 error or 102 warning)
            '   8                      code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            
            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
                                   ESeverity.siError, 8, strErrMsg, "", ""
        End If
    Next
GoTo Cleanup


ErrorHandler:
    MsgBox Err.Description
    
Cleanup:
    Set oCallBack = Nothing
    Set oCollection = Nothing
    Set oPlnJoint = Nothing
    Set oPlnJointProps = Nothing
End Sub

Private Function WeldSideString(lWeldSide As JointWeldSide) As String
    Const Method = "WeldSideString"
    On Error GoTo ErrorHandler
    
    Select Case lWeldSide
        Case PLN_WeldSide_Molded
            WeldSideString = "Molded"
        Case PLN_WeldSide_AntiMolded
            WeldSideString = "Anti-Molded"
        Case PLN_WeldSide_Reference
            WeldSideString = "Reference"
        Case PLN_WeldSide_NonReference
            WeldSideString = "Non-Reference"
        Case Else
            WeldSideString = "Unknown"
    End Select
    
Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Function CalculateWeldSide(ByVal oPlnJoint As IJPlnJoint) As JointWeldSide
    Const sMETHOD As String = "CalculateWeldSide"
    On Error GoTo ErrorHandler
    
    'going to default to weld side molded
    CalculateWeldSide = PLN_WeldSide_Molded

    'get a wrapper for the planning joint
    Dim oPlnJointHelper As IJPlnJointHelper
    Set oPlnJointHelper = New PlnJointHelper
    
    Set oPlnJointHelper.object = oPlnJoint
    
    Dim oGeomHelper As PlnGeometryHelper.PlnGeometryHelper
    Set oGeomHelper = New PlnGeometryHelper.PlnGeometryHelper
    
    Dim oPlateWrap As StructDetailObjects.PlatePart
    Dim oPort As IJPort
    Dim oPosition As IJDPosition
    
    Dim dMinPos As Double
    Dim dsrcX As Double, dsrcY As Double, dsrcZ As Double, dinX As Double, dinY As Double, dinZ As Double
    Dim dPts1() As Double, dPts2() As Double, dPars1() As Double, dPars2() As Double
    Dim lNumPts As Long
    Dim oComplexString As IJComplexString
    Dim oComplexCurve As IJCurve
    Dim oSurfaceBody As IJDModelBody
    
    'need to get the joined part
    Dim oJoinedPart As Object
    Set oJoinedPart = oPlnJointHelper.JoinedPart
    
    Dim oGeomSrvs As IGeometryServices
    Set oGeomSrvs = New IngrGeom3D.GeometryFactory
        
    Dim oPos As IJDPosition, oClosePos As IJDPosition
    Set oPos = New DPosition
    
    If TypeOf oJoinedPart Is IJPlate Then
        Set oPlateWrap = New StructDetailObjects.PlatePart
        Set oPlateWrap.object = oJoinedPart
        
        Dim strMoldedSide As String
        strMoldedSide = oPlateWrap.MoldedSide
        
        'get the proper port geometry
        If strMoldedSide = "Base" Then
            Set oPort = oPlateWrap.BasePort(BPT_Base)
        ElseIf strMoldedSide = "Offset" Then
            Set oPort = oPlateWrap.BasePort(BPT_Offset)
        Else
            GoTo Cleanup
        End If
        
        'check the distance between the planning joint and the "Molded" side port
        Set oComplexString = oPlnJoint
        
        oComplexString.GetCurve 1, oComplexCurve
        oComplexCurve.EndPoints dsrcX, dsrcY, dsrcZ, dinX, dinY, dinZ
        oPos.Set dsrcX, dsrcY, dsrcZ
        
        Set oSurfaceBody = oPort.Geometry
        oSurfaceBody.GetMinimumDistanceFromPosition oPos, oClosePos, dMinPos
        
'        oSurface.DistanceBetween oComplexCurve, dMinPos, dsrcX, dsrcY, dsrcZ, dinX, dinY, dinZ, lNumPts, dPts1, dPts2, dPars1, dPars2
'        oGeomHelper.GetDistanceFromCurveToSurface ocurve1, oSurface, oPosition, dDistance
        If dMinPos > oGeomSrvs.DistTolerance Then
            'the planning joint is not on the "Molded" side of the plate
            CalculateWeldSide = PLN_WeldSide_AntiMolded
        End If
        
    ElseIf TypeOf oJoinedPart Is IJProfile Then
    
        'Get curve from the joint
        Dim oCurve As IJCurve
        Set oCurve = oPlnJoint
    
        'Get the Start and End Parameters of Curve
        Dim dStartPar As Double
        Dim dEndPar As Double
        oCurve.ParamRange dStartPar, dEndPar
    
        'Get the Parameter for the MidPoint
        Dim dMidPar As Double
        dMidPar = (dStartPar + dEndPar) / 2
    
        'Get the MidPoint on the Curve
        Dim dX As Double
        Dim dY As Double
        Dim dZ As Double
        oCurve.Position dMidPar, dX, dY, dZ
    
        Dim oMidPoint As IJDPosition
        Set oMidPoint = New DPosition
        oMidPoint.Set dX, dY, dZ
        
        'Get the primary and secondary orientation vectors of the profile at the mid point
        Dim oPrimOrientVec      As IJDVector
        Dim oSecOrientVec       As IJDVector
        Dim oOrigin             As IJDPosition
        
        Dim oSDProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport
        Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
        
        Set oSDProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport
        Set oSDPartSupport = oSDProfilePartSupport
        
        Set oSDPartSupport.Part = oJoinedPart
        oSDProfilePartSupport.GetOrientation oMidPoint, oSecOrientVec, oPrimOrientVec, oOrigin
        
        'Normalize the secondary orientation vector and reverse its direction to get
        'a vector along the reference side
        oSecOrientVec.Length = -1
        
        'Get Normal vectors to the faces adjacent to the mounting face, i.e., the faces
        'on which the joints lie
        
        Dim oRefSideNormal As IJDVector     'This doesnot imply the normal to the face on
                                            'Reference side. Face on reference side is to
                                            'be found out through dot product.
                                            
        'Get Normal to a face on one of the sides of the mounting face
        Dim eThicknessSide  As ThicknessSide
        eThicknessSide = SideA
        
        Dim oRefPort As IJPort
        Set oRefPort = oSDProfilePartSupport.GetFacePortAdjacentToMountingFace(eThicknessSide)
        
        Dim oPartInfo As GSCADStructGeomUtilities.PartInfo
        Dim bApproxUsed As Boolean
        
        Set oPartInfo = New GSCADStructGeomUtilities.PartInfo
        Set oRefSideNormal = oPartInfo.GetPortNormal(oRefPort, bApproxUsed)
        
        'Get face port on the other side
        eThicknessSide = SideB

        Dim oNonRefPort As IJPort
        Set oNonRefPort = oSDProfilePartSupport.GetFacePortAdjacentToMountingFace(eThicknessSide)
        
        'Normalize the vector
        oRefSideNormal.Length = 1
        
        Dim eDotProduct        As EDotProd
        eDotProduct = DotProduct(oSecOrientVec, oRefSideNormal)
        
        If eDotProduct = Negative Then
            Set oRefPort = oNonRefPort
        End If
    
        'check the distance between the planning joint and the "Reference" side port
        Set oComplexString = oPlnJoint
        
        oComplexString.GetCurve 1, oComplexCurve
        oComplexCurve.EndPoints dsrcX, dsrcY, dsrcZ, dinX, dinY, dinZ
        oPos.Set dsrcX, dsrcY, dsrcZ
        
        Set oSurfaceBody = oRefPort.Geometry
        oSurfaceBody.GetMinimumDistanceFromPosition oPos, oClosePos, dMinPos
        
'        oSurface.DistanceBetween oComplexCurve, dMinPos, dsrcX, dsrcY, dsrcZ, dinX, dinY, dinZ, lNumPts, dPts1, dPts2, dPars1, dPars2
'        oGeomHelper.GetDistanceFromCurveToSurface ocurve2, oSurface1, oPosition, dDistance
        If dMinPos > oGeomSrvs.DistTolerance Then
            'the planning joint is not on the "Reference" side of the profile
            CalculateWeldSide = PLN_WeldSide_NonReference
        Else
            CalculateWeldSide = PLN_WeldSide_Reference
        End If
    End If
    
Cleanup:
    Set oJoinedPart = Nothing
    
    Set oPlateWrap = Nothing
    Set oPort = Nothing
    Set oPosition = Nothing
    
    Set oGeomHelper = Nothing
    Set oCurve = Nothing
    Set oMidPoint = Nothing
    Set oPrimOrientVec = Nothing
    Set oSecOrientVec = Nothing
    Set oOrigin = Nothing
    Set oSDProfilePartSupport = Nothing
    Set oSDPartSupport = Nothing
    Set oRefSideNormal = Nothing
    Set oRefPort = Nothing
    Set oNonRefPort = Nothing
    Set oGeomSrvs = Nothing
    Set oComplexCurve = Nothing
    Set oComplexString = Nothing
    Set oSurfaceBody = Nothing
    
    Exit Function

ErrorHandler:
    MsgBox Err.Description
End Function


'Finds out if the dot product is postive, negative or zero
Private Function DotProduct(oTangent As IJDVector, oAxis As IJDVector) As EDotProd
    Const sMETHOD As String = "IsDotProductZero"
    On Error GoTo ErrorHandler

    Dim oGeomSrvs           As IGeometryServices
    Dim dTolerance          As Double

    Set oGeomSrvs = New IngrGeom3D.GeometryFactory
    dTolerance = oGeomSrvs.DistTolerance
    
    If (oTangent.Dot(oAxis)) < dTolerance Then
        If Abs(oTangent.Dot(oAxis)) < dTolerance Then
            DotProduct = Zero
        Else
            DotProduct = Negative
        End If
    Else
        DotProduct = Positive
    End If
    
    Set oGeomSrvs = Nothing

    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
End Function
