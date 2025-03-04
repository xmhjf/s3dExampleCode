Attribute VB_Name = "ConnectionUtilities"
Option Explicit

'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : ConnectionUtilities.bas
'
'Author : RP
'
'Description :
'    Module for common connection constants/utilities
'
'History:
'
' 06/09/02   CE      Added GetIncidentMemberQuadrant() to help
'                               determine cope req based on member incidence
' 03/31/03   JS      Added attribute constants
' 08/07/03   MH      Added manual offset attribute constants
' 08/13/03   MH      Added IsPGSame
' 02/02/04   JS      Added IsAssemblyConnectionInConflictWithAnother
'                               for TR#53161; validating the AC to ensure
'                               another does not exist between either the
'                               supporting or the supported members.
'
' 02/16/04   RP      Add code to take care mirror of supporting member.
'                             Modified GetMiterPlanePosition() to return end port
'                             position of supported1 if supped2 CP lines don't
'                             intersect supped1 side surfaces.
'                             Modified GetCutDistAndPosFromPlane() to calculate
'                             distances from the other end
' 04/14/04  RP       Added CheckForGussetPlateAsmConn() and CheckForCornerBraceAsmConn()
' 04/26/06  AS       Added GetSectionWidthAndDepth, IsInterfaceInAttrList, IsAttrInInterface, IsSurfaceTypeAcceptable
' 06/06/06  MH       Removed AreMembersEndConnected  ( V3 Members )
' 06/13/06  RP       Lot of changes due to impact from curved members. Modified some existing methods
'                    and added new routins to take care of impacts (DI#84001)
' 10/03/06  AS       Added IsOperatorTypeValidForTrim to check for oper type for Surface Trim
'
' 01/25/07  RP       Added GetTangentLineAtCPandPosition() to fix TR#113075
' 01/31/07  AS       TR#109129 Added IsSurfaceTypeAcceptableForSurfaceTrim
' 08/12/07  RP       TR#124941 Added code to check return code form IJSurface->Intersect() in
'                    GetMiterPlanePosition().
' 04/22/09  RP       CR#36434 -modified GetCommonConnection()for FCs to work for Cans
' 26/05/16  APK      TR-294290,TR-294291 : Validation Checks were added to avoid record exceptions in IsCope1NeededByShapeAndIncidence() and IsCope2NeededByShapeAndIncidence()
'*****************************************************************************

Private Const MODULE = "ConnectionUtilities"

' Custom user attribute interfaces and attributes
Public Const UASPSINTERFACE_FCCenterlineRotation = "IJUASPSFCCenterlineRotation"
Public Const UA_FCCenterlineRotation_Edge = "Edge"
Public Const UA_FCCenterlineRotation_Reflect = "Reflect"

Public Const UASPSINTERFACE_FCSeatedFlushRotation = "IJUASPSFCSeatedFlushRotation"
Public Const UA_FCSeatedFlushRotation_Edge = "Edge"
Public Const UA_FCSeatedFlushRotation_Reflect = "Reflect"

Public Const UASPSINTERFACE_FCOffset = "IJUASPSFCOffset"
Public Const UA_FCOffset_Side = "Side"
Public Const UA_FCOffset_Offset = "Offset"
Public Const UA_FCOffset_OffsetDirection = "OffsetDirection"
Public Const LARGE_EDGE = 1# ' large edge in meters. Changed to 1m due to curved members as having this larger
'than the arc radius can cause problems

Public Const UASPSINTERFACE_FCManualOffset = "IJUASPSFCManualOffset"
Public Const UA_FCManualOffset_XOffset = "XOffset"
Public Const UA_FCManualOffset_YOffset = "YOffset"
Public Const UA_FCManualOffset_ZOffset = "ZOffset"
Public Const UA_FCManualOffset_CoordinateSystem = "CoordinateSystem"

Public Const ISPSMemberSystemStartEndNotify = "{E91186FF-051C-4c50-88A2-A8A3E062ADBC}"
Public Const ISPSMemberSystemEndEndNotify = "{43A92A8F-7D34-4038-A9BA-A25898C7C7DC}"
Public Const ISPSMemberSystemSuppingNotify2 = "{C155EED1-B0D8-41a3-B7A0-9526ACD67E2D}"
Public Const ISPSAxisEndPort = "{D1136F2C-51A0-4FD2-A6FE-E0E7DD59EC8A}"
Public Const ConstIJSurface = "{7D82F810-D270-11D1-9558-0060973D4824}"
Public Const ConstIJGeometry = "{96eb9676-6530-11d1-977f-080036754203}"
Public Const CONST_ISPSMemberPartGeometry = "{83FA7F63-3A6F-40BB-81C3-87EBC961661C}"
Public Const CONST_ISPSSplitAxisEndPort = "{A8A52E08-5933-45F9-8C41-F3D92507D0B7}"
Public Const CONST_ISPSSplitAxisAlongPort = "{C1110DD4-AA5D-46F4-8A3B-C98454D55D49}"
Public Const CONST_CAToMemberRelationCLSID = "{45E4020F-F8D8-47A1-9B00-C9570C1E0B17}"
Public Const CONST_IJDOccurrence = "{274317DB-0F9D-11D2-94AD-080036CD8E03}"
Public Const CONST_CustomPlatePart = "{A46498E6-9116-42B1-8A18-031415C07428}"

Public Const distTol = 0.0001
Public Const angleTol = 0.0001
                         
'*************************************************************************
'Function
'
'<IsSuppedAxisInSuppingXZPlane>
'
'Abstract
'
'<Function checks if Supported Member is in Supporting Members XZ Plane>
'
'Arguments
'
'<Supported and Supporting Members as MemberPartPrismatic>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsSuppedAxisInSuppingXZPlane(oSupped As ISPSMemberPartPrismatic, oSupping As ISPSMemberPartPrismatic) As Boolean
Const MT = "IsSuppedAxisInSuppingXZPlane"
    On Error GoTo ErrorHandler
    
    Dim dVals(16) As Double
    Dim oMat As IJDT4x4
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim oLine As IJLine
    Dim uvX As Double, uvY As Double, uvZ As Double
    
    IsSuppedAxisInSuppingXZPlane = False
    
    Dim pIGeomFact As New GeometryFactory
    Set oLine = pIGeomFact.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0, 0, 0)
    'Set oLine = New Line3d
    'get local coordinate system for  the supporting member
    
    oSupping.Rotation.GetTransform oMat
    oSupped.Axis.EndPoints sX, sY, sZ, eX, eY, eZ

    
    ' supported member line
    oLine.SetStartPoint sX, sY, sZ
    oLine.SetEndPoint eX, eY, eZ
    
    'transform supported member line to the supporting coordinate system
    oMat.Invert 'global to local transformation
    oLine.Transform oMat
    
    oLine.GetDirection uvX, uvY, uvZ
    
    If Abs(uvY) < distTol Then ' if supported axis along supporting local XZ plane, uvY is zero
        IsSuppedAxisInSuppingXZPlane = True
    End If
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function
'*************************************************************************
'Function
'
'<IsSuppedAxisInSuppingXYPlane>
'
'Abstract
'
'<Function checks if Supported Member is in Supporting Members XY Plane>
'
'Arguments
'
'<Supported and Supporting Members as MemberPartPrismatic>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsSuppedAxisInSuppingXYPlane(oSupped As ISPSMemberPartPrismatic, oSupping As ISPSMemberPartPrismatic) As Boolean
Const MT = "IsSuppedAxisInSuppingXYPlane"
    On Error GoTo ErrorHandler
    
    Dim dVals(16) As Double
    Dim oMat As IJDT4x4
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim oLine As IJLine
    Dim uvX As Double, uvY As Double, uvZ As Double
    
    IsSuppedAxisInSuppingXYPlane = False
    
    Dim pIGeomFact As New GeometryFactory
    Set oLine = pIGeomFact.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0, 0, 0)
    'Set oLine = New Line3d
    'get local coordinate system for  the supporting member
    
    oSupping.Rotation.GetTransform oMat
    oSupped.Axis.EndPoints sX, sY, sZ, eX, eY, eZ

    
    ' supported member line
    oLine.SetStartPoint sX, sY, sZ
    oLine.SetEndPoint eX, eY, eZ
    
    'transform supported member line to the supporting coordinate system
    oMat.Invert 'global to local transformation
    oLine.Transform oMat
    
    oLine.GetDirection uvX, uvY, uvZ
    
    If Abs(uvZ) < distTol Then ' if Supped axis along supporting local XY plane, uvZ is zero
        IsSuppedAxisInSuppingXYPlane = True
    End If
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function
                         
'*************************************************************************
'Function
'
'<IsSuppedAxisInSuppingXZPlane>
'
'Abstract
'
'<Function checks if Supported Member is in Supporting Members XZ Plane>
'
'Arguments
'
'<Supported and Supporting Members as MemberPartPrismatic>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************

 

'*************************************************************************
'Function
'
'<GetEndPortCloseToSupportingMemebr>
'
'Abstract
'
'<Function determines which end is closer to the Supporting Member>
'
'Arguments
'
'<Supported and Supporting Members Logical Axis>
'
'Return
'
'<SPSMemberAxisPortIndex Constants>
'
'Exceptions
'***************************************************************************
Public Function GetEndPortCloseToSupportingMemebr(oSuppd As ISPSLogicalAxis, oSupping As ISPSLogicalAxis) As SPSMemberAxisPortIndex
  Const MT = "GetEndPortCloseToSupportingMemebr"
    
    On Error GoTo ErrorHandler
    Dim sX0 As Double, sY0 As Double, sZ0 As Double
    Dim eX0 As Double, eY0 As Double, eZ0 As Double
    Dim sX1 As Double, sY1 As Double, sZ1 As Double
    Dim eX1 As Double, eY1 As Double, eZ1 As Double
    Dim dist0 As Double, dist1 As Double, cosA As Double, distNorm0 As Double, distNorm1 As Double
    
    Dim oSuppdPos0 As IJDPosition, oSuppdPos1 As IJDPosition
    Dim oSuppingPos0 As IJDPosition, oSuppingPos1 As IJDPosition
    Dim oVec As IJDVector, oVec0 As IJDVector, oVec1 As IJDVector
    
    Set oVec = New DVector
    Set oVec0 = New DVector
    Set oVec1 = New DVector
    Set oSuppdPos0 = New DPosition
    Set oSuppdPos1 = New DPosition
    Set oSuppingPos0 = New DPosition
    Set oSuppingPos1 = New DPosition
    
    
    oSuppd.GetLogicalStartPoint sX0, sY0, sZ0
    oSuppd.GetLogicalEndPoint eX0, eY0, eZ0
    oSupping.GetLogicalStartPoint sX1, sY1, sZ1
    oSupping.GetLogicalEndPoint eX1, eY1, eZ1
    
    
    'The end with lesser distance is assumed to be cut against the supporting member
    oSuppdPos0.x = sX0
    oSuppdPos0.y = sY0
    oSuppdPos0.z = sZ0
    oSuppdPos1.x = eX0
    oSuppdPos1.y = eY0
    oSuppdPos1.z = eZ0
    
    oSuppingPos0.x = sX1
    oSuppingPos0.y = sY1
    oSuppingPos0.z = sZ1
    oSuppingPos1.x = eX1
    oSuppingPos1.y = eY1
    oSuppingPos1.z = eZ1
    
    Set oVec = oSuppingPos1.Subtract(oSuppingPos0) 'vector from supporting start to supporting end
    Set oVec0 = oSuppdPos0.Subtract(oSuppingPos0) 'vector from supported start to supporting start
    Set oVec1 = oSuppdPos1.Subtract(oSuppingPos0) 'vector from supported end to supporting start
    dist0 = oVec0.Length 'distance from supported start to supporting end
    dist1 = oVec1.Length 'distance from supported end to supporting start
    
    If Abs(dist0) < distTol Then ' .1 mm
        GetEndPortCloseToSupportingMemebr = SPSMemberAxisStart
    ElseIf Abs(dist1) < distTol Then ' .1 mm
        GetEndPortCloseToSupportingMemebr = SPSMemberAxisEnd
    Else
        oVec.Length = 1 'unit vector from supporting start to supporting end
        oVec0.Length = 1 'unit vector from supported start to supporting start
        oVec1.Length = 1 'unit vector from supported end to supporting start
        cosA = oVec.Dot(oVec0) 'cosine of the angle between supporting member and vector from supporting start to supported start
        distNorm0 = dist0 * Sqr(1 - cosA * cosA) ' perpendicular distance from start of supported to supporting member
        cosA = oVec.Dot(oVec1) ' cosine of the angle between supporting member and vector from suppoting start to supported end
        distNorm1 = dist1 * Sqr(1 - cosA * cosA) ' perpendicular distance from end of supported to supporting member */
        
        If (distNorm1 >= distNorm0) Then 'start of supported is connected to the supporting member
            GetEndPortCloseToSupportingMemebr = SPSMemberAxisStart
        Else  ' end of supported is connected to the supporting member
            GetEndPortCloseToSupportingMemebr = SPSMemberAxisEnd
        End If
    End If
    Set oVec = Nothing
    Set oVec0 = Nothing
    Set oVec1 = Nothing
    Set oSuppdPos0 = Nothing
    Set oSuppdPos1 = Nothing
    Set oSuppingPos0 = Nothing
    Set oSuppingPos1 = Nothing
  Exit Function
ErrorHandler:  HandleError MODULE, MT
    
End Function

'*************************************************************************
'Function
'
'<IsSupportedAxisInPositiveY>
'
'Abstract
'
'<Checks if the Supported Member Axis is in Positive Y direction of Supporting>
'
'Arguments
'
'<Supported and Supporting MemberPartPrismatic, SPSMemberAxisPortIndex>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsSupportedAxisInPositiveY(oSupporting As ISPSMemberPartPrismatic, oSupported As ISPSMemberPartPrismatic, portIdx As SPSMemberAxisPortIndex) As Boolean
Const MT = "IsSupportedAxisInPositiveY"
    On Error GoTo ErrorHandler
    
    Dim dVals(16) As Double
    Dim oMat As IJDT4x4, oMatSupped As IJDT4x4
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim oVec As New DVector
    Dim uvX As Double, uvY As Double, uvZ As Double
    Dim oPosAlongAxis As IJDPosition

    
    IsSupportedAxisInPositiveY = False
    

    'get local coordinate system for  the supporting member
    Set oPosAlongAxis = GetConnectionPositionOnSupping(oSupporting, oSupported, portIdx)
    
    oSupporting.Rotation.GetTransformAtPosition oPosAlongAxis.x, oPosAlongAxis.y, oPosAlongAxis.z, oMat, Nothing
    oSupported.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    
    ' supported member line
    If portIdx = SPSMemberAxisStart Then
        oSupported.Rotation.GetTransformAtPosition sX, sY, sZ, oMatSupped, Nothing
        oVec.Set oMatSupped.IndexValue(0), oMatSupped.IndexValue(1), oMatSupped.IndexValue(2)
    Else
        oSupported.Rotation.GetTransformAtPosition eX, eY, eZ, oMatSupped, Nothing
        oVec.Set oMatSupped.IndexValue(0), oMatSupped.IndexValue(1), oMatSupped.IndexValue(2)
        oVec.[Scale] -1 'we want the direction which is away from the connection point
    End If
    
    'transform supported tangent line to the supporting coordinate system
    oMat.Invert 'global to local transformation
    Set oVec = oMat.TransformVector(oVec)
    
    oVec.Get uvX, uvY, uvZ
    
    If uvY >= 0# Then  ' if supported axis on the  supporting positive Y  side, uvy >0.0
        IsSupportedAxisInPositiveY = True
    End If
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function
'*************************************************************************
'Function
'
'<IsSupportedAxisInPositiveZ>
'
'Abstract
'
'<Checks if the Supported Member Axis is in Positive Z direction of Supporting>
'
'Arguments
'
'<Supported and Supporting MemberPartPrismatic, SPSMemberAxisPortIndex>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsSupportedAxisInPositiveZ(oSupporting As ISPSMemberPartPrismatic, oSupported As ISPSMemberPartPrismatic, portIdx As SPSMemberAxisPortIndex) As Boolean
Const MT = "IsSupportedAxisInPositiveZ"
    On Error GoTo ErrorHandler
    
    Dim dVals(16) As Double
    Dim oMat As IJDT4x4, oMatSupped As IJDT4x4
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim oLine As IJLine
    Dim uvX As Double, uvY As Double, uvZ As Double
    Dim oVec As New DVector
    Dim oPosAlongAxis As IJDPosition
    
    IsSupportedAxisInPositiveZ = False
    
    'get local coordinate system for  the supporting member


    oSupported.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    
    ' supported member line
    If portIdx = SPSMemberAxisStart Then
        oSupported.Rotation.GetTransformAtPosition sX, sY, sZ, oMatSupped, Nothing
        oVec.Set oMatSupped.IndexValue(0), oMatSupped.IndexValue(1), oMatSupped.IndexValue(2)
    Else
        oSupported.Rotation.GetTransformAtPosition eX, eY, eZ, oMatSupped, Nothing
        oVec.Set oMatSupped.IndexValue(0), oMatSupped.IndexValue(1), oMatSupped.IndexValue(2)
        oVec.[Scale] -1 'we want the direction which is away from the connection point
    End If
    
    'get local coordinate system for  the supporting member
    Set oPosAlongAxis = GetConnectionPositionOnSupping(oSupporting, oSupported, portIdx)
    
    oSupporting.Rotation.GetTransformAtPosition oPosAlongAxis.x, oPosAlongAxis.y, oPosAlongAxis.z, oMat, Nothing

    'transform supported tangent line to the supporting coordinate system
    oMat.Invert 'global to local transformation
    Set oVec = oMat.TransformVector(oVec)
    
    oVec.Get uvX, uvY, uvZ
    
    If uvZ >= 0# Then  ' if supported axis on the  supporting +ve Z  side, uvz >0.0
        IsSupportedAxisInPositiveZ = True
    End If
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function

'*************************************************************************
'Function
'
'<GetIncidentMemberQuadrant>
'
'Abstract
'
'<determines which quadrant one member lies with respect to other>
'
'Arguments
'
'<Supported and Supporting MemberPartPrismatic, SPSMemberAxisPortIndex>
'
'Return
'
'<Integer specifying the qaudrant>
'
'Exceptions
'***************************************************************************
Public Function GetIncidentMemberQuadrant(oSupporting As ISPSMemberPartPrismatic, oSupported As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex) As Integer
Const MT = "GetIncidentMemberQuadrant"
    On Error GoTo ErrorHandler
    Dim dVals(16) As Double
    Dim oMat As IJDT4x4, oMatSupped As IJDT4x4
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    Dim oVec As New DVector
    Dim uvX As Double, uvY As Double, uvZ As Double
    Dim oConnPos As IJDPosition
    
    GetIncidentMemberQuadrant = 0
    
    oSupported.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    'the supported member may be curved, so need to pass the position to get the transform
    'The first 3 values of the transform corresponds to the tangent at the passed in position
    If iEnd = SPSMemberAxisStart Then 'member aleady runs away from us
        oSupported.Rotation.GetTransformAtPosition sX, sY, sZ, oMatSupped, Nothing
        oVec.Set oMatSupped.IndexValue(0), oMatSupped.IndexValue(1), oMatSupped.IndexValue(2)
    Else
        oSupported.Rotation.GetTransformAtPosition eX, eY, eZ, oMatSupped, Nothing
        oVec.Set oMatSupped.IndexValue(0), oMatSupped.IndexValue(1), oMatSupped.IndexValue(2)
        oVec.[Scale] -1 'we want the direction which is away from the connection point
    End If
    
    'get local coordinate system for  the supporting member
    'first get the connection point on the supporting member by projection and intersection (if they don't
    'touch each other already, a rare case)
    Set oConnPos = GetConnectionPositionOnSupping(oSupporting, oSupported, iEnd)
    
    If Not oConnPos Is Nothing Then
        'get the transform at the position (ignore knuckle)
        oSupporting.Rotation.GetTransformAtPosition oConnPos.x, oConnPos.y, oConnPos.z, oMat, Nothing
        
        'transform supported tangent line to the supporting coordinate system
        oMat.Invert 'global to local transformation
        Set oVec = oMat.TransformVector(oVec)
    
        
        oVec.Get uvX, uvY, uvZ
        'CS is in members local X along axis, Y = -X and Z up
        'Quadrants are 1-4 with 1 at 12 o'clock going clockwise on 45 deg like an X
        If Abs(uvZ) >= Abs(uvY) Then
          If uvZ > 0 Then
            GetIncidentMemberQuadrant = 1
          Else
            GetIncidentMemberQuadrant = 3
          End If
        Else
          If uvY > 0 Then
            GetIncidentMemberQuadrant = 4
          Else
            GetIncidentMemberQuadrant = 2
          End If
        End If
    End If
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<IsCutbackNeededByShape>
'
'Abstract
'
'<Fucntion determines wether the member should be Coped or have cutback>
'
'Arguments
'
'<Supporting MemberPartPrismatic>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsCutbackNeededByShape(pISPSMemSupporting As ISPSMemberPartPrismatic) As Boolean
  Const MT = "IsCutbackNeededByShape"
  Dim pIJCSSupporting As IJCrossSection, IsNeeded As Boolean
  On Error GoTo ErrorHandler

  IsNeeded = True
  If Not pISPSMemSupporting Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = pISPSMemSupporting.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set pIJCSSupporting = OSPSCrossSection.definition
            If Not pIJCSSupporting Is Nothing Then
                Dim strSectionType As String
                strSectionType = pIJCSSupporting.Type
            End If
        End If
    End If
  Select Case strSectionType
    Case "HSSC", "CS", "PIPE"
      IsNeeded = False 'no Cutback for a round shape, it will be coped
    Case Else
      'Err.Raise
  End Select
  IsCutbackNeededByShape = IsNeeded
  Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<IsCope1NeededByShapeAndIncidence>
'
'Abstract
'
'<Determines a cope needed for sections having Top Flange>
'
'Arguments
'
'<SupportingMemberPartPrismatic, Quadrant as integer>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsCope1NeededByShapeAndIncidence(pISPSMemSupporting As ISPSMemberPartPrismatic, iQuadrant As Integer) As Boolean
    Const MT = "IsCope1NeededByShapeAndIncidence"
    Dim pIJCSSupporting As IJCrossSection, IsNeeded As Boolean
    On Error GoTo ErrorHandler
    
    Dim bMirror As Boolean
    
    bMirror = pISPSMemSupporting.Rotation.Mirror
    IsNeeded = False
    If Not pISPSMemSupporting Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = pISPSMemSupporting.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set pIJCSSupporting = OSPSCrossSection.definition
            If Not pIJCSSupporting Is Nothing Then
                Dim strSectionType As String
                strSectionType = pIJCSSupporting.Type
            End If
        End If
    End If
    
    Select Case strSectionType
        Case "HSSC", "CS", "PIPE"
            IsNeeded = True
        Case "W", "S", "HP", "M"
            If iQuadrant = 2 Or iQuadrant = 4 Then IsNeeded = True
        Case "L"
            IsNeeded = False
        Case "C", "MC"
            If (iQuadrant = 2) And (bMirror = False) Then
                IsNeeded = True
            ElseIf (iQuadrant = 4) And (bMirror = True) Then
                IsNeeded = True
            End If
        Case "WT", "MT", "ST"
            If iQuadrant = 2 Or iQuadrant = 3 Or iQuadrant = 4 Then IsNeeded = True
        Case "2L"
            IsNeeded = False
        Case "RS", "HSSR"
            IsNeeded = False
        Case Else
    End Select
    IsCope1NeededByShapeAndIncidence = IsNeeded
    Exit Function
ErrorHandler:        HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<IsCope2NeededByShapeAndIncidence>
'
'Abstract
'
'<Determines a cope needed for sections having Bottom Flange>
'
'Arguments
'
'<SupportingMemberPartPrismatic, Quadrant as integer>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsCope2NeededByShapeAndIncidence(pISPSMemSupporting As ISPSMemberPartPrismatic, iQuadrant As Integer) As Boolean
    Const MT = "IsCope2NeededByShapeAndIncidence"
    Dim pIJCSSupporting As IJCrossSection, IsNeeded As Boolean
    On Error GoTo ErrorHandler
    Dim bMirror As Boolean
    
    bMirror = pISPSMemSupporting.Rotation.Mirror

    IsNeeded = False
    If Not pISPSMemSupporting Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = pISPSMemSupporting.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set pIJCSSupporting = OSPSCrossSection.definition
            If Not pIJCSSupporting Is Nothing Then
                Dim strSectionType As String
                strSectionType = pIJCSSupporting.Type
            End If
        End If
    End If
    Select Case strSectionType
        Case "HSSC", "CS", "PIPE"
            IsNeeded = False
        Case "W", "S", "HP", "M"
            If iQuadrant = 2 Or iQuadrant = 4 Then IsNeeded = True
        Case "L"
            If iQuadrant = 1 Then
                IsNeeded = True
            ElseIf (iQuadrant = 2) And (bMirror = False) Then
                IsNeeded = True
            ElseIf (iQuadrant = 4) And (bMirror = True) Then
                IsNeeded = True
            End If
        Case "C", "MC"
            If (iQuadrant = 2) And (bMirror = False) Then
                IsNeeded = True
            ElseIf (iQuadrant = 4) And (bMirror = True) Then
                IsNeeded = True
            End If
        Case "WT", "MT", "ST"
            IsNeeded = False
        Case "2L"
            If iQuadrant = 1 Or iQuadrant = 2 Or iQuadrant = 4 Then IsNeeded = True
        Case "RS", "HSSR"
            IsNeeded = False
        Case Else
    End Select
    IsCope2NeededByShapeAndIncidence = IsNeeded
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<IsMemberAxesColinearAndEndToEnd>
'
'Abstract
'
'<Determines if the Member Axis lines are collinear and connected at the ends>
'
'Arguments
'
'<Supporting and Supported Member Axis as IJLine>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesColinearAndEndToEnd(oCurve1 As IJCurve, oCurve2 As IJCurve) As Boolean
  Const MT = "IsMemberAxesColinearAndEndToEnd"
    On Error GoTo ErrorHandler
    IsMemberAxesColinearAndEndToEnd = False
    
    Dim oLine1 As IJLine, oLine2 As IJLine
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim sX1 As Double, sY1 As Double, sZ1 As Double
    Dim sX2 As Double, sY2 As Double, sZ2 As Double
    Dim eX1 As Double, eY1 As Double, eZ1 As Double
    Dim eX2 As Double, eY2 As Double, eZ2 As Double
    
    Dim oVec1 As New DVector, oVec2 As New DVector
    Dim dist1 As Double, dist2 As Double, dist3 As Double, dist4 As Double
    Dim dx As Double, dy As Double, dz As Double
    Dim cosA As Double
    Dim tol As Double

    If (TypeOf oCurve1 Is IJLine) And (TypeOf oCurve2 Is IJLine) Then
        
        Set oLine1 = oCurve1
        Set oLine2 = oCurve2
        oLine1.GetDirection x1, y1, z1
        oLine2.GetDirection x2, y2, z2
        
        oVec1.Set x1, y1, z1
        oVec2.Set x2, y2, z2
        cosA = oVec1.Dot(oVec2)
        tol = Abs(cosA) - 1#
        
        If Abs(tol) < distTol Then
            oLine1.GetStartPoint sX1, sY1, sZ1
            oLine1.GetEndPoint eX1, eY1, eZ1
            oLine2.GetStartPoint sX2, sY2, sZ2
            oLine2.GetEndPoint eX2, eY2, eZ2
            
            If IsMemberAxesEndToEnd(oLine1, oLine2) Then
                oVec1.Set eX1 - sX1, eY1 - sY1, eZ1 - sZ1
                oVec2.Set eX2 - sX2, eY2 - sY2, eZ2 - sZ2
                oVec1.Length = 1
                oVec2.Length = 1
                cosA = oVec1.Dot(oVec2)
                tol = Abs(cosA) - 1#
                If Abs(tol) < distTol Then
                    GoTo Process
                End If
            End If
            
        End If
    End If
Exit Function
Process:
    IsMemberAxesColinearAndEndToEnd = True
Exit Function
ErrorHandler:    HandleError MODULE, MT
End Function
 
'*************************************************************************
'Function
'
'<IsMemberAxesColinear>
'
'Abstract
'
'<Determines if the Member Axis lines are collinear>
'
'Arguments
'
'<Supporting and Supported Member Axis as IJLine>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesColinear(oLine1 As IJLine, oLine2 As IJLine) As Boolean
  Const MT = "IsMemberAxesColinear"
  On Error GoTo ErrorHandler
'    Currently only checking with dotproduct to see if they are in same direction

    IsMemberAxesColinear = False
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim sX1 As Double, sY1 As Double, sZ1 As Double
    Dim sX2 As Double, sY2 As Double, sZ2 As Double
    Dim eX1 As Double, eY1 As Double, eZ1 As Double
    Dim eX2 As Double, eY2 As Double, eZ2 As Double
    
    Dim oVec1 As New DVector, oVec2 As New DVector
    Dim dist1 As Double, dist2 As Double, dist3 As Double, dist4 As Double
    Dim dx As Double, dy As Double, dz As Double
    Dim cosA As Double
    Dim tol As Double

        
    oLine1.GetDirection x1, y1, z1
    oLine2.GetDirection x2, y2, z2
    
    oVec1.Set x1, y1, z1
    oVec2.Set x2, y2, z2
    cosA = oVec1.Dot(oVec2)
    tol = Abs(cosA) - 1#
 
    If Abs(tol) < distTol Then
        GoTo Process
    End If

Exit Function
Process:
    IsMemberAxesColinear = True
Exit Function
ErrorHandler:    HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<IsMemberAxesEndToEnd>
'
'Abstract
'
'<Determines if the Member Axis lines are connected at the ends>
'
'Arguments
'
'<Supporting and Supported Member Axis as IJLine>>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesEndToEnd(oLine1 As IJLine, oLine2 As IJLine) As Boolean
Const MT = "IsMemberAxesEndToEnd"
On Error GoTo ErrorHandler
    IsMemberAxesEndToEnd = False
    
    If (oLine1.Length > distTol) And (oLine2.Length > distTol) Then
        Dim sX1 As Double, sY1 As Double, sZ1 As Double
        Dim sX2 As Double, sY2 As Double, sZ2 As Double
        Dim eX1 As Double, eY1 As Double, eZ1 As Double
        Dim eX2 As Double, eY2 As Double, eZ2 As Double
        Dim dx As Double, dy As Double, dz As Double
        Dim dist1 As Double
    
        oLine1.GetStartPoint sX1, sY1, sZ1
        oLine1.GetEndPoint eX1, eY1, eZ1
        oLine2.GetStartPoint sX2, sY2, sZ2
        oLine2.GetEndPoint eX2, eY2, eZ2
                    
        dx = sX1 - sX2
        dy = sY1 - sY2
        dz = sZ1 - sZ2
        dist1 = Sqr(dx * dx + dy * dy + dz * dz)
        
        If Abs(dist1) < distTol Then
            GoTo Process
        Else
            dx = sX1 - eX2
            dy = sY1 - eY2
            dz = sZ1 - eZ2
            If Abs(Sqr(dx * dx + dy * dy + dz * dz)) < distTol Then
                GoTo Process
            End If
        End If
        
        dx = eX1 - sX2
        dy = eY1 - sY2
        dz = eZ1 - sZ2
        dist1 = Sqr(dx * dx + dy * dy + dz * dz)
        If Abs(dist1) < distTol Then
            GoTo Process
        Else
            dx = eX1 - eX2
            dy = eY1 - eY2
            dz = eZ1 - eZ2
            If Abs(Sqr(dx * dx + dy * dy + dz * dz)) < distTol Then
                GoTo Process
            End If
        End If
        
    End If
    
Exit Function
Process:
    IsMemberAxesEndToEnd = True
Exit Function
ErrorHandler:    HandleError MODULE, MT
End Function
    
'*************************************************************************
'Function
'
'<IsMemberAxesAtRightAngles>
'
'Abstract
'
'<determines if MemberAxis are at right angles>
'
'Arguments
'
'<Supporting and Supported Members as IJLine>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesAtRightAngles(oLine1 As IJLine, oLine2 As IJLine) As Boolean
  Const MT = "IsMemberAxesAtRightAngles"
  On Error GoTo ErrorHandler
  IsMemberAxesAtRightAngles = False

  Dim x1 As Double, y1 As Double, z1 As Double
  Dim x2 As Double, y2 As Double, z2 As Double
  Dim oVec1 As New DVector, oVec2 As New DVector
  Dim cosA As Double
  
  
  oLine1.GetDirection x1, y1, z1
  oLine2.GetDirection x2, y2, z2
  
  oVec1.Set x1, y1, z1
  'normalize vector
  oVec1.Length = 1
  oVec2.Set x2, y2, z2
  'normalize vector
  oVec2.Length = 1
  cosA = oVec1.Dot(oVec2)
  If Abs(cosA) < distTol Then
    IsMemberAxesAtRightAngles = True
  End If
  
  Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetMiterPlanePosition>
'
'Abstract
'
'<Determines the Plane Position for placing Miter connection>
'
'Arguments
'
'<Supported 1 and 2 SplitAxisPort>
'
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetMiterPlanePosition(oSupped1Port As ISPSSplitAxisPort, oSupped2Port As ISPSSplitAxisPort) As IJDPosition
  Const MT = "GetMiterPlanePosition"
  On Error GoTo ErrorHandler

  Dim oGeom3DFactory As New GeometryFactory
  Dim x1 As Double, y1 As Double, z1 As Double
  Dim X0 As Double, Y0 As Double, Z0 As Double
  Dim i As Integer, j As Integer
  Dim oPlanePos(1 To 4) As IJDPosition
  Dim oPlaneNorm(1 To 4) As IJDVector
  Dim oCPLine(1 To 4) As IJLine
  Dim oCPLine1(1 To 4) As IJLine
  Dim oPlane(1 To 4) As IJPlane
  Dim oSurf As IJSurface
  Dim colIntrsectnElms As IJElements
  Dim code As Geom3dIntersectConstants
  Dim intersectionPoint As IJPoint
  Dim oPos As New DPosition, oPos1 As New DPosition, oPos2 As New DPosition
  Dim length1#, length2#
  Dim oMat1 As IJDT4x4, oMat2 As IJDT4x4
  Dim dist#, dist1#, dist2#, maxDist#
  Dim x2#, y2#, z2#
  Dim oMiterPlanePos As New DPosition
  Dim oCmplx As ComplexString3d
  Dim curveElms As IJElements
  Dim oLine3d(1 To 4) As Line3d
  Dim idx As Integer, nextI As Integer
  Dim oCPLinePos(1 To 4) As New DPosition
  Dim sX1#, sY1#, sZ1#, sX2#, sY2#, sZ2#, eX1#, eY1#, eZ1#, eX2#, eY2#, eZ2#
  Dim uvX#, uvY#, uvZ#
  Dim bSupped1Start As Boolean, bSupped2Start As Boolean
  
  oSupped1Port.Part.Axis.EndPoints sX1, sY1, sZ1, eX1, eY1, eZ1
  
  'initialize the miterplane position to the port location
  If oSupped1Port.portIndex = SPSMemberAxisStart Then
    oMiterPlanePos.Set sX1, sY1, sZ1
    bSupped1Start = True
  Else
    oMiterPlanePos.Set eX1, eY1, eZ1
  End If
      
  'initialize return value to oMiterPlanePos
  
  Set GetMiterPlanePosition = oMiterPlanePos
  
  If oSupped2Port.portIndex = SPSMemberAxisStart Then
    bSupped2Start = True
  End If
  
  
  For i = 1 To 4
    'get bounding side planes (based on cross section range box) of supped1
    Set oPlanePos(i) = GetSectionRangeBoxSidePos(oSupped1Port.Part, i, oMiterPlanePos)
    Set oPlaneNorm(i) = GetMembSidePlaneNormal(oSupped1Port.Part, i, oMiterPlanePos)
    Set oPlane(i) = oGeom3DFactory.Planes3d.CreateByPointNormal(Nothing, oPlanePos(i).x, oPlanePos(i).y, oPlanePos(i).z, _
                    oPlaneNorm(i).x, oPlaneNorm(i).y, oPlaneNorm(i).z)
  Next i
  'get  lines (based on range box) for supped2
  Set oCPLine(1) = GetTangentLineAtCP(oSupped2Port.Part, 7, bSupped2Start)
  Set oCPLine(2) = GetTangentLineAtCP(oSupped2Port.Part, 9, bSupped2Start)
  Set oCPLine(3) = GetTangentLineAtCP(oSupped2Port.Part, 3, bSupped2Start)
  Set oCPLine(4) = GetTangentLineAtCP(oSupped2Port.Part, 1, bSupped2Start)
 
  'get lines for (based on range box) supped1
  Set oCPLine1(1) = GetTangentLineAtCP(oSupped1Port.Part, 7, bSupped1Start)
  Set oCPLine1(2) = GetTangentLineAtCP(oSupped1Port.Part, 9, bSupped1Start)
  Set oCPLine1(3) = GetTangentLineAtCP(oSupped1Port.Part, 3, bSupped1Start)
  Set oCPLine1(4) = GetTangentLineAtCP(oSupped1Port.Part, 1, bSupped1Start)
 
 
 'Add boundaries for side planes of supped1

  Set curveElms = New JObjectCollection
  
  For i = 1 To 4
      nextI = i Mod 4
      oCPLine1(i).GetStartPoint sX1, sY1, sZ1
      oCPLine1(i).GetDirection uvX, uvY, uvZ
      oCPLine1(i).GetEndPoint eX1, eY1, eZ1
      oCPLine1(nextI + 1).GetStartPoint sX2, sY2, sZ2
      oCPLine1(nextI + 1).GetEndPoint eX2, eY2, eZ2
      'make the lines really long so that the CP lines of supped2 intersects the supped1 plane
      'bounded by these lines
      uvX = 100 * uvX
      uvY = 100 * uvY
      uvZ = 100 * uvZ
      oCPLinePos(1).Set sX1 - uvX, sY1 - uvY, sZ1 - uvZ
      oCPLinePos(2).Set eX1 + uvX, eY1 + uvY, eZ1 + uvZ
      oCPLinePos(3).Set eX2 + uvX, eY2 + uvY, eZ2 + uvZ
      oCPLinePos(4).Set sX2 - uvX, sY2 - uvY, sZ2 - uvZ
      For idx = 1 To 4
        Dim nextIdx As Integer
        nextIdx = idx Mod 4
        Set oLine3d(idx) = oGeom3DFactory.Lines3d.CreateBy2Points(Nothing, oCPLinePos(idx).x, oCPLinePos(idx).y, oCPLinePos(idx).z, _
        oCPLinePos(nextIdx + 1).x, oCPLinePos(nextIdx + 1).y, oCPLinePos(nextIdx + 1).z)
        
        curveElms.Add oLine3d(idx)
        Set oLine3d(idx) = Nothing
      Next idx
      Set oCmplx = oGeom3DFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
      oPlane(i).AddBoundary oCmplx
      Set oCmplx = Nothing
      curveElms.Clear
  Next i
              
    oSupped1Port.Part.Rotation.GetTransformAtPosition oMiterPlanePos.x, oMiterPlanePos.y, oMiterPlanePos.z, oMat1, Nothing
    'global to local
    oMat1.Invert
    oSupped2Port.Part.PointAtEnd(oSupped2Port.portIndex).GetPoint sX2, sY2, sZ2
    oSupped2Port.Part.Rotation.GetTransformAtPosition sX2, sY2, sZ2, oMat2, Nothing
    'global to local
    oMat2.Invert
      
  

    maxDist = 0
    'for each supped2 line
    For i = 1 To 4
      length1 = oCPLine(i).Length
      length2 = oCPLine1(i).Length
      oCPLine(i).Infinite = True
      'for each supped1 surface
      For j = 1 To 4
          dist = 0
          dist1 = 0
          dist2 = 0
          Set oSurf = oPlane(j)
          'check line surface intersection
          oSurf.Intersect oCPLine(i), colIntrsectnElms, code
          ' if intersects
          If (Not colIntrsectnElms Is Nothing) And (code = ISECT_UNKNOWN) Then
            If colIntrsectnElms.Count > 0 Then
                ' get intersection point
                Set intersectionPoint = colIntrsectnElms.Item(1)
                intersectionPoint.GetPoint X0, Y0, Z0
                oPos.Set X0, Y0, Z0
                ' get dist1, from the other end of supped1
                Set oPos1 = oMat1.TransformPosition(oPos)
                oPos1.Get x1, y1, z1
                If oSupped1Port.portIndex = SPSMemberAxisEnd Then
                    dist1 = length1 + x1
                Else
                    dist1 = length1 - x1
                End If
                'get dist2, from the other end of supped2
                Set oPos2 = oMat2.TransformPosition(oPos)
                oPos2.Get x2, y2, z2
                If oSupped2Port.portIndex = SPSMemberAxisEnd Then
                    dist2 = length2 + x2
                Else
                    dist2 = length2 - x2
                End If
                'get sum of distances from other ends of supped1 and supped2
                dist = dist1 + dist2
                If dist > maxDist Then
                'intersection point with max dist is the miter position
                   maxDist = dist
                   oMiterPlanePos.Set X0, Y0, Z0
                End If
            End If
          End If
      Next j
    Next i
    
    Set GetMiterPlanePosition = oMiterPlanePos
  
  For i = 1 To 4
    Set oPlanePos(i) = Nothing
    Set oPlaneNorm(i) = Nothing
    Set oPlane(i) = Nothing
    Set oCPLine(i) = Nothing
    Set oCPLine1(i) = Nothing
    Set oCPLinePos(i) = Nothing
  Next i
  Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetMiterPlaneNormal>
'
'Abstract
'
'<calculates the normal of the Miter Plane>
'
'Arguments
'
'<Supported 1 and 2 ports as SplitAxisPort>
'
'Return
'
'vector as IJDVector
'
'Exceptions
'***************************************************************************
Public Function GetMiterPlaneNormal(oSupped1Port As ISPSSplitAxisPort, oSupped2Port As ISPSSplitAxisPort) As IJDVector
  Const MT = "GetMiterPlaneNormal"
  On Error GoTo ErrorHandler

  Dim oLine1 As IJLine, oLine2 As IJLine
  Dim oGeom3DFactory As New GeometryFactory
  Dim oBisector As IJDVector
  Dim oVec1 As IJDVector, oVec2 As IJDVector
  Dim oVec As IJDVector, oNorm As IJDVector
  Dim x1 As Double, y1 As Double, z1 As Double
  Dim X0 As Double, Y0 As Double, Z0 As Double
  Dim mat As IJDT4x4
  
  oSupped1Port.Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1

  
  If oSupped1Port.portIndex = SPSMemberAxisEnd Then
  
      oSupped1Port.Part.Rotation.GetTransformAtPosition x1, y1, z1, mat, Nothing
      Set oLine1 = oGeom3DFactory.Lines3d.CreateByPtVectLength(Nothing, 0, 0, 0, -mat.IndexValue(0), _
      -mat.IndexValue(1), -mat.IndexValue(2), 100)
  Else
      oSupped1Port.Part.Rotation.GetTransformAtPosition X0, Y0, Z0, mat, Nothing
      Set oLine1 = oGeom3DFactory.Lines3d.CreateByPtVectLength(Nothing, 0, 0, 0, mat.IndexValue(0), _
      mat.IndexValue(1), mat.IndexValue(2), 100)

  End If
  

  
  oSupped2Port.Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1

  If oSupped2Port.portIndex = SPSMemberAxisEnd Then
      oSupped2Port.Part.Rotation.GetTransformAtPosition x1, y1, z1, mat, Nothing
      Set oLine2 = oGeom3DFactory.Lines3d.CreateByPtVectLength(Nothing, 0, 0, 0, -mat.IndexValue(0), _
      -mat.IndexValue(1), -mat.IndexValue(2), 100)
  Else
      oSupped2Port.Part.Rotation.GetTransformAtPosition X0, Y0, Z0, mat, Nothing
      Set oLine2 = oGeom3DFactory.Lines3d.CreateByPtVectLength(Nothing, 0, 0, 0, mat.IndexValue(0), _
      mat.IndexValue(1), mat.IndexValue(2), 100)
  End If
  
  Set oBisector = GetBisector(oLine1, oLine2)
  Set oVec1 = New DVector
  Set oVec2 = New DVector
  
  oLine1.GetDirection X0, Y0, Z0
  oVec1.Set X0, Y0, Z0
  oVec1.Length = 1#
  oLine2.GetDirection x1, y1, z1
  oVec2.Set x1, y1, z1
  oVec2.Length = 1#
  Set oVec = oVec1.Cross(oVec2)
  oVec.Length = 1#
  Set oNorm = oBisector.Cross(oVec)
  oNorm.Length = 1#
  Set GetMiterPlaneNormal = oNorm
  Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function

'*************************************************************************
'Function
'
'<GetBisector>
'
'Abstract
'
'<calculates the Bisector line of given 2 lines>
'
'Arguments
'
'<Lines 1 and 2 as IJLine>
'
'Return
'
'<vector as IJDVector>
'
'Exceptions
'***************************************************************************
Public Function GetBisector(oLine1 As IJLine, oLine2 As IJLine) As IJDVector
  Const MT = "GetBisector"
  On Error GoTo ErrorHandler
   
   Dim oPos0 As IJDPosition, oPos1 As IJDPosition
   Dim oVec00 As IJDVector, oVec01 As IJDVector, oVec10 As IJDVector, oVec11 As IJDVector
   Dim oVec0 As IJDVector, oVec1 As IJDVector, oVec As IJDVector
   Dim dLength00 As Double, dLength01 As Double, dLength10 As Double, dLength11 As Double
   Dim dLength0 As Double, dLength1 As Double, dLength As Double
   Dim sX0 As Double, sY0 As Double, sZ0 As Double, sX1 As Double, sY1 As Double, sZ1 As Double
   Dim eX0 As Double, eY0 As Double, eZ0 As Double, eX1 As Double, eY1 As Double, eZ1 As Double
   Dim oBisector As IJDVector
   
   Set oVec00 = New DVector
   Set oVec01 = New DVector
   Set oVec10 = New DVector
   Set oVec11 = New DVector
   
   
   Set oPos0 = New DPosition
   Set oPos1 = New DPosition

    oLine1.GetStartPoint sX0, sY0, sZ0
    oLine1.GetEndPoint eX0, eY0, eZ0
    oLine2.GetStartPoint sX1, sY1, sZ1
    oLine2.GetEndPoint eX1, eY1, eZ1
    
    oVec00.Set sX0, sY0, sZ0
    oVec01.Set eX0, eY0, eZ0
    oVec10.Set sX1, sY1, sZ1
    oVec11.Set eX1, eY1, eZ1
  
    dLength00 = oVec00.Length
    dLength01 = oVec01.Length
    dLength10 = oVec10.Length
    dLength11 = oVec11.Length
   
    If dLength00 > dLength01 Then
        Set oVec0 = oVec00
        oPos0.Set sX0, sY0, sZ0
    Else
        Set oVec0 = oVec01
        oPos0.Set eX0, eY0, eZ0
    End If
    If dLength10 > dLength11 Then
        Set oVec1 = oVec10
        oPos1.Set sX1, sY1, sZ1
    Else
        Set oVec1 = oVec11
        oPos1.Set eX1, eY1, eZ1
    End If
   dLength0 = oVec0.Length
   dLength1 = oVec1.Length
   If dLength0 > dLength1 Then
       oVec1.Length = 1#
       oVec1.Length = dLength0
   Else
       oVec0.Length = 1#
       oVec0.Length = dLength1
   End If
   
   Set oVec = oVec1.Subtract(oVec0)
   dLength = oVec.Length
   oVec.Length = 1#
   dLength = dLength / 2#
   oVec.Length = dLength
   Set oBisector = oVec0.Add(oVec)
   oBisector.Length = 1#
   Set GetBisector = oBisector
  Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function


'*************************************************************************
'Function
'
'<IsMemberAxesParallel>
'
'Abstract
'
'<Determines if MemberAxes are Parallel>
'
'Arguments
'
'<2 MemberAxes as IJLine>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesParallel(oCurve1 As IJCurve, oCurve2 As IJCurve) As Boolean
  Const MT = "IsMemberAxesParallel"
  On Error GoTo ErrorHandler
  IsMemberAxesParallel = False
  Dim oLine1 As IJLine, oLine2 As IJLine
  Dim x1 As Double, y1 As Double, z1 As Double
  Dim x2 As Double, y2 As Double, z2 As Double
  Dim oVec1 As IJDVector, oVec2 As IJDVector
  Dim cosA As Double
  Dim tol As Double
  Set oVec1 = New DVector
  Set oVec2 = New DVector
  
  If (TypeOf oCurve1 Is IJLine) And (TypeOf oCurve2 Is IJLine) Then
    Set oLine1 = oCurve1
    Set oLine2 = oCurve2

    oLine1.GetDirection x1, y1, z1
    oLine2.GetDirection x2, y2, z2
    
    oVec1.Set x1, y1, z1
    oVec2.Set x2, y2, z2
    cosA = oVec1.Dot(oVec2)
    tol = Abs(cosA) - 1#
    If Abs(tol) < distTol Then
      IsMemberAxesParallel = True
    End If
  End If
  Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetBoundingRectangleForMemberEnd>
'
'Abstract
'
'<determines the bounding rectangle of the Member End>
'
'Arguments
'
'<MemberPartPrismatic,port as SPSMemberAxisPortIndex,FacePos as IJDPosition>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub GetBoundingRectangleForMemberEnd(oMemb As ISPSMemberPartPrismatic, iPort As SPSMemberAxisPortIndex, oFacePos() As IJDPosition)
  Const MT = "GetBoundingRectangleForMemberEnd"
    On Error GoTo ErrorHandler
    
    Dim oPos As IJDPosition
    Dim oProfileBO As ISPSCrossSection
    Dim yOffset As Double, zOffset As Double, yOffsetCP As Double, zOffsetCP As Double
    Dim oMat As IJDT4x4
    Dim eX#, eY#, eZ#, sX#, sY#, sZ#, dblLength#
    Dim oVec As New DVector
    Dim cp As Long
    
    Set oPos = New DPosition
    
    oMemb.Rotation.GetTransform oMat 'Get the member coordinate system
    Set oProfileBO = oMemb.CrossSection
    If iPort = SPSMemberAxisEnd Then
        oMemb.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
        oVec.Set eX - sX, eY - sY, eZ - sZ
        dblLength = oVec.Length
    Else
        dblLength = 0#
    End If
    cp = oProfileBO.CardinalPoint
    oProfileBO.GetCardinalPointOffset cp, yOffsetCP, zOffsetCP 'Returns Rad x and y of active CP which is member local -y and z.
    
    oProfileBO.GetCardinalPointOffset 1, yOffset, zOffset 'Returns Rad x and y of CP=1 which is member local -y and z.
    yOffset = yOffset - yOffsetCP ' get the y offset from current CP
    zOffset = zOffset - zOffsetCP ' get the z offset from current CP
    oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
    Set oFacePos(0) = oMat.TransformPosition(oPos)
    
    oProfileBO.GetCardinalPointOffset 3, yOffset, zOffset
    yOffset = yOffset - yOffsetCP ' get the y offset from current CP
    zOffset = zOffset - zOffsetCP ' get the z offset from current CP
    oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
    Set oFacePos(1) = oMat.TransformPosition(oPos)
    
    oProfileBO.GetCardinalPointOffset 9, yOffset, zOffset
    yOffset = yOffset - yOffsetCP ' get the y offset from current CP
    zOffset = zOffset - zOffsetCP ' get the z offset from current CP
    oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
    Set oFacePos(2) = oMat.TransformPosition(oPos)
    
    oProfileBO.GetCardinalPointOffset 7, yOffset, zOffset
    yOffset = yOffset - yOffsetCP ' get the y offset from current CP
    zOffset = zOffset - zOffsetCP ' get the z offset from current CP
    oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
    Set oFacePos(3) = oMat.TransformPosition(oPos)
    
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
    
End Sub
'*************************************************************************
'Function
'
'<IsFaceOnThePositiveSideOfPlane>
'
'Abstract
'
'<determines if the face is on the positive side of the plane>
'
'Arguments
'
'<X,Y,Z Positions as double, nX,nY,nZ normals as doubles, position of plane as IJDPosition
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsFaceOnThePositiveSideOfPlane(oX As Double, oY As Double, oZ As Double, nx As Double, ny As Double, nz As Double, oPos() As IJDPosition) As Boolean
  Const MT = "IsFaceOnThePositiveSideOfPlane"
    On Error GoTo ErrorHandler
    Dim oPlnNormal As New DVector, oVec As New DVector
    Dim cosA As Double
    Dim i As Integer
    IsFaceOnThePositiveSideOfPlane = False
      
    oPlnNormal.Set nx, ny, nz
    oPlnNormal.Length = 1# 'in case it was not normalized
    If (UBound(oPos) - LBound(oPos) + 1 <> 4) Then 'we need exactly 4 points
        Exit Function
    End If
    For i = LBound(oPos) To UBound(oPos)
        oVec.Set oPos(i).x - oX, oPos(i).y - oY, oPos(i).z - oZ 'vector from plane root point to first face point
        If oVec.Length <> 0# Then ' to prevent divide by 0
            oVec.Length = 1# ' normalize the vector
        End If
        cosA = oPlnNormal.Dot(oVec)
        If cosA <= 0# Then ' the point is not on the positive side of the plane,, so the face is not on the positive side
            Exit Function
        End If
    Next i
    IsFaceOnThePositiveSideOfPlane = True
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function

'*************************************************************************
'Function
'
'<IsMemberAxesParallelToPlane>
'
'Abstract
'
'<Determines if the MemberAxes is parallel to a Plane >
'
'Arguments
'
'<MemberAxes as IJLine, plane as IJPlane>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsMemberAxesParallelToPlane(oLine As IJLine, oPlane As IJPlane) As Boolean
  Const MT = "IsMemberAxesParallelToPlane"
  On Error GoTo ErrorHandler
  IsMemberAxesParallelToPlane = False
  

  Dim x1 As Double, y1 As Double, z1 As Double
  Dim x2 As Double, y2 As Double, z2 As Double
  Dim oVec1 As IJDVector, oVec2 As IJDVector
  Dim cosA As Double
  Set oVec1 = New DVector
  Set oVec2 = New DVector
  

  oLine.GetDirection x1, y1, z1
  
  oVec1.Set x1, y1, z1
  oVec1.Length = 1# ' normalize the vector
  oPlane.GetNormal x2, y2, z2 ' get plane normal
  oVec2.Set x2, y2, z2
  oVec2.Length = 1# ' normalize the vector
  cosA = oVec1.Dot(oVec2)
  If Abs(cosA) < distTol Then
    IsMemberAxesParallelToPlane = True
  End If
  
  Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'The returned matrix would transform a point in RAD coord sys to the global coord sys
'*************************************************************************
'Function
'
'<CreateCSToMembTransform>
'
'Abstract
'
'<Creates a IJDT4x4 transformation from 2D Cross section to 3D Member coordinate system>
'
'Arguments
'
'<MemberTransformation Matrix as IJDT4x4, reflect as Boolean>
'
'Return
'
'<Transformation matrix IJDT4x4 >
'
'Exceptions
'***************************************************************************
Public Function CreateCSToMembTransform(pI4x4Mem As IJDT4x4, Optional reflect As Boolean = False) As IJDT4x4
 Const MT = "CreateCSToMembTransform"
On Error GoTo ErrorHandler

  Dim tmpI4x4 As New DT4x4
  Dim pI4x4 As New DT4x4


  tmpI4x4.LoadMatrix pI4x4Mem
  pI4x4.LoadIdentity
  'loads 0 0 -1 0 to get a local 2d CS flipped into member local CS where x along axis, Z up Y perp web
  '     -1 0  0 0
  '      0 1  0 0
  '      0 0  0 1
  'Means local x = Member -Y
  '      local y = Member Z
  '      local z = Member -X
  
  pI4x4.IndexValue(0) = 0
  If reflect = True Then
    pI4x4.IndexValue(1) = 1
  Else
    pI4x4.IndexValue(1) = -1
  End If
  pI4x4.IndexValue(5) = 0
  pI4x4.IndexValue(6) = 1
  pI4x4.IndexValue(8) = -1
  pI4x4.IndexValue(10) = 0
  
  tmpI4x4.MultMatrix pI4x4
  Set CreateCSToMembTransform = tmpI4x4
  Set pI4x4 = Nothing
Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function

'The returned matrix would transform a point in RAD coord sys to the member coord sys
'*************************************************************************
'Function
'
'<CreateCSToMembLocalTransform>
'
'Abstract
'
'<Creates a IJDT4x4 transformation from 2D Cross section to 3D Member coordinate system without Reflect Option>
'
'Arguments
'
'<MemberTransformation Matrix as IJDT4x4, reflect as Boolean>
'
'Return
'
'<Transformation matrix IJDT4x4 >
'
'Exceptions
'***************************************************************************
Public Function CreateCSToMembLocalTransform(Optional reflect As Boolean = False) As IJDT4x4
 Const MT = "CreateCSToMembLocalTransform"
On Error GoTo ErrorHandler


  Dim pI4x4 As New DT4x4
  pI4x4.LoadIdentity
  'loads 0 0 -1 0 to get a local 2d CS flipped into member local CS where x along axis, Z up Y perp web
  '     -1 0  0 0
  '      0 1  0 0
  '      0 0  0 1
  'Means local x = Member -Y
  '      local y = Member Z
  '      local z = Member -X
  
  pI4x4.IndexValue(0) = 0
  
  If reflect = True Then
    pI4x4.IndexValue(1) = 1
  Else
    pI4x4.IndexValue(1) = -1
  End If
 
  pI4x4.IndexValue(5) = 0
  pI4x4.IndexValue(6) = 1
  pI4x4.IndexValue(8) = -1
  pI4x4.IndexValue(10) = 0
  
  Set CreateCSToMembLocalTransform = pI4x4
  Set pI4x4 = Nothing
Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<ConnectSmartOccurrence>
'
'Abstract
'
'<Adds the object to the ReferencesCollection Object>
'
'Arguments
'
'<SmartOccurrence as Object, refcoll object As IJDReferencesCollection>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub ConnectSmartOccurrence(pSO As IJSmartOccurrence, pRefColl As IJDReferencesCollection)
Const MT = "ConnectSmartOccurrence"
 On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pSO
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    pCollectionHelper.Add pRefColl, "RC", pRelationshipHelper
    Set pRevision = New JRevision
    pRevision.AddRelationship pRelationshipHelper
  
 Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'
'<GetRefCollFromSmartOccurrence>
'
'Abstract
'
'<Gets the object from the ReferencesCollection Object>
'
'Arguments
'
'<SmartOccurrence as Object>
'
'Return
'
'<refcoll object As IJDReferencesCollection>
'
'Exceptions
'***************************************************************************
Public Function GetRefCollFromSmartOccurrence(pSO As IJSmartOccurrence) As IJDReferencesCollection
Const MT = "GetRefCollFromSmartOccurrence"
 On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pSO
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    If Not pCollectionHelper Is Nothing Then
        If pCollectionHelper.Count = 1 Then
            Set GetRefCollFromSmartOccurrence = pCollectionHelper.Item("RC")
        End If
    End If
  
 Exit Function
ErrorHandler: HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<IsSOOverridden>
'
'Abstract
'
'<Checks if the occurrence attributes are Empty>
'
'Arguments
'
'<attrCol as CollectionProxy is the collection of attributes on a Interface>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsSOOverridden(attrCol As CollectionProxy) As Boolean
Const MT = "IsSOOverridden"
On Error GoTo ErrorHandler
Dim i As Integer
Dim oAttr As IJDAttribute
  IsSOOverridden = False
  For i = 1 To attrCol.Count
    Set oAttr = attrCol.Item(i)
    If Not IsEmpty(oAttr.Value) Then
        IsSOOverridden = True
        Set oAttr = Nothing
        Exit For
    End If
    Set oAttr = Nothing
  Next i
 
 Exit Function
ErrorHandler: HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<CopyValuesToSOFromItem>
'
'Abstract
'
'<Copies defintion values to occurrence values>
'
'Arguments
'
'<occurrence and defintion collection as CollectionProxy is the collection of attributes on a Interface>>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub CopyValuesToSOFromItem(soCol As CollectionProxy, itmCol As CollectionProxy)
Const MT = "CopyValuesToSOFromItem"
On Error GoTo ErrorHandler
    Dim i As Integer

    If soCol.Count <> itmCol.Count Then
        Exit Sub
    End If
    For i = 1 To soCol.Count
        If soCol.Item(i).AttributeInfo.Name = itmCol.Item(i).AttributeInfo.Name Then
            soCol.Item(i).Value = itmCol.Item(i).Value
        End If
    Next i
     
    Exit Sub

ErrorHandler: HandleError MODULE, MT

End Sub

'*************************************************************************
'Function
'
'<GetCPLine>
'
'Abstract
'
'<Gets the line of the MemberPartPrismatic given a cardinal point>
'
'Arguments
'
'<Member part as MemberPartPrismatic and Cardinal point as integer>
'
'Return
'
'<line as IJLine>
'
'Exceptions
'***************************************************************************
Public Function GetCPLine(oMemb As ISPSMemberPartPrismatic, cp As Long) As IJLine
  Const MT = "GetCPLine"
    On Error GoTo ErrorHandler
    Dim curCP As Long
    Dim oProfileBO As ISPSCrossSection
    Dim pGeometryFactory As New GeometryFactory
    Dim oPos As New DPosition
    Dim oMat As IJDT4x4
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double
    
    Set oProfileBO = oMemb.CrossSection
    curCP = oProfileBO.CardinalPoint
    
    oProfileBO.GetCardinalPointOffset curCP, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP
    oProfileBO.GetCardinalPointOffset cp, xOffset, yOffset 'Returns Rad x and y of CP requested
    oPos.x = xOffset - xOffsetCP ' x offset of current cp from requested CP
    oPos.y = yOffset - yOffsetCP ' y offset of current cp from requested CP
    oPos.z = 0
    
    oMemb.Rotation.GetTransform oMat
    oMat.IndexValue(12) = 0
    oMat.IndexValue(13) = 0
    oMat.IndexValue(14) = 0
    Set oMat = CreateCSToMembTransform(oMat, oMemb.Rotation.Mirror)
    Set oPos = oMat.TransformPosition(oPos)
    oMemb.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    sX = sX + oPos.x
    sY = sY + oPos.y
    sZ = sZ + oPos.z
    eX = eX + oPos.x
    eY = eY + oPos.y
    eZ = eZ + oPos.z

    
    Set GetCPLine = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, sX, sY, sZ, eX, eY, eZ)
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT
    
End Function

'*************************************************************************
'Function
'
'<PerpDistOfPosFromPlane>
'
'Abstract
'
'<Calculates the perpendicular distance of a position from the plane>
'
'Arguments
'
'<position as IJPlane , rootpoint of plane as IJDPosition, planenormal as IJDVector>
'
'Return
'
'<distance as Double>
'
'Exceptions
'***************************************************************************
Public Function PerpDistOfPosFromPlane(oPos As IJDPosition, oPlaneRoot As IJDPosition, oPlaneNormal As IJDVector) As Double
  Const MT = "PerpDistOfPosFromPlane"
    On Error GoTo ErrorHandler
    Dim oVec As New DVector
    Dim dist#
    oVec.Set oPos.x - oPlaneRoot.x, oPos.y - oPlaneRoot.y, oPos.z - oPlaneRoot.z
    'normalize the vector, if not done before passing in
    oPlaneNormal.Length = 1
    'get the projected length of oVec along plane normal
    dist = oVec.Dot(oPlaneNormal)
    PerpDistOfPosFromPlane = Abs(dist)
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function



'*************************************************************************
'Function
'
'<IsSuppedSeatedOnSupping>
'
'Abstract
'
'<Determines if a Supported Member is Seated on Supporting Member>
'
'Arguments
'
'<Supported and Supporting objects as MemberPartPrismatic>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsSuppedSeatedOnSupping(oSupped As ISPSMemberPartPrismatic, oSupping As ISPSMemberPartPrismatic) As Boolean
  Const MT = "IsSuppedSeatedOnSupping"
    On Error GoTo ErrorHandler
    Dim suppingCP As Long
    Dim oProfileBO As ISPSCrossSection
    Dim pGeometryFactory As New GeometryFactory
    Dim oPos As New DPosition
    Dim oMat As IJDT4x4
    Dim sX#, sY#, sZ#
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double
    Dim oLine(1 To 4) As IJLine
    Dim idx As Long
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, Width As Double
    Dim halfDepth As Double, halfWidth As Double
    Dim dist As Double
    Dim diff As Double
    Dim bSeated As Boolean
    Dim bInXZPlane As Boolean, bInXYPlane As Boolean
    Dim bAllNegativeSide As Boolean ' flag to check on which side of supporting is supported
    Dim bNegativeSide As Boolean ' flag to check on which side of supporting the current point is
    
    
    IsSuppedSeatedOnSupping = False
    
    If (TypeOf oSupped Is ISPSMemberPartLinear) And (TypeOf oSupping Is ISPSMemberPartLinear) Then
        'get CP1 line of supported
        Set oLine(1) = GetCPLine(oSupped, 1)
        'get CP3 line of supported
        Set oLine(2) = GetCPLine(oSupped, 3)
        'get CP7 line of supported
        Set oLine(3) = GetCPLine(oSupped, 7)
        'get CP9 line of supported
        Set oLine(4) = GetCPLine(oSupped, 9)
        
        Set oProfileBO = oSupping.CrossSection
        suppingCP = oProfileBO.CardinalPoint
        Set oSectionAttrbs = oProfileBO.Definition
        
        depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
        Width = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
        
        halfDepth = depth / 2
        halfWidth = Width / 2
        
        oProfileBO.GetCardinalPointOffset suppingCP, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP
        oProfileBO.GetCardinalPointOffset 5, xOffset, yOffset 'Returns Rad x and y of CP5
        oPos.x = xOffsetCP - xOffset ' x offset of current cp from  CP5
        oPos.y = yOffsetCP - yOffset ' y offset of current cp from CP5
        oPos.z = 0
        
        Set oMat = CreateCSToMembLocalTransform(oSupping.Rotation.Mirror)
        'CP5 offset in member local coordinates
        Set oPos = oMat.TransformPosition(oPos)
        oSupping.Rotation.GetTransform oMat
        oMat.Invert
        bSeated = True
        bInXZPlane = IsSuppedAxisInSuppingXZPlane(oSupped, oSupping)
        bInXYPlane = IsSuppedAxisInSuppingXYPlane(oSupped, oSupping)
        If bInXZPlane Or bInXYPlane Then
            For idx = 1 To 4
                'from global to supporting coord system
                oLine(idx).Transform oMat
                oLine(idx).GetStartPoint sX, sY, sZ
                'get the endpoints with CP5 as origin
                If bInXZPlane Then
                    'since Supped Axis is in Supping XZ Plane oPos.Y gives the perpendicular dist to line from supping CP5
                    dist = oPos.y + sY
                    diff = Abs(dist) - (halfWidth - distTol) ' minus distTol needs when dist and halfWidth are almost equal. We don't want the diff to be negative in this case
                ElseIf bInXYPlane Then
                    'since Supped Axis is in Supping XY Plane oPos.Z gives the perpendicular dist to line from supping CP5
                    dist = oPos.z + sZ
                    diff = Abs(dist) - (halfDepth - distTol) ' minus distTol needs when dist and halfDepth are almost equal. We don't want the diff to be negative in this case
                End If
                If idx = 1 Then ' initialize bAllNegativeSide
                    If dist < 0 Then
                        bAllNegativeSide = True
                    End If
                End If
                bNegativeSide = False ' initialize bNegativeSide
                If dist < 0 Then ' check if current point is on the negative side
                    bNegativeSide = True
                End If
                            
                If (diff < 0) Or (bAllNegativeSide <> bNegativeSide) Then ' the (dist  < halfdepth or Halfwidth) or the points are on either side
                    bSeated = False
                    Exit For
                End If
            Next idx
            If bSeated Then
                IsSuppedSeatedOnSupping = True
            End If
        End If
    End If
    
    For idx = 1 To 4
            Set oLine(idx) = Nothing
    Next idx
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<GetAxesPlaneNormal>
'
'Abstract
'
'<Calculates the Normal vector of the plane containing two MemberAxes>
'
'Arguments
'
'<2 Lines as IJLines>
'
'Return
'
'<Vector as IJDVector>
'
'Exceptions
'***************************************************************************
Public Function GetAxesPlaneNormal(oLine1 As IJLine, oLine2 As IJLine) As IJDVector
  Const MT = "GetAxesPlaneNormal"
    On Error GoTo ErrorHandler
    Dim nx#, ny#, nz#
    Dim oVec1 As New DVector, oVec2 As New DVector, oVec3 As IJDVector
    
    oLine1.GetDirection nx, ny, nz
    oVec1.Set nx, ny, nz
    oVec1.Length = 1
    oLine2.GetDirection nx, ny, nz
    oVec2.Set nx, ny, nz
    oVec2.Length = 1
    On Error Resume Next
    'Cross product is undefined when ovec1 and ovec2 are parallel. This may be an error
    Set oVec3 = oVec1.Cross(oVec2)
    On Error GoTo ErrorHandler
    If Not oVec3 Is Nothing Then
        oVec3.Length = 1
        Set GetAxesPlaneNormal = oVec3
    End If
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<ProjectPosToPlane>
'
'Abstract
'
'<Projects given position on to a plane to return the position on that plane>
'
'Arguments
'
'<position as IJDPostion,rootposition of plane as IJDPosition, normal as IJDVector>
'
'Return
'
'<position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function ProjectPosToPlane(oPos As IJDPosition, oPlaneRoot As IJDPosition, oPlaneNormal As IJDVector) As IJDPosition
  Const MT = "ProjectPosToPlane"
    On Error GoTo ErrorHandler
    Dim oVec As New DVector
    Dim oProjectedPos As New DPosition
    Dim dist#
    oVec.Set oPlaneRoot.x - oPos.x, oPlaneRoot.y - oPos.y, oPlaneRoot.z - oPos.z
    'normalize the vector, if not done before passing in
    oPlaneNormal.Length = 1
    'get the projected length of oVec along plane normal
    dist = oPlaneNormal.Dot(oVec)
    'prject the position to the plane
    oProjectedPos.x = dist * oPlaneNormal.x + oPos.x
    oProjectedPos.y = dist * oPlaneNormal.y + oPos.y
    oProjectedPos.z = dist * oPlaneNormal.z + oPos.z

    Set ProjectPosToPlane = oProjectedPos
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function

'*************************************************************************
'Function
'
'<ProjectLineToPlane>
'
'Abstract
'
'<Projects given line on to a plane to return the line on that plane>
'
'Arguments
'
'<line as IJLine,rootposition as IJDPosition,normal as IJDVector>
'
'Return
'
'<line as IJLine>
'
'Exceptions
'***************************************************************************
Public Function ProjectLineToPlane(oLine As IJLine, oPlaneRoot As IJDPosition, oPlaneNormal As IJDVector) As IJLine
  Const MT = "ProjectLineToPlane"
    On Error GoTo ErrorHandler
    Dim pGeometryFactory As New GeometryFactory
    Dim oProjectedLine As IJLine
    Dim oStartPos As New DPosition, oEndPos As New DPosition
    Dim x#, y#, z#

    oLine.GetStartPoint x, y, z
    oStartPos.Set x, y, z
    
    oLine.GetEndPoint x, y, z
    oEndPos.Set x, y, z
    
    Set oStartPos = ProjectPosToPlane(oStartPos, oPlaneRoot, oPlaneNormal)
    Set oEndPos = ProjectPosToPlane(oEndPos, oPlaneRoot, oPlaneNormal)
    
    Set oProjectedLine = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oStartPos.x, oStartPos.y, oStartPos.z, oEndPos.x, oEndPos.y, oEndPos.z)
    Set ProjectLineToPlane = oProjectedLine
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function

'*************************************************************************
'Function
'
'<GetXVector>
'
'Abstract
'
'<returns the X vector of the MemberPartPrismatic from its Transformation Matrix>
'
'Arguments
'
'<Member as ISPSMemberPartPrismatic>
'
'Return
'
'<vector as IJDVector>
'
'Exceptions
'***************************************************************************
Public Function GetXVector(oMemb As ISPSMemberPartPrismatic, Optional oPosAlong As IJDPosition) As IJDVector
  Const MT = "GetXVector"
    On Error GoTo ErrorHandler
    Dim oMat As IJDT4x4
    Dim oXVec As New DVector
    If oPosAlong Is Nothing Then
        oMemb.Rotation.GetTransform oMat
    Else
        oMemb.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMat, Nothing
    End If
    'get X vector from the transformation matrix
    oXVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    'normalize
    oXVec.Length = 1
    Set GetXVector = oXVec
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetYVector>
'
'Abstract
'
'<returns the Y vector of the MemberPartPrismatic from its Transformation Matrix>
'
'Arguments
'
'<Member as ISPSMemberPartPrismatic>
'
'Return
'
'<vector as IJDVector>
'
'Exceptions
'***************************************************************************
Public Function GetYVector(oMemb As ISPSMemberPartPrismatic, Optional oPosAlong As IJDPosition) As IJDVector
  Const MT = "GetYVector"
    On Error GoTo ErrorHandler
    Dim oMat As IJDT4x4
    Dim oYVec As New DVector
    
    If Not oPosAlong Is Nothing Then
        oMemb.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMat, Nothing
    Else
        oMemb.Rotation.GetTransform oMat
    End If
    'get Y vector from the transformation matrix
    oYVec.Set oMat.IndexValue(4), oMat.IndexValue(5), oMat.IndexValue(6)
    'normalize
    oYVec.Length = 1
    Set GetYVector = oYVec
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<GetZVector>
'
'Abstract
'
'<returns the Z vector of the MemberPartPrismatic from its Transformation Matrix>
'
'Arguments
'
'<Member as ISPSMemberPartPrismatic>
'
'Return
'
'<vector as IJDVector>
'
'Exceptions
'***************************************************************************
Public Function GetZVector(oMemb As ISPSMemberPartPrismatic, Optional oPosAlong As IJDPosition) As IJDVector
  Const MT = "GetZVector"
    On Error GoTo ErrorHandler
    Dim oMat As IJDT4x4
    Dim oZvec As New DVector
    
    If Not oPosAlong Is Nothing Then
        oMemb.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMat, Nothing
    Else
        oMemb.Rotation.GetTransform oMat
    End If
    'get Z vector from the transformation matrix
    oZvec.Set oMat.IndexValue(8), oMat.IndexValue(9), oMat.IndexValue(10)
    'normalize
    oZvec.Length = 1
    Set GetZVector = oZvec
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<AreVectorsParallel>
'
'Abstract
'
'<Determines if the two vectors are parallel>
'
'Arguments
'
'<2 vectors as IJDVector>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function AreVectorsParallel(oVec1 As IJDVector, oVec2 As IJDVector) As Boolean
  Const MT = "AreVectorsParallel"
    On Error GoTo ErrorHandler
    Dim cosA As Double
    Dim diff As Double
    'normalize vectors
    oVec1.Length = 1
    oVec2.Length = 1
    
    cosA = oVec1.Dot(oVec2)
    'if the vectors are at 180 deg , cosA=-1
    diff = Abs(cosA) - 1
    If Abs(diff) < angleTol Then
        AreVectorsParallel = True
    End If
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<GetMembSidePlane>
'
'Abstract
'
'<Given a qudrant, returns the side plane of the member>
'
'Arguments
'
'<Memb as MemberPartPrismatic,quadrant as integer,oRootPos As IJDPosition ,oNormal as IJDVector>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub GetMembSidePlane(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer, ByRef oRootPos As IJDPosition, ByRef oNormal As IJDVector)
  Const MT = "GetMembSidePlane"
    On Error GoTo ErrorHandler
    Set oRootPos = GetMembSidePos(oMemb, iQuadrant)
    Set oNormal = GetMembSidePlaneNormal(oMemb, iQuadrant)
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'
'<GetMembSidePlaneNormal>
'
'Abstract
'
'<Given a qudrant and position along the axis  returns the side plane normal for the member>
'if position along the axis is not given the normal returned is for the start of the member
'
'Arguments
'
'<Memb as MemberPartPrismatic,quadrant as integer>
'
'Return
'
'<Normal as IJDVector>
'
'Exceptions
'***************************************************************************
Public Function GetMembSidePlaneNormal(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer, Optional oPosAlong As IJDPosition) As IJDVector
  Const MT = "GetMembSidePlaneNormal"
    On Error GoTo ErrorHandler
    
    Dim cp As Long
    Dim oProfileBO As ISPSCrossSection
    Dim oMat As IJDT4x4
    Dim xOffset As Double, yOffset As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
    Dim oVec As New DVector
    Dim bMirror  As Boolean
    
    'iQuadrant=1 return top plane normal, cp =7,8,9
    'iQuadrant=2 return right plane normal cp =9,6,3
    'iQuadrant=3 return bottom plane normal cp=3,2,1
    'iQuadrant=4 return left plane normal cp = 1,4,7
    
    bMirror = oMemb.Rotation.Mirror
    If iQuadrant = 1 Then
        cp = 8
    ElseIf iQuadrant = 2 Then
        cp = 6
    ElseIf iQuadrant = 3 Then
        cp = 2
    ElseIf iQuadrant = 4 Then
        cp = 4
    End If
    
    Set oProfileBO = oMemb.CrossSection

    oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns in x and y of the CP5 in cross section coordinates
    oProfileBO.GetCardinalPointOffset cp, xOffset, yOffset 'Returns in x and y of the CP5 in cross section coordinates
    oVec.x = xOffset - xOffsetCP5 ' x offset of current cp from CP5
    oVec.y = yOffset - yOffsetCP5 ' y offset of current cp from CP5
    oVec.z = 0
    If Not oPosAlong Is Nothing Then
        oMemb.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMat, Nothing
    Else
        oMemb.Rotation.GetTransform oMat
    End If
    'remove translation component
    oMat.IndexValue(12) = 0
    oMat.IndexValue(13) = 0
    oMat.IndexValue(14) = 0
    
    Set oMat = CreateCSToMembTransform(oMat)
    'transform to global
    Set oVec = oMat.TransformVector(oVec)
    oVec.Length = 1
    Set GetMembSidePlaneNormal = oVec
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<GetMembSidePos>
'
'Abstract
'
'<Given a quadrant,position along the axis returns a point on the face of the section depending upon  section type>
'when oPosAlong is not given the position returned is at the start of the meber
'
'<Position returned here is on the face of the member at the start point>
'
'Arguments
'
'<Member as MemberPartPrismatic,quadrant as integer, optional oPosAlong as IJPosition>
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetMembSidePos(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer, Optional oPosAlong As IJDPosition) As IJDPosition
  Const MT = "GetMembSidePos"
    On Error GoTo ErrorHandler

    Dim curCP As Long, cp As Long
    Dim oProfileBO As ISPSCrossSection
    Dim oPos As IJDPosition
    Dim oMat As New DT4x4
    Dim xOffsetCP5 As Double, yOffsetCP5 As Double, xOffsetCP As Double, yOffsetCP As Double
    Dim oVec As New DVector
    
    If Not oMemb Is Nothing Then
        Set oProfileBO = oMemb.CrossSection
        If Not oProfileBO Is Nothing Then
            Dim oCrossection As IJCrossSection
            Set oCrossection = oProfileBO.definition
            If Not oCrossection Is Nothing Then
                Dim strSectionType As String
                strSectionType = oCrossection.Type
            End If
        End If
    End If
    Select Case strSectionType
        Case "HSSC", "CS", "PIPE", "RS", "HSSR"
            Set oPos = GetMembSidePosForRect(oMemb, iQuadrant)
        Case "W", "S", "HP", "M"
            Set oPos = GetMembSidePosForW(oMemb, iQuadrant)
        Case "L"
            Set oPos = GetMembSidePosForL(oMemb, iQuadrant)
        Case "C", "MC"
            Set oPos = GetMembSidePosForC(oMemb, iQuadrant)
        Case "WT", "MT", "ST"
            Set oPos = GetMembSidePosForT(oMemb, iQuadrant)
        Case "2L"
            Set oPos = GetMembSidePosFor2L(oMemb, iQuadrant)
        Case Else
            'for any unknown cross section, retun pos based on range box
            Set oPos = GetMembSidePosForRect(oMemb, iQuadrant)
    End Select

    curCP = oProfileBO.CardinalPoint
    
    oProfileBO.GetCardinalPointOffset curCP, xOffsetCP, yOffsetCP 'Returns Rad x and y of current  CP
    oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of the CP5
    xOffsetCP5 = xOffsetCP5 - xOffsetCP ' x offset of  cp5 from current CP
    yOffsetCP5 = yOffsetCP5 - yOffsetCP ' y offset of  cp5 from current CP
    
    'oPos was calculated based on CP5. Transform such that it is from the current CP
    oVec.Set xOffsetCP5, yOffsetCP5, 0
    oMat.LoadIdentity
    oMat.Translate oVec
    Set oPos = oMat.TransformPosition(oPos)
    
    'transform to global
    If Not oPosAlong Is Nothing Then
        oMemb.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMat, Nothing
    Else
        oMemb.Rotation.GetTransform oMat
    End If
    Set oMat = CreateCSToMembTransform(oMat, oMemb.Rotation.Mirror)
    Set GetMembSidePos = oMat.TransformPosition(oPos)
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetSectionRangeBoxSidePos>
'
'Abstract
'
'<Given a quadrant and position along the axis, returns a point on the face of the section rangebox>
'<If position along the axis is not given point on the face at the start of the member is returned>
'Arguments
'
'<Member as MemberPartPrismatic,quadrant as integer, position along the axis as IJDPosition>
'
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetSectionRangeBoxSidePos(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer, Optional oPosAlong As IJDPosition) As IJDPosition
  Const MT = "GetSectionRangeBoxSidePos"
    On Error GoTo ErrorHandler

    Dim curCP As Long, cp As Long
    Dim oProfileBO As ISPSCrossSection
    Dim oPos As IJDPosition
    Dim oMat As New DT4x4
    Dim xOffsetCP5 As Double, yOffsetCP5 As Double, xOffsetCP As Double, yOffsetCP As Double
    Dim oVec As New DVector

    Set oProfileBO = oMemb.CrossSection

    Set oPos = GetMembSidePosForRect(oMemb, iQuadrant)
    curCP = oProfileBO.CardinalPoint

    oProfileBO.GetCardinalPointOffset curCP, xOffsetCP, yOffsetCP 'Returns Rad x and y of current  CP
    oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5 'Returns Rad x and y of the CP5
    xOffsetCP5 = xOffsetCP5 - xOffsetCP ' x offset of  cp5 from current CP
    yOffsetCP5 = yOffsetCP5 - yOffsetCP ' y offset of  cp5 from current CP
    
    'oPos was calculated based on CP5. Transform such that it is from the current CP
    oVec.Set xOffsetCP5, yOffsetCP5, 0
    oMat.LoadIdentity
    oMat.Translate oVec
    Set oPos = oMat.TransformPosition(oPos)
    
    'transform to global
    If Not oPosAlong Is Nothing Then
        oMemb.Rotation.GetTransformAtPosition oPosAlong.x, oPosAlong.y, oPosAlong.z, oMat, Nothing
    Else
        oMemb.Rotation.GetTransform oMat
    End If
    Set oMat = CreateCSToMembTransform(oMat, oMemb.Rotation.Mirror)
    Set GetSectionRangeBoxSidePos = oMat.TransformPosition(oPos)
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetMembSidePosForW>
'
'Abstract
'
'<Given a quadrant, returns a point on the face of the W section>
'
'Arguments
'
'<Member as MemberPartPrismatic,quadrant as integer>
'
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetMembSidePosForW(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer) As IJDPosition
  Const MT = "GetMembSidePosForW"
    On Error GoTo ErrorHandler
    
    Dim oProfileBO As ISPSCrossSection
    Dim oPos As New DPosition
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, tWeb As Double
    
    Set oProfileBO = oMemb.CrossSection
    Set oSectionAttrbs = oMemb.CrossSection.Definition
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value

    ' position based on CP5 as origin
    Select Case iQuadrant
        Case 1
                oPos.Set 0, depth / 2, 0
        Case 2
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set tWeb / 2, 0, 0
                Else
                    oPos.Set -tWeb / 2, 0, 0
                End If
        Case 3
                oPos.Set 0, -depth / 2, 0
        Case 4
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set -tWeb / 2, 0, 0
                Else
                    oPos.Set tWeb / 2, 0, 0
                End If
    End Select
    Set GetMembSidePosForW = oPos
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<GetMembSidePosForC>
'
'Abstract
'
'<Given a quadrant, returns a point on the face of the Channel section>
'
'Arguments
'
'<Member as MemberPartPrismatic,quadrant as integer>
'
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetMembSidePosForC(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer) As IJDPosition
  Const MT = "GetMembSidePosForC"
    On Error GoTo ErrorHandler
    
    Dim oProfileBO As ISPSCrossSection
    Dim oPos As New DPosition
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, tWeb As Double, Width As Double
    
    Set oProfileBO = oMemb.CrossSection
    Set oSectionAttrbs = oMemb.CrossSection.Definition
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    Width = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value

    ' position based on CP5 as origin
    Select Case iQuadrant
        Case 1
                oPos.Set 0, depth / 2, 0
        Case 2
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set -Width / 2 + tWeb, 0, 0
                Else
                    oPos.Set -Width / 2, 0, 0
                End If
        
        Case 3
                oPos.Set 0, -depth / 2, 0
        Case 4
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set -Width / 2, 0, 0
                Else
                    oPos.Set -Width / 2 + tWeb, 0, 0
                End If
    End Select
    Set GetMembSidePosForC = oPos
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<GetMembSidePosForL>
'
'Abstract
'
'<Given a quadrant, returns a point on the face of the Angle section>
'
'Arguments
'
'<Member as MemberPartPrismatic,quadrant as integer>
'
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetMembSidePosForL(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer) As IJDPosition
  Const MT = "GetMembSidePosForL"
    On Error GoTo ErrorHandler
    
    Dim oProfileBO As ISPSCrossSection
    Dim oPos As New DPosition
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, Width As Double, tFlange As Double
    
    Set oProfileBO = oMemb.CrossSection
    Set oSectionAttrbs = oMemb.CrossSection.Definition
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    Width = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    
    ' position based on CP5 as origin
    Select Case iQuadrant
        Case 1
                oPos.Set 0, -depth / 2 + tFlange, 0
        Case 2
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set -Width / 2 + tFlange, 0, 0 'thickness of web and thickness of flange same for angle
                Else
                    oPos.Set -Width / 2, 0, 0
                End If
        Case 3
                oPos.Set 0, -depth / 2, 0
        Case 4
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set -Width / 2, 0, 0
                Else
                    oPos.Set -Width / 2 + tFlange, 0, 0 'thickness of web and thickness of flange same for angle
                End If
    End Select
    Set GetMembSidePosForL = oPos
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetMembSidePosForT>
'
'Abstract
'
'<Given a quadrant, returns a point on the face of the T section>
'
'Arguments
'
'<Member as MemberPartPrismatic,quadrant as integer>
'
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetMembSidePosForT(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer) As IJDPosition
  Const MT = "GetMembSidePosForT"
    On Error GoTo ErrorHandler
    
    Dim oProfileBO As ISPSCrossSection
    Dim oPos As New DPosition
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, tWeb As Double, tFlange As Double
    
    Set oProfileBO = oMemb.CrossSection
    Set oSectionAttrbs = oMemb.CrossSection.Definition
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    
    ' position based on CP5 as origin
    Select Case iQuadrant
        Case 1
                oPos.Set 0, depth / 2, 0
        Case 2
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set tWeb / 2, 0, 0
                Else
                    oPos.Set -tWeb / 2, 0, 0
                End If
        Case 3
                oPos.Set 0, depth / 2 - tFlange, 0
        Case 4
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set -tWeb / 2, 0, 0
                Else
                    oPos.Set tWeb / 2, 0, 0
                End If
    End Select
    Set GetMembSidePosForT = oPos
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetMembSidePosFor2L>
'
'Abstract
'
'<Given a quadrant, returns a point on the face of the 2L section>
'
'Arguments
'
'<Member as MemberPartPrismatic,quadrant as integer>
'
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetMembSidePosFor2L(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer) As IJDPosition
  Const MT = "GetMembSidePosFor2L"
    On Error GoTo ErrorHandler
    
    Dim oProfileBO As ISPSCrossSection
    Dim oPos As New DPosition
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, tFlange As Double, sBtoB As Double
    
    Set oProfileBO = oMemb.CrossSection
    Set oSectionAttrbs = oMemb.CrossSection.Definition
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    sBtoB = oSectionAttrbs.CollectionOfAttributes("IJUA2L").Item("bb").Value
    
    ' position based on CP5 as origin
    Select Case iQuadrant
        Case 1
                oPos.Set 0, -depth / 2 + tFlange, 0
        Case 2
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set tFlange + sBtoB / 2, 0, 0 'thickness of web and thickness of flange same for angle
                Else
                    oPos.Set -tFlange - sBtoB / 2, 0, 0
                End If
        Case 3
                oPos.Set 0, -depth / 2, 0
        Case 4
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set -tFlange - sBtoB / 2, 0, 0 'thickness of web and thickness of flange same for angle
                Else
                    oPos.Set tFlange + sBtoB / 2, 0, 0
                End If
    End Select
    Set GetMembSidePosFor2L = oPos
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetMembSidePosForRect>
'
'Abstract
'
'<Given a quadrant, returns a point on the face of the circular , rectangular section or on the range box>
'
'Arguments
'
'<Member as MemberPartPrismatic,quadrant as integer>
'
'Return
'
'<Position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function GetMembSidePosForRect(oMemb As ISPSMemberPartPrismatic, iQuadrant As Integer) As IJDPosition
  Const MT = "GetMembSidePosForRect"
    On Error GoTo ErrorHandler
    
    Dim oProfileBO As ISPSCrossSection
    Dim oPos As New DPosition
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, Width As Double
    
    Set oProfileBO = oMemb.CrossSection
    Set oSectionAttrbs = oMemb.CrossSection.Definition
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    Width = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    ' position based on CP5 as origin
    Select Case iQuadrant
        Case 1
                oPos.Set 0, depth / 2, 0
        Case 2
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set Width / 2, 0, 0
                Else
                    oPos.Set -Width / 2, 0, 0
                End If
        Case 3
                oPos.Set 0, -depth / 2, 0
        Case 4
                If oMemb.Rotation.Mirror = False Then
                    oPos.Set -Width / 2, 0, 0
                Else
                    oPos.Set Width / 2, 0, 0
                End If

    End Select
    Set GetMembSidePosForRect = oPos
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetCutDistAndPosFromPlane>
'
'Abstract
'
'<Calculates the position and distance of the cutplane from the end port>
'
'Arguments
'
'<Memb as MemberPartPrismatic,port as SPSMemberAxisPortIndex,plane as IJPlane,oPos as IJDPosition, dist as Double>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub GetCutDistAndPosFromPlane(oMemb As ISPSMemberPartPrismatic, iPort As SPSMemberAxisPortIndex, oSurface As IJSurface, ByRef oPos As IJDPosition, ByRef dist As Double)
  Const MT = "GetCutDistAndPosFromPlane"
    On Error GoTo ErrorHandler
    Dim oLine(1 To 4) As IJLine
    Dim idx As Integer
    Dim code As Geom3dIntersectConstants
    Dim colElm As IJElements
    Dim oPoint As IJPoint
    Dim oCutPos As New DPosition
    Dim distToPoint As Double, tmpDist As Double
    Dim x1#, y1#, z1#, x2#, y2#, z2#, dx#, dy#, dz#
    Dim bStartPoint As Boolean
    
    
    If iPort = SPSMemberAxisStart Then
        bStartPoint = True
    Else
        bStartPoint = False
    End If
    
    ' get tangential bounding lines for the end
    Set oLine(1) = GetTangentLineAtCP(oMemb, 1, bStartPoint)
    Set oLine(2) = GetTangentLineAtCP(oMemb, 3, bStartPoint)
    Set oLine(3) = GetTangentLineAtCP(oMemb, 7, bStartPoint)
    Set oLine(4) = GetTangentLineAtCP(oMemb, 9, bStartPoint)
    
    oMemb.PointAtEnd(iPort).GetPoint x1, y1, z1
    oCutPos.Set x1, y1, z1
    distToPoint = 0
    For idx = 1 To 4
        tmpDist = 0
        'get the other end point
        If iPort = SPSMemberAxisStart Then
            oLine(idx).GetEndPoint x1, y1, z1
        Else
            oLine(idx).GetStartPoint x1, y1, z1
        End If
        'make the line infinite
        oLine(idx).Infinite = True
        oSurface.Intersect oLine(idx), colElm, code
        If colElm.Count > 0 Then
            Set oPoint = colElm.Item(1)
            oPoint.GetPoint x2, y2, z2
            dx = x2 - x1
            dy = y2 - y1
            dz = z2 - z1
            tmpDist = Sqr(dx * dx + dy * dy + dz * dz)
            'the cutback position is that which is least distant from the other end
            If (idx = 1) Or (tmpDist < distToPoint) Then
                distToPoint = tmpDist
                oCutPos.Set x2, y2, z2
            End If
        End If
    Next idx
    Set oPos = oCutPos
    dist = distToPoint
    
    For idx = 1 To 4
        Set oLine(idx) = Nothing
    Next idx
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'
'<GetCopePoint>
'
'Abstract
'
'<calculates the copepoint using the Supporting and Supported Members>
'
'Arguments
'
'<Supporting and Supported as MemberPartPrismatic,end as SPSMemberAxisPortIndex, Plane as IJPlane, bTop as Boolean , oCopePoint As IJPoint>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub GetCopePoint(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oPlane As IJPlane, bTop As Boolean, ByRef oCopePoint As IJPoint)
  Const MT = "GetCopePoint"
    On Error GoTo ErrorHandler
    Dim oLine(1 To 4) As IJLine
    Dim idx As Integer
    Dim oSurf As IJSurface
    Dim code As Geom3dIntersectConstants
    Dim colElm As IJElements
    Dim oPoint(1 To 4) As IJPoint
    Dim dist1 As Double, dist2 As Double
    Dim x#, y#, z#
    Dim oMat As IJDT4x4
    Dim oPos2 As New DPosition
    Dim index As Integer
    Dim bStartPoint As Boolean
    Dim oPosAlongAxis As IJDPosition
    
    
    If iEnd = SPSMemberAxisStart Then
        bStartPoint = True
    Else
        bStartPoint = False
    End If
    ' get bounding lines
    Set oLine(1) = GetTangentLineAtCP(oSupped, 1, bStartPoint)
    Set oLine(2) = GetTangentLineAtCP(oSupped, 3, bStartPoint)
    Set oLine(3) = GetTangentLineAtCP(oSupped, 7, bStartPoint)
    Set oLine(4) = GetTangentLineAtCP(oSupped, 9, bStartPoint)
    
    Set oSurf = oPlane
    
    Set oPosAlongAxis = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    oSupping.Rotation.GetTransformAtPosition oPosAlongAxis.x, oPosAlongAxis.y, oPosAlongAxis.z, oMat, Nothing
    oMat.Invert

    dist1 = 0
    dist2 = 0
    
    For idx = 1 To 4
        oLine(idx).Infinite = True
        oSurf.Intersect oLine(idx), colElm, code
        If Not colElm Is Nothing Then
            If colElm.Count > 0 Then
                Set oPoint(idx) = colElm.Item(1)
                oPoint(idx).GetPoint x, y, z
                oPos2.Set x, y, z
                'transform to supporting local coordinates
                Set oPos2 = oMat.TransformPosition(oPos2)
                oPos2.Get x, y, z
                If idx = 1 Then
                    dist1 = z
                    dist2 = z
                    index = idx
                End If
                If bTop Then
                    If z > dist1 Then
                        dist1 = z
                        index = idx
                    End If
                Else
                    If z < dist2 Then
                        dist2 = z
                        index = idx
                    End If
                End If
            End If
        End If
    Next idx
    
    Set oCopePoint = oPoint(index) 'this may be nothing if no tangent line intersected the surface
    'caller needs to check the return value for nothing
    For idx = 1 To 4
        Set oLine(idx) = Nothing
        Set oPoint(idx) = Nothing
    Next idx
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'
'<TransformPosToGlobal>
'
'Abstract
'
'<Transforms position from Member local to global>
'
'Arguments
'
'<Member as MemberPartPrismatic, position to be transformed, optional position along the axis>
'The third parameter if passed in is used for obtaining the transformation matrix
'
'Return
'
'<position as IJDPosition>
'
'Exceptions
'***************************************************************************
Public Function TransformPosToGlobal(oMemb As ISPSMemberPartPrismatic, oInputPos As IJDPosition, Optional oPosAlongAxis As IJDPosition) As IJDPosition
Const MT = "TransformPosToGlobal"
  On Error GoTo ErrorHandler
    Dim oProfileBO As ISPSCrossSection
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
    Dim oMat1 As New DT4x4, oMat2 As DT4x4
    Dim oVec As New DVector

    
    Set oProfileBO = oMemb.CrossSection
    oProfileBO.GetCardinalPointOffset oProfileBO.CardinalPoint, xOffsetCP, yOffsetCP 'Returns x and y of the current CP in RAD coordinates, which is member negative y and z.
    oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5
    xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
    yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
    
    oVec.Set xOffset, yOffset, 0
    oMat1.LoadIdentity
    oMat1.Translate oVec
    
    If oPosAlongAxis Is Nothing Then
        oMemb.Rotation.GetTransform oMat2
    Else
        oMemb.Rotation.GetTransformAtPosition oPosAlongAxis.x, oPosAlongAxis.y, oPosAlongAxis.z, oMat2, Nothing
    End If
    'from cross section coordinates to member coordinates
    Set oMat2 = CreateCSToMembTransform(oMat2, oMemb.Rotation.Mirror)
    oMat2.MultMatrix oMat1
    Set TransformPosToGlobal = oMat2.TransformPosition(oInputPos)
  Exit Function
ErrorHandler:    HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
'<GetCrossSectionLines>
'
'Abstract
'
'<Get lines connecting Cross section points at either end of the member>
'
'Arguments
'
'<Member as MemberPartPrismatic,member end as SPSMemberAxisPortIndex,  lines as IJLine>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub GetCrossSectionLines(oMemb As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex, oCornerLines() As IJLine)
  Const MT = "GetCrossSectionLines"
    On Error GoTo ErrorHandler
    
    Dim oSectionAttrbs As IJDAttributes
    Dim depth As Double, tFlange As Double, tWeb As Double, bFlange As Double
    Dim oPos() As IJDPosition
    Dim xVec As New DVector
    Dim membLength As Double
    Dim i As Integer
    Dim oMat As IJDT4x4
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#, nx#, ny#, nz#
    Dim pGeometryFactory As New GeometryFactory
    Dim oProfile As IJCrossSection
    Dim uBnd As Integer
    Dim bTob#
    Dim oPosAlongAxis As New DPosition
    
     If Not oMemb Is Nothing Then
        Dim OSPSCrossSection As ISPSCrossSection
        Set OSPSCrossSection = oMemb.CrossSection
        If Not OSPSCrossSection Is Nothing Then
            Set oProfile = OSPSCrossSection.definition
            If Not oProfile Is Nothing Then
                Dim strSectionType As String
                strSectionType = oProfile.Type
            End If
        End If
    End If
    
    Set oSectionAttrbs = oMemb.CrossSection.Definition
    'for now assume that the section type is W
    'get section attributes
    depth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    bFlange = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    On Error Resume Next ' some sections don't support the interfaces below
    tWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
    tFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
    On Error GoTo ErrorHandler
    membLength = oMemb.Axis.Length
    oMemb.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    
    If iEnd = SPSMemberAxisStart Then
        oMemb.Rotation.GetTransformAtPosition sX, sY, sZ, oMat, Nothing
        oPosAlongAxis.Set sX, sY, sZ
    Else
        oMemb.Rotation.GetTransformAtPosition eX, eY, eZ, oMat, Nothing
        oPosAlongAxis.Set eX, eY, eZ
    End If
    
   
    Select Case strSectionType
        Case "HSSC", "CS", "PIPE", "RS", "HSSR"
            uBnd = 4
            ReDim oPos(1 To uBnd)
            For i = 1 To uBnd
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, -oPos(2).y, 0
            oPos(4).Set oPos(1).x, -oPos(1).y, 0
        
        Case "W", "S", "HP", "M"
            uBnd = 12
            ReDim oPos(1 To uBnd)
            For i = 1 To uBnd
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, oPos(2).y - tFlange, 0
            oPos(4).Set tWeb / 2, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -oPos(4).y, 0
            oPos(6).Set oPos(3).x, -oPos(3).y, 0
            oPos(7).Set oPos(2).x, -oPos(2).y, 0
            oPos(8).Set oPos(1).x, -oPos(1).y, 0
            oPos(9).Set -oPos(6).x, oPos(6).y, 0
            oPos(10).Set -oPos(5).x, oPos(5).y, 0
            oPos(11).Set oPos(10).x, -oPos(10).y, 0
            oPos(12).Set oPos(9).x, -oPos(9).y, 0
        
        Case "L"
        
            uBnd = 6
            ReDim oPos(1 To uBnd)
            For i = 1 To uBnd
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set oPos(1).x - tWeb, oPos(1).y, 0
            oPos(3).Set oPos(2).x, -depth / 2 + tFlange, 0
            oPos(4).Set -oPos(1).x, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -oPos(1).y, 0
            oPos(6).Set oPos(1).x, -oPos(1).y, 0
        
        Case "C", "MC"
            uBnd = 8
            ReDim oPos(1 To uBnd)
            For i = 1 To uBnd
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, oPos(2).y - tFlange, 0
            oPos(4).Set oPos(1).x - tWeb, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -oPos(4).y, 0
            oPos(6).Set oPos(3).x, -oPos(3).y, 0
            oPos(7).Set oPos(2).x, -oPos(2).y, 0
            oPos(8).Set oPos(1).x, -oPos(1).y, 0

        Case "WT", "MT", "ST"
            uBnd = 8
            ReDim oPos(1 To uBnd)
            For i = 1 To uBnd
                Set oPos(i) = New DPosition
            Next i
            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, oPos(2).y - tFlange, 0
            oPos(4).Set tWeb / 2, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -oPos(4).y, 0
            
            oPos(6).Set -oPos(5).x, oPos(5).y, 0
            oPos(7).Set -oPos(4).x, oPos(4).y, 0
            oPos(8).Set -oPos(3).x, oPos(3).y, 0
        Case "2L"
            bTob = oSectionAttrbs.CollectionOfAttributes("IJUA2L").Item("bb").Value
            uBnd = 8
            ReDim oPos(1 To uBnd)
            For i = 1 To uBnd
                Set oPos(i) = New DPosition
            Next i
            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -(tWeb / 2 + bTob / 2), depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, -(depth / 2 - tFlange), 0
            oPos(4).Set bFlange / 2, oPos(3).y, 0
            oPos(5).Set oPos(4).x, -depth / 2, 0
            oPos(6).Set -oPos(5).x, oPos(5).y, 0
            oPos(7).Set -oPos(4).x, oPos(4).y, 0
            oPos(8).Set -oPos(3).x, oPos(3).y, 0
        Case Else
            'unknown cross section. Create based on range box
            uBnd = 4
            ReDim oPos(1 To uBnd)
            For i = 1 To uBnd
                Set oPos(i) = New DPosition
            Next i

            'oPos(1) is CP7 and going clockwise
            oPos(1).Set -bFlange / 2, depth / 2, 0
            oPos(2).Set -oPos(1).x, oPos(1).y, 0
            oPos(3).Set oPos(2).x, -oPos(2).y, 0
            oPos(4).Set oPos(1).x, -oPos(1).y, 0
    End Select
    
    xVec.Set -oMat.IndexValue(0), -oMat.IndexValue(1), -oMat.IndexValue(2)
    xVec.Length = membLength
    
    ReDim oCornerLines(1 To uBnd)
    For i = 1 To uBnd
        'transform Position to global
        Set oPos(i) = TransformPosToGlobal(oMemb, oPos(i), oPosAlongAxis)
        If iEnd = SPSMemberAxisEnd Then
            Set oPos(i) = oPos(i).Offset(xVec)
        End If
        
        oPos(i).Get sX, sY, sZ
        Set oCornerLines(i) = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, sX, sY, sZ, oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2), membLength)
    Next i
    
    For i = 1 To UBound(oPos)
        Set oPos(i) = Nothing
    Next i
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
    
End Sub

'*************************************************************************
'Function
'
'<GetIntersectionPoints>
'
'Abstract
'
'<>
'
'Arguments
'
'<>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub GetIntersectionPoints(oCornerLines() As IJLine, oSurface As IJSurface, oInterSecPoints() As IJDPosition)
  Const MT = "GetIntersectionPoints"
    On Error GoTo ErrorHandler
    Dim uBnd As Integer, lBnd As Integer
    Dim i As Integer
    Dim colElms As IJElements
    Dim code As Geom3dIntersectConstants
    Dim oPoint As IJPoint
    Dim x#, y#, z#
    Dim pGeometryFactory As New GeometryFactory
    Dim oLine As IJLine
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
        
    lBnd = LBound(oCornerLines)
    uBnd = UBound(oCornerLines)
        
    ReDim oInterSecPoints(lBnd To uBnd)
    For i = lBnd To uBnd
        oCornerLines(i).GetStartPoint sX, sY, sZ
        oCornerLines(i).GetEndPoint eX, eY, eZ
        Set oLine = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, sX, sY, sZ, eX, eY, eZ)
        oLine.Infinite = True
        oSurface.Intersect oLine, colElms, code
        If Not colElms Is Nothing Then
            If colElms.Count > 0 Then
                Set oPoint = colElms.Item(1)
                oPoint.GetPoint x, y, z
                Set oInterSecPoints(i) = New DPosition
                oInterSecPoints(i).Set x, y, z
            End If
        End If
    Next i
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub

'The input edge points (They correspond to edges of a closed 2d shape) are clipped to the cope positions
'*************************************************************************
'Function
'
'<ClipEdgePosToYDepth>
'
'Abstract
'
'<>
'
'Arguments
'
'<>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub ClipEdgePosToYDepth(oInputPos() As IJDPosition, oTopClipPos As IJDPosition, oBottomClipPos As IJDPosition, oCopeTransform As IJDT4x4, oClippedPos() As IJDPosition)
  Const MT = "ClipEdgePosToYDepth"
    On Error GoTo ErrorHandler
    
    Dim uBnd As Integer, lBnd As Integer
    Dim oLine() As IJLine, oLineLocal() As IJLine
    Dim pGeometryFactory As New GeometryFactory
    Dim nextIdx As Integer
    Dim oSurfHigh As IJSurface, oSurfLow As IJSurface
    Dim i As Integer
    Dim colElms As IJElements
    Dim code As Geom3dIntersectConstants
    Dim oPoint As IJPoint
    Dim cnt As Integer
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim yHigh#, yLow#, yMax#, yMin#
    Dim xTop#, yTop#, zTop#, xBot#, yBot#, zBot#
    Dim oPos As IJDPosition
    Dim x#, y#, z#, nx#, ny#, nz#
    Dim oMatLocal As New DT4x4
        
        
    oMatLocal.LoadMatrix oCopeTransform
    oMatLocal.Invert 'global to local
    
    Set oPos = oMatLocal.TransformPosition(oTopClipPos)
    yMax = oPos.y
        
    Set oPos = oMatLocal.TransformPosition(oBottomClipPos)
    yMin = oPos.y
    
    
    nx = oCopeTransform.IndexValue(4)
    ny = oCopeTransform.IndexValue(5)
    nz = oCopeTransform.IndexValue(6)
    oTopClipPos.Get xTop, yTop, zTop
    oBottomClipPos.Get xBot, yBot, zBot
    
    If yMax > yMin Then
        yHigh = yMax
        yLow = yMin
    '    'now create planes that correspond to clipping planes at yHigh and yLow
        Set oSurfHigh = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, xTop, yTop, zTop, nx, ny, nz)
        Set oSurfLow = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, xBot, yBot, zBot, nx, ny, nz)
    
    Else
        yHigh = yMin
        yLow = yMax
    '    'now create planes that correspond to clipping planes at yHigh and yLow
        Set oSurfHigh = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, xBot, yBot, zBot, nx, ny, nz)
        Set oSurfLow = pGeometryFactory.Planes3d.CreateByPointNormal(Nothing, xTop, yTop, zTop, nx, ny, nz)
    End If
        
    lBnd = LBound(oInputPos)
    uBnd = UBound(oInputPos)
    
    ReDim oLine(lBnd To uBnd)
    ReDim oLineLocal(lBnd To uBnd)
    'create lines connecting the corner points
    For i = lBnd To uBnd
        nextIdx = i Mod uBnd
        Set oLine(i) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oInputPos(i).x, oInputPos(i).y, oInputPos(i).z, _
        oInputPos(nextIdx + 1).x, oInputPos(nextIdx + 1).y, oInputPos(nextIdx + 1).z)
        'create another line as a copy of the above and transform to local
        Set oLineLocal(i) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oInputPos(i).x, oInputPos(i).y, oInputPos(i).z, _
        oInputPos(nextIdx + 1).x, oInputPos(nextIdx + 1).y, oInputPos(nextIdx + 1).z)
        oLineLocal(i).Transform oMatLocal 'global to local
    Next i
    cnt = 0
    'what if the array, oClippedPos, is not initialized? That is all points are above the yDepth and bClipTop=true.

    For i = lBnd To uBnd
        oLineLocal(i).GetStartPoint sX, sY, sZ
        oLineLocal(i).GetEndPoint eX, eY, eZ
        If (sY < yHigh) Or (eY < yHigh) Then
            If (sY >= yHigh) Then 'one point is below yDepth
                oSurfHigh.Intersect oLine(i), colElms, code
                If Not colElms Is Nothing Then
                    If colElms.Count > 0 Then
                        Set oPoint = colElms.Item(1)
                        oPoint.GetPoint x, y, z
                        oLine(i).SetStartPoint x, y, z
                    End If
                End If
            ElseIf (eY >= yHigh) Then 'one point is below yDepth
                oSurfHigh.Intersect oLine(i), colElms, code
                If Not colElms Is Nothing Then
                    If colElms.Count > 0 Then
                        Set oPoint = colElms.Item(1)
                        oPoint.GetPoint x, y, z
                        oLine(i).SetEndPoint x, y, z
                    End If
                End If
            End If
            If (sY > yLow) Or (eY > yLow) Then
                If (sY <= yLow) Then 'one point is above yDepth
                    oSurfLow.Intersect oLine(i), colElms, code
                    If Not colElms Is Nothing Then
                        If colElms.Count > 0 Then
                            Set oPoint = colElms.Item(1)
                            oPoint.GetPoint x, y, z
                            'allocate for the points
                            cnt = cnt + 2
                            ReDim Preserve oClippedPos(1 To cnt)
                            Set oClippedPos(cnt - 1) = New DPosition
                            Set oClippedPos(cnt) = New DPosition
                            
                            oClippedPos(cnt - 1).Set x, y, z
                            oLine(i).GetEndPoint x, y, z
                            oClippedPos(cnt).Set x, y, z
                        End If
                    End If
                ElseIf (eY <= yLow) Then 'one point is above yDepth
                    oSurfLow.Intersect oLine(i), colElms, code
                    If Not colElms Is Nothing Then
                        If colElms.Count > 0 Then
                            Set oPoint = colElms.Item(1)
                            oPoint.GetPoint x, y, z
                            'allocate for the points
                            cnt = cnt + 2
                            ReDim Preserve oClippedPos(1 To cnt)
                            Set oClippedPos(cnt - 1) = New DPosition
                            Set oClippedPos(cnt) = New DPosition
                            oClippedPos(cnt).Set x, y, z
                            oLine(i).GetStartPoint x, y, z
                            oClippedPos(cnt - 1).Set x, y, z
                        End If
                    End If
                Else 'both points are above yDepth
                    'allocate for the points
                    cnt = cnt + 2
                    ReDim Preserve oClippedPos(1 To cnt)
                    Set oClippedPos(cnt - 1) = New DPosition
                    Set oClippedPos(cnt) = New DPosition
                    oLine(i).GetStartPoint x, y, z
                    oClippedPos(cnt - 1).Set x, y, z
                    oLine(i).GetEndPoint x, y, z
                    oClippedPos(cnt).Set x, y, z
                End If
            End If
        End If
    Next i
    
    For i = lBnd To uBnd
        oLine(i) = Nothing
        oLineLocal(i) = Nothing
    Next i
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
'creates lines from the positions to the other end parallel to member axis. On which side the positions, are determined by iPort.
'The positions may not be exactly at the end, but close to the end corresponding to iPort
'*************************************************************************
'Function
'
'<GetMembLineFromEndPos>
'
'Abstract
'
'<>
'
'Arguments
'
'<>
'
'Return
'
'Exceptions
'***************************************************************************
Public Sub GetMembLineFromEndPos(oMemb As ISPSMemberPartPrismatic, iPort As SPSMemberAxisPortIndex, oEndPos() As IJDPosition, oMembLine() As IJLine)
  Const MT = "GetMembLineFromEndPos"
    On Error GoTo ErrorHandler
    Dim lBnd As Integer, uBnd As Integer
    Dim oPos1 As New DPosition, oPos2 As New DPosition
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim pGeometryFactory As New GeometryFactory
    Dim i As Integer
    Dim oVec As New DVector
    Dim oMatGlobal As IJDT4x4, oMatLocal As IJDT4x4
    
    
    lBnd = LBound(oEndPos)
    uBnd = UBound(oEndPos)
    oMemb.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    
    'the member may be curved. so we imagine a linear member along the
    'tangent at the end with length say,10m
    
    'so let us compute the imaginary other end 10m away
    If iPort = SPSMemberAxisStart Then 'get the other end
        oMemb.Rotation.GetTransformAtPosition sX, sY, sZ, oMatGlobal, Nothing
        oVec.Set oMatGlobal.IndexValue(0), oMatGlobal.IndexValue(1), oMatGlobal.IndexValue(2)
        oVec.Length = 10
        oPos1.Set sX, sY, sZ
        Set oPos1 = oPos1.Offset(oVec)
        
    Else 'get the other end
        oMemb.Rotation.GetTransformAtPosition eX, eY, eZ, oMatGlobal, Nothing
        'reverse direction so that it is going from the connected end
        oVec.Set -oMatGlobal.IndexValue(0), -oMatGlobal.IndexValue(1), -oMatGlobal.IndexValue(2)
        oVec.Length = 10
        oPos1.Set eX, eY, eZ
        Set oPos1 = oPos1.Offset(oVec)
    End If
    'remove translation component from matrix
    oMatGlobal.IndexValue(12) = 0
    oMatGlobal.IndexValue(13) = 0
    oMatGlobal.IndexValue(14) = 0
    Set oMatLocal = oMatGlobal.Clone()
    oMatLocal.Invert
    Set oPos1 = oMatLocal.TransformPosition(oPos1) 'transform to local
    'allocat memory for the lines
    ReDim oMembLine(lBnd To uBnd)
    For i = lBnd To uBnd
        Set oPos2 = oMatLocal.TransformPosition(oEndPos(i)) 'transform to local
        ' Make Y  and Z of other end same as this end
        oPos1.y = oPos2.y
        oPos1.z = oPos2.z
        If iPort = SPSMemberAxisStart Then
            Set oMembLine(i) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos2.x, oPos2.y, oPos2.z, oPos1.x, oPos1.y, oPos1.z)
        Else
            Set oMembLine(i) = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos1.x, oPos1.y, oPos1.z, oPos2.x, oPos2.y, oPos2.z)
        End If
        'local to global
        oMembLine(i).Transform oMatGlobal
    Next i
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
'The lines are of diffrent length but have the same origin. oPos correspond to points on the lines on the other side of the origin
'The function below returns the position which is closest to the origin
'*************************************************************************
'Function
'
'<GetPosClosestToOtherEnd>
'
'Abstract
'
'<>
'
'Arguments
'
'<>
'
'Return
'
'Exceptions
'***************************************************************************
Public Function GetPosClosestToOtherEnd(oInputPos() As IJDPosition, oLine() As IJLine, iEnd As SPSMemberAxisPortIndex) As IJDPosition
  Const MT = "GetPosClosestToOtherEnd"
    On Error GoTo ErrorHandler
    Dim lBnd As Integer, uBnd As Integer
    Dim i As Integer, minDistIndex As Integer
    Dim x#, y#, z#
    Dim dist#, minDist#
    Dim oPos As New DPosition
    
    lBnd = LBound(oLine)
    uBnd = UBound(oLine)
    
    For i = lBnd To uBnd
        If iEnd = SPSMemberAxisStart Then
            oLine(i).GetEndPoint x, y, z
        Else
            oLine(i).GetStartPoint x, y, z
        End If
        oPos.Set x, y, z
        dist = oPos.DistPt(oInputPos(i))
        If i = lBnd Then
            minDist = dist
            minDistIndex = i
        End If
        If dist < minDist Then
            minDist = dist
            minDistIndex = i
        End If
    Next i
    oInputPos(minDistIndex).Get x, y, z
    oPos.Set x, y, z
    Set GetPosClosestToOtherEnd = oPos
    Exit Function
ErrorHandler:      HandleError MODULE, MT

End Function
'returns the trnasformation matrix for the cope
'*************************************************************************
'Function
'
'<GetCopeTransform>
'
'Abstract
'
'<>
'
'Arguments
'
'<>
'
'Return
'
'Exceptions
'***************************************************************************
Public Function GetCopeTransform(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, iEnd As SPSMemberAxisPortIndex) As IJDT4x4
  Const MT = "GetCopeTransform"
    On Error GoTo ErrorHandler
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim oCopeXvec As New DVector, oCopeYvec As IJDVector, oCopeZvec As IJDVector
    Dim oMat As New DT4x4
    Dim oPosAlong As New DPosition
    
    oSupped.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    'get supped axis direction
    If iEnd = SPSMemberAxisStart Then
        oSupped.Rotation.GetTransformAtPosition sX, sY, sZ, oMat, Nothing
        oCopeXvec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    Else
        oSupped.Rotation.GetTransformAtPosition eX, eY, eZ, oMat, Nothing
        oCopeXvec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        oCopeXvec.[Scale] -1 'reverse direction
    End If
    'vector is from the connected end to the other end
    oCopeXvec.Length = 1 'normalize
    
    Set oPosAlong = GetConnectionPositionOnSupping(oSupping, oSupped, iEnd)
    Set oCopeYvec = GetZVector(oSupping, oPosAlong)
    Set oCopeZvec = oCopeXvec.Cross(oCopeYvec)
    oCopeZvec.Length = 1
    Set oCopeYvec = oCopeZvec.Cross(oCopeXvec)
    oCopeYvec.Length = 1
    oMat.LoadIdentity
    'x axis
    oMat.IndexValue(0) = oCopeXvec.x
    oMat.IndexValue(1) = oCopeXvec.y
    oMat.IndexValue(2) = oCopeXvec.z
    'y axis
    oMat.IndexValue(4) = oCopeYvec.x
    oMat.IndexValue(5) = oCopeYvec.y
    oMat.IndexValue(6) = oCopeYvec.z
    'z axis
    oMat.IndexValue(8) = oCopeZvec.x
    oMat.IndexValue(9) = oCopeZvec.y
    oMat.IndexValue(10) = oCopeZvec.z
    Set GetCopeTransform = oMat
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function

'IsPGSame returns False unless both PGs can be obtained and are the same.
'*************************************************************************
'Function
'
'<IsPGSame>
'
'Abstract
'
'<Determines if Permission Group of 2 objects is same>
'
'Arguments
'
'<2 objs as Object>
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsPGSame(oObj1 As Object, oObj2 As Object) As Boolean

    Const MT = "IsPGSame"
    On Error GoTo ErrorHandler
    
    Dim iIJDObject As IJDObject
    Dim pg1 As Long, pg2 As Long

    IsPGSame = False
    
    If oObj1 Is Nothing Then
        GoTo wrapup
    ElseIf Not TypeOf oObj1 Is IJDObject Then
        GoTo wrapup
    Else
        Set iIJDObject = oObj1
        pg1 = iIJDObject.PermissionGroup
    End If

    If oObj2 Is Nothing Then
        GoTo wrapup
    ElseIf Not TypeOf oObj2 Is IJDObject Then
        GoTo wrapup
    Else
        Set iIJDObject = oObj2
        pg2 = iIJDObject.PermissionGroup
    End If

    If pg1 = pg2 Then
        IsPGSame = True
    End If

wrapup:
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function

'IsFCCleared returns whether the non-RefColl FC relations are set to Nothing
'*************************************************************************
'Function
'
'<IsFCCleared>
'
'Abstract
'
'<Determines whether joint relations are cleared, as when the type was set to Undefined>
'<No check can be made for the RefColl because setting to a definition can occur after
'selection ByRule, which uses the RefColl.  And no check is made for the current Definition
'type either, since the type will always be set when SetRelatedObjects is called.>
'
'Arguments
'
'<FC>  the FC of interest.
'
'Return
'
'<Boolean>
'
'Exceptions
'***************************************************************************
Public Function IsFCCleared(oFCinput As ISPSFrameConnection) As Boolean

    Const MT = "IsFCCleared"
    On Error GoTo ErrorHandler
    
    Dim oObj1 As Object, oObj2 As Object
    Dim oFC As ISPSFrameConnection
    Dim eleEndMemberSystems As IJElements
    
    IsFCCleared = True
    Exit Function
    
    IsFCCleared = False
    
    Set oFC = oFCinput
    oFC.Joint.GetPointOn oObj1, oObj2
    
    If Not oObj1 Is Nothing Then
        GoTo wrapup
    End If
    If Not oObj2 Is Nothing Then
        GoTo wrapup
    End If
    
    Set eleEndMemberSystems = oFC.Joint.EndMemberSystems
    If eleEndMemberSystems.Count <> 1 Then
        GoTo wrapup
    End If
    If Not eleEndMemberSystems.Item(1) Is oFC.MemberSystem Then
        GoTo wrapup
    End If

    IsFCCleared = True

wrapup:
    IsFCCleared = True  'IsFCCleared is obsolete and will be removed soon.
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'
' IsAssemblyConnectionInConflictWithAnother
'
'Abstract
'
' Checks/removes the ports from the custom assembly if another Assembly
'   connection is already attached to the ports. This function then disables
'   the AC and the users only option is to delete it. This should only be
'   the case on a copy operation.
'
'Arguments
'
' StructConn - structural connection custom assembly
'
'Return
'
' True  - Assembly connection is in conflict with an already existing AC
' False - Assembly connection is not in conflict with antoher AC
'
'Exceptions
'***************************************************************************
Public Function IsAssemblyConnectionInConflictWithAnother(oStructConn As IJAppConnection) As Boolean
    Const MT = "IsAssemblyConnectionInConflictWithAnother"
    Const CONST_IJPort As String = "{5CF7C404-546D-11D2-B328-080036024603}"
    
    Dim colPorts As IJElements
    Dim oPort As ISPSSplitAxisPort
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim oAppConn As Object
    Dim attachedAssemblyConnectionCnt As Long
    Dim cnt As Long, nPortCnt As Long
    Dim oCustParent As IJDMemberObjects

    IsAssemblyConnectionInConflictWithAnother = False
    
    ' Retrieve the port inputs of the custom assembly occurrence
    oStructConn.enumPorts colPorts
    ' Verify we do not have an assembly connection already attached to the end ports
    '   because if we do then this asssembly connection needs to have its relations
    '   to the ports severed
    For nPortCnt = 1 To colPorts.Count
        'check if the port is from a member part, otherwise continue
        If TypeOf colPorts.Item(nPortCnt) Is ISPSSplitAxisPort Then
            Set oPort = colPorts.Item(nPortCnt)
            ' Only check end ports because along ports can have many ACs related
            '   to themselves; whereas end ports can only have one AC
            If oPort.portIndex <> SPSMemberAxisAlong Then
                Set pRelationHelper = oPort
                Set pCollectionHelper = pRelationHelper.CollectionRelations(CONST_IJPort, "ConnHasPorts_DEST")
                If Not pCollectionHelper Is Nothing Then
                    attachedAssemblyConnectionCnt = 0
                    For cnt = 1 To pCollectionHelper.Count
                        Set oAppConn = pCollectionHelper.Item(cnt)
                        If TypeOf oAppConn Is IJStructAssemblyConnection Then
                            attachedAssemblyConnectionCnt = attachedAssemblyConnectionCnt + 1
                        End If
                    Next cnt
                    
                    ' If more than one AC attached to a port then clear the relations and
                    '   inform calling function that we discovered a conflict
                    If attachedAssemblyConnectionCnt > 1 Then
                        For cnt = 1 To colPorts.Count
                            oStructConn.removePort colPorts.Item(cnt)
                        Next cnt
                        'remove all child objects. Otherwise
                        'during a paste restore child objects may not be computed
                        'correctly for the paste restored assembly connection - TR#78489
                        Set oCustParent = oStructConn
                        For cnt = 1 To oCustParent.Count
                            oCustParent.Remove cnt
                        Next cnt
                        
                        IsAssemblyConnectionInConflictWithAnother = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next nPortCnt

    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'CheckForCornerBraceAsmConn
'
'Abstract
'Checks if supported and  supporting members are oriented properly
'
'Arguments
'Supported member part, first supporting member part,second supporting member part
'
'Return
'      TRUE - Members are oriented properly
'      FALSE - They are not oriented properly
'Exceptions
'
'***************************************************************************
Public Function CheckForCornerBraceAsmConn(oSupped As ISPSMemberPartPrismatic, oSupping1 As ISPSMemberPartPrismatic, _
oSupping2 As ISPSMemberPartPrismatic) As Boolean
    Const METHOD = "CheckForCornerBraceAsmConn"
    On Error GoTo ErrorHandler

    Dim oYVecSupped As New DVector, oZVecSupped As New DVector
    Dim oYVecSupping1 As New DVector, oZVecSupping1 As New DVector
    Dim oYVecSupping2 As New DVector, oZVecSupping2 As New DVector
    Dim bIsNeeded As Boolean

    If TypeOf oSupped Is ISPSMemberPartCurve Or TypeOf oSupping1 Is ISPSMemberPartCurve _
    Or TypeOf oSupping2 Is ISPSMemberPartCurve Then
        CheckForCornerBraceAsmConn = False
        Exit Function
    End If
    ' check if supping1 and supping2 are at right angles and all three members
    ' are in the same plane and their weak or strong axes are parallel
    If IsMemberAxesAtRightAngles(oSupping1.Axis, oSupping2.Axis) Then ' Make sure supporting members are at right angles
        ' Get Y and Z Vectors of supported member
        Set oYVecSupped = GetYVector(oSupped)
        Set oZVecSupped = GetZVector(oSupped)
        ' Get Y and Z Vectors of 1st supporting member
        Set oYVecSupping1 = GetYVector(oSupping1)
        Set oZVecSupping1 = GetZVector(oSupping1)
        ' Get Y and Z Vectors of 2nd supporting member
        Set oYVecSupping2 = GetYVector(oSupping2)
        Set oZVecSupping2 = GetZVector(oSupping2)
        '
        If AreVectorsParallel(oYVecSupped, oYVecSupping1) Or AreVectorsParallel(oYVecSupped, oZVecSupping1) Then
            ' Weak axis of supported and weak or strong axis of 1st supporting memver are parallel
            If AreVectorsParallel(oYVecSupped, oYVecSupping2) Or AreVectorsParallel(oYVecSupped, oZVecSupping2) Then
                ' Weak axis of supported and weak or strong axis of 2nd supporting member are parallel
                ' That means all members are in the same plane and with proper rotation
                bIsNeeded = True
            End If
        ElseIf AreVectorsParallel(oZVecSupped, oYVecSupping1) Or AreVectorsParallel(oZVecSupped, oZVecSupping1) Then
            ' Strong axis of supported and weak or strong axis of 1st supporting member are parallel
            If AreVectorsParallel(oZVecSupped, oYVecSupping2) Or AreVectorsParallel(oZVecSupped, oZVecSupping2) Then
                ' Strong axis of supported and weak or strong axis of 2nd supporting member are parallel
                ' That means all members are in the same plane and with proper rotation
                bIsNeeded = True
            End If
        End If
    End If
    CheckForCornerBraceAsmConn = bIsNeeded
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function
'*************************************************************************
'Function
'CheckForGussetPlateAsmConn
'
'Abstract
'Checks if supported and supporting member are oriented properly
'
'Arguments
'Supported member part, supporting member part
'
'Return
'      TRUE - Members are oriented properly
'      FALSE - They are not oriented properly
'Exceptions
'
'***************************************************************************

Public Function CheckForGussetPlateAsmConn(oSupped As ISPSMemberPartPrismatic, oSupping As ISPSMemberPartPrismatic) As Boolean
    Const METHOD = "CheckForGussetPlateAsmConn"
    On Error GoTo ErrorHandler
    Dim oYVecSupped As New DVector, oZVecSupped As New DVector
    Dim oYVecSupping As New DVector, oZVecSupping As New DVector
    Dim bIsNeeded As Boolean
    
    'check if Supping and Supported have their weak or strong axes  parallel
    'Get Y and Z Vectors of supported
    Set oYVecSupped = GetYVector(oSupped)
    Set oZVecSupped = GetZVector(oSupped)
    ' Get Y and Z Vectors of supporting
    Set oYVecSupping = GetYVector(oSupping)
    Set oZVecSupping = GetZVector(oSupping)

    If AreVectorsParallel(oYVecSupped, oYVecSupping) Or AreVectorsParallel(oYVecSupped, oZVecSupping) Then
        ' Weak axis of supported and weak or strong axis supporting are parallel
        bIsNeeded = True
    ElseIf AreVectorsParallel(oZVecSupped, oYVecSupping) Or AreVectorsParallel(oZVecSupped, oZVecSupping) Then
        ' Strong axis of supported and weak or strong axis supporting are parallel
        bIsNeeded = True
    End If
    CheckForGussetPlateAsmConn = bIsNeeded
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function


'*************************************************************************
'Function
'GetEndPort
'
'Abstract
'Returns an end port from the input collection of ports
'
'Arguments
'Elements collection
'
'Return
'ISPSSplitAxisEndPort
'
'Exceptions
'
'***************************************************************************

Public Function GetEndPort(oPortCol As IJElements) As ISPSSplitAxisEndPort
    Const METHOD = "GetEndPort"
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim oPort As IJPort
    If Not oPortCol Is Nothing Then
        For i = 1 To oPortCol.Count
            Set oPort = oPortCol.Item(i)
            If TypeOf oPort Is ISPSSplitAxisEndPort Then
                Set GetEndPort = oPort
                Exit For
            End If
        Next i
    End If
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

'*************************************************************************
'Function
'GetCornerGussetInputPorts
'
'Abstract
'Returns the input Ports for the corner gusset assembly connection. The first port
'returned is the supported port, second is primary supporting port and third is
'secondary supporting port. The ports in the input port collection may be in any order.
'If one of the supporting member has the other two point on to them, that is returned as
'primary supporting. Otherwise we check one of the supporting is a column. If so that is
'returned as primary supporting. Otherwise the first supporting in the input collection
'is returned as primary supporting.
'
'Arguments
'Ports collection
'
'Return
'Supported member port, primary supporting port, Secondary supporting port
'
'Exceptions
'
'***************************************************************************

Public Sub GetCornerGussetInputPorts(oPortCol As IJElements, oSuppedPort As ISPSSplitAxisPort, _
oSupping1Port As ISPSSplitAxisPort, oSupping2Port As ISPSSplitAxisPort)
    Const METHOD = "GetCornerGussetInputPorts"
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim oPort As IJPort, oPort1 As IJPort, oPort2 As IJPort
    Dim oSupping1Part As ISPSMemberPartPrismatic
    Dim oSupping2Part As ISPSMemberPartPrismatic
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim dist1#, dist2#, dist11#, dist12#, dist21#, dist22#, normDist#
    Dim X0#, Y0#, Z0#, x1#, y1#, z1#
    Dim oNormal As New DVector, oVec1 As IJDVector, oVec2 As IJDVector
    Dim oLine As IJLine
    
    Dim oConnPoint As IJPoint
    Dim oEndPoint As IJPoint
    
    If Not oPortCol Is Nothing Then
        If oPortCol.Count = 3 Then
            For i = 1 To oPortCol.Count
                If TypeOf oPortCol.Item(i) Is ISPSSplitAxisEndPort Then
                    Set oPort = oPortCol.Item(i)
                    Exit For
                End If
            Next i
        End If
    End If
    If Not oPort Is Nothing Then
        If oPort Is oPortCol.Item(1) Then
            Set oPort1 = oPortCol.Item(2)
            Set oPort2 = oPortCol.Item(3)
        ElseIf oPort Is oPortCol.Item(2) Then
            Set oPort1 = oPortCol.Item(1)
            Set oPort2 = oPortCol.Item(3)
        Else
            Set oPort1 = oPortCol.Item(1)
            Set oPort2 = oPortCol.Item(2)
        End If
        
        If (TypeOf oPort1 Is ISPSSplitAxisAlongPort) And (TypeOf oPort2 Is ISPSSplitAxisAlongPort) Then
            Set oSuppedPort = oPort
            Set oSupping1Port = oPort1
            Set oSupping2Port = oPort2
            Set oSuppedPart = oSuppedPort.Part
            Set oSupping1Part = oSupping1Port.Part
            Set oSupping2Part = oSupping2Port.Part
            
            'get normal vector to the plane of supped,supping1 and supping2
            Set oVec1 = GetXVector(oSuppedPart)
            Set oVec2 = GetXVector(oSupping1Part)
            
            'get the normal to the plane containing supped and supping1
            Set oNormal = oVec1.Cross(oVec2)
            oNormal.Length = 1
            
            'get the connection point
            Set oConnPoint = oSuppedPort.Port.Geometry
            oConnPoint.GetPoint X0, Y0, Z0
            
            'get closest end to oConnPoint for first supporting
            Set oEndPoint = oSupping1Part.PointAtEnd(SPSMemberAxisStart)
            oEndPoint.GetPoint x1, y1, z1
            'vector from connection point to end point of supporting1
            oVec1.Set x1 - X0, y1 - Y0, z1 - Z0
            'distance from connection point
            dist11 = oVec1.Length

            Set oEndPoint = oSupping1Part.PointAtEnd(SPSMemberAxisEnd)
            oEndPoint.GetPoint x1, y1, z1
            oVec2.Set x1 - X0, y1 - Y0, z1 - Z0
            'distance from connection point
            dist12 = oVec2.Length
           
            If dist11 < dist12 Then ' get the lesser distance
                'distance from connection point in the direction of plane normal
                normDist = oVec1.Dot(oNormal)
                'gets distance from connection point along the plane
                dist1 = Sqr(dist11 * dist11 - normDist * normDist)
            Else
                'distance from connection point in the direction of plane normal
                normDist = oVec2.Dot(oNormal)
                'gets distance from connection point along the plane
                dist1 = Sqr(dist12 * dist12 - normDist * normDist)
            End If
            
            'get closest end to oConnPoint from second supporting
            Set oEndPoint = oSupping2Part.PointAtEnd(SPSMemberAxisStart)
            oEndPoint.GetPoint x1, y1, z1
            'vector from connection point to end point of supporting1
            oVec1.Set x1 - X0, y1 - Y0, z1 - Z0
            'distance from connection point
            dist21 = oVec1.Length

            Set oEndPoint = oSupping2Part.PointAtEnd(SPSMemberAxisEnd)
            oEndPoint.GetPoint x1, y1, z1
            oVec2.Set x1 - X0, y1 - Y0, z1 - Z0
            'distance from connection point
            dist22 = oVec2.Length
           
            If dist21 < dist22 Then ' get the lesser distance
                'distance from connection point in the direction of plane normal
                normDist = oVec1.Dot(oNormal)
                'gets distance from connection point along the plane
                dist2 = Sqr(dist21 * dist21 - normDist * normDist)
            Else
                'distance from connection point in the direction of plane normal
                normDist = oVec2.Dot(oNormal)
                'gets distance from connection point along the plane
                dist2 = Sqr(dist22 * dist22 - normDist * normDist)
            End If
            
            If Abs(dist1 - dist2) < distTol Then
                ' both supporting members have their ends at the connection point
                'check one of them is column. If so make it supporting1
                If oSupping1Part.MemberType.TypeCategory <> 2 And oSupping2Part.MemberType.TypeCategory = 2 Then
                    Set oSupping1Port = oPort2
                    Set oSupping2Port = oPort1
                End If
            ElseIf dist2 > dist1 Then
                ' supporting2 has its end farhter from the connection point
                'probabbly supporting1 and supported are point on to supporting2
                'make supporting2 as supporting1
                Set oSupping1Port = oPort2
                Set oSupping2Port = oPort1
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


'*************************************************************************
'Function
'GetSpliceMiterInputPorts
'
'Abstract
'Returns an end port from the input collection of ports
'
'Arguments
'Elements collection
'
'Return
'ISPSSplitAxisPort,ISPSSplitAxisPort
'
'Exceptions
'
'***************************************************************************

Public Sub GetSpliceMiterInputPorts(oPortCol As IJElements, oSupped1Port As ISPSSplitAxisPort, oSupped2Port As ISPSSplitAxisPort)
    Const METHOD = "GetSpliceMiterInputPorts"
    On Error GoTo ErrorHandler
    Dim oNorm As IJDVector, oVec1 As New DVector, oVec2 As New DVector
    Dim X0#, Y0#, Z0#, x1#, y1#, z1#
    Dim biggestValue#
    
    If Not oPortCol Is Nothing Then
        If oPortCol.Count = 2 Then
            If (TypeOf oPortCol.Item(1) Is ISPSSplitAxisEndPort) And (TypeOf oPortCol.Item(2) Is ISPSSplitAxisEndPort) Then
                Set oSupped1Port = oPortCol.Item(1)
                Set oSupped2Port = oPortCol.Item(2)
            End If
        End If
    End If

    If oSupped1Port Is Nothing Or oSupped2Port Is Nothing Then
        Exit Sub
    End If

    If oSupped1Port.portIndex = SPSMemberAxisStart Then
        oSupped1Port.Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
    Else
        oSupped1Port.Part.Axis.EndPoints x1, y1, z1, X0, Y0, Z0
    End If
    'vector from connected end to the other end of supped1
    oVec1.Set x1 - X0, y1 - Y0, z1 - Z0
    oVec1.Length = 1 'normalize
    
    If oSupped2Port.portIndex = SPSMemberAxisStart Then
        oSupped2Port.Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
    Else
        oSupped2Port.Part.Axis.EndPoints x1, y1, z1, X0, Y0, Z0
    End If
    'vector from connected end  to the other end of supped2
    oVec2.Set x1 - X0, y1 - Y0, z1 - Z0
    oVec2.Length = 1 'normalize

    Set oNorm = oVec1.Cross(oVec2)
    If oNorm.Length < distTol Then

        'vec1 and vec2 are colinear.  find largest single absolute value of oVec1
        'if it is positive, then part1 is on right of part2, and we leave order the same.
        
        biggestValue = oVec1.x
        If Abs(oVec1.y) > Abs(biggestValue) Then
            biggestValue = oVec1.y
        End If
        If Abs(oVec1.z) > Abs(biggestValue) Then
            biggestValue = oVec1.z
        End If

    ' else not colinear.
    ' looking at the pair, which is on the left ?
    
    ' left is defined by looking at the pair along a global axis, mostly Z-axis if possible.
    ' if vec1 cross vec2 is positive, then vec2 is on the left.

    Else

        oNorm.Length = 1
        
        ' use Z value of normal unless the members are really in a vertical plane.
        If Abs(oNorm.z) < distTol Then     ' is vertical plane
        
            biggestValue = oNorm.x
            If Abs(oNorm.y) > Abs(biggestValue) Then
                biggestValue = oNorm.y
            End If

        Else
            biggestValue = oNorm.z
        End If
    
    End If

    If biggestValue > 0 Then
        Set oSupped1Port = oPortCol.Item(2)
        Set oSupped2Port = oPortCol.Item(1)
    End If
        
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Function CheckForReadOnlyAccess(obj As IJDObject) As Boolean
Const METHOD = "CheckForReadOnlyAccess"
On Error GoTo ErrHandler
    
    CheckForReadOnlyAccess = False
    
    If obj.AccessControl And acUpdate Then
        CheckForReadOnlyAccess = False
    Else
        CheckForReadOnlyAccess = True
    End If
    
Exit Function
ErrHandler:
    HandleError MODULE, METHOD
End Function

'Sub GetCommonConnection(oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem,
'                               ByRef oPointConn As IJPoint, ByRef lngRelated As Long)
'
'if oMS1 is related to oMS2 with a FrameConnection or a SplitConnection, this will return that connection.
'if they are not related, then oPointConn is returned as Nothing.
'
'lngRelated     Meaning
' 0             not related
' 1             oMS2 is dependent on oMS1.  oPointConn is a FC at an end of oMS2
' 2             oMS1 is dependent on oMS2.  oPointConn is a FC at an end of oMS1
' 3             oPointConn is a split connection.  dependency not clear !
'
Public Sub GetCommonConnection(oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem, ByRef oPointConn As IJPoint, ByRef lngRelated As Long)
    Const METHOD = "GetCommonConnection"
    On Error GoTo ErrHandler

    Dim oFC As ISPSFrameConnection
    Dim IHStatus As SPSFCInputHelperStatus
    Dim O1 As Object, O2 As Object
    
    Dim eleSplits As IJElements, eleInputs As IJElements
    Dim ii As Long, jj As Long
    Dim iSC As ISPSSplitMemberConnection
    
    Set oPointConn = Nothing
    lngRelated = 0

    'check for invalid inputs.
    If oMS1 Is Nothing Or oMS2 Is Nothing Then
        GoTo wrapup
    End If
    If oMS1 Is oMS2 Then
        GoTo wrapup
    End If

    'check if they are related via a common FC
    'first check if oMS2 is dependent on oMS1
    Set oFC = oMS2.FrameConnectionAtEnd(SPSMemberAxisStart)
    oFC.InputHelper.GetRelatedObjects oFC, O1, O2


    'oMS2 does depend on oMS1
    If oMS1 Is O1 Or oMS1 Is O2 Then
        Set oPointConn = oFC
        lngRelated = 1          'oMS2 does depend on oMS1
        GoTo wrapup
    End If
    Set O1 = Nothing
    Set O2 = Nothing

    Set oFC = oMS2.FrameConnectionAtEnd(SPSMemberAxisEnd)
    oFC.InputHelper.GetRelatedObjects oFC, O1, O2

    If oMS1 Is O1 Or oMS1 Is O2 Then
        Set oPointConn = oFC
        lngRelated = 1          'oMS2 does depend on oMS1
        GoTo wrapup
    End If

    'second check if oMS1 is dependent on oMS2
    Set oFC = oMS1.FrameConnectionAtEnd(SPSMemberAxisStart)
    oFC.InputHelper.GetRelatedObjects oFC, O1, O2

    If oMS2 Is O1 Or oMS2 Is O2 Then
        Set oPointConn = oFC
        lngRelated = 2          'oMS1 does depend on oMS2
        GoTo wrapup
    End If
    Set O1 = Nothing
    Set O2 = Nothing

    Set oFC = oMS1.FrameConnectionAtEnd(SPSMemberAxisEnd)
    oFC.InputHelper.GetRelatedObjects oFC, O1, O2

    If oMS2 Is O1 Or oMS2 Is O2 Then
        Set oPointConn = oFC
        lngRelated = 2          'oMS1 does depend on oMS2
        GoTo wrapup
    End If
    Set O1 = Nothing
    Set O2 = Nothing

    Set eleSplits = oMS1.SplitConnections
    For ii = 1 To eleSplits.Count
        Set iSC = eleSplits(ii)
        Set eleInputs = iSC.InputObjects
        For jj = 1 To eleInputs.Count
            If eleInputs(jj) Is oMS2 Then
                Set oPointConn = iSC
                lngRelated = 3  ' we return 3 when it is a split conn.
                GoTo wrapup
            End If
        Next jj
    Next ii

    ' It would be redundant to check oMS2's list of split Connections...!

wrapup:
    Exit Sub

ErrHandler:
    HandleError MODULE, METHOD
End Sub

'This function gets supporting members for a FC via the RefColl.  No data integrity checking is made, and if the joint
'is constrained by grid lines, those are not returned either.
'
Public Sub GetRelatedObjectsForFC(oFC As Object, ByRef oObj1 As Object, ByRef oObj2 As Object)
    Const METHOD = "GetRelatedObjectsForFC"

    On Error GoTo ErrHandler
    Dim nRC As Long
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    
    Set oObj1 = Nothing
    Set oObj2 = Nothing

    Set oRC = GetRefCollNoCreate(oFC)
    If oRC Is Nothing Then
        Exit Sub
    End If

    nRC = oRC.IJDEditJDArgument.GetCount
    
    'EntityByIndex(1) is always the supported member.
    
    If nRC >= 2 Then
        Set oObj1 = oRC.IJDEditJDArgument.GetEntityByIndex(2)
    End If
    
    'AxisEnd with shared joint uses EntityByIndex(3) for the supporting AxisEndPort
    'VCB uses EntityByIndex(3) for the secondary supporting member
    If nRC >= 3 Then
        Set oObj2 = oRC.IJDEditJDArgument.GetEntityByIndex(3)
        If Not oObj2 Is Nothing Then
            If Not TypeOf oObj2 Is ISPSMemberSystem Then
                Set oObj2 = Nothing
            End If
        End If
    End If

    Exit Sub

ErrHandler:
    HandleError MODULE, METHOD
    Err.Clear
End Sub

Public Function GeometricalCondition(oFC As ISPSFrameConnection, oSuppingCurve As IJCurve) As Long

    ' evaluate the supported member's logical axis against the supporting curve object.
    '
    ' values returned are according to the following:
    ' bit 0 is set if PointOn to supping curve
    ' bit 1 is set if PointOn at either end of supping curve
    ' bit 2 is set if parallel
    Const METHOD = "GeometricalCondition"

    Dim Value As Long                                   ' return value
    Dim distTol As Double
    Dim FCx As Double, FCy As Double, FCz As Double     ' FC end of supping logical axis
    Dim FC2x As Double, FC2y As Double, FC2z As Double  ' non-FC end of supping logical axis
    Dim portIndex As SPSMemberAxisPortIndex
    Dim iSuppedLogicalAxis As ISPSLogicalAxis
    Dim SingSx As Double, SingSy As Double, SingSz As Double, SingEx As Double, SingEy As Double, SingEz As Double

    On Error GoTo ErrorHandler

    Value = 0
    
    portIndex = oFC.WPO.portIndex
    Set iSuppedLogicalAxis = oFC.MemberSystem.LogicalAxis
    
    If portIndex = SPSMemberAxisStart Then
        iSuppedLogicalAxis.GetLogicalStartPoint FCx, FCy, FCz
        iSuppedLogicalAxis.GetLogicalEndPoint FC2x, FC2y, FC2z
    Else
        iSuppedLogicalAxis.GetLogicalStartPoint FC2x, FC2y, FC2z
        iSuppedLogicalAxis.GetLogicalEndPoint FCx, FCy, FCz
    End If
    
    oSuppingCurve.EndPoints SingSx, SingSy, SingSz, SingEx, SingEy, SingEz
    
    distTol = 0.000001

    ' check for end-matched first.  if not, check for pointOn.
    If Abs(FCx - SingSx) < distTol And Abs(FCy - SingSy) < distTol And Abs(FCz - SingSz) < distTol Then
        Value = 3       ' set bits 0 and 1
    ElseIf Abs(FCx - SingEx) < distTol And Abs(FCy - SingEy) < distTol And Abs(FCz - SingEz) < distTol Then
        Value = 3       ' set bits 0 and 1
    ElseIf oSuppingCurve.IsPointOn(FCx, FCy, FCz) Then
        Value = 1       ' set only bit 0
    End If

    ' parallel can only exist if both are lines...
    If TypeOf iSuppedLogicalAxis Is IJLine And TypeOf oSuppingCurve Is IJLine Then
        Dim oVecSupped As New DVector, oVecSupping As New DVector
        Dim oVecCrossP As IJDVector
            
        oVecSupped.Set FC2x - FCx, FC2y - FCy, FC2z - FCz
        oVecSupping.Set SingEx - SingSx, SingEy - SingSy, SingEz - SingSz
        
        ' if either is degen, not parallel
        If oVecSupped.Length > distTol And oVecSupping.Length > distTol Then
            oVecSupped.Length = 1
            oVecSupping.Length = 1
            Set oVecCrossP = oVecSupped.Cross(oVecSupping)
            If oVecCrossP.Length < distTol Then
                Value = Value + 4           ' set bit 2 for parallel
            End If
        End If
    End If

    GeometricalCondition = Value
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    GeometricalCondition = 0
    Exit Function

End Function

'*************************************************************************
'Function
'GetSectionWidthAndDepth
'
'Abstract
'Returns a crosssection width and depth
'
'Arguments
'MemberPart
'
'Return
'Double width, double depth
'
'Exceptions
'
'***************************************************************************
Sub GetSectionWidthAndDepth(oMemberPart As ISPSMemberPartPrismatic, dWidth As Double, dDepth As Double)
Const METHOD = "GetSectionWidthAndDepth"
    'MsgBox "in " + METHOD
    Dim oXSection As ISPSCrossSection
    Dim oAttributes As IJDAttributes
    
    Set oXSection = oMemberPart.CrossSection
    Set oAttributes = oXSection.Definition
    
    dWidth = oAttributes.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    dDepth = oAttributes.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    
    Exit Sub
        
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    Exit Sub
    
End Sub



'*************************************************************************
'Function
'IsSurfaceTypeAcceptable
'
'Abstract
' checks to see if a suraface tpe is acceptable.
' supports IJPlane, IJSurface or IJDynamicSurfaceFInd
' Used by GenSurfAC
'Arguments
'port
'
'Return
'Boolean
'
'***************************************************************************
Function IsSurfaceTypeAcceptable(oSurface As Object) As Boolean
Const METHOD = "IsSurfaceTypeAcceptable"
    Dim bIsGood As Boolean
    
    bIsGood = False
    
    If TypeOf oSurface Is IJSurface Then
        bIsGood = True
    ElseIf TypeOf oSurface Is IJPlane Then
        bIsGood = True
    ElseIf TypeOf oSurface Is IJDynamicSurfaceFind Then
        bIsGood = True
    End If
    
    IsSurfaceTypeAcceptable = bIsGood
    Exit Function
        
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    Exit Function
    
End Function

'*************************************************************************
'Abstract
'<Gets the line of the MemberPartPrismatic geometric Mid Section. Shoiuld be used instead of a hard-coded CP
' as sometimes the CP line may not intersect the surface and best guess is the geometric center>
'
'Arguments
'<Member part as MemberPartPrismatic>
'
'Return
'<line as IJLine>
'***************************************************************************
Public Function GetLineFromCrossSectionMid(oMemb As ISPSMemberPartPrismatic) As IJLine
    Dim dWidth As Double, dDepth As Double
    Dim oGeomFactory As New GeometryFactory
    Dim oMidPos As IJPoint
    Dim dMidX As Double, dMidY As Double, dMidZ As Double
    Dim oMat As IJDT4x4
    Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
        
    Call GetSectionWidthAndDepth(oMemb, dWidth, dDepth)
    'For the RAD2D symbol which is in xy plane
    dMidX = dWidth / 2#
    dMidY = dDepth / 2#
    dMidZ = 0#
    
    Set oMidPos = oGeomFactory.Points3d.CreateByPoint(Nothing, dMidX, dMidY, dMidZ)
    
    oMemb.Rotation.GetTransform oMat
    oMat.IndexValue(12) = 0#
    oMat.IndexValue(13) = 0#
    oMat.IndexValue(14) = 0#
    Set oMat = CreateCSToMembTransform(oMat, oMemb.Rotation.Mirror)
    Set oMidPos = oMat.TransformPosition(oMidPos)
    oMemb.Axis.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    dStartX = dStartX + oMidPos.x
    dStartY = dStartY + oMidPos.y
    dStartZ = dStartZ + oMidPos.z
    dEndX = dEndX + oMidPos.x
    dEndY = dEndY + oMidPos.y
    dEndZ = dEndZ + oMidPos.z
    
    Set GetLineFromCrossSectionMid = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ)
    
    Set oGeomFactory = Nothing
End Function
'*************************************************************************
'Abstract
'<Gets the connection point on the supporting member axis. The function gets the tangent at the
'supported member end, creates a line from it and then intersects this line with supporting axis>
'
'
'Arguments
'<Supporting Member part as MemberPartPrismatic>
'<Supported Member part as MemberPartPrismatic>
'<Supported end  as SPSMemberAxisPortIndex>
'
'Return
'<Position as IJDPosition> may be nothing if there is no intersection
'***************************************************************************


Public Function GetConnectionPositionOnSupping( _
                            oSupping As ISPSMemberPartCommon, _
                            oSupped As ISPSMemberPartCommon, _
                            iEnd As SPSMemberAxisPortIndex) As IJDPosition
Const METHOD = "GetConnectionPositionOnSupping"
    On Error GoTo ErrorHandler
    
    Dim x As Double, y As Double, z As Double
    Dim uX As Double, uY As Double, uZ As Double
    Dim oPos As IJDPosition
    Dim oLine As IJLine
    Dim transf1 As DT4x4
    Dim transf2 As DT4x4
    Dim oGeomFactory As New GeometryFactory
    
    oSupped.PointAtEnd(iEnd).GetPoint x, y, z
    oSupped.Rotation.GetTransformAtPosition x, y, z, transf1, transf2
    'ignore transf2 as this is the end of the part
    
    'get the tanget at this end of the supported member
    uX = transf1.IndexValue(0)
    uY = transf1.IndexValue(1)
    uZ = transf1.IndexValue(2)

   
    'create a line of 1 m length along the tangent
    Set oLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, uX, uY, uZ, 1)

   
    'we need to intersect this line with the supporting axis
    Set oPos = GetCurveLineIntersection(oSupping.Axis, oLine)
    
    
    If Not oPos Is Nothing Then
        Set GetConnectionPositionOnSupping = oPos
    End If
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function



'*************************************************************************
'Function
'
'<GetTangentLineAtCP>
'
'Abstract
'
'<Gets the tangent line for start or end of the MemberPartPrismatic at a given  cardinal point>
'<the length of the line is hard coded to 100m. >
'
'Arguments
'
'<Member part as MemberPartPrismatic , Cardinal point as integer, bStartPoint as boolean>
'
'Return
'
'<line as IJLine>
'
'Exceptions
'***************************************************************************
Public Function GetTangentLineAtCP(oMemb As ISPSMemberPartPrismatic, cp As Long, bStartPoint As Boolean) As IJLine
  Const MT = "GetTangentLineAtCP"
    On Error GoTo ErrorHandler
    Dim curCP As Long
    Dim oProfileBO As ISPSCrossSection
    Dim pGeometryFactory As New GeometryFactory
    Dim oPos As New DPosition
    Dim oMat As IJDT4x4
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double
    Dim oVec As New DVector
    
    Set oProfileBO = oMemb.CrossSection
    curCP = oProfileBO.CardinalPoint
    
    oProfileBO.GetCardinalPointOffset curCP, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP
    oProfileBO.GetCardinalPointOffset cp, xOffset, yOffset 'Returns Rad x and y of CP requested
    oPos.x = xOffset - xOffsetCP ' x offset of current cp from requested CP
    oPos.y = yOffset - yOffsetCP ' y offset of current cp from requested CP
    oPos.z = 0
    
    oMemb.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    If bStartPoint Then
        oMemb.Rotation.GetTransformAtPosition sX, sY, sZ, oMat, Nothing
        oVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        oVec.Length = 1
    Else
        oMemb.Rotation.GetTransformAtPosition eX, eY, eZ, oMat, Nothing
        oVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        oVec.Length = 1
        oVec.[Scale] -1 ' reverse the direction
    End If
    
    
    Set oMat = CreateCSToMembTransform(oMat, oMemb.Rotation.Mirror)
    Set oPos = oMat.TransformPosition(oPos)

    'if it is the end of the member , then make the line start 1 m away towards the start end of the member
    If bStartPoint = False Then
        oVec.Length = 100
        Set oPos = oPos.Offset(oVec)
        oVec.[Scale] -1 ' reverse the direction
        'so the line that we create is in the same direction as the member
    End If

   
    Set GetTangentLineAtCP = pGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, oPos.x, oPos.y, oPos.z, oVec.x, oVec.y, oVec.z, 100)
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT
    
End Function

'*************************************************************************
'Function
'
'<GetTangentLineAtCPAndPosition>
'
'Abstract
'
'<Gets the tangent to the MemberPartPrismatic at a given  cardinal point and location>
'<  Returns line with ends  20m apart on either side of input position and along the tangent >
'
'Arguments
'
'<Member part as MemberPartPrismatic , Cardinal point as integer, x as double, y as double, z as double>
'
'Return
'
'<line as IJLine>
'
'Exceptions
'***************************************************************************
Public Function GetTangentLineAtCPAndPosition(oMemb As ISPSMemberPartPrismatic, cp As Long, x As Double, y As Double, z As Double) As IJLine
  Const MT = "GetTangentLineAtCPAndPosition"
    On Error GoTo ErrorHandler
    Dim curCP As Long
    Dim oProfileBO As ISPSCrossSection
    Dim pGeometryFactory As New GeometryFactory
    Dim oPos As New DPosition, oPos1 As New DPosition
    Dim oMat As IJDT4x4
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double
    Dim oVec As New DVector
    
    Set oProfileBO = oMemb.CrossSection
    curCP = oProfileBO.CardinalPoint
    
    oProfileBO.GetCardinalPointOffset curCP, xOffsetCP, yOffsetCP 'Returns Rad x and y of the current CP
    oProfileBO.GetCardinalPointOffset cp, xOffset, yOffset 'Returns Rad x and y of CP requested
    oPos.x = xOffset - xOffsetCP ' x offset of current cp from requested CP
    oPos.y = yOffset - yOffsetCP ' y offset of current cp from requested CP
    oPos.z = 0
    
    oMemb.Rotation.GetTransformAtPosition x, y, z, oMat, Nothing
    oVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    oVec.Length = 1
    
    
    Set oMat = CreateCSToMembTransform(oMat, oMemb.Rotation.Mirror)
    Set oPos = oMat.TransformPosition(oPos)

    oVec.Length = -10
    Set oPos = oPos.Offset(oVec)
    oVec.[Scale] -2 ' reverse the direction and double the length
    Set oPos1 = oPos.Offset(oVec)
    'so at the end we have 2 positions 20m apart on either side of input position and along the tangent
   
    Set GetTangentLineAtCPAndPosition = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos.x, oPos.y, oPos.z, oPos1.x, oPos1.y, oPos1.z)
    
    Exit Function
ErrorHandler:      HandleError MODULE, MT
    
End Function


Public Function GetCurve(oSupping As ISPSMemberPartPrismatic, oSupped As ISPSMemberPartPrismatic, _
iEnd As SPSMemberAxisPortIndex) As IJCurve
Const METHOD = "GetCurve"
    On Error GoTo ErrorHandler
    
    Dim oConnPoint As IJPoint
    Dim x As Double, y As Double, z As Double
    Dim uX As Double, uY As Double, uZ As Double
    Dim transf1 As DT4x4, transf2 As DT4x4
    Dim oGeomFactory As New GeometryFactory
    Dim oLine As IJLine
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingPart As ISPSMemberPartPrismatic
    Dim oCurve As IJCurve
    Dim oPos As IJDPosition
    
    Set oCurve = oSupping.Axis
    
    If (TypeOf oCurve Is IJArc) Or (TypeOf oCurve Is IJLine) Then
        Set GetCurve = oCurve
        Exit Function
    End If
    
    If TypeOf oCurve Is IJComplexString Then
        'need to check for error here as the geometry may be a line
        Set oConnPoint = oSupped.AxisPort(iEnd).Port.Geometry
        oConnPoint.GetPoint x, y, z
        
        oSupped.Rotation.GetTransformAtPosition x, y, z, transf1, transf2
        'ignore transf2 as this is the end of the part
        
        'get the tanget at this end of the supported member
        uX = transf1.IndexValue(0)
        uY = transf1.IndexValue(1)
        uZ = transf1.IndexValue(2)
    
       
        'create a line of 1 m length along the tangent
        Set oLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, uX, uY, uZ, 1)
        
        'we need to intersect this line with the supporting axis
        Set oPos = GetCurveLineIntersection(oCurve, oLine)

        If Not oPos Is Nothing Then
            Dim oCmplxCurve As IJComplexString
            Dim colCurves As IJElements
            Dim ii As Long, cnt As Long
            Set oCmplxCurve = oCurve
            oCmplxCurve.GetCurves colCurves
            cnt = colCurves.Count
            For ii = 1 To colCurves.Count
                Set oCurve = colCurves.Item(ii)
                If oCurve.IsPointOn(oPos.x, oPos.y, oPos.z) Then
                    Set GetCurve = oCurve
                    Exit Function
                End If
            Next
        End If
    End If
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function
'*************************************************************************
'Function
'
'<GetAxisCurveAtPosition>
'
'Abstract
'
'<Gets the curve segment at the specified x,y,z position on the member>
'
'
'Arguments
'
'<x as double , y as double, z as double, oMemb as ISPSMemberPartPrismatic>
'
'Return
'oCurve as IJCurve
'Exception
'Caller needs to handle exceptions

Public Function GetAxisCurveAtPosition(x As Double, y As Double, z As Double, oMemb As ISPSMemberPartCommon) As IJCurve
Const METHOD = "GetCurveAtEnd"
    On Error GoTo ErrorHandler
    Dim ii As Long
    Dim cnt As Long
    Dim oCurve As IJCurve
    Dim colCurves As IJElements
    Dim oCmplxCurve As IJComplexString
    
    If Not oMemb Is Nothing Then
        Set oCurve = oMemb.Axis
    Else
        Exit Function
    End If
    If Not oCurve Is Nothing Then
        If (TypeOf oCurve Is IJArc) Or (TypeOf oCurve Is IJLine) Then
            Set GetAxisCurveAtPosition = oCurve
            Exit Function
        ElseIf TypeOf oCurve Is IJComplexString Then
            Set oCmplxCurve = oCurve
            oCmplxCurve.GetCurves colCurves
            cnt = colCurves.Count
            For ii = 1 To colCurves.Count
                Set oCurve = colCurves.Item(ii)
                If oCurve.IsPointOn(x, y, z) Then
                    Set GetAxisCurveAtPosition = oCurve
                    Exit Function
                End If
            Next
        End If
    End If
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Function

'*************************************************************************
'
'<GetCurveLineIntersection>
'
'Abstract
'
'<Intersects the line with curve and returns the intersection point which is closest
'to the specified end of the line. If no end is specified, start end is assumed>
'
'Arguments
'
'<oCurve As IJCurve, oLine As IJLine, Optional bStartEnd As Boolean = True >
'
'Return
'
'<IJDPosition>
'
'Exceptions
'*************************************************************************
Public Function GetCurveLineIntersection(oCurve As IJCurve, oLine As IJLine, Optional bStartEnd As Boolean = True) As IJDPosition
Const METHOD = "GetCurveLineIntersection"
    On Error GoTo ErrorHandler
    
    Dim oStartPos As New DPosition, oEndPos As New DPosition, oInterscnPos As New DPosition
    Dim oTmpLine As IJLine
    Dim oGeomFactory As New GeometryFactory
    Dim numIntrscn As Long, numOverlaps As Long, retCode As Geom3dIntersectConstants
    Dim idx As Long
    Dim oRefPos As IJDPosition
    Dim minDistance#, dist#
    Dim x#, y#, z#, inX#, inY#, inZ#
    Dim intrscnPoints() As Double
    
    
    oLine.GetStartPoint x, y, z
    oStartPos.Set x, y, z
    
    oLine.GetEndPoint x, y, z
    oEndPos.Set x, y, z
    
    If bStartEnd = True Then
        Set oRefPos = oStartPos
    Else
        Set oRefPos = oEndPos
    End If
    
    Set oTmpLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStartPos.x, oStartPos.y, oStartPos.z, _
    oEndPos.x, oEndPos.y, oEndPos.z)
    oTmpLine.Infinite = True
    oCurve.Intersect oTmpLine, numIntrscn, intrscnPoints, numOverlaps, retCode
    If numIntrscn > 0 Then
        minDistance = 10000#  'default to a large number
        For idx = 1 To numIntrscn
            oInterscnPos.Set intrscnPoints(idx * 3 - 3), intrscnPoints(idx * 3 - 2), intrscnPoints(idx * 3 - 1)
            dist = oInterscnPos.DistPt(oRefPos)
            If dist < minDistance Then
                minDistance = dist
                Set GetCurveLineIntersection = oInterscnPos.Clone()
            End If
        Next
    Else 'no intersection
        'get minimum distance position on the curve
    
        oCurve.DistanceBetween oLine, minDistance#, x, y, z, inX, inY, inZ
        oInterscnPos.Set x, y, z
        Set GetCurveLineIntersection = oInterscnPos
    End If
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD

End Function

'*************************************************************************
'Abstract: IsOperatorTypeValidForTrim finds out if the given operator type is valid for Surface Trim
'
'Arguments: operator - Object
'
'Return: Boolean
'
'Overview: Checks type of the Operator- Should be one of the following
'          1. If port then its connectable should be an assy child which includes all parts (or detailing objects)
'               and Port type should be face and,
'          2. Geometry of the operator should support plane or surface or dynamicsurfacefind
'*************************************************************************
Public Function IsOperatorTypeValidForTrim(oOperator As Object) As Boolean
Const METHOD = "IsPortFace"
    On Error GoTo ErrorHandler

    IsOperatorTypeValidForTrim = False
    
    Dim oPort As IJPort
    Dim oSupgObj As Object
    
    'First do the validation by type
    If TypeOf oOperator Is IJPort Then
        Dim oConnectable As Object
        Dim ePortType As PortType
        Set oPort = oOperator
        ePortType = oPort.Type
        
        Set oConnectable = oPort.Connectable
        If TypeOf oConnectable Is IJAssemblyChild Then
            ' the ports connectable is an assy child
            If PortFace = ePortType Then
                'The port is a face
                Set oSupgObj = oPort.Geometry
            Else
                GoTo AfterGeometryValidation
            End If
        Else
            GoTo AfterGeometryValidation
        End If
    Else
       Set oSupgObj = oOperator
    End If

    If Not oSupgObj Is Nothing Then
        If TypeOf oSupgObj Is IJPlane Or TypeOf oSupgObj Is IJSurface Or TypeOf oSupgObj Is IJDynamicSurfaceFind Then
            'Operator's geometry supports one of the surafce IFs
            IsOperatorTypeValidForTrim = True
        End If
    End If

AfterGeometryValidation:
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

'*************************************************************************
'Function:
'
' GetRefCollNoCreate
'
'Abstract
'
' Returns the existing ReferencesCollection Object attached to the
'   Frame Connection. This function differs from the common GetRefColl
'   function where it will not create a reference collection if the
'   reference collection does not exist.
'
'Arguments:
'
' FrameConnection business object
'
'Return:
'
' refcoll object As IJDReferencesCollection
'
'Exceptions
'***************************************************************************
Public Function GetRefCollNoCreate(pFC As Object) As IJDReferencesCollection
    Const MT = "GetRefCollNoCreate"
    Const INTERFACE_IJSmartOccurrence = "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}"
    Const RELATION_Role = "toArgs_O"
    Const RELATION_Name = "RC"
   
    On Error GoTo ErrorHandler
    Set GetRefCollNoCreate = Nothing
    
    ' Traverse the relation from the SO to the RefColl
   
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    
    If Not pFC Is Nothing Then
        Set pRelationHelper = pFC
        Set pCollectionHelper = pRelationHelper.CollectionRelations(INTERFACE_IJSmartOccurrence, RELATION_Role)
        If Not pCollectionHelper Is Nothing Then
            If pCollectionHelper.Count = 1 Then
                Set GetRefCollNoCreate = pCollectionHelper.Item(RELATION_Name)
            End If
        End If
    End If

    Exit Function
 
ErrorHandler:
    HandleError MODULE, MT
End Function

'*************************************************************************
'Function
'IsSurfaceTypeAcceptableForSurfaceTrim
'
'Abstract
' checks to see if a suraface tpe is acceptable for SUrfaceTrim.
' supports IJPlane, IJSurface or IJDynamicSurfaceFInd
' Used by GenSurfAC
'Arguments
'port from member
' Supporting Surface
'Return
'Boolean
'
'***************************************************************************
Function IsSurfaceTypeAcceptableForSurfaceTrim(oSupdPort As Object, oSurface As Object) As Boolean
Const METHOD = "IsSurfaceTypeAcceptable"
    IsSurfaceTypeAcceptableForSurfaceTrim = False
    
    'First check based on surface type
    If Not IsOperatorTypeValidForTrim(oSurface) Then
        Exit Function
    End If
    
    'Now check based on gemetric condition
    Dim oMemberFactory As New SPSMemberFactory
    Dim oFeatureServices As ISPSMemberFeatureServices
    Dim WhichEnd As SPSMemberAxisPortIndex
    Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices
    Dim allOK As Boolean
    Dim oSplitAxisPort As ISPSSplitAxisPort
    Dim oSupdPart As ISPSMemberPartPrismatic
    
    allOK = False
    Set oSplitAxisPort = oSupdPort
    If Not oSplitAxisPort Is Nothing Then
        Set oSupdPart = oSplitAxisPort.Part
        'Now do the geometric validation
        oFeatureServices.GetTrimSurfacePartGeometricCondition oSupdPart, oSurface, WhichEnd, allOK
    End If
    
    Set oMemberFactory = Nothing
    If allOK = False Then
        Exit Function
    End If
    
    IsSurfaceTypeAcceptableForSurfaceTrim = True
    Exit Function
        
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    Exit Function
    
End Function
