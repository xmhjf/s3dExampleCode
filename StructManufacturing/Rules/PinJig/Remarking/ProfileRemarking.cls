VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileRemarking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: ProfileRemarking
'
' Description:  Example to show the remarking lines due to profiles
'               that are connected to the supported plates, on the pinjig
'
' Author: Anand Hariharan
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigRule.ProfileRemarking"

Implements IJDMfgPinJigRulePerRemarkingType

Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IMSCoreCollections.IJElements
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler

    Set IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5020, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IMSCoreCollections.IJElements) As String

End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IMSCoreCollections.IJElements) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler

    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport

    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject

    Dim oMfgUtil As IJMfgUtilSurface
    Set oMfgUtil = New MfgUtilSurface
    
    Dim oMfgMGhelper As New MfgMGHelper

    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim NormalX As Double
    Dim NormalY As Double
    Dim NormalZ As Double
    Dim RootPtX As Double
    Dim RootPtY As Double
    Dim RootPtZ As Double
    oPinJig.GetBasePlane NormalX, NormalY, NormalZ, _
                         RootPtX, RootPtY, RootPtZ
    
    Dim oBasePlane As IJPlane
    Set oBasePlane = New Plane3d
    oBasePlane.DefineByPointNormal RootPtX, RootPtY, RootPtZ, _
                                   NormalX, NormalY, NormalZ
    
    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates

    Dim SideSupported() As Long
    SideSupported = oMfgUtil.GetSidesOfPlatesFacingPlane(oPlateColl, oBasePlane)

    Dim Iter As Integer
    ' For each supported plate ...
    For Iter = 1 To oPlateColl.Count
                
        Set oPartSupport.Part = oPlateColl.Item(Iter)
        
        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        Dim ThisPortColl As Collection
        Dim OtherPortColl As Collection
        ' ... get all objects connected to plate with a physical connection
        oPartSupport.GetConnectedObjects ConnectionPhysical, _
                                         ConnectedObjColl, ConnectionsColl, _
                                         ThisPortColl, OtherPortColl

        Dim i As Integer
        ' For each connected object ...
        For i = 1 To ConnectedObjColl.Count
            ' ... check if this connected object is another supported plate
            If TypeOf ConnectedObjColl.Item(i) Is IJProfilePart Then

                Dim oAppCon As IJAppConnection
                Set oAppCon = ConnectionsColl.Item(i)
                    
                Dim ConnType As ContourConnectionType
                Dim WhichPartCross As Boolean
                oPartSupport.GetConnectionTypeForContour oAppCon, ConnType, WhichPartCross

            ' ... check if this connected object connects with supported plate with a TEE type connection
                If ConnType = PARTSUPPORT_CONNTYPE_TEE Then ' (probably always true for profiles?)
                    
                    Dim oPort As IJStructPort
                    Set oPort = ThisPortColl.Item(i)
                    
            ' ... check if this connected object is on pin jig's remarking side
                    If ((oPort.ContextID & CTX_OFFSET) And SideSupported(Iter) = CTX_BASE) Or _
                       ((oPort.ContextID & CTX_BASE) And SideSupported(Iter) = CTX_OFFSET) Then
                       
                        Dim oSystemParent As IJSystem
                        Set oSystemParent = GetSystemParent(ConnectedObjColl.Item(i))
                       
            ' ... get the desired portion of the geometry (SideA/SideB for physical or navigate to logical or ....)
                        Dim oWireBody As IJWireBody
                Dim oThickDir As IJDVector
                oPartSupport.GetConnectionContour_Tee ConnectionsColl.Item(i), oWireBody, oThickDir, SideA

                        Dim oCS As IJComplexString
                        oMfgMGhelper.WireBodyToComplexString oWireBody, oCS
                        
                Dim oNI As IJNamedItem
                        Set oNI = oSystemParent

                Dim RemarkingLineName As String
                If Not oNI Is Nothing Then
                    ' Decorate your name here (e.g., prefix with "Deck" etc.)
                    RemarkingLineName = oNI.Name
                    ' Decorate your name here (e.g., suffix with "SideA" etc.)
                End If

                        ' ******************************* IMPORTANT NOTE ******************************* '
                        ' The Pin jig semantic expects Connection Geometry (with proper attribution).    '
                        ' So create a MfgGeom3d object (that has both geometry and attribution).         '
                        ' The semantic will use these objects to create connection remarking lines and   '
                        ' merge them based on MONIKERS and NAMES of these objects.                       '
                        ' ******************************* IMPORTANT NOTE ******************************* '

                    CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, _
                                              oSystemParent, ReturnColl, _
                                          RemarkingLineName, STRMFG_PinJig_Remarking_Profile

                        Set oSystemParent = Nothing
                Set oNI = Nothing
                Set oWireBody = Nothing
                        Set oThickDir = Nothing
                Set oCS = Nothing
                        
                    End If ' end check for remarking side
                    
                    Set oPort = Nothing
                    
                End If ' end check for TEE connection
                
                Set oAppCon = Nothing
                
            End If ' end check for supported plate
            
        Next ' end looping around connected objects
                
    Next ' end looping around supported plates

    Set IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry = ReturnColl

    Set oBasePlane = Nothing
    Set oMfgMGhelper = Nothing
    Set oMfgUtil = Nothing
    Set oPlateColl = Nothing
    Set ReturnColl = Nothing
    Set oPartSupport = Nothing
    Set oPinJig = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5021, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function
