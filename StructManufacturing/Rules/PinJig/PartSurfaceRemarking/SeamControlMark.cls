VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamControlMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigPartSurfLCRule
' Module:SeamControlMark
'
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigPartSurfLCRule.SeamControlMark::"

Implements IJDMfgPinJigRulePerRemarkingType

Private Function CalculateOffsetFromMarginsAndRootGap(MarginValues() As Double, _
                                                      PortsWithMargin As IJElements, _
                                                      RefPort As Object, _
                                                      RootGapValue As Double, _
                                                      MinOffset As Double) As Double
                                                      
    On Error GoTo ErrorHandler
    Const METHOD As String = "CalculateOffsetFromMarginsAndRootGap"
    
    Dim ReturnOffset As Double
    ReturnOffset = MinOffset + Abs(RootGapValue)

    Dim TotalMargin As Double
    TotalMargin = 0#
    
    Dim i As Integer
    Dim lPortsCount     As Long
    
    lPortsCount = PortsWithMargin.Count
    For i = 1 To lPortsCount
        If PortsWithMargin(i) Is RefPort Then
           TotalMargin = TotalMargin + MarginValues(i - 1) ' lower bound of Margin values is 0
        End If
    Next
    
    If TotalMargin < 0.8 * MinOffset Then
        ReturnOffset = ReturnOffset - TotalMargin
    End If

    CalculateOffsetFromMarginsAndRootGap = ReturnOffset
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function

Private Function GetAverageNormalOfPortSurface(oPort As IJPort) As IJDVector
    Const METHOD = "GetAverageNormalOfPortSurface"
    On Error GoTo ErrorHandler
    
    Dim TopoLocate As IJTopologyLocate
    Set TopoLocate = New TopologyLocate
    
    Dim ApproxPoint As IJDPosition
    Dim ApproxNormal As IJDVector
    TopoLocate.FindApproxCenterAndNormal oPort.Geometry, ApproxPoint, ApproxNormal
    
    Set GetAverageNormalOfPortSurface = ApproxNormal
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IJElements

End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As String
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities"
    On Error GoTo ErrorHandler
    
    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsDesignSeam]"
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler
    
    Const MIN_OFFSET As Double = 0.1

    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oMfgUtilWrapper As New GSCADMathGeom.MfgGeomUtilWrapper
    Dim oMfgMGhelper As New GSCADMathGeom.MfgMGHelper
    
    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim RemarkingSurface As IJSurfaceBody
    Set RemarkingSurface = oPinJig.RemarkingSurface
    
    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates
    
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn

    Dim Iter As Integer
    ' For each supported plate ...
    For Iter = 1 To oPlateColl.Count
                
        Set oPartSupport.Part = oPlateColl.Item(Iter)
        
        Dim MarginValues() As Double
        Dim oMarginPortGeom As IJElements
        
        Dim PrimaryDirection As Object, SecndryDirection As Object
        Dim PrimaryShrinkage As Double, SecndryShrinkage As Double
        oMfgGeomHelper.GetMarginAndShrinkageFromPlatePart oPlateColl.Item(Iter), _
                                                          MarginValues, _
                                                          oMarginPortGeom, _
                                                          PrimaryShrinkage, _
                                                          PrimaryDirection, _
                                                          SecndryShrinkage, _
                                                          SecndryDirection
                                                          
        
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
            ' ... check if this connected object is another plate
            If TypeOf ConnectedObjColl.Item(i) Is IJPlatePart Then
                Dim oAppCon As IJAppConnection
                Set oAppCon = ConnectionsColl.Item(i)
                    
                Dim ConnType As ContourConnectionType
                Dim WhichPartCross As Boolean
                oPartSupport.GetConnectionTypeForContour oAppCon, ConnType, WhichPartCross
                
            ' ... check if this connected object connects with supported plate with a BUTT type connection
                If ConnType = PARTSUPPORT_CONNTYPE_BUTT Then
            ' ... get geometry portions of this BUTT connection that are on the remarking surface
                    Dim RootGapValue As Double
                    Set oSDPhysicalConn.object = ConnectionsColl.Item(i)
                    oSDPhysicalConn.GetBevelParameterValue "RootGap", RootGapValue, DoubleType
                    Set oSDPhysicalConn.object = Nothing

                    Dim oPCgeoms As IJElements
                    Set oPCgeoms = GetPCPortionOnSurface(ConnectionsColl.Item(i), _
                                                         RemarkingSurface)
                    
            ' ... calculate offset considering margins and root gap.
                    Dim OffsetValue As Double
                    OffsetValue = CalculateOffsetFromMarginsAndRootGap(MarginValues, _
                                                                       oMarginPortGeom, _
                                                                       ThisPortColl.Item(i), _
                                                                       RootGapValue, _
                                                                       MIN_OFFSET)
                                                                       
            ' ... determine offset direction as *OTHER* port's normal.
                    Dim OffsetDir As IJDVector
                    Set OffsetDir = GetAverageNormalOfPortSurface(OtherPortColl.Item(i))
                    
            ' ... offset each of these geometries, and create output.
                    Dim oCS As IJComplexString
                    For Each oCS In oPCgeoms
                    
                      On Error Resume Next
                      
                        Dim oWire As IJWireBody
                        oMfgMGhelper.ComplexStringToWireBody oCS, oWire
                        If oWire Is Nothing Then GoTo NextPCgeom
                        
                        Dim oOffsetWire As IJWireBody
                        Set oOffsetWire = oMfgUtilWrapper.OffsetCurve(RemarkingSurface, _
                                                                      oWire, OffsetDir, _
                                                                      OffsetValue, True)
                        If oOffsetWire Is Nothing Then GoTo NextPCgeom
                        
                        Dim oOffsetCSColl As IJElements
                        Set oOffsetCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oOffsetWire)
                        If oOffsetCSColl Is Nothing Then GoTo NextPCgeom
                        If oOffsetCSColl.Count = 0 Then GoTo NextPCgeom
                        
                        Dim oOffsetCS As IJComplexString
                        For Each oOffsetCS In oOffsetCSColl
                            CreateMfgGeom3dObject oOffsetCS, STRMFG_PinJigRemarkingLine3D, _
                                                  Nothing, ReturnColl, "SeamControlMark"
                            Set oOffsetCS = Nothing
                        Next
                      
NextPCgeom:
                        Err.Clear
                      On Error GoTo ErrorHandler
                        Set oWire = Nothing
                        Set oOffsetWire = Nothing
                        Set oOffsetCSColl = Nothing
                    Next
                    
                    Set oPCgeoms = Nothing
                    Set OffsetDir = Nothing
                End If ' end check for BUTT connection
                
                Set oAppCon = Nothing
            End If ' end check for connected object being plate
        Next ' end looping around connected objects
        
        oMarginPortGeom.Clear
        Set oMarginPortGeom = Nothing
        
        Set PrimaryDirection = Nothing
        Set SecndryDirection = Nothing
    Next ' end looping around supported plates
    
    Set IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry = ReturnColl
    
CleanUp:
    Set oPlateColl = Nothing
    Set ReturnColl = Nothing
    Set oPartSupport = Nothing
    Set oPinJig = Nothing
    Set oMfgGeomHelper = Nothing
    Set oMfgUtilWrapper = Nothing
    Set oMfgMGhelper = Nothing
    Set oSDPhysicalConn = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function
