VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPinJigReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CpinJigReport.cls
''
''  Description : Populates a text file with PinJig related data
''
''
''  Author      : Intergraph Development.
''
''  History     :
''               Initial Creation   -
''
''
''**************************************************************************************
Implements IJDCustomReport

Private Const MODULE = "CustomPinJigReports.CPinJigReport"
Private Const M_PI = 3.14159265358979
Private Enum m_MarginType
    HeatingMargin = 1
    BendingMargin = 2
    EndFaceMargin = 3
    IntercoastalMargin = 4
    AssyMarginSelected = 5
    AssyMarginDeselected = 6
End Enum
'- PinJig ------------------------------------------------------'
Private Const IJPinJig = "{FE221533-5879-11D5-B86E-0000E2300200}"


Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
            
    If pElements.Count > 0 Then
        Dim objPinJig As Object
        Dim oSelectedObj As Object
        For Each oSelectedObj In pElements
            If Not TypeOf oSelectedObj Is IJPinJig Then
                'This is plate. Get Pin Jigs from plate
                Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                Dim oColl As IJElements
                Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJPinJig)
                
                'For each pinjig, report the information
                Dim oPinJig As Object
                For Each oPinJig In oColl
                    ReportPinJigObjectInformation oPinJig, strFileName
                Next
            Else
                ReportPinJigObjectInformation oSelectedObj, strFileName
            End If
        Next
        eCustomReportStatus = StrMfgProcessFinished
    End If
            
End Sub

Private Sub ReportPinJigObjectInformation(oObject As Object, strLogFile As String)
 Const METHOD = "ReportPinJigObjectInformation"
    On Error GoTo ErrorHandler
    
    Dim oPinJig             As IJDPinJigReport
    Set oPinJig = oObject
    If oPinJig Is Nothing Then
        GoTo Cleanup
    End If
    
    'Open log file
    Dim oStream As TextStream
    Dim oFSO As FileSystemObject
    Set oFSO = New FileSystemObject
    Set oStream = oFSO.OpenTextFile(strLogFile, ForAppending, True)
    Set oFSO = Nothing
    
    oStream.WriteLine "Reporting Pin Jig: " & Now & vbNewLine

    'Show PinJig Information
    ShowPinJigInformation oPinJig, oStream
    
    ShowRemarkingLineInformation oObject, oStream
    
    Dim oIntxnPoint         As IJElements
    Set oIntxnPoint = oPinJig.GetEntity(IntersectionPointCollection)
    
    Dim iCount              As Integer
    Dim oPrevPoint As IJDJigIntxnPointReport
    For iCount = 1 To oIntxnPoint.Count
        Dim oPoint          As IJDJigIntxnPointReport
        Set oPoint = oIntxnPoint.Item(iCount)
        
        If oPoint Is Nothing Then
            GoTo Cleanup
        End If
 
        oStream.WriteLine "Information on Jig Intersection point No." & iCount & " of " & oIntxnPoint.Count
        ShowJigIntxPtInformation oPoint, oStream
        If Not oPrevPoint Is Nothing Then
            Dim xDistance As Double
            Dim yDistance As Double
            
            oPoint.GetFloorOffsets PointToPoint, oPrevPoint, xDistance, yDistance
            oStream.WriteLine "    Offsets from previous intersection point    X: " & xDistance & "  Y:" & yDistance & vbCrLf
        End If
        Set oPrevPoint = oPoint
        Set oPoint = Nothing
    Next
    Set oPrevPoint = Nothing

    oIntxnPoint.Clear
    Set oIntxnPoint = Nothing
 
    Dim oPins               As IJElements
    Set oPins = oPinJig.GetEntity(MfgPinCollection)
    
    Dim oPrevPin            As IJDMfgPinReport
    For iCount = 1 To oPins.Count
        Dim oPin            As IJDMfgPinReport
        Set oPin = oPins.Item(iCount)
        
        If oPin Is Nothing Then
            GoTo Cleanup
        End If
 
        oStream.WriteLine "Information on Pin No." & iCount & " of " & oPins.Count
        ShowMfgPinInformation oPin, oStream
        
        If Not oPrevPin Is Nothing Then
            oPin.GetFloorOffsets PinToPin, oPrevPin, xDistance, yDistance
            oStream.WriteLine "    Offsets from previous pin    X: " & xDistance & "  Y:" & yDistance & vbCrLf
        End If
        
        Set oPrevPin = oPin
        Set oPin = Nothing
    Next
    Set oPrevPin = Nothing
    
    ShowMountingAngleInformation oPinJig, oStream
 
    oStream.WriteLine "============================= End of report ==================================="
    
    'Close Log File
    oStream.Close
    
    oPins.Clear
    Set oPins = Nothing
    
    Set oPinJig = Nothing
    
Cleanup:

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub ShowPinJigInformation(oPinJigReport As IJDPinJigReport, oStream As TextStream)
    Const METHOD = "ShowPinJigInformation"
     
    On Error GoTo ErrorHandler
    oStream.WriteLine "================================================================"
    oStream.WriteLine " "

'
'   Get different lengths.
'
    Dim HorSpacingBetweenPins As Double
    oPinJigReport.getLength HorPinIntervalSpacing, HorSpacingBetweenPins
    oStream.WriteLine "Horizontal spacing (interval) between Pins : " & HorSpacingBetweenPins
    
    Dim VerSpacingBetweenPins As Double
    oPinJigReport.getLength VerPinIntervalSpacing, VerSpacingBetweenPins
    oStream.WriteLine "Vertical spacing (interval) between Pins   : " & VerSpacingBetweenPins
    oStream.WriteLine " "
    
    ' Old API to get straight diagonal lengths without shrinkage/margin
    
    Dim StraightDiagOPDP As Double
    oPinJigReport.getLength PanelDiagonal_OP_To_DP, StraightDiagOPDP
    oStream.WriteLine "Straight distance between Origin Point and Diagonal Point corners : " & StraightDiagOPDP
    
    Dim StraightDiagXPYP As Double
    oPinJigReport.getLength PanelDiagonal_XP_To_YP, StraightDiagXPYP
    oStream.WriteLine "Straight distance between Horizontal Point to origin and Vertical Point to origin corners : " & StraightDiagXPYP
    oStream.WriteLine " "

    ' New (preferred) API to get diagonal lengths
    
    ' Some code to dump the margin & shrinkage values for all supported plates.
    
    oStream.WriteLine "Margin and Shrinkage inputs for the new API (obtained from Model)"
    
    Dim oPinJig As IJPinJig
    Set oPinJig = oPinJigReport
    
    Dim oSupportedPlates As IJElements
    Set oSupportedPlates = oPinJig.SupportedPlates
    
    Dim oGeomHelper As IJMfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
                                         
    Dim oPlate As Object
    For Each oPlate In oSupportedPlates
    
        Dim MarginValues() As Double
        Dim oMarginPortGeom As IJElements
        
        Dim PrimShrDir As IJDVector, SecShrDir As IJDVector
        Dim PrimShrVal As Double, SecShrVal As Double
        
        oGeomHelper.GetMarginAndShrinkageFromPlatePart oPlate, MarginValues, oMarginPortGeom, _
                                                       PrimShrVal, PrimShrDir, SecShrVal, SecShrDir
            
        Dim oPlateNI As IJNamedItem
        Set oPlateNI = oPlate
        
        Dim PlateName As String
        PlateName = oPlateNI.name
        
        If UBound(MarginValues) - LBound(MarginValues) < 0 Then
            oStream.WriteLine vbCrLf & PlateName & " does not have any margins defined!"
        Else
            oStream.WriteLine vbCrLf & PlateName & " has the following margins defined:"
            Dim Foo As Integer
            For Foo = LBound(MarginValues) To UBound(MarginValues)
                oStream.WriteLine "(" & Foo & ") " & MarginValues(Foo)
            Next
        End If
        
        
        Erase MarginValues
        
        
        oStream.WriteLine ""
        If Not PrimShrDir Is Nothing Then
            oStream.WriteLine PlateName & " has a Primary shrinkage value   = " & PrimShrVal & " along (" & _
                        PrimShrDir.x & ", " & PrimShrDir.y & ", " & PrimShrDir.z & ")"
        End If
        
        If Not SecShrDir Is Nothing Then
            oStream.WriteLine PlateName & " has a Secondary shrinkage value = " & SecShrVal & " along (" & _
                        SecShrDir.x & ", " & SecShrDir.y & ", " & SecShrDir.z & ")"
        End If
    Next
    
    Dim HorShrDir As New DVector
    HorShrDir.Set 1, 0, 0 ' UGLY HACK not actually done within Pin Jig API.
    
    Dim VerShrDir As New DVector
    VerShrDir.Set 0, 1, 0 ' UGLY HACK not actually done within Pin Jig API.
    
    Dim HorShrVal As Double, VerShrVal As Double
    oGeomHelper.GetShrinkageFromAssembly oSupportedPlates, _
                                         HorShrDir, VerShrDir, _
                                         HorShrVal, VerShrVal
    
    oStream.WriteLine " "
    oStream.WriteLine "Amortized Shrinkage Value for all supported plates together:"
    oStream.WriteLine "Shrinkage value   = " & HorShrVal & " along Horizontal (" & _
                HorShrDir.x & ", " & HorShrDir.y & ", " & HorShrDir.z & ")"
    oStream.WriteLine "Shrinkage value = " & VerShrVal & " along Vertical (" & _
                VerShrDir.x & ", " & VerShrDir.y & ", " & VerShrDir.z & ")"
                                         
    Set oPinJig = Nothing
    Set oSupportedPlates = Nothing
    Set oGeomHelper = Nothing
    
    ' End code to dump the margin & shrinkage values for all supported plates.
                                         
    Dim MarginTypesAffectingCorners() As Long
    Dim ConsiderShrinkage As Boolean

'   SYNOPSIS:
'    Dim GirthUp As Double, StraightUp As Double, GirthDown As Double, StraightDown As Double
'    oPinJigReport.GetPinJigDiagonalsWithShrinkageMargin MarginTypesAffectingCorners, ConsiderShrinkage, _
'                                                        GirthUp, StraightUp, GirthDown, StraightDown

    Dim Dist(1 To 4, 1 To 4) As Double
    
    ' Get diagonal lengths without shrinkage/margin
    Erase MarginTypesAffectingCorners ' Make sure that array is empty
    ConsiderShrinkage = False
    
    oPinJigReport.GetPinJigDiagonalsWithShrinkageMargin MarginTypesAffectingCorners, ConsiderShrinkage, _
                                                        Dist(1, 1), Dist(1, 2), Dist(1, 3), Dist(1, 4)

    ' Get diagonal lengths considering only shrinkage, no margin
    ConsiderShrinkage = True
    
    oPinJigReport.GetPinJigDiagonalsWithShrinkageMargin MarginTypesAffectingCorners, ConsiderShrinkage, _
                                                        Dist(2, 1), Dist(2, 2), Dist(2, 3), Dist(2, 4)

    ' Get diagonal lengths considering only margin, no shrinkage
    ReDim MarginTypesAffectingCorners(1 To 2) As Long
    MarginTypesAffectingCorners(1) = m_MarginType.IntercoastalMargin
    MarginTypesAffectingCorners(2) = m_MarginType.BendingMargin
    
    ConsiderShrinkage = False
    
    oPinJigReport.GetPinJigDiagonalsWithShrinkageMargin MarginTypesAffectingCorners, ConsiderShrinkage, _
                                                        Dist(3, 1), Dist(3, 2), Dist(3, 3), Dist(3, 4)

    ' Get diagonal lengths considering both margin and shrinkage
    ConsiderShrinkage = True
    
    oPinJigReport.GetPinJigDiagonalsWithShrinkageMargin MarginTypesAffectingCorners, ConsiderShrinkage, _
                                                        Dist(4, 1), Dist(4, 2), Dist(4, 3), Dist(4, 4)

    Dim DistStr(1 To 4) As String
    Dim i As Integer, j As Integer
    For i = 1 To 4
        For j = 1 To 4
            DistStr(i) = DistStr(i) & Format(Round(Dist(j, i), 3), " 0#.000 \|")
        Next
    Next
    
    oStream.WriteLine " "
    oStream.WriteLine "+------------------------------------------------+--------+--------+--------+--------+"
    oStream.WriteLine "|                                                | No S/M | Only S | Only M | S & M  |"
    oStream.WriteLine "+------------------------------------------------+--------+--------+--------+--------+"
    oStream.WriteLine "| Origin Point to Diagonal Point dist   - Girth  :" & DistStr(1)
    oStream.WriteLine "| Origin Point to Diagonal Point dist- Straight  :" & DistStr(2)
    oStream.WriteLine "| Vertical Point to Hori Point dist     - Girth  :" & DistStr(3)
    oStream.WriteLine "| Vertical Point to Hori Point dist  - Straight  :" & DistStr(4)
    oStream.WriteLine "+------------------------------------------------+--------+--------+--------+--------+"
    oStream.WriteLine " "
    oStream.WriteLine " "
    
'
'   Get the different angles.
'

    Dim JigFloorNormalAndXaxis As Double
    JigFloorNormalAndXaxis = oPinJigReport.GetAngle(BasePlane_YZplane)
    oStream.WriteLine "Angle between Jig Floor and YZ plane : " & JigFloorNormalAndXaxis

    Dim JigFloorNormalAndYaxis As Double
    JigFloorNormalAndYaxis = oPinJigReport.GetAngle(BasePlane_ZXplane)
    oStream.WriteLine "Angle between Jig Floor and ZX plane : " & JigFloorNormalAndYaxis

    Dim JigFloorNormalAndZaxis As Double
    JigFloorNormalAndZaxis = oPinJigReport.GetAngle(BasePlane_XYplane)
    oStream.WriteLine "Angle between Jig Floor and XY plane : " & JigFloorNormalAndZaxis
    oStream.WriteLine " "

'
'   Get different entities associated with Pin Jig.
'
    'oPinJigReport.GetEntity HorizontalCoordinateLine
    'oPinJigReport.GetEntity VerticalCoordinateLine
    'oPinJigReport.GetEntity JigFloor
    
    oStream.WriteLine " "

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub ShowRemarkingLineInformation(oObject As Object, oStream As TextStream)
    Const METHOD = "ShowRemarkingLineInformation"
    On Error GoTo ErrorHandler
    
    oStream.WriteLine "----------------------------------------------------------------"
    oStream.WriteLine " "
    
    Dim oPinJig As IJDPinJigReport
    Set oPinJig = oObject
    
    Dim oRemarkUtil As IJDJigRemarkLineUtil
    Set oRemarkUtil = oPinJig.GetEntity(RemarkingLineHandler)
    
    Dim o3dRemarkLineColl As IJElements
    Set o3dRemarkLineColl = oRemarkUtil.GetCollection(AllRemarkingLines, JigSurfaceRemarking)
    
    oStream.WriteLine " "
    oStream.WriteLine "There are " & o3dRemarkLineColl.Count & " 3D remarking lines."
    oStream.WriteLine " "
    
    Dim LineName As IJNamedItem
    Dim LineObj As Object
    Dim LineIdx As Integer
    
    LineIdx = 1
    For Each LineObj In o3dRemarkLineColl
        Dim eRemarkType As JigRemarkingTopologyType
        eRemarkType = oRemarkUtil.GetTypeOfRemarkingLine(LineObj)
        
        Set LineName = LineObj
        oStream.WriteLine "[" & LineIdx & "] " & LineName.name & " is a 3D " & StringForRemarkType(eRemarkType) & " remarking line"
        LineIdx = LineIdx + 1
    Next
    Set o3dRemarkLineColl = Nothing

    Dim HorLines As IJElements
    Set HorLines = oRemarkUtil.GetCollection(AllRemarkingLines, HorizontalJigFloorRemark)
    
    Dim VerLines As IJElements
    Set VerLines = oRemarkUtil.GetCollection(AllRemarkingLines, VerticalJigFloorRemark)
    
    oStream.WriteLine " "
    oStream.WriteLine "There are " & HorLines.Count & " horizontal and " & VerLines.Count & " vertical remarking lines."
    
    Dim oIntPts As IJElements
    Dim IntPtObj As Object
    Dim oIntPt As IJDJigIntxnPointReport
    
    Dim oClosestMfgPin As IJDMfgPinReport
    Dim PinName As String
    Dim HorDist As Double
    Dim VerDist As Double
    
    For Each LineObj In HorLines
        Set LineName = LineObj
        Set oIntPts = oRemarkUtil.GetIntxPtsOnRemarkingLine(LineObj)
        
        eRemarkType = oRemarkUtil.GetTypeOfRemarkingLine(LineObj)
        
        oStream.WriteLine " "
        oStream.WriteLine LineName.name & " is a horizontal " & StringForRemarkType(eRemarkType) & " remarking line with " & oIntPts.Count & " intersection points:"
        
        LineIdx = 1
        For Each IntPtObj In oIntPts
            Set oIntPt = IntPtObj
            oIntPt.GetEntity NearestManufacturingPin, oClosestMfgPin, PinName
            
            oIntPt.GetFloorOffsets PointToPin, oClosestMfgPin, HorDist, VerDist
            oStream.WriteLine " "
            oStream.WriteLine "[ " & LineIdx & "-a ] Offset (" & HorDist & ", " & VerDist & ") from Pin " & PinName
            
            Dim ClosestRow As Long, ClosestCol As Long
            oClosestMfgPin.GetPinGridIndices ClosestRow, ClosestCol
            
            Dim NamedPos As DPosition
            
            Set NamedPos = oRemarkUtil.GetNamedPointOnRemarkingLine(LineObj, PinName)
            If Not NamedPos Is Nothing Then
                oStream.WriteLine "[ " & LineIdx & "-b ] Pt on Remark line closest to " & PinName & ": (" & _
                           CStr(NamedPos.x) & ", " & CStr(NamedPos.y) & ", " & CStr(NamedPos.z)
                Set NamedPos = Nothing
            End If
            
            Set NamedPos = oRemarkUtil.GetNamedPointOnRemarkingLine(LineObj, "R " & CStr(ClosestRow))
            If Not NamedPos Is Nothing Then
                oStream.WriteLine "[ " & LineIdx & "-c ] Pt on Remark line along Pin Row " & CStr(ClosestRow) & ": (" & _
                           CStr(NamedPos.x) & ", " & CStr(NamedPos.y) & ", " & CStr(NamedPos.z)
                Set NamedPos = Nothing
            End If
            
            Set NamedPos = oRemarkUtil.GetNamedPointOnRemarkingLine(LineObj, "col" & CStr(ClosestCol))
            If Not NamedPos Is Nothing Then
                oStream.WriteLine "[ " & LineIdx & "-d ] Pt on Remark line along Pin Col " & CStr(ClosestCol) & ": (" & _
                           CStr(NamedPos.x) & ", " & CStr(NamedPos.y) & ", " & CStr(NamedPos.z)
                Set NamedPos = Nothing
            End If
            
            LineIdx = LineIdx + 1
            Set oIntPt = Nothing
            Set oClosestMfgPin = Nothing
        Next
    Next
    HorLines.Clear
    Set HorLines = Nothing
    
    For Each LineObj In VerLines
        Set LineName = LineObj
        Set oIntPts = oRemarkUtil.GetIntxPtsOnRemarkingLine(LineObj)
        
        eRemarkType = oRemarkUtil.GetTypeOfRemarkingLine(LineObj)
        
        oStream.WriteLine " "
        oStream.WriteLine LineName.name & " is a vertical " & StringForRemarkType(eRemarkType) & " remarking line with the following " & oIntPts.Count & " intersection points:"
        
        LineIdx = 1
        For Each IntPtObj In oIntPts
            Set oIntPt = IntPtObj
            oIntPt.GetEntity NearestManufacturingPin, oClosestMfgPin, PinName
            
            oIntPt.GetFloorOffsets PointToPin, oClosestMfgPin, HorDist, VerDist
            Set oClosestMfgPin = Nothing
            
            oStream.WriteLine "[" & LineIdx & "] Offset (" & HorDist & ", " & VerDist & ") from Pin " & PinName
            LineIdx = LineIdx + 1
            Set oIntPt = Nothing
        Next
    Next
    VerLines.Clear
    Set VerLines = Nothing
    
    oStream.WriteLine " "
    
    Set LineObj = Nothing
    Set LineName = Nothing
    Set oObject = Nothing
    Set oRemarkUtil = Nothing
    Set oPinJig = Nothing

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description

End Sub

Private Sub ShowJigIntxPtInformation(oJigXptReport As IJDJigIntxnPointReport, oStream As TextStream)
    Const METHOD = "ShowJigIntxPtInformation"
     
    On Error GoTo ErrorHandler
    oStream.WriteLine "----------------------------------------------------------------"
    oStream.WriteLine " "

'
'   Get Jig Intersection point's remarking lines & coordinates.
'
    
Label_HorRemarkingLine:
    On Error GoTo Label_VerRemarkingLine
    Dim oHorRemarkLine As Object
    Dim HorRemarkName As String
    oJigXptReport.GetEntity JigIntxPointEntityReportType.HorizontalRemarkingParent, oHorRemarkLine, HorRemarkName
    oStream.WriteLine "    Horizontal remarking line responsible for Intersection point : " & HorRemarkName
    
Label_VerRemarkingLine:
    On Error GoTo Label_BaseCoordinate
    Dim oVerRemarkLine As Object
    Dim VerRemarkName As String
    oJigXptReport.GetEntity JigIntxPointEntityReportType.VerticalRemarkingParent, oVerRemarkLine, VerRemarkName
    oStream.WriteLine "    Vertical remarking line responsible for Intersection point   : " & VerRemarkName
    oStream.WriteLine " "

Label_BaseCoordinate:
    On Error GoTo Label_CuspAngle:
    Dim LowerX As Double
    Dim LowerY As Double
    Dim LowerZ As Double
    oJigXptReport.GetCoordinate BaseCoordinate, LowerX, LowerY, LowerZ
    oStream.WriteLine "    Intersection of above 2 lines PROJECTED onto jig floor : (" & LowerX & ", " & LowerY & ", " & LowerZ & ")"
    
    Dim UpperX As Double
    Dim UpperY As Double
    Dim UpperZ As Double
    oJigXptReport.GetCoordinate UpperPosition, UpperX, UpperY, UpperZ
    oStream.WriteLine "    Coordinates of intersection of above 2 remarking lines : (" & UpperX & ", " & UpperY & ", " & UpperZ & ")"
    oStream.WriteLine " "

Label_GirthLengths:
    On Error GoTo Label_CuspAngle:
    Dim HorGirLenNoShrink As Double, VerGirLenNoShrink As Double
    Dim HorGirLenWithShrk As Double, VerGirLenWithShrk As Double
    ' SUGGESTION: IN PRODUCTION CODE USE
       'oJigXptReport.GetMeasurement HorizontalGirthLength
       'oJigXptReport.GetMeasurement VerticalGirthLength
       'oJigXptReport.GetMeasurement HorGirthLength_NoShrinkage
       'oJigXptReport.GetMeasurement VerGirthLength_NoShrinkage
    ' WARNING: BELOW NOT TO BE USED IN PRODUCTION CODE.  IN PRODUCTION CODE USE ABOVE
    Dim oJigXptBO As IJJigIntersectPoint
    Set oJigXptBO = oJigXptReport
    oJigXptBO.GetGirthLengths HorGirLenWithShrk, VerGirLenWithShrk, HorGirLenNoShrink, VerGirLenNoShrink
    oStream.WriteLine "    Horizontal Girth Length not considering Shrinkage : " & HorGirLenNoShrink
    oStream.WriteLine "    Horizontal Girth Length considering Shrinkage     : " & HorGirLenWithShrk
    oStream.WriteLine "    Vertical Girth Length not considering Shrinkage   : " & VerGirLenNoShrink
    oStream.WriteLine "    Vertical Girth Length considering Shrinkage       : " & VerGirLenWithShrk
    Set oJigXptBO = Nothing
    oStream.WriteLine " "
    
Label_CuspAngle:
    On Error GoTo Label_IsJigCornerPoint:
    Dim angle As Double
    angle = oJigXptReport.GetAngle(CuspAngle)
    If angle > 0 Then
        oStream.WriteLine "    The cusp angle here is: " & angle
    End If
        
Label_IsJigCornerPoint:
    On Error GoTo Label_IsPlateCornerPoint:
    If oJigXptReport.GetBoolean(IsJigCornerPoint) Then
        oStream.WriteLine "    This is a jig corner point! "
    End If

Label_IsPlateCornerPoint:
    On Error GoTo Label_IsRemarkingIntersectionPoint:
    If oJigXptReport.GetBoolean(IsPlateCornerPoint) Then
        oStream.WriteLine "    This is a Plate corner point! "
    End If

Label_IsRemarkingIntersectionPoint:
    On Error GoTo Label_IsCuspPoint:
    If oJigXptReport.GetBoolean(IsRemarkingIntersectionPoint) Then
        oStream.WriteLine "    This is a remarking intersection point! "
    End If

Label_IsCuspPoint:
    If oJigXptReport.GetBoolean(IsCuspPoint) Then
        oStream.WriteLine "    This is a cusp point! "
    End If

    Set oHorRemarkLine = Nothing
    Set oVerRemarkLine = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub ShowMfgPinInformation(oMfgPinReport As IJDMfgPinReport, oStream As TextStream)
    Const METHOD = "ShowMfgPinInformation"
     
    On Error GoTo ErrorHandler
    oStream.WriteLine "----------------------------------------------------------------"
    oStream.WriteLine " "

    
    Dim LowerX As Double
    Dim LowerY As Double
    Dim LowerZ As Double
    Dim UpperX As Double
    Dim UpperY As Double
    Dim UpperZ As Double
    Dim diff As Double
    oMfgPinReport.GetCoordinate JigFloorCoordinate, LowerX, LowerY, LowerZ
    oStream.WriteLine "    Base coordinates of Pin (on jig floor)                        : (" & LowerX & ", " & LowerY & ", " & LowerZ & ")"
    
    oMfgPinReport.GetCoordinate SupportedPoint, UpperX, UpperY, UpperZ
    oStream.WriteLine "    Coordinates of Pin (Supported Point)                          : (" & UpperX & ", " & UpperY & ", " & UpperZ & ")"
    
    oStream.WriteLine "    Distance between upper and lower point (should be = height)   : " & _
    Sqr((LowerX - UpperX) * (LowerX - UpperX) + (LowerY - UpperY) * (LowerY - UpperY) + (LowerZ - UpperZ) * (LowerZ - UpperZ))
    
    Dim PinHeight As Double
    PinHeight = oMfgPinReport.GetDistance(MfgPinDistanceReportType.BetweenPinBaseAndPinTip)
    oStream.WriteLine "    Height of pin                                                 : " & PinHeight
    oStream.WriteLine " "
    
    If PinHeight > 0 Then
        Dim PinHeadHorAngle As Double
        PinHeadHorAngle = oMfgPinReport.GetAngle(HorAngleAtPinHead)
        oStream.WriteLine "    Angle between Surface-Tangent at contact-point and horizontal : " & PinHeadHorAngle
        
        Dim PinHeadVerAngle As Double
        PinHeadVerAngle = oMfgPinReport.GetAngle(VerAngleAtPinHead)
        oStream.WriteLine "    Angle between Surface-Tangent at contact-point and vertical   : " & PinHeadVerAngle
        oStream.WriteLine " "
    
        Dim SeamDistance As Double
        SeamDistance = oMfgPinReport.GetDistance(DistanceFromNearestSeam)
        oStream.WriteLine "    Distance of Pin from nearest seam                             : " & SeamDistance
        oStream.WriteLine " "
        
        Dim ButtDistance As Double
        ButtDistance = oMfgPinReport.GetDistance(DistanceFromNearestButt)
        oStream.WriteLine "    Distance of Pin from nearest butt                             : " & ButtDistance
        oStream.WriteLine " "
        
'        Dim Entity As Object
'        Dim Plate As IJPlate
'        Set Entity = oMfgPinReport.GetEntity(SupportedPlate)  'Currently not used for anything
'        Set Plate = Entity
'        If Not Entity Is Nothing Then Set Entity = Nothing
'        If Not Plate Is Nothing Then Set Plate = Nothing
    End If
        
'    Dim CornerProperty As Boolean
'    CornerProperty = oMfgPinReport.GetBoolean(IsCornerPin)
'    If CornerProperty Then
'        oStream.WriteLine "    Pin is a corner pin! :-)"
'    Else
'        oStream.WriteLine "    Pin is not a corner pin  :-("
'    End If
    oStream.WriteLine " "
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub ShowMountingAngleInformation(oPinJigReport As IJDPinJigReport, oStream As TextStream)
    Const METHOD = "ShowMountingAngleInformation"
    
    Dim oPinJig As IJPinJig
    Set oPinJig = oPinJigReport
    
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    Dim oBasePlaneNormal As IJDVector
    Set oBasePlaneNormal = New DVector
    
    oPinJig.GetBasePlane dNormalX, dNormalY, dNormalZ, dRootX, dRootY, dRootZ
    
    oBasePlaneNormal.Set dNormalX, dNormalY, dNormalZ
    oBasePlaneNormal.length = 1#
    
    Dim oXVector As DVector, oYVector As DVector, oZVector As DVector
    Set oXVector = New DVector
    Set oYVector = New DVector
    Set oZVector = New DVector

    oXVector.Set 1, 0, 0
    oYVector.Set 0, 1, 0
    oZVector.Set 0, 0, 1
    
    'normalize the vectors
    oXVector.length = 1#
    oYVector.length = 1#
    oZVector.length = 1#
    
    Dim dx As Double, dy As Double, dz As Double
     
    dx = Abs(oXVector.Dot(oBasePlaneNormal))
    dy = Abs(oYVector.Dot(oBasePlaneNormal))
    dz = Abs(oZVector.Dot(oBasePlaneNormal))
    
    Dim oParallelVector As IJDVector, oPerpendicularVector1 As IJDVector, oPerpendicularVector2 As IJDVector
     
    If (dx > dy) And (dx > dz) Then
        Set oParallelVector = oXVector
        Set oPerpendicularVector1 = oYVector
        Set oPerpendicularVector2 = oZVector
    ElseIf (dy > dx) And (dy > dz) Then
        Set oParallelVector = oYVector
        Set oPerpendicularVector1 = oXVector
        Set oPerpendicularVector2 = oZVector
    ElseIf (dz > dx) And (dz > dy) Then
        Set oParallelVector = oZVector
        Set oPerpendicularVector1 = oXVector
        Set oPerpendicularVector2 = oYVector
     End If
         
    Dim oFrameColl1 As IJElements
    Dim oFrameColl2 As IJElements
     
    Set oFrameColl1 = GetReferencePlanesInPinJigRange(oPinJig, oPerpendicularVector1)
    If oFrameColl1 Is Nothing Then
        Exit Sub
    Else
        Set oFrameColl2 = GetReferencePlanesInPinJigRange(oPinJig, oPerpendicularVector2)
        If oFrameColl2 Is Nothing Then
            Exit Sub
        End If
    End If
   
    
    oStream.WriteLine " "
    oStream.WriteLine "Reporting Mounting Angle of the frame relative to PinJig base plane normal at given position: "
    
    'Create Dummpy Position
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    oPosition.Set 0, 0, 0 'Dummy Position. Since we pass the intersector
   
   'Get the Middle Frame as Intersector
    
    Dim iMid As Integer
    iMid = oFrameColl2.Count / 2
   
    Dim oIntersector As Object
    Set oIntersector = oFrameColl2.Item(iMid)
    
    Dim oNI As IJNamedItem
    Set oNI = oIntersector
    
    oStream.WriteLine "Intersecting frame to get the mounting position : " & oNI.name
    oStream.WriteLine "----------------------------------------------------------------"
    oStream.WriteLine " "
    
    Dim dConcaveRadius As Double, dConvexRadius As Double
    dConcaveRadius = 0.2
    dConvexRadius = 0.1
         
    Dim oMountingPosition As IJDPosition
    Dim dMountingAngle As Double, dMountingAngle1 As Double
    Dim iCount As Integer
    
    For iCount = 1 To oFrameColl1.Count

        dMountingAngle = oPinJigReport.GetMountingAngle(oFrameColl1.Item(iCount), AssemblyJig, oPosition, oMountingPosition, oIntersector)
        
        Set oNI = oFrameColl1.Item(iCount)
        oStream.WriteLine "Mounting Angle of " & oNI.name & " : " & Round(dMountingAngle * (180 / M_PI), 2)
        
        oPinJigReport.GetMountingAngleBasedOnInputs oFrameColl1.Item(iCount), AssemblyJig, oPosition, dConvexRadius, dConcaveRadius, oMountingPosition, dMountingAngle, dMountingAngle1, oIntersector
        
        oStream.WriteLine "Mounting Angles of " & oNI.name & " based on  - " & "Concave radius : " & dConcaveRadius & " ; Convex radius : " & dConvexRadius
        oStream.WriteLine "Angle1 :" & Round(dMountingAngle * (180 / M_PI), 2) & " ; Angle2 :" & Round(dMountingAngle1 * (180 / M_PI), 2)
        
        'oMountingPosition can be null if the plane doesn't intersect.
        On Error Resume Next
        oStream.WriteLine "Mounting Position : (" & Round(oMountingPosition.x, 2) & ", " & Round(oMountingPosition.y, 2) & ", " & Round(oMountingPosition.z, 2) & ")"
        oStream.WriteLine " " & vbCrLf
    
        Dim oRemarkUtil As IJDJigRemarkLineUtil
        Set oRemarkUtil = oPinJigReport.GetEntity(RemarkingLineHandler)
        
        oStream.WriteLine "*** Testing new API with " & oNI.name & " ****"
        
        Dim oLowerSeamRemark As IJMfgGeom3d
        Set oLowerSeamRemark = oRemarkUtil.GetCollection(LowerContourJigRemarking, ProjectedJigRemarking).Item(1)
        
        Dim oSeam As IJSeam
        For Each oSeam In oPinJig.ActiveEntity.GetSeams(RemarkingLowerSeam)
            Set oNI = oSeam
            oStream.WriteLine "Angle computed in plane whose normal is perpendicular to tangent of seam <<" & oNI.name & ">> evaluated at root point"
            
            Dim oXpt As IJJigIntersectPoint
            For Each oXpt In oRemarkUtil.GetIntxPtsOnRemarkingLine(oLowerSeamRemark)
                oXpt.GetUpperPosition dx, dy, dz
                oPosition.Set dx, dy, dz
    
                oStream.WriteLine "Plane root point: (" & Round(oPosition.x, 2) & ", " & Round(oPosition.y, 2) & ", " & Round(oPosition.z, 2) & ")"
                
                dMountingAngle = oPinJigReport.GetMountingAngle(oFrameColl1.Item(iCount), SupportedPart, oPosition, oMountingPosition, oSeam)
                
                oStream.WriteLine "Angle evaluated at Position : (" & Round(oMountingPosition.x, 2) & ", " & Round(oMountingPosition.y, 2) & _
                                  ", " & Round(oMountingPosition.z, 2) & ") = " & Round(dMountingAngle * (180 / M_PI), 2)
                
                oPinJigReport.GetMountingAngleBasedOnInputs oFrameColl1.Item(iCount), SupportedPart, oPosition, _
                                                            dConvexRadius, dConcaveRadius, oMountingPosition, _
                                                            dMountingAngle, dMountingAngle1, oSeam
                
                oStream.WriteLine "Angle evaluated at Position : (" & Round(oMountingPosition.x, 2) & ", " & Round(oMountingPosition.y, 2) & _
                                  ", " & Round(oMountingPosition.z, 2) & ")" & vbCrLf & "Concave = " & Round(dMountingAngle * (180 / M_PI), 2) & _
                                  ", Convex = " & Round(dMountingAngle1 * (180 / M_PI), 2)
            Next
            Set oSeam = Nothing
        Next
        
        Set oRemarkUtil = Nothing
        Set oLowerSeamRemark = Nothing
    Next iCount
        
     oStream.WriteLine "================================================================"

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

    
Public Function GetReferencePlanesInPinJigRange(oPinJig As IJPinJig, oDirectionVec As IJDVector) As IJElements
    Const METHOD As String = "GetReferencePlanesInPinJigRange"
    
    On Error GoTo ErrorHandler

    Dim BOwithFS As IJDMfgFrameSystem
    Set BOwithFS = oPinJig
    
    Dim oPinJigFrameSys As IHFrameSystem
    Set oPinJigFrameSys = BOwithFS.FrameSysParent
    
    ' If no frame system was set on Pin jig (directly/indirectly), then nothing to do!
    If oPinJigFrameSys Is Nothing Then
        Set GetReferencePlanesInPinJigRange = Nothing
        Exit Function
    End If
    
    Dim oFramesColl As IJElements
    Set oFramesColl = GetReferencePlanesFromAllCSinRange(oPinJig, oDirectionVec)
    
    If oFramesColl Is Nothing Then
        Set GetReferencePlanesInPinJigRange = Nothing
        Exit Function
    End If
     
     '''Uncomment below lines to consider all frames in range
     '''If no frame system was set on Pin jig then send all the frames in range
''    If oPinJigFrameSys Is Nothing Then
''        Set GetReferencePlanesInPinJigRange = oFramesColl
''        Exit Function
''    End If
    
    ' Filter this collection to return only those from Pin Jig's FS.
    Dim oFrameSet As IJElements
    Set oFrameSet = New JObjectCollection

    Dim nIndex As Long
    For nIndex = 1 To oFramesColl.Count
    
        Dim oFrame As IHFrame
        Set oFrame = oFramesColl.Item(nIndex)
        
        Dim oFrameAxis As IHFrameAxis
        Set oFrameAxis = oFrame.FrameAxis
        
        Dim oFrameSystem As IHFrameSystem
        Set oFrameSystem = oFrameAxis.FrameSystem

        If oFrameSystem Is oPinJigFrameSys Then
            oFrameSet.Add oFramesColl.Item(nIndex)
        End If

        Set oFrame = Nothing
        Set oFrameAxis = Nothing
        Set oFrameSystem = Nothing
        
    Next nIndex

    ' Send the data back
    Set GetReferencePlanesInPinJigRange = oFrameSet
    
    Set oFrameSet = Nothing
    Set oPinJigFrameSys = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function GetReferencePlanesFromAllCSinRange(oPinJig As IJPinJig, oDirectionVec As IJDVector) As IJElements
    Const METHOD As String = "GetReferencePlanesFromAllCSinRange"
    On Error GoTo ErrorHandler
    
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    
    Dim oFramesColl As New Collection
    oGeomHelper.GetReferencePlanesInRange oPinJig.RemarkingSurface, oDirectionVec, oFramesColl

    If oFramesColl Is Nothing Then
        Set GetReferencePlanesFromAllCSinRange = Nothing
        GoTo Cleanup
    End If
    
    Dim oFrameSet As IJElements
    Set oFrameSet = New JObjectCollection

    Dim nIndex As Long
    For nIndex = 1 To oFramesColl.Count
        oFrameSet.Add oFramesColl.Item(nIndex)
    Next nIndex
    
    Set GetReferencePlanesFromAllCSinRange = oFrameSet
    
Cleanup:
    Set oGeomHelper = Nothing
    Set oFrameSet = Nothing
    Set oFramesColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function StringForRemarkType(eRemarkType As JigRemarkingTopologyType) As String
    Select Case eRemarkType
    
        Case ContourJigRemarking:
            StringForRemarkType = "Contour"
    
        Case SeamJigRemarking:
            StringForRemarkType = "Seam"
    
        Case PlateJigRemarking:
            StringForRemarkType = "Plate"
    
        Case ProfileJigRemarking:
            StringForRemarkType = "Profile"
    
        Case TransverseJigRemarking:
            StringForRemarkType = "Transverse Frame"
    
        Case LongitudinalJigRemarking:
            StringForRemarkType = "Longitudinal Frame"
    
        Case ButtockJigRemarking:
            StringForRemarkType = "Buttock Frame"
    
        Case FrameJigRemarking:
            StringForRemarkType = "Frame (Generic?!)"
    
        Case NavalArchJigRemarking:
            StringForRemarkType = "Naval Arch"
    
        Case UserAddedJigRemarking:
            StringForRemarkType = "User Mark"
            
        Case Default:
            StringForRemarkType = "UNKNOWN TYPE?!"
    
    End Select
End Function


