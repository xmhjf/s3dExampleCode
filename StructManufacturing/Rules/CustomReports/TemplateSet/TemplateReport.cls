VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CIngrTemplateSetXML1.cls
''
''  Description : Populates the XML DOM Document with Template related data for Sample 1
''
''
''  Author      : Intergraph
''
''  History     :
''               Initial Creation   -
''
''
''**************************************************************************************

Implements IJDCustomReport

Private Const MODULE = "CustomReport.TemplateReport"

Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler
       
    If pElements.Count > 0 Then
        Dim objTemplateSet As Object
        Dim oSelectedObj As Object
        For Each oSelectedObj In pElements
            If Not TypeOf oSelectedObj Is IJDMfgTemplateSet Then
                'This is plate. Get TemplateSets from plate
                Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                Dim oColl As IJElements
                Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJDMfgTemplateSet)
                
                'For each TemplateSet, report the information
                Dim oTemplateSet As Object
                For Each oTemplateSet In oColl
                    ReportTemplateSetInformation oTemplateSet, strFileName
                Next
            Else
                ReportTemplateSetInformation oSelectedObj, strFileName
            End If
        Next
        eCustomReportStatus = StrMfgProcessFinished
    End If
    
WrapUp:
    Exit Sub
    
ErrorHandler:
    eCustomReportStatus = StrMfgErrorUnknown
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub ReportTemplateSetInformation(oObject As Object, strLogFile As String)
 Const METHOD = "ReportTemplateSetInformation"
    On Error GoTo ErrorHandler
    
    Dim oTemplate           As IJMfgTemplate
    Dim oTemplateSet        As IJDMfgTemplateSet
    Dim oElements           As IJElements
    Dim iCount              As Integer

    Set oTemplateSet = oObject
 
    Dim oTemplateSetRpt As IJMfgTemplateSetReport
    Set oTemplateSetRpt = oTemplateSet
       
    If oTemplateSet Is Nothing Then
        Exit Sub
    End If
    
    'Open log file
    Dim oStream As TextStream
    Dim oFSO As FileSystemObject
    Set oFSO = New FileSystemObject
    Set oStream = oFSO.OpenTextFile(strLogFile, ForWriting, True)
    Set oFSO = Nothing
    
    Dim nTemplateGroupCount As Integer
    nTemplateGroupCount = oTemplateSetRpt.GroupsCount
    
    oStream.WriteLine "Number of Template Groups: " & nTemplateGroupCount
    
    Dim nGroupNumber As Integer
    
    For nGroupNumber = 1 To nTemplateGroupCount
    
        Set oElements = oTemplateSet.GetTemplatesSorted(GirthBetweenBaseCtlPointAndAftPoint, nGroupNumber)
        
        oStream.WriteLine "Number of Templates in Group" & nGroupNumber & "  : " & oElements.Count
        
        Dim oSDPlateWrapper As StructDetailObjects.PlatePart
        Set oSDPlateWrapper = New StructDetailObjects.PlatePart
        Dim oMfgChild As IJMfgChild
        Set oMfgChild = oTemplateSet
        Set oSDPlateWrapper.object = oMfgChild.getParent
        
        Dim oNavalArchRefCurveCol                     As Collection
        Set oNavalArchRefCurveCol = New Collection
    
        If oSDPlateWrapper.plateType = Hull Then
            Dim i               As Integer
            Dim oTempColl       As Collection
            'Get the Naval Arch reference curves that intersect the part.Pass both
            'the side of plate
            
            oSDPlateWrapper.ReferenceCurves PlateBaseSide, oTempColl
            For i = 1 To oTempColl.Count
                Dim oRefCurveData As IJRefCurveData
                Set oRefCurveData = oTempColl.Item(i)
                
                'Need to restrict the types supported(We do not want to support GRID and MARK referenc curves in MFG)
                If oRefCurveData.Type = JSRCOS_REFERENCE Or _
                    oRefCurveData.Type = JSRCOS_TANGENT Or _
                    oRefCurveData.Type = JSRCOS_KNUCKLE Or _
                    oRefCurveData.Type = JSRCOS_UNKNOWN Then
                    oNavalArchRefCurveCol.Add oRefCurveData
                End If
            Next
            Set oTempColl = Nothing
        End If
    
        For iCount = 1 To oElements.Count
    
            oStream.WriteLine " "
            oStream.WriteLine "Template # " & iCount
    
            Set oTemplate = oElements.Item(iCount)
    
            If oTemplate Is Nothing Then
                GoTo CleanUp
            End If
    
            ReportTemplateInformation oTemplate, oNavalArchRefCurveCol, oStream
            Set oTemplate = Nothing
    
        Next
    Next nGroupNumber
    

    
#If SHOW_BACKSET Then
    Dim oVec               As IJDVector
    Set oVec = New DVector
    Dim oBacksetValues() As Double
    Dim oAValues() As Double
    Dim oLValues() As Double
    Dim oDValues() As Double
    oVec.Set 1#, 0#, 0#
    Call oTemplateSetRpt.GetBackSet(oVec, oBacksetValues(), oAValues(), oLValues(), oDValues())
    
    oStream.WriteLine " "
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE SET REPORT - GET BACKSET "
    oStream.WriteLine "================================================================"
    For iCount = LBound(oBacksetValues) To UBound(oBacksetValues)
        oStream.WriteLine "Template # " & iCount
        oStream.WriteLine "A Values: " & oAValues(iCount)
        oStream.WriteLine "L Values: " & oLValues(iCount)
        oStream.WriteLine "D Values: " & oDValues(iCount)
        oStream.WriteLine "BackSet Values: " & oBacksetValues(iCount)
    Next
#End If
   
    oStream.WriteLine " "
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE SET REPORT - GET CHORD HEIGHT "
    oStream.WriteLine "================================================================"
    
 
    Dim dAccumulatedOffset()  As Double
    Dim dGirthLength()        As Double
    Dim dHeight()             As Double
    Dim oPosOnStLineElems     As IJElements
    Dim oPosOnCurveElems      As IJElements
    
    Dim oPosOnStraightLine  As IJDPosition
    Dim oPosOnCurve         As IJDPosition
 
    
    oTemplateSetRpt.GetChordHeights 1#, dAccumulatedOffset, dGirthLength, dHeight, oPosOnStLineElems, oPosOnCurveElems
    oStream.WriteLine " "
    oStream.WriteLine "Length (Str):" & vbTab & "Length(Curve):" & vbTab & "Height:" & vbTab & vbTab & "Point(Str)" & vbTab & vbTab & vbTab & "Point(Curve)"
    
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim str4 As String
    Dim str5 As String
    Dim str6 As String
    Dim str7 As String
    Dim str8 As String
    Dim str9 As String
    
    For iCount = LBound(dAccumulatedOffset) To UBound(dAccumulatedOffset)
 
         Set oPosOnStraightLine = oPosOnStLineElems.Item(iCount)
         Set oPosOnCurve = oPosOnCurveElems.Item(iCount)
         
         str1 = Round(dAccumulatedOffset(iCount), 4)
         str2 = Round(dGirthLength(iCount), 4)
         str3 = Round(dHeight(iCount), 4)
         str4 = Round(oPosOnStraightLine.x, 4)
         str5 = Round(oPosOnStraightLine.y, 4)
         str6 = Round(oPosOnStraightLine.z, 4)
         str7 = Round(oPosOnCurve.x, 4)
         str8 = Round(oPosOnCurve.y, 4)
         str9 = Round(oPosOnCurve.z, 4)
         
         oStream.WriteLine str1 & vbTab & vbTab & str2 & vbTab & vbTab & str3 & vbTab & vbTab & str4 & ", " & str5 & ", " & str6 & vbTab & vbTab & str7 & ", " & str8 & ", " & str9
    Next iCount

    oStream.WriteLine " "
    oStream.WriteLine "============================End of report ==========="
    

    'Close Log File
    oStream.Close

CleanUp:
    For i = 1 To oNavalArchRefCurveCol.Count
        oNavalArchRefCurveCol.Remove (i)
    Next
    Set oNavalArchRefCurveCol = Nothing

    Set oTemplateSetRpt = Nothing
    Set oSDPlateWrapper = Nothing
    Set oElements = Nothing
    Set oTemplateSet = Nothing
 

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub ReportTemplateInformation(oTemplate As IJMfgTemplate, oNavalArchRefCurveCol As Collection, oStream As TextStream)
     Const METHOD = "ReportTemplateInformation"
     
    On Error Resume Next

    Dim oTemplateReport As IJMfgTemplateReport
    Dim oPos As IJDPosition
    Dim i As Integer
    
    Set oTemplateReport = oTemplate
   
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE REPORT - POINTS "
    oStream.WriteLine "================================================================"
        
    Set oPos = oTemplateReport.GetPoint(AftSeamPoint)
    oStream.WriteLine "AftSeamPoint                           : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
    
    Set oPos = oTemplateReport.GetPoint(ForeSeamPoint)
    oStream.WriteLine "ForeSeamPoint                          : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
    
    Set oPos = oTemplateReport.GetPoint(SightLinePoint)
    oStream.WriteLine "SightLinePoint                         : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
   
    Set oPos = oTemplateReport.GetPoint(UpperEndMarkingLinePoint)
    oStream.WriteLine "UpperEndMarkingLinePoint               : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
      
    Set oPos = oTemplateReport.GetPoint(UpperEndTopLinePoint)
    oStream.WriteLine "UpperEndTopLinePoint                   : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
          
    Set oPos = oTemplateReport.GetPoint(UpperSeamPoint)
    oStream.WriteLine "UpperSeamPoint                         : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
    
    Set oPos = oTemplateReport.GetPoint(UpperSeamTopLinePoint)
    oStream.WriteLine "UpperSeamTopLinePoint                  : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
    
    Set oPos = oTemplateReport.GetPoint(BaseControlPoint)
    oStream.WriteLine "BaseControlPoint                       : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
      
    Set oPos = oTemplateReport.GetPoint(BaseControlTopLinePoint)
    oStream.WriteLine "BaseControlTopLinePoint                : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)

    Set oPos = oTemplateReport.GetPoint(LowerSeamPoint)
    oStream.WriteLine "LowerSeamPoint                         : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
    
    Set oPos = oTemplateReport.GetPoint(LowerSeamTopLinePoint)
    oStream.WriteLine "LowerSeamTopLinePoint                  : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
    
    Set oPos = oTemplateReport.GetPoint(LowerEndMarkingLinePoint)
    oStream.WriteLine "LowerEndMarkingLinePoint               : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
      
    Set oPos = oTemplateReport.GetPoint(LowerEndTopLinePoint)
    oStream.WriteLine "LowerEndTopLinePoint                   : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
    
    Set oPos = oTemplateReport.GetPoint(ControlLineAftEndPoint)
    oStream.WriteLine "ControlLineAftEndPoint                 : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
    
    Set oPos = oTemplateReport.GetPoint(ControlLineForeEndPoint)
    oStream.WriteLine "ControlLineForeEndPoint                : " & Round(oPos.x, 4) & ", " & Round(oPos.y, 4) & ", " & Round(oPos.z, 4)
  
    Set oPos = Nothing
    
'
'   Get the distance checked.
'
        
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE REPORT - DISTANCE "
    oStream.WriteLine "================================================================"
    oStream.WriteLine "GirthAtLowerSeam                                         : " & Round(oTemplateReport.GetDistance(GirthAtLowerSeam), 4)
    oStream.WriteLine "GirthAtUpperSeam                                         : " & Round(oTemplateReport.GetDistance(GirthAtUpperSeam), 4)
    oStream.WriteLine "GirthAtBaseControlLine                                   : " & Round(oTemplateReport.GetDistance(GirthAtBaseControlLine), 4)
    oStream.WriteLine "GirthBetweenUpperSeamAndBaseCtlLine                      : " & Round(oTemplateReport.GetDistance(GirthBetweenUpperSeamAndBaseCtlLine), 4)
    oStream.WriteLine "GirthBetweenBaseCtlPointAndAftPoint                      : " & Round(oTemplateReport.GetDistance(GirthBetweenBaseCtlPointAndAftPoint), 4)
    oStream.WriteLine "GirthBetweenBaseCtlPointAndForePoint                     : " & Round(oTemplateReport.GetDistance(GirthBetweenBaseCtlPointAndForePoint), 4)
    oStream.WriteLine "GirthBetweenBaseCtlPointAndControlLineAftEndPoint        : " & Round(oTemplateReport.GetDistance(GirthBetweenBaseCtlPointAndControlLineAftEndPoint), 4)
    oStream.WriteLine "GirthBetweenBaseCtlPointAndControlLineForeEndPoint       : " & Round(oTemplateReport.GetDistance(GirthBetweenBaseCtlPointAndControlLineForeEndPoint), 4)

    oStream.WriteLine "BetweenLowerSeamAndBaseCtlLine                           : " & Round(oTemplateReport.GetDistance(BetweenLowerSeamAndBaseCtlLine), 4)
    oStream.WriteLine "BetweenUpperSeamAndBaseCtlLine                           : " & Round(oTemplateReport.GetDistance(BetweenUpperSeamAndBaseCtlLine), 4)
    oStream.WriteLine "BetweenLowerSeamAndClosestTemplate                       : " & Round(oTemplateReport.GetDistance(BetweenLowerSeamAndClosestTemplate), 4)
    oStream.WriteLine "BetweenUpperSeamAndClosestTemplate                       : " & Round(oTemplateReport.GetDistance(BetweenUpperSeamAndClosestTemplate), 4)
    oStream.WriteLine "BetweenLowerSeamAndBaseCtlLineOnTopLine                  : " & Round(oTemplateReport.GetDistance(BetweenLowerSeamAndBaseCtlLineOnTopLine), 4)
    oStream.WriteLine "BetweenBaseCtlLineAndUpperSeamOnTopLine                  : " & Round(oTemplateReport.GetDistance(BetweenBaseCtlLineAndUpperSeamOnTopLine), 4)
    oStream.WriteLine "BetweenLowerSeamAndUpperSeamOnTopLine                    : " & Round(oTemplateReport.GetDistance(BetweenLowerSeamAndUpperSeamOnTopLine), 4)
    oStream.WriteLine "BetweenLowerEndAndLowerSeamOnTopLine                     : " & Round(oTemplateReport.GetDistance(BetweenLowerEndAndLowerSeamOnTopLine), 4)
    oStream.WriteLine "BetweenUpperSeamAndUpperEndOnTopLine                     : " & Round(oTemplateReport.GetDistance(BetweenUpperSeamAndUpperEndOnTopLine), 4)
    oStream.WriteLine "BetweenLowerEndAndUpperEndOnTopLine                      : " & Round(oTemplateReport.GetDistance(BetweenLowerEndAndUpperEndOnTopLine), 4)

    oStream.WriteLine "BetweenSeamAndTemplate                                   : " & Round(oTemplateReport.GetDistance(BetweenSeamAndTemplate), 4)
    oStream.WriteLine "(This is linear distance between Aft butt - BCL intersection and Template BCL bottom point, measured/projected along base plane)" & vbCrLf

    oStream.WriteLine "FromStartingEdgeToBaseCtlLine                            : " & Round(oTemplateReport.GetDistance(FromStartingEdgeToBaseCtlLine), 4)
    oStream.WriteLine "HU_HeightFromGroundFloor                                 : " & Round(oTemplateReport.GetDistance(HU_HeightFromGroundFloor), 4)
    oStream.WriteLine "HL_HeightFromGroundFloor                                 : " & Round(oTemplateReport.GetDistance(HL_HeightFromGroundFloor), 4)

#If SHOW_BACKSET Then
    oStream.WriteLine "BackSet                                                  : " & Round(oTemplateReport.GetDistance(BackSet), 4)
#End If

    oStream.WriteLine "MaxHeightFromButtLine ( F )                              : " & Round(oTemplateReport.GetDistance(MaxHeightFromButtLine), 4)
    
    oStream.WriteLine " "
    oStream.WriteLine "New enums "
    oStream.WriteLine "BetweenLowerEndAndBaseCtlLine                            : " & Round(oTemplateReport.GetDistance(BetweenLowerEndAndBaseCtlLine), 4)
    oStream.WriteLine "BetweenUpperEndAndBaseCtlLine                            : " & Round(oTemplateReport.GetDistance(BetweenUpperEndAndBaseCtlLine), 4)
    oStream.WriteLine "BetweenLowerEndAndLowerSeam                              : " & Round(oTemplateReport.GetDistance(BetweenLowerEndAndLowerSeam), 4)
    oStream.WriteLine "BetweenUpperSeamAndUpperEnd                              : " & Round(oTemplateReport.GetDistance(BetweenUpperSeamAndUpperEnd), 4)
    oStream.WriteLine "BetweenLowerEndAndUpperEnd                               : " & Round(oTemplateReport.GetDistance(BetweenLowerEndAndUpperEnd), 4)
    oStream.WriteLine "BetweenLowerSeamAndUpperSeam                             : " & Round(oTemplateReport.GetDistance(BetweenLowerSeamAndUpperSeam), 4)
    oStream.WriteLine "BetweenLowerSeamAndBaseCtlLine                           : " & Round(oTemplateReport.GetDistance(BetweenLowerSeamAndBaseCtlLine), 4)
    oStream.WriteLine "BetweenUpperSeamAndBaseCtlLine                           : " & Round(oTemplateReport.GetDistance(BetweenUpperSeamAndBaseCtlLine), 4)
    
    oStream.WriteLine "GirthBetweenLowerEndAndBaseCtlLine                       : " & Round(oTemplateReport.GetDistance(GirthBetweenLowerEndAndBaseCtlLine), 4)
    oStream.WriteLine "GirthBetweenUpperEndAndBaseCtlLine                       : " & Round(oTemplateReport.GetDistance(GirthBetweenUpperEndAndBaseCtlLine), 4)
    oStream.WriteLine "GirthBetweenLowerEndAndLowerSeam                         : " & Round(oTemplateReport.GetDistance(GirthBetweenLowerEndAndLowerSeam), 4)
    oStream.WriteLine "GirthBetweenUpperSeamAndUpperEnd                         : " & Round(oTemplateReport.GetDistance(GirthBetweenUpperSeamAndUpperEnd), 4)
    oStream.WriteLine "GirthBetweenLowerEndAndUpperEnd                          : " & Round(oTemplateReport.GetDistance(GirthBetweenLowerEndAndUpperEnd), 4)
    oStream.WriteLine "GirthBetweenLowerSeamAndUpperSeam                        : " & Round(oTemplateReport.GetDistance(GirthBetweenLowerSeamAndUpperSeam), 4)
    oStream.WriteLine "GirthBetweenLowerSeamAndBaseCtlLine                      : " & Round(oTemplateReport.GetDistance(GirthBetweenLowerSeamAndBaseCtlLine), 4)
    oStream.WriteLine "GirthBetweenUpperSeamAndBaseCtlLine                      : " & Round(oTemplateReport.GetDistance(GirthBetweenUpperSeamAndBaseCtlLine), 4)

'
' Get the reference curve information
'
    If Not oNavalArchRefCurveCol Is Nothing Then
        oStream.WriteLine "================================================================"
        oStream.WriteLine " TEMPLATE REPORT - GET REFERENCE "
        oStream.WriteLine "================================================================"
        oStream.WriteLine "Number of Reference Curves                               : " & oNavalArchRefCurveCol.Count
        oStream.WriteLine " "
        
        Dim dRBB          As Double
        Dim dRBH          As Double
        
        For i = 1 To oNavalArchRefCurveCol.Count
          dRBB = 0#
          dRBH = 0#
         oTemplateReport.GetReference oNavalArchRefCurveCol.Item(i), dRBB, dRBH
         oStream.WriteLine "RB.B                                                     : " & dRBB
         oStream.WriteLine "RB.H                                                     : " & dRBH
        Next
    End If
    
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE REPORT - GETANGLE "
    oStream.WriteLine "================================================================"
    oStream.WriteLine "AttachedAngle                            : " & Round(oTemplateReport.GetAngle(AttachedAngle), 4)
    oStream.WriteLine "  "

'
'   Get the Seam names checked.
'
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE REPORT - GETSEAM "
    oStream.WriteLine "================================================================"
    oStream.WriteLine "TemplateAftSeamName                      : " & oTemplateReport.GetSeam(TemplateAftSeam)
    oStream.WriteLine "TemplateFwdSeamName                      : " & oTemplateReport.GetSeam(TemplateForwardSeam)
    oStream.WriteLine "TemplateUpperSeamName                    : " & oTemplateReport.GetSeam(TemplateUpperSeam)
    oStream.WriteLine "TemplateLowerSeamName                    : " & oTemplateReport.GetSeam(TemplateLowerSeam)
    oStream.WriteLine " "

'
'   Get the Seam names checked.
'
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE REPORT - GETSEAMENTITY"
    oStream.WriteLine "================================================================"
    
    Dim oName As IJNamedItem
    
    Set oName = oTemplateReport.GetSeamEntity(TemplateAftSeam)
    oStream.WriteLine "TemplateAftSeamName                      : " & oName.name
    
    Set oName = oTemplateReport.GetSeamEntity(TemplateForwardSeam)
    oStream.WriteLine "TemplateFwdSeamName                      : " & oName.name
    
    Set oName = oTemplateReport.GetSeamEntity(TemplateUpperSeam)
    oStream.WriteLine "TemplateUpperSeamName                    : " & oName.name
    
    Set oName = oTemplateReport.GetSeamEntity(TemplateLowerSeam)
    oStream.WriteLine "TemplateLowerSeamName                    : " & oName.name
    oStream.WriteLine " "

'
'   Get the intenval distance checked.
'
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE REPORT - GET OFFSET HEIGHT "
    oStream.WriteLine "================================================================"

    Dim dInterval As Double
    Dim mInterval As Double
    
    Dim dHeight  As Double

    mInterval = oTemplateReport.GetDistance(BetweenLowerEndAndUpperEndOnTopLine)
    dInterval = 0#
    
    While dInterval < mInterval
        oTemplateReport.GetOffsetHeight dInterval, dHeight
        oStream.WriteLine "Interval   : " & dInterval & " Value            : " & Round(dHeight, 4)
        dInterval = dInterval + 0.2
    Wend
    
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE REPORT - GET OFFSET HEIGHT LIST "
    oStream.WriteLine "================================================================"
    
    Dim dOffsets() As Double, dHeights() As Double
    Dim nSteps As Long
    
    nSteps = Int(mInterval / 0.2)
    
    ReDim dOffsets(0 To nSteps)
    Dim cnt As Integer
    
    For cnt = 0 To nSteps
        dOffsets(cnt) = cnt * 0.2
    Next
        
    dHeights = oTemplateReport.GetOffsetHeightList(dOffsets)
     
    For cnt = 0 To nSteps
        oStream.WriteLine "Interval   : " & Round(dOffsets(cnt), 4) & "Value            : " & Round(dHeights(cnt), 4)
    Next
    
    oStream.WriteLine " "
    oStream.WriteLine "================================================================"
    oStream.WriteLine " TEMPLATE REPORT - GET CHORD HEIGHT "
    oStream.WriteLine "================================================================"
    
    Dim dAccumulatedOffset()  As Double
    Dim dGirthLength()        As Double
  '  Dim dHeights()             As Double
    Dim oPosOnStLineElems     As IJElements
    Dim oPosOnCurveElems      As IJElements
    Dim iCount              As Integer
    Dim oPosOnStraightLine  As IJDPosition
    Dim oPosOnCurve         As IJDPosition
    
    oTemplateReport.GetChordHeights 1#, dAccumulatedOffset, dGirthLength, dHeights, oPosOnStLineElems, oPosOnCurveElems
    
    oStream.WriteLine "Length (Str):" & vbTab & "Length(Curve):" & vbTab & "Height:" & vbTab & vbTab & "Point(Str)" & vbTab & vbTab & vbTab & "Point(Curve)"
    
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim str4 As String
    Dim str5 As String
    Dim str6 As String
    Dim str7 As String
    Dim str8 As String
    Dim str9 As String
    
    For iCount = LBound(dAccumulatedOffset) To UBound(dAccumulatedOffset)
 
         Set oPosOnStraightLine = oPosOnStLineElems.Item(iCount)
         Set oPosOnCurve = oPosOnCurveElems.Item(iCount)
         
         str1 = Round(dAccumulatedOffset(iCount), 4)
         str2 = Round(dGirthLength(iCount), 4)
         str3 = Round(dHeights(iCount), 4)
         str4 = Round(oPosOnStraightLine.x, 4)
         str5 = Round(oPosOnStraightLine.y, 4)
         str6 = Round(oPosOnStraightLine.z, 4)
         str7 = Round(oPosOnCurve.x, 4)
         str8 = Round(oPosOnCurve.y, 4)
         str9 = Round(oPosOnCurve.z, 4)
         
         oStream.WriteLine str1 & vbTab & vbTab & str2 & vbTab & vbTab & str3 & vbTab & vbTab & str4 & ", " & str5 & ", " & str6 & vbTab & vbTab & str7 & ", " & str8 & ", " & str9
    Next iCount

    oStream.WriteLine " "
    oStream.WriteLine " "

    Set oTemplateReport = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub


