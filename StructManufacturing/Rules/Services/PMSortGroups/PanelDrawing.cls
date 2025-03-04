VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PanelDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Implements IJPMSortGroups

Dim m_objBins As StrMfgDrawingAVP.CBins
Dim m_objRects As StrMfgDrawingAVP.CRectangles

Dim m_objRects1 As RAD2D.Groups
Dim m_objRect1 As RAD2D.Group

Dim m_objRects2 As RAD2D.Groups
Dim m_objRect2 As RAD2D.Group

Dim m_objApp As Object
Dim m_objRadApp As RAD2D.Application
Dim m_objDoc As RAD2D.Document
Dim m_objSheets As RAD2D.Sheets
Dim m_objOrigSheet As RAD2D.Sheet

Dim m_dBinMinX As Double, m_dBinMinY As Double, m_dBinMaxX As Double, m_dBinMaxY As Double
Dim m_dBinCenterX As Double, m_dBinCenterY As Double

Dim m_lRotations() As Long

Const TOLERANCE As Double = 0.000001

Private Sub IJPMSortGroups_SortGroups(oSheet As Object)

    Dim objRect As CRectangle
    Dim lCount As Long, i As Long
    Dim oColl As New Collection
    Dim objRect1 As RAD2D.Group
    
    Set m_objOrigSheet = oSheet
    
    'Get rid of bins and rectangles if they exist.
    Set m_objBins = Nothing
    Set m_objRects = Nothing
    
    'Initialize m_objBins and m_objRects to new objects.
    Set m_objBins = New CBins
    Set m_objRects = New CRectangles
    
    Set m_objRects1 = m_objOrigSheet.Groups
    
    ReDim m_lRotations(1 To 1)
    
    '************** Panel Drawing **********************'
    
    Dim ii As Integer
    
    Dim oMainGroup As RAD2D.Group
    Set oMainGroup = m_objRects1.Item(1)
    
    Dim sMainGroupGUID As String
    Dim oMainGroupAttr As RAD2D.Attribute
    
    Dim sPartGUID As String
    Dim dShrinkageMainValue As Double, dShrinkageSubValue As Double
    
    ' Index 3 is SMS_PART_INFO||PART_GUID
    Set oMainGroupAttr = oMainGroup.AttributeSets("General").Item(3)
    sMainGroupGUID = oMainGroupAttr.Value
    
    Dim oPOM As IJDPOM
    Dim o2dTransMat As IJDT4x4, o3dTransMat As IJDT4x4
    
    Set o2dTransMat = New DT4x4
    
    Dim oCmnAppUtil As IJDCmnAppGenericUtil
    Dim oAccessMiddle As IJDAccessMiddle
    
    Set oCmnAppUtil = New CmnAppGenericUtil
    oCmnAppUtil.GetActiveConnection oAccessMiddle
    
    Set oPOM = GetPOM
    
    ' Get the MFG object
    Dim oMoniker As IMoniker
    Dim oMainGroupPart As Object
    Dim oMainGroupMFGPart As IJMfgPlatePart
    
    Set oMoniker = oPOM.MonikerFromDbIdentifier("{" & sMainGroupGUID & "}")
    Set oMainGroupMFGPart = oPOM.GetObject(oMoniker)
    
    ' Get UV Marks
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
    Set oMfgPlateWrapper.object = oMainGroupMFGPart
    Dim oGeomCol2d As IJMfgGeomCol2d
    Set oGeomCol2d = oMfgPlateWrapper.GetFinal2dGeometries
    
    Dim oMfgMainGeom2d As IJMfgGeom2d, oMfgGeom2d As IJMfgGeom2d, oMfgMainGeom3d As IJMfgGeom2d
    '  Get 2D Transformation Matrix
    Dim oMfgUtilMathGeom    As IJMfgUtilMathGeom
    Dim oMfgMGHelper As New MfgMGHelper
    
    Dim jj As Integer
    For jj = 1 To oGeomCol2d.Getcount
        Set oMfgGeom2d = oGeomCol2d.GetGeometry(jj)
        If oMfgGeom2d.GetGeometryType = STRMFG_UV_MARK And oMfgGeom2d.GetSubGeometryType = STRMFG_UV_MARK Then
            Set oMfgMainGeom3d = oMfgGeom2d
        End If

        If oMfgGeom2d.GetGeometryType = STRMFG_UV_MARK And oMfgGeom2d.GetSubGeometryType <> STRMFG_UV_MARK Then
            Set oMfgMainGeom2d = oMfgGeom2d
        End If
    Next jj
    
    '*** LOG THE DEBUG DATA ***'
    'Dim FileNumber
    'Dim slogFileName As String
    'slogFileName = "C:\Temp\PanelDrawing.txt"
    'FileNumber = FreeFile
    'Open slogFileName For Append As #FileNumber
    'Write #FileNumber, "PART 1:"
    
    Dim dXMin As Double, dYMin As Double, dXMax As Double, dYMax As Double
    oMainGroup.Range dXMin, dYMin, dXMax, dYMax
    
    'Write #FileNumber, "Range: dXMin:" & dXMin & " dYMin:" & dYMin & "dXMax :" & dXMax & " dYMax:" & dYMax
    '**************************'
    
    Dim oMfgMainGeom3DCS As IJComplexString, oMfgMainGeom2DCS As IJComplexString
    Set oMfgUtilMathGeom = New MfgMathGeom
    
    oMfgMGHelper.CloneComplexString oMfgMainGeom3d.GetGeometry, oMfgMainGeom3DCS
    oMfgMGHelper.CloneComplexString oMfgMainGeom2d.GetGeometry, oMfgMainGeom2DCS

    Dim oOrigTrans3DMat As IJDT4x4, oOrigTrans2DMat As IJDT4x4, oOrigTrans2DMatInv As IJDT4x4, oTransMat As IJDT4x4
    Dim oOriginLineStr As IJComplexString
    Dim oOriginLS As IJLineString
    Set oOriginLineStr = New ComplexString3d
    
    Set oOriginLS = New LineString3d
    
    Dim PointsArray(0 To 8) As Double
    
    PointsArray(0) = 0
    PointsArray(1) = 0.01
    PointsArray(2) = 0
    
    PointsArray(3) = 0
    PointsArray(4) = 0
    PointsArray(5) = 0
    
    PointsArray(6) = 0.01
    PointsArray(7) = 0
    PointsArray(8) = 0
    
    oOriginLS.SetPoints 3, PointsArray
    oOriginLineStr.AddCurve oOriginLS, True

    Set oOrigTrans3DMat = oMfgUtilMathGeom.ComputeFeatureTransf(oMfgMainGeom3DCS, oOriginLineStr)
    Set oOrigTrans2DMat = oMfgUtilMathGeom.ComputeFeatureTransf(oMfgMainGeom2DCS, oOriginLineStr)
    Set oOrigTrans2DMatInv = oMfgUtilMathGeom.ComputeFeatureTransf(oMfgMainGeom2DCS, oOriginLineStr)
    
    oOrigTrans2DMatInv.Invert
    
    Dim oDoubleMat(16) As Double
    
    ' Process subsequent groups
    For ii = 2 To m_objRects1.Count
        
        '''''        1. For each group, get its part oid.
        '''''        2. From oid, get the part#1 object.
        '''''        3. Get geometriesafterunfold (3D). Get UV Mark#1. Get TransMatrix#1.
        '''''        4. Get the part#2 object from next group.
        '''''        5. Get geometriesafterunfold (3D). Get UV Mark#2. Get TransMatrix#2.
        '''''        6. Align part#2 such that it's UV Mark#2 has same orientation as UV Mark#1
        Dim var1 As Variant
        Dim oGroupAttribute As RAD2D.Attribute
        Dim oNextGroupPart As IJMfgPlatePart
        
        Dim oAttrSet As AttributeSet
        Set oAttrSet = m_objRects1.Item(ii).AttributeSets("General")
        
        ' Index 3 is SMS_PART_INFO||PART_GUID
        'Set oGroupAttribute = m_objRects1.Item(ii).AttributeSets("General").Item(3)
        Dim k As Integer
        For k = 1 To oAttrSet.Count
            If oAttrSet.Item(k).Name = "SMS_PART_INFO||PART_GUID" Then
                sPartGUID = oAttrSet.Item(k).Value
            ElseIf oAttrSet.Item(k).Name = "SMS_PART_SHRINKAGE_INFO||MAIN_VALUE" Then
                dShrinkageMainValue = oAttrSet.Item(k).Value
            ElseIf oAttrSet.Item(k).Name = "SMS_PART_SHRINKAGE_INFO||SUB_VALUE" Then
                dShrinkageSubValue = oAttrSet.Item(k).Value
            End If
        Next k
        
        'Write #FileNumber, "***** PART # *******'" & ii
        
        'sPartGUID = oGroupAttribute.Value
        
        Set oMoniker = oPOM.MonikerFromDbIdentifier("{" & sPartGUID & "}")
        Set oNextGroupPart = oPOM.GetObject(oMoniker)
                
        ' Get UV Marks
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oNextGroupPart
        Set oGeomCol2d = oMfgPlateWrapper.GetFinal2dGeometries
        
        Dim oMfgNextGeom2d As IJMfgGeom2d, oMfgNextGeom3d As IJMfgGeom2d
        Dim oClonedCS As IJComplexString
        
        For jj = 1 To oGeomCol2d.Getcount
            Set oMfgGeom2d = oGeomCol2d.GetGeometry(jj)
            If oMfgGeom2d.GetGeometryType = STRMFG_UV_MARK And oMfgGeom2d.GetSubGeometryType = STRMFG_UV_MARK Then
                Set oMfgNextGeom3d = oMfgGeom2d
            End If
GetNext3dGeom:
            If oMfgGeom2d.GetGeometryType = STRMFG_UV_MARK And oMfgGeom2d.GetSubGeometryType <> STRMFG_UV_MARK Then
                Set oMfgNextGeom2d = oMfgGeom2d
            End If
        Next jj
        
        
        Dim oMfgNextGeom3DCS As IJComplexString, oMfgNextGeom2DCS As IJComplexString
        oMfgMGHelper.CloneComplexString oMfgNextGeom3d.GetGeometry, oMfgNextGeom3DCS
        oMfgMGHelper.CloneComplexString oMfgNextGeom2d.GetGeometry, oMfgNextGeom2DCS
        
        oMfgMGHelper.CloneComplexString oMfgNextGeom3DCS, oClonedCS
        oClonedCS.Transform oOrigTrans3DMat
        
        Dim o3DPlateToPlateTransMat As IJDT4x4
        Set o3DPlateToPlateTransMat = oMfgUtilMathGeom.ComputeFeatureTransf(oOriginLineStr, oClonedCS)
                
        Dim xTrans As Double, yTrans As Double
        
        ' Modify the matrix
        xTrans = o3DPlateToPlateTransMat.IndexValue(12)
        yTrans = o3DPlateToPlateTransMat.IndexValue(13)
        
        If dShrinkageMainValue <> 0 Then
            xTrans = xTrans * dShrinkageMainValue
        End If
        If dShrinkageSubValue <> 0 Then
            yTrans = yTrans * dShrinkageSubValue
        End If
        
        o3DPlateToPlateTransMat.IndexValue(12) = xTrans
        o3DPlateToPlateTransMat.IndexValue(13) = yTrans
        
        Set o2dTransMat = oMfgUtilMathGeom.ComputeFeatureTransf(oMfgNextGeom2d, oMfgMainGeom2d)
        o2dTransMat.MultMatrix oOrigTrans2DMat
        o2dTransMat.MultMatrix o3DPlateToPlateTransMat
        o2dTransMat.MultMatrix oOrigTrans2DMatInv
        
        Dim oGeom As IJDGeometry
        Dim oTransform As Transform
        
        Set m_objRadApp = oSheet.Application.Radapplication
        
        Set oTransform = m_objRadApp.GeometryServices.CreateTransform
        'm_objRects1.Item(ii).Transform oTransform
        

        o2dTransMat.Get oDoubleMat(0)
        oTransform.Set4x4Matrix oDoubleMat
        
        m_objRects1.Item(ii).Transform oTransform
    Next ii
    
    
    'Set ScalingGroup attribute to "PanelLayer" in order to put this group into
    '   .. the rectangle with layer = "PanelLayer"
    Dim oPanelDrawingGroup As Group
    Set oPanelDrawingGroup = MakeOneGroupToScale(m_objOrigSheet)
    
    Dim oAttributeSet As AttributeSet
    Set oAttributeSet = oPanelDrawingGroup.AttributeSets.Add("Panel")
    
    Dim oAttribute As RAD2D.Attribute
    Set oAttribute = oAttributeSet.Add("ScalingGroup", igAttrTypeString)
    
    oAttribute.Value = "Panel"
    
    'Ungroup the entire group
    'oPanelDrawingGroup.Ungroup
    
    'm_objDoc.SelectSet.RemoveAll
    m_objOrigSheet.Document.SelectSet.RemoveAll
    '***************************************************'
    
    
    '************** Side Drawing **********************'
    Dim bSuccess As Boolean
    'Check if any of rectangles in Background has Side as attribute. If yes, create parts drawing
    For i = 1 To m_objOrigSheet.Background.Rectangles2d.Count
        If UCase(m_objOrigSheet.Background.Rectangles2d.Item(i).AttributeSets("Panel").Item("Name")) = "SIDE" Then
            bSuccess = True
            Exit For
        Else
        End If
    Next i
    
    If bSuccess Then
    'Copy 1st item
    Dim objPart1 As Group
    
    'Since oPanelDrawingGroup contains all original part groups, use it instead of m_objRects1
    Set objPart1 = oPanelDrawingGroup.Item(1).Duplicate(0, 0)

    'Add to selectset
    objPart1.Select
    
    'Get range of first part
    Dim dMinX As Double, dMinY As Double, dMaxX As Double, dMaxY As Double
    Dim dNewMinX As Double, dNewMinY As Double, dNewMaxX As Double, dNewMaxY As Double
    
    objPart1.Range dMinX, dMinY, dMaxX, dMaxY
    
    'Set m_objRects1 = m_objOrigSheet.Groups
    
    For ii = 2 To oPanelDrawingGroup.Count
    
        Set m_objRect1 = oPanelDrawingGroup.Item(ii)
        
        'Get the range of new object
        m_objRect1.Range dNewMinX, dNewMinY, dNewMaxX, dNewMaxY
        
        'Copy and move the object so its min matches the point vertically above the prev groups min
        Set m_objRect2 = m_objRect1.Duplicate(dMinX - dNewMinX, dMaxY - dNewMinY)
        
        'Set the range as prev groups range
        m_objRect2.Range dMinX, dMinY, dMaxX, dMaxY
        
        'Add the object into select set
        m_objRect2.Select
    
    Next ii
    
    'Create a new group for side/parts
    Dim oSideGroup As Group
    Set oSideGroup = m_objOrigSheet.Groups.Add()
    
    'Set ScalingGroup attribute to "SideLayer" in order to put this group into
    '   .. the rectangle with layer = "SideLayer"
    Set oAttributeSet = Nothing
    Set oAttributeSet = oSideGroup.AttributeSets.Add("Panel")
    
    Set oAttribute = Nothing
    Set oAttribute = oAttributeSet.Add("ScalingGroup", igAttrTypeString)
    
    oAttribute.Value = "Side"
    End If

    '**************************************************'
    Exit Sub
ErrorHandler:


End Sub

'********************************************************************
' Routine: MakeOneGroupToScale
' Abstract: This method groups all drawingobjects of sheet into a single group for scaling.
'********************************************************************
Private Function MakeOneGroupToScale(oSheet As Sheet) As Group
    On Error GoTo ErrorHandler
    Const METHOD As String = "MakeOneGroupToScale"
    
    Dim oSymbol         As RAD2D.Symbol2d
    Dim oGroup          As RAD2D.Group
    Dim oGroups         As RAD2D.Groups
    Dim oSelectSet      As RAD2D.SelectSet
    Dim iSizeOfArray    As Integer
    Dim i               As Integer
    
    m_bProcBeforeDwgObjSel = False
    iSizeOfArray = oSheet.DrawingObjects.Count
    If iSizeOfArray <> 1 Then
        For i = 1 To iSizeOfArray
            If (oSheet.DrawingObjects.Item(i).AttributeSets.Count <> 0) Or _
            (oSheet.DrawingObjects.Item(i).Type = igGroup) Or _
            (oSheet.DrawingObjects.Item(i).Type = igTextBox) Or _
            (oSheet.DrawingObjects.Item(i).Type = igPoint2d) Or _
            (oSheet.DrawingObjects.Item(i).Type = igArc2d) Or _
            (oSheet.DrawingObjects.Item(i).Type = igLine2d) Then
                oSheet.DrawingObjects.Item(i).Select
            Else
                iSizeOfArray = iSizeOfArray
            End If
        Next
        For Each oSymbol In oSheet.Symbols
            oSymbol.Select
        Next oSymbol
        
        Set oGroups = oSheet.Groups
        'TR-185029
        'The below line Triggers a call-back to BeforeDrawingObjectsSelected,
        'which bounces back into GetAttributeValue,
        'This in turn triggers a call to jengine.dll!Collection::FillBuffer
        'on line #535 which triggers some RAD asserts.
        'Only happens when loading a second set of part(s) from an XML file,
        'after the Part Monitor has already been loaded.
        'Added a public boolean value that allows controled access to
        'the BeforeDrawingObjectsSelected routine,
        'and also to the AfterDrawingObjectsSelected routine.
        Set oGroup = oGroups.Add()
        m_oDoc.SelectSet.RemoveAll
    End If
    m_bProcBeforeDwgObjSel = True
    Set MakeOneGroupToScale = oGroup
CleanUp:
    Set oSymbol = Nothing
    Set oGroup = Nothing
    Set oGroups = Nothing
    Set oSelectSet = Nothing

    Exit Function
ErrorHandler:

    Resume Next
End Function


Private Sub Start()

    Dim objRect As StrMfgDrawingAVP.CRectangle
    Dim lCount As Long, i As Long
    Dim oColl As New Collection
    Dim objRect1 As RAD2D.Group
    
    'May want to get rid of all additional sheets besides m_objOrigSheet at this time.
'    With m_objSheets
'
'        'MsgBox .Item(1).Name
'        'MsgBox .Item(2).Name
'
'        lCount = .Count
'        For i = lCount To 2 Step -1
'
'            .Item(i).Delete
'
'        Next i
'
'    End With

    'Get rid of bins and rectangles if they exist.
    Set m_objBins = Nothing
    Set m_objRects = Nothing
    
    'Initialize m_objBins and m_objRects to new objects.
    Set m_objBins = New StrMfgDrawingAVP.CBins
    Set m_objRects = New StrMfgDrawingAVP.CRectangles
    
    Set m_objRects1 = m_objOrigSheet.Groups
    
    ReDim m_lRotations(1 To 1)
    
    'Get the coordinates of the initial RAD bin.
    GetBinCoordinates

    'm_objBins will initially be empty.  We must add one bin to get started.  It will
    '   automatically add an initial greenspace to itself that will be the size of the bin.
    m_objBins.AddByDefiningCoordinates m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY
    
    'We need to initially populate m_objRects with all the rectangles
    '   we need to eventually place.
    Set oColl = PopulateRectanglesCollection
    
    'Now, sort the rectangles according to the bin's location preferences.
    'm_objRects.Sort
    Set oColl = MergeSort(oColl, 1, m_objRects1.Count)
    
    'SortGroupsInSheet
    
    'Add each rectangle in turn to a bin.
    For Each objRect1 In oColl
    
        HandleNewRectangle objRect1
        
    Next objRect1
    
    'Display all the calculated results.
    'Commit

End Sub


Private Function PopulateRectanglesCollection() As Collection

    On Error Resume Next
    
    Dim objRADRects As RAD2D.Rectangles2d
    Dim objRADRect As RAD2D.Rectangle2d
    Dim objAVPRect As StrMfgDrawingAVP.CRectangle
    Dim objAttSets As RAD2D.AttributeSets
    Dim objAttSet As RAD2D.AttributeSet
    Dim i As Long, lCount As Long
    
    
    Dim objRADRects1 As RAD2D.Groups
    Dim objRADRect1 As RAD2D.Group
    Dim objAVPRect1 As RAD2D.Group
    Dim objAVPRects1 As RAD2D.Groups
    Dim oGrpCol As New Collection
    'Rnd -1
    
    'Randomize CLng(Text2.Text)
    
    Randomize
    
    'Get all rectangles off the sheet, make StrMfgDrawingAVPRectangles out of them,
    '   and place them in m_objRects collection.
    'During the loop, also set the preferences on each rectangle.
    
    Set objRADRects1 = m_objOrigSheet.Groups ' GET ALL GROUPS FROM SHEET HERE
    
    If Check2.Value = vbUnchecked Then
    
        For Each objRADRect1 In objRADRects1
        
            'NEEDS_WORK:
            'Eventually, we will check to see if a particular preferences attribute set exists.
            '   If it does, then we proceed and pass this attribute set into SetRectanglePreferences.
            '   If not, then we go to the next rectangle.
            
            'If the "Rectangle" attribute set exists, then we have a valid rectangle.
            With objRADRect1
            
                Set objAttSets = .AttributeSets
                Set objAttSet = objAttSets.Item("Bin")
                
                If objAttSet Is Nothing Then
                
                    'objAVPRect is group of selected items
'                    m_objRects1.Item.Select
'                    Set objAVPRect1 = m_objRects1.Add()
'                    SetRectanglePreferences objAVPRect
                        
                    oGrpCol.Add objRADRect1
                    
                    
                Else
                
                    Set objAttSet = Nothing
                
                End If
                
            End With
        
        Next objRADRect1
        
    Else
    
        '*** NEED TO BE WORKED ON - NINAD ***'
    
        lCount = CLng(Text1.Text)
        For i = 1 To lCount
        
            'Set objAVPRect = m_objRects.AddByDimensions(Rnd * 0.10795, Rnd * 0.1397)
            Set objAVPRect = m_objRects.AddByDimensions(Rnd * 0.08795 + 0.02, Rnd * 0.1197 + 0.02)
            'Set objAVPRect = m_objRects.AddByDimensions(0.01, 0.01)
        
        Next i
    
    End If
    
    Set PopulateRectanglesCollection = oGrpCol
    
End Function

Private Sub SetRectanglePreferences(ByVal objRect As StrMfgDrawingAVP.CRectangle)

    'NEEDS_WORK:
    'This will eventually probably take an attribute set as well.
    
End Sub



Private Sub GetBinCoordinates()

    On Error Resume Next

'    Dim objAttSets As RAD2D.AttributeSets
'    Dim objAttSet As RAD2D.AttributeSet
'
'    Dim objRects1 As RAD2D.Groups
'    Dim objRect1 As RAD2D.Group
'
'
'    Set objRects1 = m_objOrigSheet.Groups
'
'    'MsgBox "Count:" & objRects1.Count
'
'    'See which rectangle on m_objOrigSheet supports the "Bin" attribute set.
'    '   This is our bin.
'    For Each objRect1 In objRects1
'
'        Set objAttSets = objRect1.AttributeSets
'        Set objAttSet = objAttSets.Item("Bin")
'
'        'If Not objAttSet Is Nothing Then
'
'            objRect1.Range m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY
'            'MsgBox "m_dBinMinX:" & m_dBinMinX & "  m_dBinMinY:" & m_dBinMinY & "   m_dBinMaxX:" & m_dBinMaxX & "   m_dBinMaxY:" & m_dBinMinX
'
'
'            m_dBinCenterX = m_dBinMinX
'            m_dBinCenterY = m_dBinMinY
'            'Exit For
'
'
'        'End If
'
'    Next objRect1
    
    
    m_dBinMinX = X_ORIGIN
    m_dBinMinY = Y_ORIGIN
    m_dBinMaxX = GLOBAL_X_WIDTH
    m_dBinMaxY = GLOBAL_Y_WIDTH
    
    m_dBinCenterX = m_dBinMinX
    m_dBinCenterY = m_dBinMinY
    
    

End Sub



Private Sub HandleNewRectangle(ByRef objRect As RAD2D.Group)
    
    Dim dRectWidth As Double, dRectHeight As Double
    Dim bFoundFit As Boolean
    Dim objBin As StrMfgDrawingAVP.CBin
    Dim objGSS As StrMfgDrawingAVP.CGreenSpaces
    Dim objGS As StrMfgDrawingAVP.CGreenSpace
    Dim lBin As Long
    Dim dHeight As Double, dWidth As Double
    
    GetRange objRect, dRectHeight, dRectWidth
    
    'Go through the greenspaces in sorted order and see what is the first one into which it
    '   will fit.  If there are no greenspaces that will accommodate this rectangle, then
    '   we need to create a new bin (which will create a corresponding initial greenspace) and
    '   place our rectangle in that one.
    lBin = 0
    ReDim Preserve m_lRotations(1 To m_objBins.Count)
    For Each objBin In m_objBins
        
        lBin = lBin + 1
        
        Set objGSS = objBin.GreenSpaces
        
        For Each objGS In objGSS
            
            bFoundFit = True
            
            With objGS
            
                Select Case True
                
                    Case .Width < dRectWidth, .Height < dRectHeight
                        
                        bFoundFit = False
                
                End Select
                
            End With
            
            If bFoundFit Then
            
                'Need to place the rectangle into the greenspace according to its preferences.
                AddRectangleToGreenSpace objBin, objGS, objRect, dRectWidth, dRectHeight
                GoTo DONE
            
            End If
            
        Next objGS
        
        If Not bFoundFit Then
        
            For Each objGS In objGSS
            
                bFoundFit = True
                
                With objGS
                
                    Select Case True
                    
                        Case .Height < dRectWidth, .Width < dRectHeight
                            
                            bFoundFit = False
                    
                    End Select
                    
                End With
                
                If bFoundFit Then
                
                    'Increment rotated rectangle count
                    m_lRotations(lBin) = m_lRotations(lBin) + 1
                
                    'Need to place the rectangle into the greenspace according to its preferences.
                    AddRectangleToGreenSpace objBin, objGS, objRect, dRectHeight, dRectWidth
                    GoTo DONE
                
                End If
                
            Next objGS
        
        End If
        
    Next objBin
    
DONE:
    
    'If bFoundFit Then
    
        'If we found a place in a greenspace for this rectangle, then we sort the
        '   greenspaces collection containing that greenspace according to location
        '   preferences.  The Else needs no sort because it deals with a greenspaces
        '   collection containing only one greenspace.
        'objGSS.Sort
    
    If Not bFoundFit Then
    
        'If we didn't find a greenspace into which this rectangle will fit, then we need to
        '   add a new bin and place it in the initial greenspace there.
        
        Set objBin = m_objBins.AddByDefiningCoordinates(m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY)
        
        Set objGSS = objBin.GreenSpaces
        
        Set objGS = objGSS.Item(1)
        
        AddRectangleToGreenSpace objBin, objGS, objRect, dRectWidth, dRectHeight
      
    End If
    
    'Now we should shrink our previous greenspaces.
    ShrinkGreenSpaces objRect, objGSS
    
    objGSS.Sort

End Sub

Private Sub AddRectangleToGreenSpace(ByVal objBin As StrMfgDrawingAVP.CBin, ByVal objGS As StrMfgDrawingAVP.CGreenSpace, ByRef objRect As RAD2D.Group, ByVal dRectWidth As Double, dRectHeight As Double)
    
    'NEEDS_WORK:
    'Need to check for objRect's preferences.  If it has no location preference, then it inherits
    '   the bin's location preferences.  Right now, we are defaulting to Bottom Then Left.
    
    'This is for bottom-then-left positioning:
    'Dim dMinX As Double, dMinY As Double, dMaxX As Double, dMaxY As Double
    'objGS.GetDefiningCoordinates dMinX, dMinY, dMaxX, dMaxY
    'Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dMinX, dMinY, dMinX + dRectWidth, dMinY + dRectHeight)
    
    'This is for center-center positioning:
    Dim dRectMinX As Double, dRectMinY As Double, dRectMaxX As Double, dRectMaxY As Double
    Dim dGSMinX As Double, dGSMinY As Double, dGSMaxX As Double, dGSMaxY As Double
    Dim dGSCenterX As Double, dGSCenterY As Double
    Dim dHalfRectWidth As Double, dHalfRectHeight As Double
    
    objGS.GetDefiningCoordinates dGSMinX, dGSMinY, dGSMaxX, dGSMaxY
    dRectMinX = m_dBinCenterX - dRectWidth / 2#
    dRectMinY = m_dBinCenterY - dRectHeight / 2#
    dRectMaxX = m_dBinCenterX + dRectWidth / 2#
    dRectMaxY = m_dBinCenterY + dRectHeight / 2#
    
    If FPGreaterThan(Abs(dGSMaxY - m_dBinCenterY), Abs(dGSMinY - m_dBinCenterY)) Then
    
        'Bottom side of greenspace is closer to center of bin than top side
        '   of greenspace.
        
        If FPGreaterThan(Abs(dGSMinX - m_dBinCenterX), Abs(dGSMaxX - m_dBinCenterX)) Then
        
            'Right side of greenspace is closer to center of bin than left side
            '   of greenspace.
            'Place rectangle at bottom right of greenspace.
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMaxX - dRectWidth, dGSMinY, dGSMaxX, dGSMinY + dRectHeight, objRect)
        
        ElseIf FPGreaterThan(Abs(dGSMaxX - m_dBinCenterX), Abs(dGSMinX - m_dBinCenterX)) Then
        
            'Left side of greenspace is closer to center of bin than right side
            '   of greenspace.
            'Place rectangle at bottom left of greenspace.
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMinX, dGSMinY, dGSMinX + dRectWidth, dGSMinY + dRectHeight, objRect)
            'set objRect = objBin.Rectangles.
            
        Else
        
            'Left and right sides of greenspace are equidistant from center of bin.
            'Place rectangle at bottom center of greenspace.
            dGSCenterX = (dGSMinX + dGSMaxX) / 2#
            dHalfRectWidth = dRectWidth / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSCenterX - dHalfRectWidth, dGSMinY, dGSCenterX + dHalfRectWidth, dGSMinY + dRectHeight, objRect)
        
        End If
        
    ElseIf FPGreaterThan(Abs(dGSMinY - m_dBinCenterY), Abs(dGSMaxY - m_dBinCenterY)) Then
    
        'Top side of greenspace is closer to center of bin than bottom side
        '   of greenspace.
        
        If FPGreaterThan(Abs(dGSMinX - m_dBinCenterX), Abs(dGSMaxX - m_dBinCenterX)) Then
        
            'Right side of greenspace is closer to center of bin than left side
            '   of greenspace.
            'Place rectangle at top right of greenspace.
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMaxX - dRectWidth, dGSMaxY - dRectHeight, dGSMaxX, dGSMaxY, objRect)
            
        ElseIf FPGreaterThan(Abs(dGSMaxX - m_dBinCenterX), Abs(dGSMinX - m_dBinCenterX)) Then
        
            'Left side of greenspace is closer to center of bin than right side
            '   of greenspace.
            'Place rectangle at top left of greenspace.
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMinX, dGSMaxY - dRectHeight, dGSMinX + dRectWidth, dGSMaxY, objRect)
            
        Else
        
            'Left and right sides of greenspace are equidistant from center of bin.
            'Place rectangle at top center of greenspace.
            dGSCenterX = (dGSMinX + dGSMaxX) / 2#
            dHalfRectWidth = dRectWidth / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSCenterX - dHalfRectWidth, dGSMaxY - dRectHeight, dGSCenterX + dHalfRectWidth, dGSMaxY, objRect)
        
        End If
    
    Else
    
        'Top and bottom sides of greenspace are equidistant from center of bin.
        
        If FPGreaterThan(Abs(dGSMinX - m_dBinCenterX), Abs(dGSMaxX - m_dBinCenterX)) Then
        
            'Right side of greenspace is closer to center of bin than left side
            '   of greenspace.
            'Place rectangle at center right of greenspace.
            dGSCenterY = (dGSMinY + dGSMaxY) / 2#
            dHalfRectHeight = dRectHeight / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMaxX - dRectWidth, dGSCenterY - dHalfRectHeight, dGSMaxX, dGSCenterY + dHalfRectHeight, objRect)
        
        ElseIf FPGreaterThan(Abs(dGSMaxX - m_dBinCenterX), Abs(dGSMinX - m_dBinCenterX)) Then
        
            'Left side of greenspace is closer to center of bin than right side
            '   of greenspace.
            'Place rectangle at center left of greenspace.
            dGSCenterY = (dGSMinY + dGSMaxY) / 2#
            dHalfRectHeight = dRectHeight / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSMinX, dGSCenterY - dHalfRectHeight, dGSMinX + dRectWidth, dGSCenterY + dHalfRectHeight, objRect)
            
        Else
        
            'Left and right sides of greenspace are equidistant from center of bin.
            'Place rectangle at center center of greenspace.
            dGSCenterX = (dGSMinX + dGSMaxX) / 2#
            dGSCenterY = (dGSMinY + dGSMaxY) / 2#
            dHalfRectWidth = dRectWidth / 2#
            dHalfRectHeight = dRectHeight / 2#
            Set objRect = objBin.Rectangles.AddByDefiningCoordinates(dGSCenterX - dHalfRectWidth, dGSCenterY - dHalfRectHeight, dGSCenterX + dHalfRectWidth, dGSCenterY + dHalfRectHeight, objRect)
        
        End If
    
    End If

End Sub







Private Sub ShrinkGreenSpaces(ByVal objRect As RAD2D.Group, ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces)

    Dim i As Long, lGSCount As Long, lContainCount As Long, lSideMask As Long
    Dim lGSArray() As Long

    'Step 1: Determine which greenspaces contain this rectangle (wholly or partly).
    'Step 2: Find how many sides (and which ones) of the rectangle in question
    '        are not completely bounded by sides of other rectangles in the bin.
    'Step 3: Do the shrinking algorithm on the greenspaces found in Step 1.
    
    
    '/***************************************************************************************\'
    'Step 1:
    
    'This refers to how many greenspaces contain the rectangle.
    lContainCount = 0
    
    lGSCount = objGSS.Count
    For i = 1 To lGSCount
    
        If RectangleInGreenSpace(objRect, objGSS.Item(i)) Then
            
            lContainCount = lContainCount + 1
            
            'lGSArray contains Long values that refer to indeces in the objGSS collection.
            '   These indeces refer to individual greenspaces that contain
            '   the rectangle in question.
            ReDim Preserve lGSArray(1 To lContainCount)
            
            lGSArray(lContainCount) = i
            
        End If
        
    Next i
    
    '\***************************************************************************************/'
    
    '/***************************************************************************************\'
    'Step 2:
    
    'lSideMask will tell us which of the four sides of the rectangle are at least
    '   partially unbounded by other rectangles.
    '   If lSideMask = 0, then no side is unbounded.
    '   If lSideMask And 1 = 1, then its top side is unbounded.
    '   If lSideMask And 2 = 2, then its bottom side is unbounded.
    '   If lSideMask And 4 = 4, then its left side is unbounded.
    '   If lSideMask And 8 = 8, then its right side is unbounded.
    lSideMask = FindUnboundedRectangleSides(objRect)

    '\***************************************************************************************/'
    
    '/***************************************************************************************\'
    'Step 3:
    
    DoShrink objGSS, lGSArray, lContainCount, objRect, lSideMask
    
    '\***************************************************************************************/'

End Sub

Private Function RectangleInGreenSpace(ByVal objRect As RAD2D.Group, ByVal objGS As StrMfgDrawingAVP.CGreenSpace) As Boolean

    Dim dRectMinX As Double, dRectMinY As Double, dRectMaxX As Double, dRectMaxY As Double
    Dim dGSMinX As Double, dGSMinY As Double, dGSMaxX As Double, dGSMaxY As Double
    
    'objRect.Range dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    
    GetRangeWithMargin objRect, dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    
    objGS.GetDefiningCoordinates dGSMinX, dGSMinY, dGSMaxX, dGSMaxY
    
'    If dRectMinX < 0.00001 Then dRectMinX = 0
'    If dRectMinY < 0.00001 Then dRectMinY = 0
    
    RectangleInGreenSpace = BoxAIntersectsBoxB(dRectMinX, dRectMinY, dRectMaxX, dRectMaxY, dGSMinX, dGSMinY, dGSMaxX, dGSMaxY)

End Function

Private Function BoxAIntersectsBoxB(ByVal dAMinX As Double, ByVal dAMinY As Double, ByVal dAMaxX As Double, ByVal dAMaxY As Double, ByVal dBMinX As Double, ByVal dBMinY As Double, ByVal dBMaxX As Double, ByVal dBMaxY As Double) As Boolean
    
    BoxAIntersectsBoxB = False
    
    Select Case True
    
        Case PointInBox(dAMinX, dAMinY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             PointInBox(dAMinX, dAMaxY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             PointInBox(dAMaxX, dAMinY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             PointInBox(dAMaxX, dAMaxY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             PointInBox(dBMinX, dBMinY, dAMinX, dAMinY, dAMaxX, dAMaxY), _
             PointInBox(dBMinX, dBMaxY, dAMinX, dAMinY, dAMaxX, dAMaxY), _
             PointInBox(dBMaxX, dBMinY, dAMinX, dAMinY, dAMaxX, dAMaxY), _
             PointInBox(dBMaxX, dBMaxY, dAMinX, dAMinY, dAMaxX, dAMaxY), _
             BoxAOverlapsBoxB(dAMinX, dAMinY, dAMaxX, dAMaxY, dBMinX, dBMinY, dBMaxX, dBMaxY), _
             BoxAOverlapsBoxB(dBMinX, dBMinY, dBMaxX, dBMaxY, dAMinX, dAMinY, dAMaxX, dAMaxY)
             
            BoxAIntersectsBoxB = True
            
    End Select
    
End Function

Private Function PointInBox(ByVal dPointX As Double, ByVal dPointY As Double, ByVal dBoxMinX As Double, ByVal dBoxMinY As Double, ByVal dBoxMaxX As Double, ByVal dBoxMaxY As Double) As Boolean
    
    PointInBox = False
    
    If FPGreaterThanOrEqualTo(dPointX, dBoxMinX) Then
    
        If FPLessThanOrEqualTo(dPointX, dBoxMaxX) Then
        
            If FPGreaterThanOrEqualTo(dPointY, dBoxMinY) Then
            
                If FPLessThanOrEqualTo(dPointY, dBoxMaxY) Then
                
                    PointInBox = True
                    
                End If
                
            End If
            
        End If
        
    End If

End Function

Private Function BoxAOverlapsBoxB(ByVal dAMinX As Double, ByVal dAMinY As Double, ByVal dAMaxX As Double, ByVal dAMaxY As Double, ByVal dBMinX As Double, ByVal dBMinY As Double, ByVal dBMaxX As Double, ByVal dBMaxY As Double) As Boolean

    BoxAOverlapsBoxB = False
    
    If FPGreaterThan(dBMinX, dAMinX) Then
    
        If FPGreaterThan(dAMaxX, dBMaxX) Then
        
            If FPGreaterThan(dAMinY, dBMinY) Then
            
                If FPGreaterThan(dBMaxY, dAMaxY) Then
                
                    BoxAOverlapsBoxB = True
                    
                End If
                
            End If
            
        End If
        
    End If

End Function

Private Function FPGreaterThanOrEqualTo(ByVal dA As Double, ByVal dB As Double) As Boolean

    If dA >= dB - TOLERANCE Then
    
        FPGreaterThanOrEqualTo = True
        
    Else
    
        FPGreaterThanOrEqualTo = False
        
    End If

End Function

Private Function FPLessThanOrEqualTo(ByVal dA As Double, ByVal dB As Double) As Boolean

    If dA <= dB + TOLERANCE Then
    
        FPLessThanOrEqualTo = True
        
    Else
    
        FPLessThanOrEqualTo = False
        
    End If

End Function

Private Function FindUnboundedRectangleSides(ByVal objRect As RAD2D.Group) As Long

    Dim dTopClearDist As Double, dBottomClearDist As Double
    Dim dLeftClearDist As Double, dRightClearDist As Double
    Dim dRectWidth As Double, dRectHeight As Double
        
    'The 4 distance variables refer to how far on each side that objRect isn't bounding any
    '   other rectangle in the bin.
    '   Initialize each one to the rectangle's corresponding side's length.
    
  
    GetRange objRect, dRectHeight, dRectWidth
    dTopClearDist = dRectWidth
    dBottomClearDist = dRectWidth
    dLeftClearDist = dRectHeight
    dRightClearDist = dRectHeight
    
    'See if objRect borders side(s) of the bin.
    FindRectangleBinAdjacency objRect, dTopClearDist, dBottomClearDist, dLeftClearDist, dRightClearDist
    
    'See if objRect borders any other rectangle.
    FindRectangleRectangleAdjacency objRect, dTopClearDist, dBottomClearDist, dLeftClearDist, dRightClearDist
    
    'If objRect's top side is unbounded, then FindUnboundedRectangleSides And 1 = 1.
    'If objRect's bottom side is unbounded, then FindUnboundedRectangleSides And 2 = 2.
    'If objRect's left side is unbounded, then FindUnboundedRectangleSides And 4 = 4.
    'If objRect's right side is unbounded, then FindUnboundedRectangleSides And 8 = 8.
    
    'Initialize FindUnboundedRectangleSides to minimum value of 0.
    FindUnboundedRectangleSides = 0
    
    If FPGreaterThan(dTopClearDist, 0#) Then
    
        FindUnboundedRectangleSides = FindUnboundedRectangleSides Or 1
        
    End If
    
    If FPGreaterThan(dBottomClearDist, 0#) Then
    
        FindUnboundedRectangleSides = FindUnboundedRectangleSides Or 2
        
    End If
    
    If FPGreaterThan(dLeftClearDist, 0#) Then
    
        FindUnboundedRectangleSides = FindUnboundedRectangleSides Or 4
        
    End If
    
    If FPGreaterThan(dRightClearDist, 0#) Then
    
        FindUnboundedRectangleSides = FindUnboundedRectangleSides Or 8
        
    End If

End Function

Private Sub FindRectangleBinAdjacency(ByVal objRect As RAD2D.Group, ByRef dTopClearDist As Double, ByRef dBottomClearDist As Double, ByRef dLeftClearDist As Double, ByRef dRightClearDist As Double)

    Dim lWhichSides As Long

    'If lWhichSides = 0, then no side of objRect borders any edge of the bin.
    'If lWhichSides And 1 = 1, then objRect's top side borders the top edge of the bin.
    'If lWhichSides And 2 = 2, then objRect's bottom side borders the bottom edge of the bin.
    'If lWhichSides And 4 = 4, then objRect's left side borders the left edge of the bin.
    'If lWhichSides And 8 = 8, then objRect's right side borders the right edge of the bin.
    If RectangleBordersBin(objRect, lWhichSides) Then
    
        If (lWhichSides And 1) = 1 Then
        
            dTopClearDist = 0#
        
        End If
        
        If (lWhichSides And 2) = 2 Then
        
            dBottomClearDist = 0#
        
        End If
        
        If (lWhichSides And 4) = 4 Then
        
            dLeftClearDist = 0#
        
        End If
        
        If (lWhichSides And 8) = 8 Then
        
            dRightClearDist = 0#
        
        End If
    
    End If

End Sub

Private Function RectangleBordersBin(ByVal objRect As RAD2D.Group, ByRef lWhichSides As Long) As Boolean

    Dim dRectMinX As Double, dRectMinY As Double, dRectMaxX As Double, dRectMaxY As Double
    
    'objRect.Range dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    
    GetRangeWithMargin objRect, dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    
    'If BoxABordersWhichSidesOfBoxB = 0, then no side of Box A borders any side of Box B.
    'If BoxABordersWhichSidesOfBoxB And 1 = 1, then bottom side of Box A borders top side of Box B.
    'If BoxABordersWhichSidesOfBoxB And 2 = 2, then top side of Box A borders bottom side of Box B.
    'If BoxABordersWhichSidesOfBoxB And 4 = 4, then right side of Box A borders left side of Box B.
    'If BoxABordersWhichSidesOfBoxB And 8 = 8, then left side of Box A borders right side of Box B.
    lWhichSides = BoxABordersWhichInternalSidesOfBoxB(dRectMinX, dRectMinY, dRectMaxX, dRectMaxY, m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY)
    
    If lWhichSides > 0 Then
    
        RectangleBordersBin = True
        
    Else
    
        RectangleBordersBin = False
    
    End If
    
End Function

Private Function BoxABordersWhichInternalSidesOfBoxB(ByVal dAMinX As Double, ByVal dAMinY As Double, ByVal dAMaxX As Double, ByVal dAMaxY As Double, ByVal dBMinX As Double, ByVal dBMinY As Double, ByVal dBMaxX As Double, ByVal dBMaxY As Double) As Long

    BoxABordersWhichInternalSidesOfBoxB = 0
    
    'Check to see if the top side of Box A borders the top side of Box B.
    If FPEqualTo(dAMaxY, dBMaxY) Then
        
        BoxABordersWhichInternalSidesOfBoxB = BoxABordersWhichInternalSidesOfBoxB Or 1
        
    End If
        
    'Check to see if the bottom side of Box A borders the bottom side of Box B.
    If FPEqualTo(dAMinY, dBMinY) Then
    
        BoxABordersWhichInternalSidesOfBoxB = BoxABordersWhichInternalSidesOfBoxB Or 2
        
    End If
    
    'Check to see if the left side of Box A borders the left side of Box B.
    If FPEqualTo(dAMinX, dBMinX) Then

        BoxABordersWhichInternalSidesOfBoxB = BoxABordersWhichInternalSidesOfBoxB Or 4
        
    End If
    
    'Check to see if the right side of Box A borders the right side of Box B.
    If FPEqualTo(dAMaxX, dBMaxX) Then
                                            
        BoxABordersWhichInternalSidesOfBoxB = BoxABordersWhichInternalSidesOfBoxB Or 8
        
    End If

End Function

Private Function FPEqualTo(ByVal dA As Double, ByVal dB As Double) As Boolean

    If Abs(dA - dB) <= TOLERANCE Then
    
        FPEqualTo = True
        
    Else
    
        FPEqualTo = False
        
    End If

End Function

Private Sub FindRectangleRectangleAdjacency(ByVal objRect As RAD2D.Group, ByRef dTopClearDist As Double, ByRef dBottomClearDist As Double, ByRef dLeftClearDist As Double, ByRef dRightClearDist As Double)
    
    Dim objRects As RAD2D.Groups, objCompareRect As RAD2D.Group
    Dim lWhichSide As Long
    Dim dHorizDistInCommon As Double, dVertDistInCommon As Double
    Dim objRects1 As Object
    
    Set objRects1 = objRect.Parent
    
    For Each objCompareRect In m_objRects1
    
        If Not objRect Is objCompareRect Then
            
            'objCompareRect can only border objRect on one side.
            'dDistInCommon is the distance objCompareRect and objRect share in common
            '   on that one side.
            If RectangleABordersRectangleB(objCompareRect, objRect, lWhichSide, dHorizDistInCommon, dVertDistInCommon) Then
            
                'If lWhichSide = 0, then no side of Box A borders any side of Box B.
                'If lWhichSide = 1, then bottom side of Box A borders top side of Box B.
                'If lWhichSide = 2, then top side of Box A borders bottom side of Box B.
                'If lWhichSide = 3, then right side of Box A borders left side of Box B.
                'If lWhichSide = 4, then left side of Box A borders right side of Box B.
                
                Select Case lWhichSide
                
                    Case 1
                    
                        dTopClearDist = dTopClearDist - dHorizDistInCommon
                    
                    Case 2
                    
                        dBottomClearDist = dBottomClearDist - dHorizDistInCommon
                    
                    Case 3
                    
                        dLeftClearDist = dLeftClearDist - dVertDistInCommon
                    
                    Case 4
                    
                        dRightClearDist = dRightClearDist - dVertDistInCommon
                    
                End Select
               
            End If
            
        End If
    
    Next objCompareRect
    
End Sub

Private Function RectangleABordersRectangleB(ByVal objRectA As RAD2D.Group, ByVal objRectB As RAD2D.Group, ByRef lWhichSide As Long, ByRef dHorizDistInCommon As Double, ByRef dVertDistInCommon As Double) As Boolean

    Dim dAMinX As Double, dAMinY As Double, dAMaxX As Double, dAMaxY As Double
    Dim dBMinX As Double, dBMinY As Double, dBMaxX As Double, dBMaxY As Double
    
    'objRectA.Range dAMinX, dAMinY, dAMaxX, dAMaxY
    'objRectB.Range dBMinX, dBMinY, dBMaxX, dBMaxY
    
    GetRangeWithMargin objRectA, dAMinX, dAMinY, dAMaxX, dAMaxY
    GetRangeWithMargin objRectB, dBMinX, dBMinY, dBMaxX, dBMaxY
    

    'If BoxABordersWhichSideOfBoxB = 0, then no side of Box A borders any side of Box B.
    'If BoxABordersWhichSideOfBoxB = 1, then bottom side of Box A borders top side of Box B.
    'If BoxABordersWhichSideOfBoxB = 2, then top side of Box A borders bottom side of Box B.
    'If BoxABordersWhichSideOfBoxB = 3, then right side of Box A borders left side of Box B.
    'If BoxABordersWhichSideOfBoxB = 4, then left side of Box A borders right side of Box B.
    lWhichSide = BoxABordersWhichSideOfBoxB(dAMinX, dAMinY, dAMaxX, dAMaxY, dBMinX, dBMinY, dBMaxX, dBMaxY, dHorizDistInCommon, dVertDistInCommon)
    
    If lWhichSide = 0 Then
    
        'We will get in here only if lWhichSide = 0.
        
        RectangleABordersRectangleB = False
                
    Else
    
        RectangleABordersRectangleB = True
    
    End If
    
End Function

Private Function BoxABordersWhichSideOfBoxB(ByVal dAMinX As Double, ByVal dAMinY As Double, ByVal dAMaxX As Double, ByVal dAMaxY As Double, ByVal dBMinX As Double, ByVal dBMinY As Double, ByVal dBMaxX As Double, ByVal dBMaxY As Double, ByRef dHorizDistInCommon As Double, ByRef dVertDistInCommon As Double) As Long

    BoxABordersWhichSideOfBoxB = 0
    
    'Check to see if Box A's width range coordinates are within Box B's width range coordinates.
    If BoxASideBordersBoxBSide(dAMinX, dAMaxX, dBMinX, dBMaxX, dHorizDistInCommon) Then
    
        'Check to see if the bottom side of Box A borders the top side of Box B.
        If FPEqualTo(dAMinY, dBMaxY) Then
            
            BoxABordersWhichSideOfBoxB = 1
            
        'Check to see if the top side of Box A borders the bottom side of Box B.
        ElseIf FPEqualTo(dAMaxY, dBMinY) Then
        
            BoxABordersWhichSideOfBoxB = 2
            
        End If
        
    End If
        
    'Check to see if Box A's height range coordinates are within Box B's height range coordinates.
    If BoxASideBordersBoxBSide(dAMinY, dAMaxY, dBMinY, dBMaxY, dVertDistInCommon) Then
    
        'Check to see if the right side of Box A borders the left side of Box B.
        If FPEqualTo(dAMaxX, dBMinX) Then
    
            BoxABordersWhichSideOfBoxB = 3
        
        'Check to see if the left side of Box A borders the right side of Box B.
        ElseIf FPEqualTo(dAMinX, dBMaxX) Then
                                                
            BoxABordersWhichSideOfBoxB = 4
            
        End If
        
    End If

End Function

Private Function BoxASideBordersBoxBSide(ByVal dAMin As Double, ByVal dAMax As Double, ByVal dBMin As Double, ByVal dBMax As Double, ByRef dDistInCommon As Double) As Boolean

    dDistInCommon = 0#
    BoxASideBordersBoxBSide = True

    If FPBetween(dAMin, dBMin, dBMax) Then
        
        'Last argument being True signifies that the first argument (dAMin) is a MINIMUM range value.
        dDistInCommon = DistanceInCommon(dAMin, dAMax, dBMin, dBMax, True)
        
    ElseIf FPBetween(dAMax, dBMin, dBMax) Then
    
        'Last argument being False signifies that the first argument (dAMax) is a MAXIMUM range value.
        dDistInCommon = DistanceInCommon(dAMin, dAMax, dBMin, dBMax, False)
    
    ElseIf FPBetween(dBMin, dAMin, dAMax) Then
    
        'We get in here only if Box A borders Box B and extends out past Box B on both sides.
        '   In this case, the distance in common is the length of Box B's side.
        dDistInCommon = dBMax - dBMin
    
    Else
    
        BoxASideBordersBoxBSide = False
    
    End If
            
End Function

Private Function FPBetween(ByVal dA As Double, dB1 As Double, dB2 As Double) As Boolean

    FPBetween = False

    If FPGreaterThanOrEqualTo(dA, dB1) Then
    
        If FPLessThanOrEqualTo(dA, dB2) Then
    
            FPBetween = True
            
        End If
        
    End If
        
End Function

Private Function DistanceInCommon(ByVal dAMin As Double, ByVal dAMax As Double, ByVal dBMin As Double, ByVal dBMax As Double, ByVal bAIsMin As Boolean) As Double
    
    If bAIsMin Then
    
        'dA is a MINIMUM range value.
       
        If FPLessThanOrEqualTo(dBMax, dAMax) Then
        
            DistanceInCommon = dBMax - dAMin
       
        Else
       
            DistanceInCommon = dAMax - dAMin
            
        End If
    
    Else
    
       'dA is a MAXIMUM range value.
       
       DistanceInCommon = dAMax - dBMin
    
    End If
    
End Function

Private Function FPGreaterThan(ByVal dA As Double, ByVal dB As Double) As Boolean

    If dA > dB + TOLERANCE Then

        FPGreaterThan = True

    Else

        FPGreaterThan = False

    End If

End Function

Private Sub DoShrink(ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces, ByRef lGSArray() As Long, ByVal lUBound As Long, ByVal objRect As RAD2D.Group, ByVal lSideMask As Long)

    Dim i As Long
    Dim dGSMinXArray() As Double, dGSMinYArray() As Double
    Dim dGSMaxXArray() As Double, dGSMaxYArray() As Double
    Dim dRectMinX As Double, dRectMinY As Double, dRectMaxX As Double, dRectMaxY As Double
    
    ReDim dGSMinXArray(1 To lUBound)
    ReDim dGSMinYArray(1 To lUBound)
    ReDim dGSMaxXArray(1 To lUBound)
    ReDim dGSMaxYArray(1 To lUBound)

    'We have which greenspaces objRect intersects.  We also have which sides of objRect
    '   are at least partially clear.
    'At the end, we should have deleted the originally intersected greenspaces.  We really
    '   don't shrink greenspaces per se...we just delete the originals and add new smaller ones.
    '   When we're determining which smaller greenspaces to add, we only add the ones that
    '   envelop others.  In other words, we want the largest area greenspaces we can get
    '   by the time we're through.

    If lSideMask = 0 Then
    
        'The rectangle fits perfectly into one greenspace.  We just need to delete that greenspace.
        objGSS.Item(1).Delete
        Exit Sub
    
    End If
        
    'Populate the intersected greenspace coordinate arrays.
    For i = 1 To lUBound
    
        With objGSS.Item(lGSArray(i))
    
            .GetDefiningCoordinates dGSMinXArray(i), dGSMinYArray(i), dGSMaxXArray(i), dGSMaxYArray(i)
            
        End With
    
    Next i
    
    'Get objRect's coordinates.
    'objRect.Range dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    GetRangeWithMargin objRect, dRectMinX, dRectMinY, dRectMaxX, dRectMaxY
    
    If (lSideMask And 1) = 1 Then
    
        'Top side of objRect is at least partially clear.  Shrink all intersected
        '   greenspaces toward the top of objRect and determine which has the largest area.
        '   Add this new greenspace.
        AddNewGreenSpace objGSS, dGSMinXArray, dGSMinYArray, dGSMaxXArray, dGSMaxYArray, lUBound, dRectMaxY, 1
                    
    End If
    
    If (lSideMask And 2) = 2 Then
    
        'Bottom side of objRect is at least partially clear.  Shrink all intersected
        '   greenspaces toward the bottom of objRect and determine which has the largest area.
        '   Add this new greenspace.
        AddNewGreenSpace objGSS, dGSMinXArray, dGSMinYArray, dGSMaxXArray, dGSMaxYArray, lUBound, dRectMinY, 2
                    
    End If
    
    If (lSideMask And 4) = 4 Then
    
        'Left side of objRect is at least partially clear.  Shrink all intersected
        '   greenspaces toward the left of objRect and determine which has the largest area.
        '   Add this new greenspace.
        AddNewGreenSpace objGSS, dGSMinXArray, dGSMinYArray, dGSMaxXArray, dGSMaxYArray, lUBound, dRectMinX, 3
                    
    End If
    
    If (lSideMask And 8) = 8 Then
    
        'Right side of objRect is at least partially clear.  Shrink all intersected
        '   greenspaces toward the right of objRect and determine which has the largest area.
        '   Add this new greenspace.
        AddNewGreenSpace objGSS, dGSMinXArray, dGSMinYArray, dGSMaxXArray, dGSMaxYArray, lUBound, dRectMaxX, 4
                    
    End If
    
    'Delete the old greenspaces.
    DeleteGreenSpaces objGSS, lGSArray, lUBound
    
End Sub

Private Sub AddNewGreenSpace(ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces, ByRef dGSMinXArray() As Double, ByRef dGSMinYArray() As Double, ByRef dGSMaxXArray() As Double, ByRef dGSMaxYArray() As Double, ByVal lUBound As Long, ByVal dRectSideCoord As Double, ByVal lWhichSide As Long)

    Dim i As Long
    Dim dNewGSMinXArray() As Double, dNewGSMinYArray() As Double
    Dim dNewGSMaxXArray() As Double, dNewGSMaxYArray() As Double
    Dim lDiscardArray() As Long
    Dim bFound As Boolean
    
    ReDim dNewGSMinXArray(1 To lUBound)
    ReDim dNewGSMinYArray(1 To lUBound)
    ReDim dNewGSMaxXArray(1 To lUBound)
    ReDim dNewGSMaxYArray(1 To lUBound)
    ReDim lDiscardArray(1 To 1)
    
    lDiscardArray(1) = 0
    
    Select Case lWhichSide
    
        Case 1  'Shrink toward top of rectangle.
        
            For i = 1 To lUBound
            
                dNewGSMinXArray(i) = dGSMinXArray(i)
                dNewGSMinYArray(i) = dRectSideCoord
                dNewGSMaxXArray(i) = dGSMaxXArray(i)
                dNewGSMaxYArray(i) = dGSMaxYArray(i)
            
            Next i
            
        Case 2  'Shrink toward bottom of rectangle.
        
            For i = 1 To lUBound
            
                dNewGSMinXArray(i) = dGSMinXArray(i)
                dNewGSMinYArray(i) = dGSMinYArray(i)
                dNewGSMaxXArray(i) = dGSMaxXArray(i)
                dNewGSMaxYArray(i) = dRectSideCoord
            
            Next i
        
        Case 3  'Shrink toward left of rectangle.
        
            For i = 1 To lUBound
            
                dNewGSMinXArray(i) = dGSMinXArray(i)
                dNewGSMinYArray(i) = dGSMinYArray(i)
                dNewGSMaxXArray(i) = dRectSideCoord
                dNewGSMaxYArray(i) = dGSMaxYArray(i)
            
            Next i
        
        Case 4  'Shrink toward right of rectangle.
        
            For i = 1 To lUBound
            
                dNewGSMinXArray(i) = dRectSideCoord
                dNewGSMinYArray(i) = dGSMinYArray(i)
                dNewGSMaxXArray(i) = dGSMaxXArray(i)
                dNewGSMaxYArray(i) = dGSMaxYArray(i)
            
            Next i
        
    End Select
    
    GetDiscardArray lUBound, dNewGSMinXArray, dNewGSMinYArray, dNewGSMaxXArray, dNewGSMaxYArray, lDiscardArray, lWhichSide
            
    AddNonEnvelopedGreenSpaces objGSS, lUBound, dNewGSMinXArray, dNewGSMinYArray, dNewGSMaxXArray, dNewGSMaxYArray, lDiscardArray
    
End Sub

Private Sub GetDiscardArray(ByVal lUBound As Long, ByRef dNewGSMinXArray() As Double, ByRef dNewGSMinYArray() As Double, ByRef dNewGSMaxXArray() As Double, ByRef dNewGSMaxYArray() As Double, ByRef lDiscardArray() As Long, ByVal lWhichSide As Long)

    Dim i As Long, j As Long

    For i = 1 To lUBound
            
        If ((lWhichSide = 1 Or lWhichSide = 2) And FPGreaterThan(dNewGSMaxYArray(i), dNewGSMinYArray(i))) Or _
           ((lWhichSide = 3 Or lWhichSide = 4) And FPGreaterThan(dNewGSMaxXArray(i), dNewGSMinXArray(i))) Then
        
            For j = 1 To lUBound
            
                If i <> j Then
                
                    If FPLessThanOrEqualTo(dNewGSMinXArray(i), dNewGSMinXArray(j)) Then
                    
                        If FPLessThanOrEqualTo(dNewGSMinYArray(i), dNewGSMinYArray(j)) Then
                        
                            If FPLessThanOrEqualTo(dNewGSMaxXArray(j), dNewGSMaxXArray(i)) Then
                            
                                If FPLessThanOrEqualTo(dNewGSMaxYArray(j), dNewGSMaxYArray(i)) Then
                                
                                    AddToDiscardArray lDiscardArray, j
                                
                                End If
                                
                            End If
                            
                        End If
                        
                    End If
                
                End If
            
            Next j
        
        Else
        
            AddToDiscardArray lDiscardArray, i
        
        End If
    
    Next i

End Sub

Private Sub AddToDiscardArray(ByRef lDiscardArray() As Long, ByVal lItemToDiscard As Long)

    Dim lDiscardArrayUBound As Long

    lDiscardArrayUBound = UBound(lDiscardArray)
    lDiscardArray(lDiscardArrayUBound) = lItemToDiscard
    lDiscardArrayUBound = lDiscardArrayUBound + 1
    ReDim Preserve lDiscardArray(1 To lDiscardArrayUBound)
    lDiscardArray(lDiscardArrayUBound) = 0

End Sub

Private Sub AddNonEnvelopedGreenSpaces(ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces, ByVal lUBound As Long, ByRef dNewGSMinXArray() As Double, ByRef dNewGSMinYArray() As Double, ByRef dNewGSMaxXArray() As Double, ByRef dNewGSMaxYArray() As Double, ByRef lDiscardArray() As Long)

    Dim i As Long, j As Long, lDiscardArrayUBound As Long
    Dim bFound As Boolean

    For i = 1 To lUBound
            
        bFound = False
        lDiscardArrayUBound = UBound(lDiscardArray)
        For j = 1 To lDiscardArrayUBound
        
            If lDiscardArray(j) = i Then
            
                bFound = True
                Exit For
            
            End If
        
        Next j
        
        If Not bFound Then
        
            objGSS.AddByDefiningCoordinates dNewGSMinXArray(i), dNewGSMinYArray(i), dNewGSMaxXArray(i), dNewGSMaxYArray(i)
        
        End If
    
    Next i

End Sub

Private Sub DeleteGreenSpaces(ByVal objGSS As StrMfgDrawingAVP.CGreenSpaces, ByRef lGSArray() As Long, ByVal lUBound As Long)

    Dim i As Long
    
    For i = lUBound To 1 Step -1
    
        With objGSS.Item(lGSArray(i))
        
            .Delete
            
        End With
    
    Next i

End Sub

Private Sub Commit()

    Dim lBinCount As Long, i As Long
    Dim dMinX As Double, dMinY As Double, dMaxX As Double, dMaxY As Double
    Dim dPtArray(0 To 7) As Double
    Dim dLostArea As Double, dBinArea As Double
    Dim objRADSheet As RAD2D.Sheet
    Dim objRADRect As RAD2D.Rectangle2d
    Dim objRADRects As RAD2D.Rectangles2d
    Dim objAVPBin As StrMfgDrawingAVP.CBin
    Dim objAVPRects As StrMfgDrawingAVP.CRectangles
    Dim objAVPRect As StrMfgDrawingAVP.CRectangle
    Dim objAVPGSS As StrMfgDrawingAVP.CGreenSpaces
    Dim objAVPGS As StrMfgDrawingAVP.CGreenSpace
    Dim objRADBoundaries As RAD2D.Boundaries2d
    Dim objRADBoundary As RAD2D.Boundary2d
    Dim objRADTextBox As RAD2D.TextBox
    Const LINEWIDTH As Double = 0.0005
    
    'm_objDoc.CalculateOff
    'm_objRadApp.ScreenUpdating = False

    'Add one sheet for every bin we need.
    i = 1
    For Each objAVPBin In m_objBins
    
        Set objRADSheet = m_objSheets.Add(CStr(i))
        Set objRADRects = objRADSheet.Rectangles2d
        
        Set objRADRect = objRADRects.AddBy2Points(m_dBinMinX, m_dBinMinY, m_dBinMaxX, m_dBinMaxY)
        
        objRADRect.LinearStyle.Color = igBlueColor
        objRADRect.LinearStyle.Width = LINEWIDTH
        
        i = i + 1
    
    Next objAVPBin
    
    lBinCount = m_objBins.Count
    For i = 1 To lBinCount
    
        Set objAVPBin = m_objBins.Item(i)
        Set objAVPRects = objAVPBin.Rectangles
        Set objRADBoundaries = m_objSheets(i + 1).Boundaries2d
        
        dBinArea = objAVPBin.Width * objAVPBin.Height
        dLostArea = dBinArea
        
        For Each objAVPRect In objAVPRects
        
            objAVPRect.GetDefiningCoordinates dMinX, dMinY, dMaxX, dMaxY
            Set objRADRect = m_objSheets.Item(i + 1).Rectangles2d.AddBy2Points(dMinX, dMinY, dMaxX, dMaxY)
            
            dPtArray(0) = dMinX
            dPtArray(1) = dMinY
            dPtArray(2) = dMaxX
            dPtArray(3) = dMinY
            dPtArray(4) = dMaxX
            dPtArray(5) = dMaxY
            dPtArray(6) = dMinX
            dPtArray(7) = dMaxY
            
            Set objRADBoundary = objRADBoundaries.AddByPoints(4, dPtArray)
            
            objRADBoundary.Style.FillName = "Solid"
            objRADBoundary.Style.FillColor = igRedColor
            
            dLostArea = dLostArea - objAVPRect.Width * objAVPRect.Height
        
        Next objAVPRect
        
        For Each objAVPRect In objAVPRects
            
            objAVPRect.GetDefiningCoordinates dMinX, dMinY, dMaxX, dMaxY
            Set objRADRect = m_objSheets.Item(i + 1).Rectangles2d.AddBy2Points(dMinX, dMinY, dMaxX, dMaxY)
            objRADRect.LinearStyle.Color = igBackgroundColor
            objRADRect.LinearStyle.Width = LINEWIDTH
        
        Next objAVPRect
        
        If Check1.Value = vbChecked Then
        
            Set objAVPGSS = objAVPBin.GreenSpaces
            
            For Each objAVPGS In objAVPGSS
            
                objAVPGS.GetDefiningCoordinates dMinX, dMinY, dMaxX, dMaxY
                Set objRADRect = m_objSheets.Item(i + 1).Rectangles2d.AddBy2Points(dMinX, dMinY, dMaxX, dMaxY)
                
                dPtArray(0) = dMinX
                dPtArray(1) = dMinY
                dPtArray(2) = dMaxX
                dPtArray(3) = dMinY
                dPtArray(4) = dMaxX
                dPtArray(5) = dMaxY
                dPtArray(6) = dMinX
                dPtArray(7) = dMaxY
                
                Set objRADBoundary = objRADBoundaries.AddByPoints(4, dPtArray)
                
                objRADBoundary.Style.FillName = "Solid"
                objRADBoundary.Style.FillColor = igGreenColor
            
            Next objAVPGS
            
            For Each objAVPGS In objAVPGSS
            
                objAVPGS.GetDefiningCoordinates dMinX, dMinY, dMaxX, dMaxY
                Set objRADRect = m_objSheets.Item(i + 1).Rectangles2d.AddBy2Points(dMinX, dMinY, dMaxX, dMaxY)
                objRADRect.LinearStyle.Color = igBackgroundColor
                objRADRect.LinearStyle.Width = LINEWIDTH
            
            Next objAVPGS
        
        End If
        
        Set objRADTextBox = m_objSheets(i + 1).TextBoxes.Add(m_dBinMaxX + 0.02, m_dBinMaxY - 0.02)
        objRADTextBox.Visible = False
        objRADTextBox.Text = "Lost Area:  " & CStr(Round(dLostArea * 100 / dBinArea, 2)) & "%"
        With objRADTextBox.Edit
        
            .SetSelect 0, 0, igTextSelectAll
            .TextSize = 0.01
            
        End With
        objRADTextBox.Visible = True
        
        Set objRADTextBox = m_objSheets(i + 1).TextBoxes.Add(m_dBinMaxX + 0.02, m_dBinMaxY - 0.05)
        objRADTextBox.Visible = False
        objRADTextBox.Text = "Number of Rotations:  " & CStr(m_lRotations(i))
        With objRADTextBox.Edit
        
            .SetSelect 0, 0, igTextSelectAll
            .TextSize = 0.01
            
        End With
        objRADTextBox.Visible = True
    
    Next i
    
    'm_objRadApp.ScreenUpdating = True
    'm_objDoc.CalculateOn

End Sub

Private Sub Form_Terminate()

    'Destroy rectangles and bins member collection objects
    
    Set m_objRects = Nothing
    Set m_objBins = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Destroy the RAD member objects
    TerminateRAD
    
End Sub

Private Sub TerminateRAD()

    Set m_objOrigSheet = Nothing
    Set m_objSheets = Nothing
    Set m_objRadApp = Nothing
    Set m_objApp = Nothing

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbEnter Then
    
        Start
    
    End If

End Sub

Private Function SortGroupsInSheet(oSheet As Sheet) As Collection

    Dim objGrp As Group
    Dim objGroups As Groups
    Dim dXMin As Double, dXMax As Double, dYMin As Double, dYMax As Double
    Dim dRange As Double, dMax As Double
    
    Dim oColl As Collection
    Dim i As Integer
        
    oSheet.Activate
    Set objGroups = oSheet.Groups
    
    Set oColl = New Collection
    
    i = 1
    
    '*** Making a collection of groups which exists on the sheet ***'
    For Each objGrp In objGroups
        
        'Assigning Key to each group in the collection
        oColl.Add objGrp, "GRP_" & i
        i = i + 1
        
    Next
    
    Set SortGroupsInSheet = MergeSort(oColl, 1, objGroups.Count)
    


End Function

Private Function MergeSort(ByRef objCollection As Collection, _
                      ByVal lStart As Long, _
                      ByVal lFinish As Long) As Collection
                      
                      
    Dim lMiddle     As Long
    Dim objStart    As Group
    Dim objFinish   As Group
    Dim dStartRange As Double
    Dim dFinishRange As Double
    Dim objTempCol As Collection
    Dim dHeight As Double, dWidth As Double
    
    
    If lFinish - lStart <= 1 Then
        
        'This means there are only 2 elements in collection to sort
        If (lFinish - lStart) > 0 Then
    
            With objCollection
            
                Set objStart = .Item(lStart)
                Set objFinish = .Item(lFinish)
                
                GetRange objStart, dHeight, dStartRange
                GetRange objFinish, dHeight, dFinishRange
                
                If FPGreaterThanOrEqualTo(dFinishRange, dStartRange) Then
                    
                    .Remove lFinish
                    .Add objFinish, , lStart
                    
                    '2nd element is bigger than 1st one. Since for Check drawing, we need descending sort,
                    '  keep the order as it is.
                    Set objTempCol = New Collection
                    objTempCol.Add objFinish
                    objTempCol.Add objStart
                Else
                    '1st element is bigger than 2nd one. Since for Check drawing, we need descending sort,
                    '  swap the elements.
                    Set objTempCol = New Collection
                    objTempCol.Add objStart
                    objTempCol.Add objFinish
                    
                End If ' modUtils.FPGreaterThanOrEqualTo(objFinish.Area, objStart.Area)
                
            End With ' objCollection
        
        Else  ' There is only one element to sort
            Set objStart = objCollection.Item(lStart)
            Set objTempCol = New Collection
            objTempCol.Add objStart
        End If ' lFinish - lStart <> 0
        
    Else
    
        lMiddle = (lStart + lFinish) \ 2
        
        MergeSort objCollection, lStart, lMiddle
        MergeSort objCollection, lMiddle + 1, lFinish
        Set objTempCol = Merge(objCollection, lStart, lMiddle, lFinish)
        
    End If ' lFinish - lStart <= 1
    
    
    Set MergeSort = objTempCol

End Function

Private Sub GetRange(objGroup As Group, dHeight As Double, dWidth As Double)

    Dim dXMin As Double, dXMax As Double, dYMin As Double, dYMax As Double
    Dim dRange As Double, dMax As Double
    
    objGroup.Range dXMin, dYMin, dXMax, dYMax
    
    dWidth = dXMax - dXMin
    dHeight = dYMax - dYMin
    
    'GetRange = dRange

End Sub

Private Sub GetRangeWithMargin(objGroup As Group, dXMin As Double, dYMin As Double, dXMax As Double, dYMax As Double)
    
    'Get X and Y co-ordinates of the Group
    objGroup.Range dXMin, dYMin, dXMax, dYMax
    
    'Add margin and return the updated co-ordinates
    dXMin = dXMin + X_MARGIN
    dYMin = dYMin + X_MARGIN
    dXMax = dXMax + Y_MARGIN
    dYMax = dYMax + Y_MARGIN
    
End Sub


Private Function Merge(ByRef objCollection As Collection, ByVal lStart As Long, _
                  ByVal lMiddle As Long, ByVal lFinish As Long) As Collection

    Dim i           As Long
    Dim lTemp       As Long
    Dim lStart1     As Long
    Dim lStart2     As Long
    Dim lFinish1    As Long
    Dim lFinish2    As Long
    Dim objStart1   As Group
    Dim objStart2   As Group
    Dim objTempCol  As Collection
    Dim dStartRange As Double
    Dim dFinishRange As Double
    Dim dHeight As Double, dWidth As Double
    
    Set objTempCol = New Collection
    
    ' Search from Start to Middle
    lStart1 = lStart
    lFinish1 = lMiddle
    
    ' Search from Middle + 1 to Finish
    lStart2 = lMiddle + 1
    lFinish2 = lFinish
    
    With objCollection
    
        Do While lStart1 <= lFinish1 And lStart2 <= lFinish2
        
            Set objStart1 = .Item(lStart1)
            Set objStart2 = .Item(lStart2)
            
            GetRange objStart1, dHeight, dStartRange
            GetRange objStart2, dHeight, dFinishRange
            
            If FPGreaterThanOrEqualTo(dStartRange, dFinishRange) Then
            
                objTempCol.Add objStart1
                lStart1 = lStart1 + 1
                
            Else
            
                objTempCol.Add objStart2
                lStart2 = lStart2 + 1
                
            End If ' FPGreaterThanOrEqualTo(objStart1.Area, objStart2.Area) ... Else
            
        Loop ' While lStart1 <= lFinish1 And lStart2 <= lFinish2
    
        If lStart1 <= lFinish1 Then
    
            For i = lStart1 To lFinish1
            
                objTempCol.Add .Item(i)
                
            Next i
            
        ElseIf lStart <= lFinish2 Then
        
            For i = lStart2 To lFinish2
            
                objTempCol.Add .Item(i)
                
            Next i
            
        End If ' lStart1 <= lFinish1 ... ElseIf lStart <= lFinish2
        
        For i = lStart To lFinish
        
            .Remove i
            
            If i <= .Count Then
                
                .Add objTempCol.Item(objTempCol.Count - (lFinish - i)), , i
                
            Else
            
                .Add objTempCol.Item(objTempCol.Count - (lFinish - i))
            
            End If ' i <= .Count
            
        Next i
                
    End With ' objCollection
    
    
    Set Merge = objTempCol
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetActiveConnection
' Purpose   : Gets the Active Connection (what else?!)
'---------------------------------------------------------------------------------------
'
Private Function GetPOM() As IJDPOM
    Const METHOD As String = "GetPOM"
    On Error GoTo ErrorHandler

    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    Dim oActiveConnection As IJDAccessMiddle
    oCmnAppGenericUtil.GetActiveConnection oActiveConnection
    
    Dim jContext As IJContext
    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'Get IJDBTypeConfiguration from the Context.
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    
    'Get the Model DataBase ID given the database type
    Dim strConnectionName As String
    strConnectionName = oDBTypeConfiguration.get_DataBaseFromDBType("Model")
    
    Set jContext = Nothing
    Set oDBTypeConfiguration = Nothing

    
    Set GetPOM = oActiveConnection.GetResourceManager(strConnectionName)
    
    Set oActiveConnection = Nothing
    Set oCmnAppGenericUtil = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function




