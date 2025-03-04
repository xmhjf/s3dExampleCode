VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssemblyPanelDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Implements IJPMSortGroups

Dim m_objRects As StrMfgDrawingAVP.CRectangles

Dim m_objRects1 As RAD2D.Groups
Dim m_objRect1 As RAD2D.Group

Dim m_objOrigSheet As RAD2D.Sheet

Private Sub IJPMSortGroups_SortGroups(oSheet As Object)
    Dim i As Long
    Dim objRect1 As RAD2D.Group
    Dim ii As Integer
    
    Set m_objOrigSheet = oSheet
    
    'Get rid of rectangles if they exist.
    Set m_objRects = Nothing
    
    'Initialize m_objRects to new objects.
    Set m_objRects = New CRectangles
    Set m_objRects1 = m_objOrigSheet.Groups
    
    '************** Panel Drawing **********************'
    '*** LOG THE DEBUG DATA ***'
    'Dim FileNumber
    'Dim slogFileName As String
    'slogFileName = "C:\Temp\PanelDrawing.txt"
    'FileNumber = FreeFile
    'Open slogFileName For Append As #FileNumber
    'Write #FileNumber, "PART 1:"
    '**************************'
    
    'Set ScalingGroup attribute to "PanelLayer" in order to put this group into
    '   .. the rectangle with layer = "PanelLayer"
    Dim oPanelDrawingGroup As Group
    Set oPanelDrawingGroup = m_objRects1.Item(1)
    
    Dim oAttributeSet As AttributeSet
    Set oAttributeSet = oPanelDrawingGroup.AttributeSets.Add("Panel")
    
    Dim oAttribute As RAD2D.Attribute
    Set oAttribute = oAttributeSet.Add("ScalingGroup", igAttrTypeString)
    
    oAttribute.Value = "Panel"
    
    m_objOrigSheet.Document.SelectSet.RemoveAll
    
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
        Dim objPart1 As Group
        If m_objRects1.Count > 1 Then
            Set objPart1 = m_objRects1.Item(2)
        
            'Add to selectset
            objPart1.Select
            
            'Get range of first part
            Dim dMinX As Double, dMinY As Double, dMaxX As Double, dMaxY As Double
            Dim dNewMinX As Double, dNewMinY As Double, dNewMaxX As Double, dNewMaxY As Double
            
            objPart1.Range dMinX, dMinY, dMaxX, dMaxY
            
            For ii = 2 To m_objRects1.Count
                Set m_objRect1 = m_objRects1.Item(ii)
                
                'Get the range of new object
                GetRangeWithMargin m_objRect1, dNewMinX, dNewMinY, dNewMaxX, dNewMaxY
                
                'Move the object so its min matches the point vertically above the prev groups min
                m_objRect1.Move dNewMinX, dNewMinY, dMinX, dMaxY
                
                'Set the range as prev groups range
                GetRangeWithMargin m_objRect1, dMinX, dMinY, dMaxX, dMaxY
                
                'Add the object into select set
                m_objRect1.Select
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
    End If
    Exit Sub
ErrorHandler:

End Sub

Private Sub Form_Terminate()
    'Destroy rectangles and bins member collection objects
    Set m_objRects = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Destroy the RAD member objects
    TerminateRAD
End Sub

Private Sub TerminateRAD()

    Set m_objOrigSheet = Nothing
End Sub

Private Sub GetRangeWithMargin(objGroup As Group, dXMin As Double, dYMin As Double, dXMax As Double, dYMax As Double)
    
    'Get X and Y co-ordinates of the Group
    objGroup.Range dXMin, dYMin, dXMax, dYMax
    
    'Add margin and return the updated co-ordinates
    dXMin = dXMin - X_MARGIN
    dYMin = dYMin - X_MARGIN
    dXMax = dXMax + Y_MARGIN
    dYMax = dYMax + Y_MARGIN
    
End Sub

