Attribute VB_Name = "ExecuteRule"
Option Explicit

Public Function ExecuteFeatureRule(oDetailedPart As Object, oResourceManager As Object) As Boolean

    On Error GoTo ErrorHandler
    
    ExecuteFeatureRule = False

    Dim oStructFeature As Object
    Dim oIJDesignParent As IJDesignParent
    Dim oNamingUtils As IJNamingUtils2
    Dim oStructEntityNaming As IJDStructEntityNaming
    Dim oStructEntityUtils As IJStructEntityUtils
    Dim oMultiRule As Object

    Set oStructEntityUtils = New StructEntityUtils
    Set oNamingUtils = oStructEntityUtils
    
    If TypeOf oDetailedPart Is IJPlate Or TypeOf oDetailedPart Is IJPlatePart Then
           
        'Do Corner feature creation stuff here
        'From the plate object, get the corner data
        Dim oCorners As IJElements
        Dim oFeatures As Collection
        Dim oCorner As New CornerData
        Dim oPlatePart As New PlatePart
                
        Set oPlatePart.object = oDetailedPart
        Set oCorners = oPlatePart.GetPlateCornerData(False)
        Set oFeatures = oPlatePart.PlateFeatures
                        
        For Each oCorner In oCorners
            Dim oEdgePort1 As IJPort
            Dim oEdgePort2 As IJPort
            Dim oFacePort As IJPort
            Dim oStructFeatureFactory As IJSDFeatureDefinition
            Dim oCornerFeature As New CornerFeature
            Dim bFoundFeature As Boolean
            Dim bCreatedInCompute As Boolean

            Set oEdgePort1 = oCorner.EdgePort1
            Set oEdgePort2 = oCorner.EdgePort2
            Set oFacePort = oCorner.FacePort
            Set oStructFeatureFactory = New SDFeatureUtils
            Set oIJDesignParent = oDetailedPart
            bFoundFeature = False
            bCreatedInCompute = False

            'Check if there is already a feature on the corner
            'if so, we will skip it
            For Each oStructFeature In oFeatures
                'Sketched features are included in the collection of plate features,
                'but they don't support IJStructFeature. We don't care about
                'sketched features so if we have an object that doesn't support
                'IJStructFeature, we skip it
                If TypeOf oStructFeature Is IJStructFeature Then
                    Dim oLocation As IJDPosition
                    Dim dAngle1 As Double
                    Dim dAngle2 As Double
                    Dim dTol As Double
    
                    dTol = 0.00001
                    Dim oIJStructFeature As IJStructFeature
                    Dim oStructFeatureType As StructFeatureTypes
                    Set oIJStructFeature = oStructFeature
                    oStructFeatureType = oIJStructFeature.get_StructFeatureType()
                    If oStructFeatureType = SF_CornerFeature Then
                        Set oCornerFeature.object = oStructFeature
                        oCornerFeature.GetLocationAndAngles oLocation, dAngle1, dAngle2
                        If oLocation.DistPt(oCorner.Location) < dTol Then
                            bFoundFeature = True
                            Set oLocation = Nothing
                            Set oCornerFeature = Nothing
                            Set oIJStructFeature = Nothing
                            Exit For
                        End If
                        
                        Set oLocation = Nothing
                        Set oCornerFeature = Nothing
                    End If
                    Set oIJStructFeature = Nothing
                End If
            Next
            Set oStructFeature = Nothing

      
            If (Not bFoundFeature) Then
                On Error Resume Next
                                
                Dim isCornerValid As Boolean
                Dim bFailed As Boolean
                
                If oMultiRule Is Nothing Then
                    Set oMultiRule = oStructEntityUtils.GetOffSetRuleFromCatalog("ValidCorners")
                End If
                
                If (Not oMultiRule Is Nothing) And (Err.Number = 0) Then
                    isCornerValid = oMultiRule.isCornerValid(oEdgePort1, oEdgePort2)
                
                    If isCornerValid Then
                        oCornerFeature.Create oResourceManager, _
                                        oFacePort, _
                                        oEdgePort1, _
                                        oEdgePort2, _
                                        GetSOFeatureRootClass(SMARTTYPE_CORNERFEATURE), _
                                        oIJDesignParent
                        
                        If Err.Number <> 0 Then
                            bFailed = True
                        End If
                        
                        On Error GoTo ErrorHandler

                        Set oStructEntityNaming = oCornerFeature.object
                        oStructEntityNaming.NamingRule = oNamingUtils.GetDefaultGenericNamingRule
                        
                        ExecuteFeatureRule = True
                    End If
                Else
                    bFailed = True
                End If
            End If
            
            Set oEdgePort1 = Nothing
            Set oEdgePort2 = Nothing
            Set oFacePort = Nothing
            Set oStructFeatureFactory = Nothing
            Set oStructFeature = Nothing
            Set oCornerFeature = Nothing
            Set oIJDesignParent = Nothing
            Set oStructEntityNaming = Nothing
            If bFailed = True Then
                Exit For
            End If
        Next
        
        Set oStructEntityUtils = Nothing
        Set oMultiRule = Nothing
        Set oNamingUtils = Nothing

        Set oPlatePart = Nothing
        Set oCorner = Nothing
        Set oCorners = Nothing
        Set oFeatures = Nothing
    End If
    
    Exit Function
    
ErrorHandler:

End Function


'------------------------------------------------------------------------------------------------------------
' Procedure (Function):
'     GetSOFeatureRootClass
'
' Description:
'     Given a SmartOccurrence feature type, performs a catalog query (SRDQuery) to retrieve
'     the root selection rule class and returns the class name.
'
' Arguments:
'     eFeatureType    Enum (SmartClassType)    The type of SmartOccurrence feature.
'------------------------------------------------------------------------------------------------------------

Private Function GetSOFeatureRootClass(eFeatureType As SmartClassType) As String
    Const METHOD = "GetSOFeatureRootClass"
    
    Dim oCatalogQuery As IJSRDQuery
    Dim oSOFeatureClassQuery As IJSmartQuery
    Dim oSORootClass As IJSmartClass
    
    'Check if we are ready to compute.  If so, proceed.
    Set oCatalogQuery = New SRDQuery
    Set oSOFeatureClassQuery = oCatalogQuery
    Set oSORootClass = oSOFeatureClassQuery.GetRootClass(eFeatureType)
    
    GetSOFeatureRootClass = oSORootClass.SCName
    
    Set oCatalogQuery = Nothing
    Set oSOFeatureClassQuery = Nothing
    Set oSORootClass = Nothing
    
    Exit Function
    
End Function


