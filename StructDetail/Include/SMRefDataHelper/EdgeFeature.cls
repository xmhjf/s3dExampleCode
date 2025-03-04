VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EdgeFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub UpdateDependentEdgeFeature(oPhyConn As Object)

Const sMETHOD = "UpdateDependentEdgeFeature"
On Error GoTo ErrHandler

    ' ----------------------
    ' Check for valid inputs
    ' ----------------------
    If Not TypeOf oPhyConn Is IJStructPhysicalConnection Then
        Exit Sub
    End If

    Dim oSDOPhysConn As New StructDetailObjects.PhysicalConn
    Set oSDOPhysConn.object = oPhyConn
    
    
    ' For each part
    ' -------------
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport
    
    Dim oFeaturesList As Collection
    Dim oFeature As IUnknown
    Dim oStructFeature As IJStructFeature
    Dim featureType As StructFeatureTypes
    Dim oFeatureLoc As IJDPosition
    Dim bForceUpdate As Boolean
    
    Dim exitflag As Boolean
    exitflag = False
    
    Dim oWB As IJWireBody
    Dim oComplexStrings As IJElements
    Dim oCurve As IJCurve
    Dim iCount As Long
    Dim jCount As Long
    Dim kCount As Long
    Dim oPlate As IJPlate
    
    For iCount = 1 To 2
        If iCount = 1 Then
            If TypeOf oSDOPhysConn.ConnectedObject1 Is IJPlate Then
                Set oPlate = oSDOPhysConn.ConnectedObject1
                If oPlate.plateType = Hull Then Exit For
            End If
            Set oPartSupport.Part = oSDOPhysConn.ConnectedObject1
        Else
            If TypeOf oSDOPhysConn.ConnectedObject2 Is IJPlate Then
                Set oPlate = oSDOPhysConn.ConnectedObject2
                If oPlate.plateType = Hull Then Exit For
            End If
            Set oPartSupport.Part = oSDOPhysConn.ConnectedObject2
        End If
        oPartSupport.GetFeatures oFeaturesList

        ' ----------------------------
        ' For each feature on the part
        ' ----------------------------
        Dim nFeatures As Long
        nFeatures = oFeaturesList.Count
        
        For jCount = 1 To nFeatures
            Set oFeature = oFeaturesList.Item(jCount)
            
            bForceUpdate = False
            
            ' -------------------------
            ' If edge feature
            ' -------------------------
            If TypeOf oFeature Is IJStructFeature Then
                Set oStructFeature = oFeature
                featureType = oStructFeature.get_StructFeatureType
                                
                Select Case featureType
                    Case SF_EdgeFeature
                        
                        Dim pIJStructFeatUtils As IJSDFeatureAttributes
                        Set pIJStructFeatUtils = New SDFeatureUtils
                        
                        Dim oppEdgePort As IJPort
                        Dim oppLocationPoint As Object
        
                        pIJStructFeatUtils.Get_Inputs_EdgeCut oStructFeature, oppEdgePort, oppLocationPoint
                        Dim oP3d As IngrGeom3D.Point3d
                        Set oP3d = New Point3d
                        
                        If TypeOf oppLocationPoint Is Point3d Then
                            Set oP3d = oppLocationPoint
                        Else 'Considering case when placed using seam or knuckle method
                            Dim oTempPos As IJDPosition
                            Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
                            Set oTopologyLocate = New TopologyLocate
        
                            Set oTempPos = oTopologyLocate.FindIntersectionPoint(oppEdgePort, oppLocationPoint)
                            oP3d.SetPoint oTempPos.x, oTempPos.y, oTempPos.z
                        End If
                        
                        Dim dx As Double
                        Dim dy As Double
                        Dim dz As Double
                        
                        oP3d.GetPoint dx, dy, dz
                        
                        Dim dMinDist As Double
                        Dim dSrcX As Double
                        Dim dSrcY As Double
                        Dim dSrCZ As Double
                        Dim dInX As Double
                        Dim dInY As Double
                        Dim dInZ  As Double
                        
                        If TypeOf oSDOPhysConn.object Is IJWireBody Then
                            Set oWB = oSDOPhysConn.object
        
                            oWB.GetComplexStrings Nothing, oComplexStrings
                            For kCount = 1 To oComplexStrings.Count
                                Set oCurve = oComplexStrings.Item(kCount)
                                
                                exitflag = oCurve.IsPointOn(dx, dy, dz)
                                
                                If exitflag = False Then
                                    oCurve.DistanceBetween oP3d, dMinDist, dSrcX, dSrcY, dSrCZ, dInX, dInY, dInZ
                                    'ReCompute Edge Feature if Point lies within a minimum distance of 10 mm from PC
                                    If dMinDist < 0.01 Then exitflag = True
                                End If
                                'Since point lies on PC, Update dependent Edge Feature
                                If exitflag = True Then Exit For
                            Next kCount
                        End If
                        
                        If exitflag = True Then
                            bForceUpdate = True
                        End If
                End Select
            End If
         
            If bForceUpdate Then
                Dim oHelper As New SMRefDataHelper.Helper
                oHelper.ForceUpdateSmartItem oFeature
                Set oHelper = Nothing
            End If
        Next jCount
    Next iCount
    
    Exit Sub

ErrHandler:
  Err.Raise LogError(Err, "RuleKeys", sMETHOD).Number
End Sub
