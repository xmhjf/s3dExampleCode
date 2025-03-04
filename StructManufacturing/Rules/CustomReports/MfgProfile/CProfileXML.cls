VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProfileXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2003,2008 Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgProfileCustomReports
''  File        : CProfileMXML.cls
''
''  Description : Populates the XML DOM Document with profiles from the ElementBroker
''
''
''  Authors     : Marcel Veldhuizen
''
''  Notes       : Time-consuming routines:
''                      - GetSymbolInformation (initialisation and compute)
''**************************************************************************************
Implements IJDCustomReport

' Constants
Private Const Module = "CustomProfileXML.CProfileXML"
Private Const PI = 3.14159265358979
Private Const OBJECT_ID = "OID"
Private Const dblFmt = "#####0.0#####"

' To enable debugging, set SM3DDEBUG to something > 0
Private Const SM3DDEBUG = 0

' Member variables
Private m_objXMLDoc             As DOMDocument
Private m_oPOM                  As IJDPOM
'- MfgProfile --------------------------------------------------'
Private Const IJMfgProfilePart = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"


'
'*************************************************************************
' * End Added by MJV                                                         *
'*************************************************************************
'

Private Sub Class_Initialize()
'    InitUOMService
    Set m_oPOM = GetPOM
End Sub

Private Sub Class_Terminate()
    Set m_objXMLDoc = Nothing
    Set m_oPOM = Nothing
End Sub

' ShellExpnSerializer has only one ElementList.
' DOMDocument is passed-in without initialization
Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    
    If pElements.Count > 0 Then
        Set m_objXMLDoc = New DOMDocument
        m_objXMLDoc.loadXML "<report> <elementList></elementList> </report>"
        Dim objTemplateSet As Object
        Dim oSelectedObj As Object
        For Each oSelectedObj In pElements
            If Not TypeOf oSelectedObj Is IJMfgProfilePart Then
                'This is profile. Get MfgProfiles from profile
                Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                Dim oColl As IJElements
                Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJMfgProfilePart)
                
                'For each TemplateSet, report the information
                Dim oMfgProfilePart As Object
                Set oMfgProfilePart = oColl.Item(1)
                ReportMfgProfileInformation oMfgProfilePart, strFileName
            Else
                ReportMfgProfileInformation oSelectedObj, strFileName
            End If
        Next
        eCustomReportStatus = StrMfgProcessFinished
    End If

Wrapup:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

' ShellExpnSerializer has only one ElementList.
' DOMDocument is passed-in without initialization
Private Sub ReportMfgProfileInformation(ByVal oMfgProfile As IJMfgProfilePart, strFileName As String)
    Const METHOD = "ReportMfgProfileInformation"
    
    Dim oIXMLDOMElementList As IXMLDOMElement
    Dim i As Integer
    Dim j As Integer
    Dim x As Integer

    Dim oSDProfileSupport As GSCADSDPartSupport.IJProfilePartSupport
    Set oSDProfileSupport = New GSCADSDPartSupport.ProfilePartSupport
    
    Dim oSDProfilePart As StructDetailObjects.ProfilePart
    Set oSDProfilePart = New StructDetailObjects.ProfilePart
                    
    Dim oFeatureCol As Collection
    Dim sProfileID As String
    Dim sSectionSize As String
    Dim sSectionType As String
    Dim sMaterialType As String
    Dim sGrade As String
    Dim sCurvature As String
    Dim sSymLocation As String
    
    Dim sShip As String
    Dim sBlock As String
    Dim sObjectID As String
    Dim sOffset As String
    Dim sBoard As String
    Dim sUpside As String
    
    Dim sFeatureA As String
    Dim sFeatureOID As String
    Dim sFeatureB As String
    Dim sParameters() As String
    Dim sParametersValue() As String
    
    Dim oObject As Object
    Dim oFeature As IJStructFeature
    Dim dDis As Double
    Dim strNewVal As String
    
    'SmartOccurence related information
    Dim oSmartOccurence As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oSymbol As IJDSymbol
       
    'Symmetry options that can be set on a part bases
    Dim mBoardSymmetry As IMSProfileEntity.JSBM_SYMMETRY
    
    'Required variables to get to the Block name
    Dim strBlockName As String
    Dim sItemName As String
         
    On Error GoTo ErrorHandler
    
    Dim oAdaptor As IJMfgXMLData
    On Error Resume Next
    Set oAdaptor = SP3DCreateObject("MfgXMLAdaptor.GNest")
    If oAdaptor Is Nothing Then
        MsgBox "ERROR: Cannot create XML-adaptor with prog-ID: MfgXMLAdaptor.GNest"
    End If
    On Error GoTo ErrorHandler
                   
    'Initialize report
    Set oIXMLDOMElementList = m_objXMLDoc.selectSingleNode("//report/elementList")
        
    ' Get detailed profile
    Dim oProfilePart As IJProfilePart
    Set oProfilePart = oMfgProfile.GetDetailedPart
            
    'Initialize SD profile-helper
    Set oSDProfilePart.object = oProfilePart
     
    'Initialize SD profile-support
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = oSDProfileSupport
    Set oSDPartSupport.Part = oSDProfilePart.object
            
    Dim partName As String
    partName = oSDProfilePart.Name
                             
    'Get mfg-entities
    Dim colFinal2d As Collection
    GetMfgProfileEntities oMfgProfile, colFinal2d
                                    
    'Get upside
    Dim webUpside As PrimaryProfileFace
            
    Dim oProcSettings As IJMfgProfileProcessSettings
    Set oProcSettings = oMfgProfile.GetProcessSettings
    sUpside = oProcSettings.ProfileUpside
    
    If sUpside = "WebLeft" Then
        webUpside = WebLeftFace
    Else
        sUpside = "WebRight"
        webUpside = WebRightFace
    End If
            
    'Get scaling-shrinkage (we use only primary shrinkage and assume it to be in length-direction of profile)
    Dim primShr As Double
    Dim secShr As Double
    Dim primDir As IJDVector
    Dim secDir As IJDVector
    Dim primPortUnk As IUnknown
    Dim secPortUnk As IUnknown
    
    primShr = 1#
    secShr = 1#
    Set primPortUnk = Nothing
    Set secPortUnk = Nothing
    
    oMfgProfile.GetTotalScalingShr Nothing, primShr, secShr, primDir, secDir, primPortUnk, secPortUnk
            
    'Get processing info
    Dim xMin As Double
    Dim xMax As Double
    Dim yMin As Double
    Dim yMax As Double
    Dim marginMin As Double
    Dim marginMax As Double
    Dim sDevelopDir As String
    
    'oDevMat is the 3D->2D transformation of UV-mark
    Dim oDevMat As IJDT4x4
    
    GetMfgProcessInfo oMfgProfile, _
                      colFinal2d, _
                      xMin, _
                      marginMin, _
                      xMax, _
                      marginMax, _
                      yMin, _
                      yMax, _
                      sDevelopDir, _
                      oDevMat
             
             
    'Retrieve the values directly derived from the profile part
    sProfileID = oSDProfilePart.Name
    sSectionSize = oSDProfilePart.SectionName
    sSectionType = oSDProfilePart.SectionType
    sMaterialType = oSDProfilePart.MaterialType
    sGrade = oSDProfilePart.Grade
    
    'Get the project name
    Dim oShipRoot As IJProjectRoot
    Set oShipRoot = oSDProfilePart.ShipModelObject
    sShip = oShipRoot.Name
      
    Dim oRange As IJRangeAlias
    Set oRange = oProfilePart
    Dim y_low As Double
    Dim y_high As Double
    y_high = oRange.GetRange.m_high.y
    y_low = oRange.GetRange.m_low.y
    
    'Get geometrical board
    If y_low >= 0 Then
        sBoard = "P"
    ElseIf y_high <= 0 Then
        sBoard = "S"
    Else
        sBoard = "C"
    End If
            
    'Get board symmetry flag
    mBoardSymmetry = oSDProfilePart.Symmetry
    
    Select Case mBoardSymmetry
    Case Symmetry_Centered
        sBoard = "C"
                
    Case Symmetry_PortOnly
        sBoard = "P"
        
    Case Symmetry_StarboardOnly
        sBoard = "S"
    
    End Select
    
    'Get the blockname in case the part is assigned to a block
    Dim oMfgParent As GSCADAsmHlpers.IJMfgParent
    Dim oNamedItem As IJNamedItem
    Dim oMfgChild As GSCADAsmHlpers.IJMfgChild
    
    If TypeOf oSDProfilePart.object Is IJMfgParent Then
        Set oMfgParent = oSDProfilePart.object
    Else
        Set oMfgParent = Nothing
    End If
    
    Dim bBlockFound As Boolean
    bBlockFound = False
    
    Do While Not bBlockFound And Not oMfgParent Is Nothing
        If TypeOf oMfgParent Is IJBlock Then
            Set oNamedItem = oMfgParent
            strBlockName = oNamedItem.Name
            bBlockFound = True
        Else
            If TypeOf oMfgParent Is IJMfgChild Then
                Set oMfgChild = oMfgParent
                Set oMfgParent = oMfgChild.getParent
            Else
                Set oMfgParent = Nothing
            End If
        End If
    Loop
    
    If bBlockFound = True Then
       sBlock = strBlockName
    Else
       sBlock = "No-Blk"
    End If
                                      
    'Retrive the OID from the Manufacturing Object
    sObjectID = GetOID(oMfgProfile)
                
    'Get curvature via stored property of mfg-profile
    Dim curvature As ProfileCurvature
    curvature = oMfgProfile.CurvatureType
            
    If curvature = PROFILE_CURVATURE_Straight Then
        sCurvature = "Straight"
    Else
        sCurvature = "NotStraight"
    End If
            
    Dim oProfileNode As IXMLDOMNode
    Dim oProfilePropNode As IXMLDOMNode
    
    'Add part-element profile property-list to report
    AddReportProfile oIXMLDOMElementList, oProfileNode, oProfilePropNode
    
    ' oWebCutPosStart and oWebCutPosEnd determines web-length
    Dim oWebCutPosStart As IJDPosition
    Set oWebCutPosStart = Nothing
    
    Dim oWebCutPosEnd As IJDPosition
    Set oWebCutPosEnd = Nothing
    
    'Get all features related to the profile
    Set oFeatureCol = Nothing
    oSDPartSupport.GetFeatures oFeatureCol
    
    'Extract only features supporting IJStructFeature
    Dim oStructFeatures As Collection
    Set oStructFeatures = New Collection
    
    Dim unkFeature As IUnknown
    Dim sName As String
    
    For j = 1 To oFeatureCol.Count
        Set unkFeature = oFeatureCol.Item(j)
        If TypeOf unkFeature Is IJStructFeature Then
            Set oFeature = unkFeature
            oStructFeatures.Add oFeature
        Else
            '@@@: How to handle e.g. sketched features???
            If TypeOf unkFeature Is IJNamedItem Then
                Set oNamedItem = unkFeature
                sName = oNamedItem.Name
            Else
                sName = "<no-name>"
            End If
            MsgBox "Not an IJStructFeature: " & sName
        End If
    Next
    
    Set unkFeature = Nothing
    Set oNamedItem = Nothing
    Set oFeature = Nothing
    Set oFeatureCol = Nothing
            
    'Retrieve all the values related to the features
    
    For Each oFeature In oStructFeatures
        
        Dim bAddMargin As Boolean
        bAddMargin = False ' Only web/flange-cuts are adjusted for margin
                    
        Dim sPortSide As String
        sPortSide = ""
        
        Select Case oFeature.get_StructFeatureType
        Case SF_WebCut
            sFeatureA = "WEBCUT"
            bAddMargin = True
            
            Dim m_WebCut As WebCut
            Set m_WebCut = New StructDetailObjects.WebCut
            Set m_WebCut.object = oFeature
            
            sPortSide = GetPortBaseOffsetSide(m_WebCut.BoundedPort)
            
            Set m_WebCut = Nothing
                                
        Case SF_EdgeFeature
            sFeatureA = "EDGE"
                
            Dim m_oEdgeFeature As EdgeFeature
            Set m_oEdgeFeature = New StructDetailObjects.EdgeFeature
            Set m_oEdgeFeature.object = oFeature
            
            Set m_oEdgeFeature = Nothing
                
        Case SF_CornerFeature
            sFeatureA = "CORNER"
            
            Dim m_oCornerFeature As CornerFeature
            Set m_oCornerFeature = New StructDetailObjects.CornerFeature
            Set m_oCornerFeature.object = oFeature
            
            Set m_oCornerFeature = Nothing
    
        Case SF_FaceFeature
            sFeatureA = "FACE"
        
        Case SF_FlangeCut
            sFeatureA = "FLANGECUT"
            bAddMargin = True
            
            Dim m_oFlangeCutFeature As FlangeCut
            Set m_oFlangeCutFeature = New StructDetailObjects.FlangeCut
            Set m_oFlangeCutFeature.object = oFeature
            
            sPortSide = GetPortBaseOffsetSide(m_oFlangeCutFeature.BoundedPort)
                                    
            Set m_oFlangeCutFeature = Nothing
    
        Case SF_Slot
            sFeatureA = "SLOT"
            
            Dim m_oSlotFeature As Slot
            Set m_oSlotFeature = New StructDetailObjects.Slot
            Set m_oSlotFeature.object = oFeature
            
            Set m_oSlotFeature = Nothing
    
        Case Else
            sFeatureA = "UNKNOWN"
        
        End Select
                        
        If sFeatureA <> "UNKNOWN" Then
        
            If TypeOf oFeature Is IJNamedItem Then
                Dim oIJNamedItem As IJNamedItem
                Set oIJNamedItem = oFeature
                sFeatureB = oIJNamedItem.Name
            Else
                sFeatureB = ""
            End If
            
            sFeatureOID = GetOID(oFeature)
            
            Dim oGNESTNode As IXMLDOMNode
            Dim rootGap As Double
            Set oGNESTNode = Nothing
            rootGap = 0
            If sFeatureA = "WEBCUT" Or sFeatureA = "FLANGECUT" Then
                ' Need bevel and root-gap for web/flange cuts
                Set oGNESTNode = GetXMLBevel(oMfgProfile, _
                                             oAdaptor, _
                                             oSDPartSupport, _
                                             oFeature, _
                                             colFinal2d)
                
                If Not oGNESTNode Is Nothing Then
                    Dim oItem As IXMLDOMNode
                    Set oItem = oGNESTNode.Attributes.getNamedItem("GAP")
                    If Not oItem Is Nothing Then
                        rootGap = 0.001 * oItem.nodeValue
                    End If
                End If
            End If
                                
            Set oSymbol = oFeature
            Set oSmartOccurence = oFeature
            Set oSmartItem = oSmartOccurence.ItemObject
            sItemName = oSmartItem.Name
                            
            Dim colParamNames As Collection
            Dim colParamValues As Collection
            Dim colDimNames As Collection
            Dim colDimValues As Collection
            Dim colRefPointNames As Collection
            Dim colRefPointPos As Collection
            
            'Get symbol-transformation
            Dim oSymMat As IJDT4x4 '2D-symbol to 3D transformation
                            
            GetSymbolInformation oSymbol, _
                                 colParamNames, _
                                 colParamValues, _
                                 colDimNames, _
                                 colDimValues, _
                                 colRefPointNames, _
                                 colRefPointPos, _
                                 oSymMat
                                 
            
            'Get feature-transformation
            Dim oFeatMat As IJDT4x4 ' 2D-symbol to 2D-StrMfg transformation
            Dim sSymOrient As String 'Symbolic feature-orientation
            Dim sFeatureLump As String ' "WEB", "TOPFLANGE" or "BOTTOMFLANGE"
            
            GetFeatureTransformation oFeature, _
                                     colFinal2d, _
                                     oSymMat, _
                                     oDevMat, _
                                     oFeatMat, _
                                     sSymOrient, _
                                     sFeatureLump
                            
            'Transform reference-points
            TransformRefPoints colRefPointPos, _
                               oFeatMat, _
                               bAddMargin, _
                               xMin, _
                               marginMin, _
                               xMax, _
                               marginMax, _
                               primShr, _
                               rootGap
            
            
            Dim oColLengthPoints As Collection
            Set oColLengthPoints = Nothing
            
            Dim bBase As Boolean
            Dim bOffset As Boolean
            
            bBase = False
            bOffset = False
            
            If sPortSide = "BASE" Then
                bBase = True
            ElseIf sPortSide = "OFFSET" Then
                bOffset = True
            End If
            
            If sFeatureA = "WEBCUT" Then
                Set oColLengthPoints = oSDProfilePart.EndCutLengthPoints_X1(bBase, bOffset, WebCut_Points)
                AdjustLengthPoints oFeature, colRefPointNames, colRefPointPos, oDevMat, oColLengthPoints
                
                UpdateWebCutPos colRefPointNames, colRefPointPos, xMin, xMax, oWebCutPosStart, oWebCutPosEnd
            ElseIf sFeatureA = "FLANGECUT" Then
                If sFeatureLump = "TOPFLANGE" Then
                    Set oColLengthPoints = oSDProfilePart.EndCutLengthPoints_X1(bBase, bOffset, FlangeCut_Points_Top)
                    AdjustLengthPoints oFeature, colRefPointNames, colRefPointPos, oDevMat, oColLengthPoints
                ElseIf sFeatureLump = "BOTTOMFLANGE" Then
                    Set oColLengthPoints = oSDProfilePart.EndCutLengthPoints_X1(bBase, bOffset, FlangeCut_Points_Bottom)
                    AdjustLengthPoints oFeature, colRefPointNames, colRefPointPos, oDevMat, oColLengthPoints
                End If
            End If
            
            ' Get symbolic location: Start(Bottom),Start(Top),End(Bottom),End(Top),unknown
            sSymLocation = GetSymLocation(sFeatureA, _
                                          xMin, _
                                          xMax, _
                                          yMin, _
                                          yMax, _
                                          colRefPointNames, _
                                          colRefPointPos)
            
            'Add feature to the XML
            AddReportFeature oProfileNode, _
                             sFeatureA, _
                             sFeatureOID, _
                             sFeatureB, _
                             sItemName, _
                             sSymLocation, _
                             sSymOrient, _
                             colParamNames, _
                             colParamValues, _
                             colDimNames, _
                             colDimValues, _
                             colRefPointNames, _
                             colRefPointPos
            
            If sFeatureA = "WEBCUT" Or sFeatureA = "FLANGECUT" Then
                'Add bevel to the XML
                AddReportBevel oProfileNode, oGNESTNode, sFeatureOID
            End If
            
        End If
                   
     Next oFeature
                 
     'Correct web start/end according to web-cuts
     
     Dim xWebMin As Double
     
     If Not oWebCutPosStart Is Nothing Then
         xWebMin = oWebCutPosStart.x
     Else
         xWebMin = xMin * primShr
         xWebMin = xWebMin - marginMin
     End If
     
     Dim xWebMax As Double
     
     If Not oWebCutPosEnd Is Nothing Then
         xWebMax = oWebCutPosEnd.x
     Else
         xWebMax = xMax * primShr
         xWebMax = xWebMax + marginMax
     End If
                  
     'Get processingLength
     Dim dProcLen As Double
     dProcLen = 0#
     
     Dim oIJMfgProfileLengths As IJMfgProfileLengths
     Set oIJMfgProfileLengths = oMfgProfile
     dProcLen = oIJMfgProfileLengths.AfterFeaturesWeb
                       
     'Add properties for the profile
     AddReportProfileProp oProfilePropNode, _
                          sShip, _
                          sProfileID, _
                          sBoard, _
                          sBlock, _
                          sObjectID, _
                          sSectionSize, _
                          sSectionType, _
                          sMaterialType, _
                          sGrade, _
                          sUpside, _
                          sCurvature, _
                          dProcLen, _
                          xWebMax - xWebMin, _
                          xWebMin, _
                          xWebMax, _
                          marginMin, _
                          marginMax, _
                          primShr, _
                          sDevelopDir
                              
    m_objXMLDoc.save strFileName

Wrapup:
    Set oIXMLDOMElementList = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
    GoTo Wrapup

End Sub



Private Sub AddReportProfile(oIXMLDOMElementList As IXMLDOMElement, _
                             ByRef oElement As IXMLDOMNode, _
                             ByRef oPropList As IXMLDOMNode)

    Const METHOD = "AddReportProfile"
    On Error GoTo ErrorHandler
    
    Set oElement = AddChildNode(oIXMLDOMElementList, "element")
    AddAttribute oElement, "elementName", "Part"
    
    Set oPropList = AddChildNode(oElement, "propertyList")
    AddAttribute oPropList, "propertyListName", "Profile"

Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub AddReportProfileProp(oIXMLDOMNode As IXMLDOMNode, _
                                 sShip As String, _
                                 sProfileID As String, _
                                 sBoard As String, _
                                 sBlock As String, _
                                 sObjectID As String, _
                                 sSectionSize As String, _
                                 sSectionType As String, _
                                 sMaterialType As String, _
                                 sGrade As String, _
                                 sUpside As String, _
                                 sCurvature As String, _
                                 dProcessingLength As Double, _
                                 dWebLength As Double, _
                                 dRefPosStart As Double, _
                                 dRefPosEnd As Double, _
                                 dMarginStart As Double, _
                                 dMarginEnd As Double, _
                                 dShrinkageFactor As Double, _
                                 sDevelopDir As String)

    Const METHOD = "AddReportProfileProp"
    On Error GoTo ErrorHandler
    
    Dim oPropList           As IXMLDOMNode
    Dim oProperty           As IXMLDOMNode
    Dim i As Integer
        
    Set oPropList = oIXMLDOMNode
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Ship"
    AddAttribute oProperty, "propertyValue", sShip

    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Block"
    AddAttribute oProperty, "propertyValue", sBlock

    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", OBJECT_ID
    AddAttribute oProperty, "propertyValue", sObjectID

    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Name"
    AddAttribute oProperty, "propertyValue", sProfileID
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Section Size"
    AddAttribute oProperty, "propertyValue", sSectionSize
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Section Type"
    AddAttribute oProperty, "propertyValue", sSectionType
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Material Type"
    AddAttribute oProperty, "propertyValue", sMaterialType
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Grade"
    AddAttribute oProperty, "propertyValue", sGrade
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Board"
    AddAttribute oProperty, "propertyValue", sBoard
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Upside"
    AddAttribute oProperty, "propertyValue", sUpside
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Curvature"
    AddAttribute oProperty, "propertyValue", sCurvature
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "ProcessingLength"
    AddAttribute oProperty, "propertyValue", dProcessingLength
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "WebLength"
    AddAttribute oProperty, "propertyValue", dWebLength
            
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "RefPosStart"
    AddAttribute oProperty, "propertyValue", dRefPosStart
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "RefPosEnd"
    AddAttribute oProperty, "propertyValue", dRefPosEnd
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "MarginStart"
    AddAttribute oProperty, "propertyValue", dMarginStart
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "MarginEnd"
    AddAttribute oProperty, "propertyValue", dMarginEnd
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "ShrinkageFactor"
    AddAttribute oProperty, "propertyValue", dShrinkageFactor
        
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "DevelopmentDir"
    AddAttribute oProperty, "propertyValue", sDevelopDir

Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
    ' do not raise error
    Err.Clear
    GoTo Cleanup
    
End Sub

Private Sub AddReportFeature(oProfileNode As IXMLDOMNode, _
                            sFeatureA As String, _
                            sFeatureOID As String, _
                            sFeatureB As String, _
                            sItemName As String, _
                            sSymLocation As String, _
                            sSymOrient As String, _
                            colParamNames As Collection, _
                            colParamValues As Collection, _
                            colDimNames As Collection, _
                            colDimValues As Collection, _
                            colRefPointNames As Collection, _
                            colRefPointPos As Collection)

    Const METHOD = "AddReportFeature"
    On Error GoTo ErrorHandler
    
    Dim oElement            As IXMLDOMNode
    Dim oPropList           As IXMLDOMNode
    Dim oProperty           As IXMLDOMNode
    Dim i As Integer
        
    Set oPropList = AddChildNode(oProfileNode, "propertyList")
    AddAttribute oPropList, "propertyListName", "Feature"
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "FeatureType"
    AddAttribute oProperty, "propertyValue", sFeatureA
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", OBJECT_ID
    AddAttribute oProperty, "propertyValue", sFeatureOID
     
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Name"
    AddAttribute oProperty, "propertyValue", sFeatureB
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "ItemName"
    AddAttribute oProperty, "propertyValue", sItemName
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "SymLocation"
    AddAttribute oProperty, "propertyValue", sSymLocation
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "SymOrientation"
    AddAttribute oProperty, "propertyValue", sSymOrient
    
    For i = 1 To colParamNames.Count
        Set oProperty = AddChildNode(oPropList, "property")
        AddAttribute oProperty, "propertyName", colParamNames(i)
        AddAttribute oProperty, "propertyValue", colParamValues(i)
    Next i
    
'    For i = 1 To colDimNames.Count
'        Set oProperty = AddChildNode(oPropList, "property")
'        AddAttribute oProperty, "propertyName", colDimNames(i)
'        AddAttribute oProperty, "propertyValue", colDimValues(i)
'    Next i
    
    Dim oPos As IJDPosition
    Dim pntName As String
    
    For i = 1 To colRefPointNames.Count
        pntName = colRefPointNames(i)
        
        If (SM3DDEBUG > 0) Or _
           (pntName <> "RefPoint_Origo" And pntName <> "RefPoint_SymbolOrigin") _
        Then
            Set oProperty = AddChildNode(oPropList, "property")
            AddAttribute oProperty, "propertyName", colRefPointNames(i)
        
            Set oPos = colRefPointPos(i)
            AddAttribute oProperty, "propertyValue", Format(oPos.x, dblFmt) & " , " & Format(oPos.y, dblFmt)
        End If
    Next i
    
Cleanup:
    Set oElement = Nothing
    Set oPropList = Nothing
    Set oProperty = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function GetXMLBevel(oMfgProfile As IJMfgProfilePart, _
                             oAdaptor As IJMfgXMLData, _
                             oSDPartSupport As IJPartSupport, _
                             oFeature As IJStructFeature, _
                             colFinal2d As Collection) As IXMLDOMNode

    Const METHOD = "GetXMLBevel"
    On Error GoTo ErrorHandler
            
    ' Get primary feature-port
    Dim oPrimPortMnk As IMoniker
    Dim colSubPortMnk As Collection
    oSDPartSupport.GetFeaturePortMonikers oFeature, oPrimPortMnk, colSubPortMnk
        
    Dim oMGHelper As GSCADMathGeom.MfgMGHelper
    Set oMGHelper = New MfgMGHelper
    Dim unkPrimPort As IUnknown
    oMGHelper.BindMoniker oSDPartSupport.Part, oPrimPortMnk, unkPrimPort
        
    ' Find a mfg-geometry with bevel-data that points to oPrimPort
    Dim OK As Boolean
    OK = False
        
    Dim countGeom As Integer
    countGeom = colFinal2d.Count
    
    Dim i As Integer
    For i = 1 To countGeom
        Dim oMfgGeom As MfgGeom2d
        Set oMfgGeom = colFinal2d.Item(i)
            
        If oMfgGeom.GetGeometryType = STRMFG_OUTER_CONTOUR Then
            Dim xmlBevel As String
            xmlBevel = ""
            On Error Resume Next
            oMfgGeom.GetBevelXML xmlBevel
            On Error GoTo ErrorHandler
            If xmlBevel <> "" Then
                Dim unkMoniker As IUnknown
                Set unkMoniker = oMfgGeom.GetMoniker
                
                If Not unkMoniker Is Nothing Then
                    Dim unkObject As IUnknown
                    Set unkObject = m_oPOM.GetObject(unkMoniker)
                
                    If unkPrimPort Is unkObject Then
                        OK = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    
    If Not OK Then
        'MsgBox "Could not find bevel for the feature"
        Exit Function
    End If
    
    Set GetXMLBevel = GetBevelNode(oMfgProfile, oAdaptor, xmlBevel)
    
Cleanup:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function


Private Sub AddReportBevel(oProfileNode As IXMLDOMNode, oGNESTNode As IXMLDOMNode, sFeatureOID As String)

    Const METHOD = "AddReportBevel"
    On Error GoTo ErrorHandler
    
    If oGNESTNode Is Nothing Then
        'MsgBox "xmlBevel not in GNEST-format"
        Exit Sub
    End If
        
    Dim oElement    As IXMLDOMNode
    Dim oPropList   As IXMLDOMNode
    Dim oProperty   As IXMLDOMNode
    Dim oItem       As IXMLDOMNode
    Dim GNESTParam  As Double
    
    Set oPropList = AddChildNode(oProfileNode, "propertyList")
    AddAttribute oPropList, "propertyListName", "Bevel"

    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "FeatureOID"
    AddAttribute oProperty, "propertyValue", sFeatureOID
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Angle_A"
    Set oItem = oGNESTNode.Attributes.getNamedItem("ANGLE_A")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * PI / 180#
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Angle_B"
    Set oItem = oGNESTNode.Attributes.getNamedItem("ANGLE_B")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * PI / 180#
    
'    Set oProperty = AddChildNode(oPropList, "property")
'    AddAttribute oProperty, "propertyName", "Angle_N"
'    Set oItem = oGNESTNode.Attributes.getNamedItem("ANGLE_N")
'    If oItem Is Nothing Then
'        GNESTParam = 0#
'    Else
'        GNESTParam = oItem.nodeValue
'    End If
'    AddAttribute oProperty, "propertyValue", GNESTParam * PI / 180#
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Angle_D"
    Set oItem = oGNESTNode.Attributes.getNamedItem("ANGLE_D")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * PI / 180#
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Angle_E"
    Set oItem = oGNESTNode.Attributes.getNamedItem("ANGLE_E")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * PI / 180#
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Chamfer_Angle_M"
    Set oItem = oGNESTNode.Attributes.getNamedItem("CHAMFER_ANGLE_M")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Chamfer_Angle_UM"
    Set oItem = oGNESTNode.Attributes.getNamedItem("CHAMFER_ANGLE_UM")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Chamfer_Depth_M"
    Set oItem = oGNESTNode.Attributes.getNamedItem("CHAMFER_DEPTH_M")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Chamfer_Depth_UM"
    Set oItem = oGNESTNode.Attributes.getNamedItem("CHAMFER_DEPTH_UM")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Depth_A"
    Set oItem = oGNESTNode.Attributes.getNamedItem("DEPTH_A")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Depth_B"
    Set oItem = oGNESTNode.Attributes.getNamedItem("DEPTH_B")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Depth_C"
    Set oItem = oGNESTNode.Attributes.getNamedItem("DEPTH_C")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Depth_D"
    Set oItem = oGNESTNode.Attributes.getNamedItem("DEPTH_D")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Depth_E"
    Set oItem = oGNESTNode.Attributes.getNamedItem("DEPTH_E")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Gap"
    Set oItem = oGNESTNode.Attributes.getNamedItem("GAP")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
    
    Set oProperty = AddChildNode(oPropList, "property")
    AddAttribute oProperty, "propertyName", "Shift"
    Set oItem = oGNESTNode.Attributes.getNamedItem("SHIFT")
    If oItem Is Nothing Then
        GNESTParam = 0#
    Else
        GNESTParam = oItem.nodeValue
    End If
    AddAttribute oProperty, "propertyValue", GNESTParam * 0.001
       
Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

'-------------------------------------------------------------------------------
'        XML Schema for Report Data :
'        <xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema">
'
'            <xsd:annotation>
'                <xsd:documentation xml:lang="en">
'                    Default Report XML Schema
'                    Copyright 2001. All rights reserved.
'                </xsd:documentation>
'            </xsd:annotation>
'
'            <xsd:element name="report" type="ReportType"/>
'
'            <xsd:complexType name="ReportType">
'            <xsd:element name="elementList"   type="ElementListType" minoccurs="0" maxoccurs="unbounded"/>
'            <xsd:attribute name="reportName" type="xsd:string" />
'            <xsd:attribute name="reportUID" type="xsd:string" required="true" />
'            </xsd:complexType>
'
'            <xsd:complexType name="ElementListType">
'            <xsd:element name="element"   type="ElementType" minoccurs="0" maxoccurs="unbounded"/>
'            <xsd:attribute name="elementListName" type="xsd:string" />
'            </xsd:complexType>
'
'            <xsd:complexType name="ElementType">
'            <xsd:element name="propertyList"   type="PropertyListType" minoccurs="0" maxoccurs="unbounded"/>
'            <xsd:attribute name="elementName" type="xsd:string" />
'            </xsd:complexType>
'
'            <xsd:complexType name="PropertyListType">
'            <xsd:element name="property" type="PropertyType" minoccurs="0" maxoccurs="unbounded"/>
'            <xsd:attribute name="propertyListName" type="xsd:string">
'            </xsd:complexType>
'
'            <xsd:complexType name="PropertyType">
'            <xsd:attribute name="propertyName" type="xsd:string" required="true"/>
'            <xsd:attribute name="propertyValue" type="xsd:string" required="true"/>
'            <xsd:attribute name="propertyType" type="xsd:string"/>
'            </xsd:complexType>
'
'        </xsd:schema>
'
'        xml Snippet:
'
'        <report reportName="#name#" reportUID="#uuid#" >
'           <elementList elementListName="#name#" >  *
'                <element elementName="#name#" > *
'                 <propertyList propertyListName="#name#" >  ?
'                 <property
'            propertyName="#name#"  propertyValue="#name#"
'                propertyType="#name#"? /> *
'                 </propertyList>
'            </element>
'           </elementList>
'        </report>
'        ·   Characters appended to elements or attributes represent the following: "?" (0 or 1), "*" (0 or more), "+" (1 or more).
'        ·   For a Shell Expansion Report, it is assumed that -only one PropertyList node will exist within an Element node.
'

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'******************************************************************
' Method : AddChildNode
'
' Description:
'   This method creates an node (of type NODE_ELEMENT) and links it to the
'   parent node.
'
' Arguments:
'   [in] Parent As IXMLDOMNode, the node to be parent of the new Node
'   [in] NodeName As String,    the name new node
'
'   Return Values:
'
'******************************************************************
Private Function AddChildNode(Parentname As IXMLDOMNode, ChildName As String) As IXMLDOMNode
    Const METHOD As String = "AddChildNode"
    On Error GoTo ErrorHandler
    Dim oChild As IXMLDOMNode
    'create childnode, and hook up to parent
    Set oChild = m_objXMLDoc.createElement(ChildName)
    Set oChild = Parentname.appendChild(oChild)
    Set AddChildNode = oChild
    
Cleanup:
    Set oChild = Nothing
    Exit Function
    
ErrorHandler:
    Set AddChildNode = Nothing
    Err.Raise Err.Number, , Err.Description
    GoTo Cleanup

End Function

'******************************************************************
' Method : AddAttribute
'
' Description:
'   This method create an attribute assigns a value and link it to the parent node.
'
' Arguments:
'   [in] Parent As IXMLDOMNode, the node to be parent of the attribute
'   [in] AttName As String,     the name of the attribute
'   [out] AttValue As String    the value of the attribute (can be blank)
'
'   Return Values:
'
'******************************************************************
Private Sub AddAttribute(Parent As IXMLDOMNode, attName As String, AttValue As Variant)
    Const METHOD As String = "AddAttribute"
    On Error GoTo ErrorHandler
    Dim oATT As IXMLDOMAttribute
    Set oATT = m_objXMLDoc.createAttribute(attName)
    Dim oNamedNodeMap As IXMLDOMNamedNodeMap
    Set oNamedNodeMap = Parent.Attributes
    oNamedNodeMap.setNamedItem oATT
    
    Dim sValue As String
    
    If VarType(AttValue) = vbDouble Then
        sValue = Format(AttValue, dblFmt)
    Else
        sValue = AttValue
    End If
    
    oATT.Value = sValue

Cleanup:
    Set oATT = Nothing
    Set oNamedNodeMap = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
    GoTo Cleanup
    
End Sub


''**************************************************************************************
'' Routine      : GetMfgProfilesCollection
'' Description  : Get all IJMfgProfilePart monikers from the broker.
''                NOTE: If a block/assembly is selected, we need to recursively pick
''                up the contained mfg-profiles
''**************************************************************************************
Private Function GetMfgProfilesCollection(oElements As GSCADStructMfgGlobals.IJElements) As IJElements
    Const METHOD = "GetMfgProfilesCollection"
    On Error GoTo ErrorHandler

    Dim oMfgProfiles As IJElements
    Set oMfgProfiles = New IMSCoreCollections.JMonikerCollection
    Set GetMfgProfilesCollection = oMfgProfiles
    
    Dim oHierarchyWalker As IJHierarchyWalker
    Set oHierarchyWalker = New GSCADHierarchyWalker.CGSCADHierarchyWalker
    
'    Dim oElements As IJElements
'    Set oElements = oBroker.GetElements
    
    Dim oSelObj As Object
    Dim oSelMnk As IMoniker
    For Each oSelObj In oElements
        If TypeOf oSelObj Is IJMfgProfilePart Then
            Set oSelMnk = m_oPOM.GetObjectName(oSelObj)
            oMfgProfiles.Add oSelMnk
        ElseIf TypeOf oSelObj Is IJAssembly Then
            Dim oAssyRoot As IJAssembly
            Set oAssyRoot = oSelObj
            
            ' Need to recursiveliy get all the children
            oHierarchyWalker.Flatten "CustomProfileXML.CWalkerAssist", _
                                     WLK_HRCHY_MFG, _
                                     True, _
                                     oAssyRoot, _
                                     False
                        
            Dim oMnkMfgProfile As IMoniker
            Set oMnkMfgProfile = oHierarchyWalker.First
            
            While Not oMnkMfgProfile Is Nothing
                oMfgProfiles.Add oMnkMfgProfile
                Set oMnkMfgProfile = oHierarchyWalker.Next
            Wend
        End If
    Next

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Private Sub GetMfgProfileEntities(oMfgProfilePart As IJMfgProfilePart, _
                                  ByRef colFinal2d As Collection)
                                  
    Const METHOD = "GetMfgProfileEntities"
    On Error GoTo ErrorHandler
        
    Set colFinal2d = Nothing
    
    If oMfgProfilePart Is Nothing Then
        Exit Sub
    End If
    
    Dim oProc2dOutput As IJMfgGeomCol2d
    Set oProc2dOutput = oMfgProfilePart.FinalGeometriesAfterProcess2D
    '
    ' Transfer MfgGeomCol3d/MfgGeomCol2d to VB-collections for faster access
    '
    Dim i As Integer
    Dim colCount As Integer
    Dim oGeom2d As MfgGeom2d
    Set colFinal2d = New Collection
    colCount = oProc2dOutput.Getcount
    'MsgBox "Final2d: " & colCount
    
    For i = 1 To colCount
        Set oGeom2d = oProc2dOutput.GetGeometry(i)
        colFinal2d.Add oGeom2d
    Next
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

''**************************************************************************************
'' Routine      : GetFeatureTransformation
'' Description  : Get symbol to after-unfold transformation. This is accomplished by
''                concatenating the symbol view-transformation with the 3D->2D feature
''                transformation from StrMfg
''
''                NOTE: In case we do not find both the 3D and 2D feature-mark, the StrMfg
''                part of the transformation is taken for (oDevOrigo,oDevU,oDevV).
''                --- THIS WORKS ONLY FOR STRAIGHT PROFILES ---
''**************************************************************************************
Private Sub GetFeatureTransformation(oFeature As IJStructFeature, _
                                     colFinal2d As Collection, _
                                     oSymMat As IJDT4x4, _
                                     oDevMat As IJDT4x4, _
                                     ByRef oFeatMat As IJDT4x4, _
                                     ByRef sSymOrient As String, _
                                     ByRef sFeatureLump As String)
    
    Const METHOD = "GetFeatureTransformation"
    On Error GoTo ErrorHandler
                                                                                          
    Dim oMathUtil As MfgMathGeom
    Set oMathUtil = New MfgMathGeom
                                                 
    Dim oGeom2d As IJMfgGeom2d
    Dim oFeatureMark3d As IJMfgGeom2d
    Dim oFeatureMark2d As IJMfgGeom2d
    
    Set oFeatureMark3d = Nothing
    Set oFeatureMark2d = Nothing
    
    Dim mnkFeature As IMoniker
    Dim punkFeature As IUnknown
    
    sFeatureLump = "WEB"
    
    For Each oGeom2d In colFinal2d
        If oGeom2d.GetGeometryType = STRMFG_FEATURE_MARK Then
            Set mnkFeature = oGeom2d.GetMoniker
            If Not mnkFeature Is Nothing Then
                Set punkFeature = BindMfgMoniker(mnkFeature)
                If oFeature Is punkFeature Then
                    If oGeom2d.GetSubGeometryType = STRMFG_FEATURE_MARK Then
                        Set oFeatureMark3d = oGeom2d
                        If oGeom2d.FaceId = 513 Then
                            sFeatureLump = "BOTTOMFLANGE"
                        ElseIf oGeom2d.FaceId = 514 Then
                            sFeatureLump = "TOPFLANGE"
                        End If
                        
                        If (Not oFeatureMark2d Is Nothing) And _
                           (Not oFeatureMark3d Is Nothing) _
                        Then
                            Exit For
                        End If
                    Else
                        Set oFeatureMark2d = oGeom2d
                        
                        If (Not oFeatureMark2d Is Nothing) And _
                           (Not oFeatureMark3d Is Nothing) _
                        Then
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If Not oFeatureMark3d Is Nothing And Not oFeatureMark2d Is Nothing Then
        ' ComputeFeatureTransf gives 3D to 2D transformation !!!
        Set oFeatMat = oMathUtil.ComputeFeatureTransf(oFeatureMark3d, oFeatureMark2d)
    Else
        'MsgBox "WARNING: No feature-mark. Using UV-mark"
        Set oFeatMat = oDevMat.Clone
    End If
    
    ' 1) oSymMat is 2D-symbol to 3D transformation
    ' 2) oFeatMat is 3D to 2D-StrMfg transformantion
    ' So after MultMatrix,oFeatMat becomes 2D-symbol to 2D-StrMfg transformation
    oFeatMat.MultMatrix oSymMat
    
    '
    ' NOTE: If oFeatMat(origo.z) is non-zero, then the feature is places on the non up-side
    '
    
    Dim imgX As String
    imgX = Convert2Axis(oFeatMat.IndexValue(0), oFeatMat.IndexValue(1), oFeatMat.IndexValue(2))
    
    Dim imgY As String
    imgY = Convert2Axis(oFeatMat.IndexValue(4), oFeatMat.IndexValue(5), oFeatMat.IndexValue(6))
    
    sSymOrient = imgX & " , " & imgY
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
                                             
End Sub

''**************************************************************************************
'' Routine      : TransformRefPoints
'' Description  : Transform feature ref-points and also,
''                take care of margin, shrinkage and root-gap.
''**************************************************************************************
Private Sub TransformRefPoints(ByRef colRefPointPos, _
                               oFeatMat As IJDT4x4, _
                               bAddMargin As Boolean, _
                               xMin As Double, _
                               marginMin As Double, _
                               xMax As Double, _
                               marginMax As Double, _
                               primShr As Double, _
                               rootGap As Double)
    
    Const METHOD = "TransformRefPoints"
    On Error GoTo ErrorHandler

    Dim oPos As IJDPosition
    Dim transPos As IJDPosition
    
    Dim xmid As Double
    xmid = 0.5 * (xMin + xMax)
    
    For Each oPos In colRefPointPos
        Set transPos = oFeatMat.TransformPosition(oPos)
        
        transPos.x = transPos.x * primShr
        
        If bAddMargin Then
            ' Adjust for margin
            If transPos.x <= xmid Then
                transPos.x = transPos.x - marginMin
            Else
                transPos.x = transPos.x + marginMax
            End If
        End If
        
        ' Adjust for root-gap
        If transPos.x <= xmid Then
            transPos.x = transPos.x + rootGap
        Else
            transPos.x = transPos.x - rootGap
        End If
        
        oPos.Set transPos.x, transPos.y, transPos.z
    Next

    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
                                                                            
End Sub

''**************************************************************************************
'' Routine      : AdjustLengthPoints
'' Description  : Adjust length-points to "processing-shape"
'' Note         : Symbol-points are located in the center of the material
''                We need to transform to one of the two sides (max)
''
'' Index in oColLengthPoints:
'' 1 Bottom side 1 (at base/offset)
'' 2 Bottom side 2 (at base/offset)
'' 3 Top side 1 (at base/offset)
'' 4 Top side 2 (at base/offset)
'' Base/Offset is context of endcut-port
''**************************************************************************************
Private Sub AdjustLengthPoints(oFeature As IJStructFeature, _
                               colRefPointNames As Collection, _
                               colRefPointPos As Collection, _
                               oDevMat As IJDT4x4, _
                               oColLengthPoints As Collection)
    Const METHOD = "AdjustLengthPoints"
    On Error GoTo ErrorHandler
    
    Dim n As Integer
    n = oColLengthPoints.Count
        
    Dim oNI As IJNamedItem
    Set oNI = oFeature
    If n < 4 Then
        Set oNI = oFeature
        MsgBox "AdjustLengthPoints: Need 4 length-points, got " & n & ", " & oNI.Name
        Exit Sub
    End If
    
    Dim botL1 As Double
    Dim botD1 As Double
    Dim topL1 As Double
    Dim topD1 As Double
    
    Dim oPos1 As IJDPosition
    Dim oPos2 As IJDPosition
    
    Set oPos1 = oDevMat.TransformPosition(oColLengthPoints(1))
    Set oPos2 = oDevMat.TransformPosition(oColLengthPoints(2))
    
    botL1 = oPos1.x
    botD1 = oPos2.x - oPos1.x
    
    Set oPos1 = oDevMat.TransformPosition(oColLengthPoints(3))
    Set oPos2 = oDevMat.TransformPosition(oColLengthPoints(4))
    
    topL1 = oPos1.x
    topD1 = oPos2.x - oPos1.x
        
    ' Make adjustmenst relative to original symbol refpoints in center
    botD1 = -Abs(botD1 * 0.5)
    topD1 = -Abs(topD1 * 0.5)
    
    Dim oPos As IJDPosition
    Dim vntName As Variant
    Dim i As Integer
    i = 0
    
    For Each vntName In colRefPointNames
        i = i + 1
        
        If vntName = "RefPoint_BottomLengthPoint" Then
            Set oPos = colRefPointPos.Item(i)
            oPos.x = oPos.x + botD1
        ElseIf vntName = "RefPoint_TopLengthPoint" Then
            Set oPos = colRefPointPos.Item(i)
            oPos.x = oPos.x + topD1
        End If
    Next
    
    Exit Sub
    
ErrorHandler:
        Err.Raise Err.Number, , Err.Description
End Sub

''**************************************************************************************
'' Routine      : UpdateWebCutPos
'' Description  : Update either webcut start- or end-position
''**************************************************************************************
Private Sub UpdateWebCutPos(colRefPointNames As Collection, _
                            colRefPointPos As Collection, _
                            xMin As Double, _
                            xMax As Double, _
                            ByRef oWebCutPosStart As IJDPosition, _
                            ByRef oWebCutPosEnd As IJDPosition)
    
    Const METHOD = "UpdateWebCutPos"
    On Error GoTo ErrorHandler

    Dim xmid As Double
    xmid = 0.5 * (xMin + xMax)
    
    Dim vntName As Variant
    Dim i As Integer
    i = 0
    
    For Each vntName In colRefPointNames
        i = i + 1
        
        If vntName = "RefPoint_BottomLengthPoint" Then
            Dim oPos As IJDPosition
            Set oPos = colRefPointPos.Item(i)
                        
            If oPos.x <= xmid Then
                Set oWebCutPosStart = oPos
            Else
                Set oWebCutPosEnd = oPos
            End If
            
            Exit Sub
        End If
    Next

    Exit Sub
    
ErrorHandler:
        Err.Raise Err.Number, , Err.Description
    
End Sub

''**************************************************************************************
'' Routine      : GetSymLocation
'' Description  : Get symbolic feature-location:
''                     - Start(Bottom)
''                     - Start(Top)
''                     - End(Bottom)
''                     - End(Top)
''                     - unknown
''**************************************************************************************
Private Function GetSymLocation(sFeatureA As String, _
                                xMin As Double, _
                                xMax As Double, _
                                yMin As Double, _
                                yMax As Double, _
                                colRefPointNames As Collection, _
                                colRefPointPos As Collection) As String
    
    Const METHOD = "GetSymLocation"
    On Error GoTo ErrorHandler

    Dim sLoc As String
    sLoc = "unknown"
        
    Dim xmid As Double
    xmid = (xMin + xMax) * 0.5

    Dim ymid As Double
    ymid = (yMin + yMax) * 0.5
    
    Dim oPos As IJDPosition
    Dim vntName As Variant
    Dim i As Integer
    i = 0
    
    For Each vntName In colRefPointNames
        i = i + 1
        
        '--------------------------------------------------------------
        ' NOTE: We use Origo (for Web/Flange-cuts and Corner-features)
        '       because SymbolOrigin is not well-defined for those.
        '       For edge-features, SymbolOrigin is well-defined.
        '--------------------------------------------------------------
        
        If sFeatureA = "EDGE" Then
            If vntName = "RefPoint_Origin" Then
                Set oPos = colRefPointPos.Item(i)
                    
                sLoc = "" 'No start or end-specification for edge-features
                    
                If oPos.y <= ymid Then
                    sLoc = sLoc & "Bottom"
                Else
                    sLoc = sLoc & "Top"
                End If
                
                GetSymLocation = sLoc
                Exit Function
            End If
        Else
            If vntName = "RefPoint_Origo" Then
                Set oPos = colRefPointPos.Item(i)
            
                If oPos.x <= xmid Then
                    sLoc = "Start"
                Else
                    sLoc = "End"
                End If
            
                If sFeatureA = "CORNER" Then
                    If oPos.y <= ymid Then
                        sLoc = sLoc & "Bottom"
                    Else
                        sLoc = sLoc & "Top"
                    End If
                End If
            
                GetSymLocation = sLoc
                Exit Function
            End If
        End If
                        
    Next

    GetSymLocation = sLoc
    Exit Function
    
ErrorHandler:
        Err.Raise Err.Number, , Err.Description
    
End Function

' Get symbolic location: Start(Bottom),Start(Top),End(Bottom),End(Top),NotSet
'sSymbolicLocation = GetSymbolicLocation(sFeatureA, WD, WL, colRefPointNames, colRefPointPos)

''**************************************************************************************
'' Routine      : GetMfgProcessInfo
'' Description  : Get positions of profile-ends after unfold (xMin,xMax) and
''                the cummulative margin-value at each end (maringMin,marginMax).
''                Also return factor for scaling-shrinkage (scalingShr).
''                Finally determines development-direction (sDevelopDir) which is
''                the 3D-direction of the profile when going from xMin to xMax (i.e.
''                the positive U-axis direction). We also return the 3D to 2D transformation
''                determined by the UV-mark as oDevMat
''
''**************************************************************************************
Private Sub GetMfgProcessInfo(oMfgProfile As IJMfgProfilePart, _
                              colFinal2d As Collection, _
                              ByRef xMin As Double, _
                              ByRef marginMin As Double, _
                              ByRef xMax As Double, _
                              ByRef marginMax As Double, _
                              ByRef yMin As Double, _
                              ByRef yMax As Double, _
                              ByRef sDevelopDir As String, _
                              ByRef oDevMat As IJDT4x4)
                              
    Const METHOD = "GetMfgProcessInfo"
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim oCurve As IJCurve
    Dim oGeom2d As MfgGeom2d
    
    ' Find xMin and xMax
    
    Dim x1 As Double
    Dim x2 As Double
    Dim y1 As Double
    Dim y2 As Double
    Dim z1 As Double
    Dim z2 As Double
        
    Dim bInit As Boolean
    bInit = True
    
    For Each oGeom2d In colFinal2d
        If oGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR Then
            Set oCurve = oGeom2d.GetGeometry
            oCurve.EndPoints x1, y1, z1, x2, y2, z2
            If bInit Then
                xMin = x1
                xMax = x1
                yMin = y1
                yMax = y1
                bInit = False
            End If
            
            If xMin > x1 Then xMin = x1
            If xMin > x2 Then xMin = x2
            If xMax < x1 Then xMax = x1
            If xMax < x2 Then xMax = x2
            
            If yMin > y1 Then yMin = y1
            If yMin > y2 Then yMin = y2
            If yMax < y1 Then yMax = y1
            If yMax < y2 Then yMax = y2
        End If
    Next
    
    ' Find all margin-definitions
        
    Dim colMarginDefs As Collection
    Set colMarginDefs = New Collection
    
    Dim oAssocRel As IJDAssocRelation
    Dim oUnkColl As Object
    Dim oTargetObjCol As IJDTargetObjectCol
    Dim colCount As Integer

    Set oAssocRel = oMfgProfile.GetDetailedPart
    If Not oAssocRel Is Nothing Then
        Set oUnkColl = oAssocRel.CollectionRelations("IJMfgParent", "IsStrMfgChildOf")
        Set oTargetObjCol = oUnkColl
        
        colCount = oTargetObjCol.Count
        
        For i = 1 To colCount
            Dim punkMfgDef As IUnknown
            Set punkMfgDef = oTargetObjCol.Item(i)
            If TypeOf punkMfgDef Is IJMfgDefinition Then
                Dim oMfgDef As IJMfgDefinition
                Set oMfgDef = punkMfgDef
                If oMfgDef.GetMfgType = STRMFG_CONSTANTMARGIN Then
                    colMarginDefs.Add oMfgDef
                End If
            End If
        Next
    End If
    
    ' Calculate accumulated magin-value at xMin and xMax
    
    marginMin = 0
    marginMax = 0
    
    Dim xmid As Double
    xmid = 0.5 * (xMax + xMin)
    
    Dim oPort As IJPort
    Dim punkPort As IUnknown
    
    If colMarginDefs.Count > 0 Then
                
        ' Find out if base-port is at xMin or xMax
        Dim bBaseAtMin As Boolean
        Dim bBaseAtMax As Boolean
        
        bBaseAtMin = False
        bBaseAtMax = False
        
        For Each oGeom2d In colFinal2d
            Dim mnkGeom2d As IMoniker
            Set mnkGeom2d = oGeom2d.GetMoniker
                        
            If (Not mnkGeom2d Is Nothing) And (oGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR) Then
                
                Set punkPort = BindMfgMoniker(mnkGeom2d)
                
                Dim oStructPort As IJStructPort
                Set oStructPort = punkPort
                
                Dim ctx As eUSER_CTX_FLAGS
                ctx = oStructPort.ContextID
                
                Dim bBase As Boolean
                Dim bOffset As Boolean
                
                If ((ctx And CTX_BASE) = CTX_BASE) Then
                    bBase = True
                ElseIf ((ctx And CTX_OFFSET) = CTX_OFFSET) Then
                    bOffset = True
                Else
                    bBase = False
                    bOffset = False
                End If
                               
                If bBase Or bOffset Then
                    Set oCurve = oGeom2d.GetGeometry
                    oCurve.EndPoints x1, y1, z1, x2, y2, z2
                
                    If (x1 < xmid) And (x2 < xmid) Then
                        If bBase Then
                            bBaseAtMin = True
                        Else
                            bBaseAtMax = True
                        End If
                        Exit For
                    ElseIf (x1 > xmid) And (x2 > xmid) Then
                        If bBase Then
                            bBaseAtMax = True
                        Else
                            bBaseAtMin = True
                        End If
                        Exit For
                    End If
                End If
            End If
        Next
    
        'For each margin, classify port as base or offset and update marginMin, marginMax
        For Each oMfgDef In colMarginDefs
            
            Dim oConstMargin As IJConstMargin
            Set oConstMargin = oMfgDef
                
            Dim marginValue As Double
            marginValue = oConstMargin.Value
            
            Set oStructPort = oMfgDef.GetPort
            ctx = oStructPort.ContextID
            
            If ((ctx And CTX_BASE) = CTX_BASE) Then
                If bBaseAtMin Then
                    marginMin = marginMin + marginValue
                ElseIf bBaseAtMax Then
                    marginMax = marginMax + marginValue
                End If
            ElseIf ((ctx And CTX_OFFSET) = CTX_OFFSET) Then
                If bBaseAtMin Then
                    marginMax = marginMax + marginValue
                ElseIf bBaseAtMax Then
                    marginMin = marginMin + marginValue
                End If
            End If
            
        Next
        
    End If
    
    ' Find 3D UV-mark and 2D UV-mark
    Dim oUVmark3d As MfgGeom2d
    Dim oUVmark2d As MfgGeom2d
    
    Set oUVmark3d = Nothing
    Set oUVmark2d = Nothing
    
    For Each oGeom2d In colFinal2d
        If oGeom2d.GetGeometryType = STRMFG_UV_MARK Then
            If oGeom2d.GetSubGeometryType = STRMFG_UV_MARK Then
                Set oUVmark3d = oGeom2d
            Else
                Set oUVmark2d = oGeom2d
            End If
            
            If (Not oUVmark2d Is Nothing) And _
               (Not oUVmark3d Is Nothing) _
            Then
                Exit For
            End If
        End If
    Next
    
    If oUVmark3d Is Nothing Or oUVmark2d Is Nothing Then
        sDevelopDir = "unknown"
    Else
        Dim oMathUtil As MfgMathGeom
        Set oMathUtil = New MfgMathGeom
        
        ' ComputeFeatureTransf gives 3D to 2D transformation !!!
        Set oDevMat = oMathUtil.ComputeFeatureTransf(oUVmark3d, oUVmark2d)
        
        ' Get 2D to 3D transformation to determine development-dir
        Dim oDevMatInv As IJDT4x4
        Set oDevMatInv = oDevMat.Clone
        oDevMatInv.Invert
        
        Dim oDevU As IJDVector
        Set oDevU = New DVector
        oDevU.Set oDevMatInv.IndexValue(0), oDevMatInv.IndexValue(1), oDevMatInv.IndexValue(2)
        
        Dim dx As Double
        Dim dy As Double
        Dim dz As Double
        
        dx = Abs(oDevU.x)
        dy = Abs(oDevU.y)
        dz = Abs(oDevU.z)
        
        If dx >= dy And dx >= dz Then
            If oDevU.x > 0 Then
                sDevelopDir = "fore"
            Else
                sDevelopDir = "aft"
            End If
        ElseIf dy >= dx And dy >= dz Then
            If oDevU.y > 0 Then
                sDevelopDir = "port"
            Else
                sDevelopDir = "starboard"
            End If
        ElseIf dz >= dx And dz >= dy Then
            If oDevU.z > 0 Then
                sDevelopDir = "upper"
            Else
                sDevelopDir = "lower"
            End If
        End If
        
    End If
    
    Exit Sub
    
ErrorHandler:
        Err.Raise Err.Number, , Err.Description
    
End Sub

''**************************************************************************************
'' Routine      : GetSymbolInformation
'' Description  : This function retrieves all necessary information from a symbol
''**************************************************************************************

Private Sub GetSymbolInformation( _
                            ByVal oSymbol As IJDSymbol, _
                            ByRef colParamNames As Collection, _
                            ByRef colParamValues As Collection, _
                            ByRef colDimNames As Collection, _
                            ByRef colDimValues As Collection, _
                            ByRef colRefPointNames As Collection, _
                            ByRef colRefPointPos As Collection, _
                            ByRef oSymMat As IJDT4x4)
    
    Const METHOD = "GetSymbolInformation"
    On Error GoTo ErrorHandler
    
    ' Do a temporary compute of symbol to get the correct RAD2D-symbol
    Dim oDefPlayerEx As IJDDefinitionPlayerEx
    Set oDefPlayerEx = oSymbol.IJDSymbolDefinition(1).IJDDefinitionPlayerEx
    
    oDefPlayerEx.OpenGame oSymbol, igUndefined
    oDefPlayerEx.ComputeConverter
    
    Dim oDefPlayer As IJDDefinitionPlayer
    Set oDefPlayer = oSymbol.IJDSymbolDefinition(1).IJDDefinitionPlayer
    
    oDefPlayer.SetInputs oSymbol, oSymbol.IJDInputsArg.GetInputs(igINPUT_ARGUMENTS_MERGE)
    oDefPlayer.Calculate oSymbol
    
    ' Get RAD2D-symbol
    Dim oRAD As RAD2D.Document
    Set oRAD = oSymbol.IJDSymbolDefinition(1).IJDServerIdentification.object
    'Dim FileName As String
    'FileName = Environ("TEMP")
    'If FileName = "" Or FileName = vbNullString Then
    '    FileName = "c:\\temp" 'Only use C:\Temp if there is a %TEMP% failure
    'End If
    'FileName = FileName & "\\feature.sym"
    'oRAD.SaveCopyAs FileName

    Dim oGeomConv As CommonSymbolUtils.STGeomConverter
    Set oGeomConv = oDefPlayerEx.Converter
        
    ' NB: Use the view-transformation to determine sketch-plane incl (U,V)-directions
    ' (output sketch-plane for edge-features does not care about (U,V)-directions, only
    ' normal and origin is set correctly
    
    Set oSymMat = oGeomConv.ViewTransform.Clone
    oSymMat.IndexValue(15) = 1# 'NOTE: StructDetail does not always set scale-factor to 1 !!!
            
    GetSymbolInputParamters oSymbol, colParamNames, colParamValues
    GetRAD2DRefPoints oRAD, colRefPointNames, colRefPointPos
    
Cleanup:
    On Error Resume Next
    oDefPlayerEx.CloseGame
    Exit Sub
    
ErrorHandler:
        Err.Raise Err.Number, , Err.Description
    GoTo Cleanup
End Sub

''**************************************************************************************
'' Routine      : GetSymbolInputParameters
'' Description  : This function retrieves symbol input-parameters.
''**************************************************************************************
Private Sub GetSymbolInputParamters( _
    ByRef oSymbol As IJDSymbol, _
    ByRef colParamNames As Collection, _
    ByRef colParamValues As Collection)
    
    Const METHOD = "GetSymbolInputParameters"
    On Error Resume Next
    
    'Create output-collections
    Set colParamNames = New Collection
    Set colParamValues = New Collection
                 
    ' We do not really neeed input-parameters
    If SM3DDEBUG = 0 Then Exit Sub
    
    Dim sParamName As String
    Dim sParamValue As String
    
    Dim oInputs As IJDInputs
    Set oInputs = oSymbol.IJDSymbolDefinition(1).IJDInputs
    
    Dim oInput As IJDInput
    
    For Each oInput In oInputs
        If oInput.IsPropertySet(igINPUT_IS_A_PARAMETER) Then
            sParamName = oInput.Name
            
            Dim oPC As IJDParameterContent
            Set oPC = oInput.IJDInputDuringGame.Argument
            
            If oPC.Type = igString Then
                sParamValue = oPC.String
            Else
                sParamValue = oPC.UomValue
            End If
            
            colParamNames.Add "InputParameter_" & sParamName
            colParamValues.Add sParamValue
        End If
    Next
    
    Exit Sub
    
ErrorHandler:
        Err.Raise Err.Number, , Err.Description
    
End Sub

''**************************************************************************************
'' Routine      : GetRAD2RefPoints
'' Description  : This function retrieves some symbol reference-points in the following
''                layers:
''
''                -  "CrossSectionPlane"
''                -  "Input"
''
''                Two special points are added:
''                RefPoint_SymbolOrigin : RAD2D symbol-origin
''                RefPoint_Origio       : (0,0)
''**************************************************************************************
Private Sub GetRAD2DRefPoints(ByRef rad2dDoc As RAD2D.Document, _
                              ByRef colRefPointNames As Collection, _
                              ByRef colRefPointPos As Collection)
    
    Const METHOD = "GetRAD2DRefPoints"
    
    Const layerLength = "CrossSectionPlane"
    Const layerInput = "Input"
    
    On Error Resume Next
    
    'Create output-collections
    Set colRefPointNames = New Collection
    Set colRefPointPos = New Collection
    
    Dim oDrawingObjects As RAD2D.DrawingObjects
    Set oDrawingObjects = rad2dDoc.ActiveSheet.DrawingObjects

    Dim oObject As Object
    Dim oDwgObject As Object
    Dim oAttrbSets As RAD2D.AttributeSets
    Dim oAttrbSet As RAD2D.AttributeSet
    Dim oAttrb As RAD2D.Attribute
    Dim oGroup As RAD2D.Group
    Dim oRefPoint As RAD2D.Point2d
    Dim oNamedRefPoint As RAD2D.Point2d
    Dim sRefPointName As String
    Dim oPos As IJDPosition
    
    Dim ii As Long
        
    For Each oObject In oDrawingObjects
        If oObject.Type = igGroup Then
            Set oGroup = oObject
            Dim layerName As String
            layerName = oGroup.Layer
                            
            ' =================================================
            ' NOTE: We do NOT check layer-name as that name may
            ' vary from version to version
            ' =================================================
            
            For ii = 1 To oGroup.Count
                Set oNamedRefPoint = Nothing
                
                Set oDwgObject = oGroup.Item(ii)
                If oDwgObject.Type = igPoint2d Then
                    Set oRefPoint = oDwgObject '' Point2d
                    Set oAttrbSets = oRefPoint.AttributeSets
                    
                    For Each oAttrbSet In oAttrbSets
                        For Each oAttrb In oAttrbSet
                            If oAttrb.AttributeType = igAttrTypeString Then
                                If (SM3DDEBUG > 0) Or _
                                   (oAttrb.Value = "BottomLengthPoint") Or _
                                   (oAttrb.Value = "TopLengthPoint") Or _
                                   (oAttrb.Value = "Origin") _
                                Then
                                    Set oNamedRefPoint = oRefPoint
                                
                                    sRefPointName = "RefPoint_" & oAttrb.Value
                            
                                    Set oPos = New DPosition
                                    oPos.Set oNamedRefPoint.x, oNamedRefPoint.y, 0#
                                
                                    'Add to output-collections
                                    colRefPointNames.Add sRefPointName
                                    colRefPointPos.Add oPos
                                
                                    Exit For
                                End If
                            End If
                        Next
                        If Not oNamedRefPoint Is Nothing Then Exit For
                    Next
                End If
            Next
                
        End If
    Next oObject
        
    Dim x As Double
    Dim y As Double
    rad2dDoc.SymbolProperties.GetOrigin x, y
    Set oPos = New DPosition
    oPos.Set x, y, 0#
    
    colRefPointNames.Add "RefPoint_SymbolOrigin"
    colRefPointPos.Add oPos
    
    Set oPos = New DPosition
    oPos.Set 0#, 0#, 0#
    
    colRefPointNames.Add "RefPoint_Origo"
    colRefPointPos.Add oPos
    
    Exit Sub
    
ErrorHandler:
        Err.Raise Err.Number, , Err.Description

End Sub

Private Function GetOID(objDB As Object) As String

    'Retrive the OID from the Manufacturing Object
    Dim oMoniker As IMoniker
    Set oMoniker = m_oPOM.GetObjectMoniker(objDB)
    GetOID = m_oPOM.DbIdentifierFromMoniker(oMoniker)
    
End Function

'
' Bin a moniker contained in a IJMfgGeom3d/IJMfgGeom2d-object.
' We know it is a global moniker (relative to the POM).
' We do not report errors for monikers that cannot be bound.
'
Private Function BindMfgMoniker(mnkMfg As IMoniker) As IUnknown
    On Error Resume Next
    Set BindMfgMoniker = Nothing
    
    If mnkMfg Is Nothing Then
        Exit Function
    End If
        
    Set BindMfgMoniker = m_oPOM.GetObject(mnkMfg)
    
    Exit Function
End Function

'Private Sub InitUOMService()
'    Const METHOD As String = "InitUOMService"
'    On Error GoTo ErrorHandler
'
'    Dim oTrader As New Trader
'
'    Set m_oUomFormat = New UomFormat
'    Set m_oUnitsOfMeasure = oTrader.Service(TKUnitsOfMeasure, "")
'
'    ' UnitTypes + Units:
'    ' Unit for unit-type UNIT_DISTANCE could be DISTANCE_METER
'    ' Unit for unit-type UNIT_ANGLE could be ANGLE_RADIAN
'
'    m_oUnitsOfMeasure.SetDefaultUnits UNIT_DISTANCE, DISTANCE_METER
'    m_oUnitsOfMeasure.SetDefaultUnits UNIT_ANGLE, ANGLE_RADIAN
'
'    m_oUnitsOfMeasure.GetDefaultUnits UNIT_DISTANCE, m_CurrentDistanceUnits
'    m_oUnitsOfMeasure.GetDefaultUnits UNIT_ANGLE, m_CurrentAngleUnits
'
'    m_oUomFormat.PrecisionType = PRECISIONTYPE_DECIMAL
'    m_oUomFormat.DecimalPrecision = 4
'    m_oUomFormat.UnitsDisplayed = False
'
'    Set oTrader = Nothing
'
'    Exit Sub
'
'ErrorHandler:
'        Err.Raise Err.Number, , Err.Description
'
'End Sub

Private Function Convert2Axis(x As Double, y As Double, z As Double) As String
    If Abs(x) >= Abs(y) And Abs(x) >= Abs(z) Then
        Convert2Axis = "X"
        If x < 0 Then Convert2Axis = "-X"
    ElseIf Abs(y) >= Abs(z) And Abs(y) >= Abs(z) Then
        Convert2Axis = "Y"
        If y < 0 Then Convert2Axis = "-Y"
    Else
        Convert2Axis = "Z"
        If z < 0 Then Convert2Axis = "-Z"
    End If
End Function

Private Function GetBevelNode(ByRef oMfgProfile As IJMfgProfilePart, ByRef oAdaptor As IJMfgXMLData, ByVal xmlBevel As String) As IXMLDOMNode

    Set GetBevelNode = Nothing
        
    xmlBevel = "<SMS_DOC>" & xmlBevel & "</SMS_DOC>"
    Dim xmlDoc As DOMDocument
    Set xmlDoc = New DOMDocument
            
    If Not xmlDoc.loadXML(xmlBevel) Then
        Exit Function
    End If
    
    Dim oBevelList As IXMLDOMNodeList
    Set oBevelList = xmlDoc.getElementsByTagName("SMS_BEVEL")
        
    Dim numBevels As Integer
    
    If Not oBevelList Is Nothing Then
        numBevels = oBevelList.length
    End If
    
    If numBevels < 1 Then
        Exit Function
    End If
        
    If numBevels = 1 Then
        Set GetBevelNode = oBevelList.Item(0)
        Exit Function
    End If
        
    Dim seglen2 As Double
    Dim maxlen2 As Double
    maxlen2 = -1
    
    Dim i As Integer
    
    For i = 0 To numBevels - 1
        Dim oBevelNode As IXMLDOMNode
        Set oBevelNode = oBevelList.Item(i)

        Dim oPointList As IXMLDOMNodeList
        Set oPointList = xmlDoc.getElementsByTagName("CVG_POINT")
        
        Dim oP0 As IXMLDOMNode
        Set oP0 = oPointList.Item(0)
        
        Dim x0 As Double
        Dim y0 As Double
        x0 = oP0.Attributes.getNamedItem("X").nodeValue
        y0 = oP0.Attributes.getNamedItem("Y").nodeValue
        
        Dim oP1 As IXMLDOMNode
        Set oP1 = oPointList.Item(1)
        
        Dim x1 As Double
        Dim y1 As Double
        x1 = oP1.Attributes.getNamedItem("X").nodeValue
        y1 = oP1.Attributes.getNamedItem("Y").nodeValue
        
        seglen2 = (x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0)
        
        If seglen2 > maxlen2 Then
            maxlen2 = seglen2
            Set GetBevelNode = oBevelNode
        End If
    Next

    Exit Function
    
End Function

'***********************************************************************
' METHOD:  GetPortBaseOffsetSide
'
' DESCRIPTION:
'       Given an IJStructPort
'       Determine if the Port is "Base" port or "Offset" port
'
'   inputs:
'           oPort As IJPort
'
'   outputs:
'           GetPortBaseOffsetSide as String
'               "BASE"      given Port is "Base" port
'               "OFFSET"    given Port is "Offset" port
'               "TOP"       given Port is Profile JXSEC_TOP port
'               "BOTTOM"    given Port is Profile JXSEC_BOTTOM port
'               "WEB_LEFT"  given Port is Profile JXSEC_WEB_LEFT port
'               "WEB_RIGHT" given Port is Profile JXSEC_WEB_RIGHT port
'               ""          given Port is not "base" or "Offset"
'
'***********************************************************************
Private Function GetPortBaseOffsetSide(oPort As IJPort) As String

    On Error Resume Next
    Dim lBaseCheck As Long
    Dim lNplusCheck As Long
    Dim lNminusCheck As Long
    Dim lOffsetCheck As Long
    
    Dim lOperatorID As Long
    
    Dim sPortSide As String
    Dim eContextID As GSCADSDPartSupport.eUSER_CTX_FLAGS
    Dim oStructPort As IMSStructConnection.IJStructPort
    
    sPortSide = ""
    If TypeOf oPort Is IJStructPort Then
        Set oStructPort = oPort
        eContextID = oStructPort.ContextID
        lBaseCheck = eContextID And GSCADSDPartSupport.CTX_BASE
        lNplusCheck = eContextID And GSCADSDPartSupport.CTX_NPLUS
        lNminusCheck = eContextID And GSCADSDPartSupport.CTX_NMINUS
        lOffsetCheck = eContextID And GSCADSDPartSupport.CTX_OFFSET
        
        If lBaseCheck <> 0 Then
            sPortSide = "BASE"
            
        ElseIf lOffsetCheck <> 0 Then
            sPortSide = "OFFSET"
                
        ElseIf lNplusCheck <> 0 Then
            sPortSide = "OFFSET"
                
        ElseIf lNminusCheck <> 0 Then
            sPortSide = "BASE"
        
        Else
            ' Connecting Port is not Base or Offset port
            ' check if Profile Top/Bottom/Left/Right port
            lOperatorID = oStructPort.OperatorID
            If lOperatorID = JXSEC_TOP Then
                sPortSide = "TOP"
            ElseIf lOperatorID = JXSEC_BOTTOM Then
                sPortSide = "BOTTOM"
            ElseIf lOperatorID = JXSEC_WEB_LEFT Then
                sPortSide = "WEB_LEFT"
            ElseIf lOperatorID = JXSEC_WEB_RIGHT Then
                sPortSide = "WEB_RIGHT"
            Else
                sPortSide = ""
            End If
        End If
            
    End If
    
    GetPortBaseOffsetSide = sPortSide
    Exit Function
        
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

