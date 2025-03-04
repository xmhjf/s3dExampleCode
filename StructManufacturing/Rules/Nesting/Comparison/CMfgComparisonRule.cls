VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMfgComparisonRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''
''   Copyright (C) 2002, Intergraph Corporation. All rights reserved.
''
''   Abstract:
''   CMfgComparisonRule.cls
''   ProgID:         {<(GSCADSTRMFGComparisonRule)>}.{<(CMfgComparisonRule)>}
''   Author:         R. Shanmugaraj
''   Creation Date:  14.jan.2002
''   Description:
''       Comparison rule for manufacturing parts.
''               -Customizable by the End User
''
''   Change History:
''   dd.mmm.yyyy     who     change description
''   -----------     ---     ------------------
''   22.04.2004     MJV      Included correct error handling
''
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'
Implements IJMfgComparisonRule
Const Module = "CMfgComparisonRule"

'*********************************************************************************************************
'Name       :   CompareMfgPlate
'
'Class      :   CMfgComparisonRule      ;Interface  :   IJMfgComparisonRule
'
'Purpose    :   Default Comparison Rule : Comparison rule for Structural Manufacturing Plates
'               - Two input Plate part is given for comparison
'
'
'Rule ID    :   Default Rule
'
'Etc. .     :   !

'User tolernces :
'                   1. dPlanarTolerance     = 0.001
'                   2. dContourTolerance    = 0.001

'*********************************************************************************************************
Private Function IJMfgComparisonRule_CompareMfgPlate(ByVal pMfgPlate1 As Object, ByVal pMfgPlate2 As Object, strRC As String) As Boolean
Const Method = "_CompareMfgPlate"
On Error GoTo ErrorHandler
        'Keep IJMfgComparisonRule_CompareMfgPlate  = False  always unless all comparison succeed.
        IJMfgComparisonRule_CompareMfgPlate = False
        Dim oMfgPlatePart1 As IJMfgPlatePart
        Dim oMfgPlatePart2 As IJMfgPlatePart
        Dim oPlatePart1 As IJPlatePart
        Dim oPlatePart2 As IJPlatePart
        
        Set oMfgPlatePart1 = pMfgPlate1
        Set oMfgPlatePart2 = pMfgPlate2
        
        Dim oPartInfo As IJDPartInfo
        'User input to check the planar plate Tolernces
        ' Planar tolerance
        Dim dPlanarTolerance As Double
        dPlanarTolerance = 0.001
        
        'Comparing the plates are planar or not
        '***************************************
        oMfgPlatePart1.GetDetailedPart oPlatePart1
        oMfgPlatePart2.GetDetailedPart oPlatePart2
        Set oPartInfo = New PartInfo

        If oPartInfo.IsPlanarPlate(oPlatePart1, dPlanarTolerance) = False Or _
            oPartInfo.IsPlanarPlate(oPlatePart2, dPlanarTolerance) = False Then

            IJMfgComparisonRule_CompareMfgPlate = False
            strRC = "Not a planar plate"
            GoTo SetToNothing
        End If



       'Getting XML Data for each plates & compare it's attributes
        Dim oXmlDom1 As DOMDocument
        Dim oXmlDom2 As DOMDocument

        'XML Node info
        '*************
        Dim oEdgeNodeList1 As IXMLDOMNodeList
        Dim oEdgeNodeList2 As IXMLDOMNodeList
        Dim oElem As IXMLDOMElement
        Dim oAttribute1 As IXMLDOMAttribute
        Dim oAttribute2 As IXMLDOMAttribute
        Set oXmlDom1 = oMfgPlatePart1.GetXMLData
        Set oXmlDom2 = oMfgPlatePart2.GetXMLData
        If oXmlDom1 Is Nothing Or oXmlDom2 Is Nothing Then
            IJMfgComparisonRule_CompareMfgPlate = False
            strRC = "XML data is not available"
            GoTo SetToNothing
        End If


        'Check for number of openings in the plate
        '*****************************************
        Dim nCurNode As Integer
        Dim nNodeCountInPlate1 As Integer
        Dim nNodeCountInPlate2 As Integer
        nNodeCountInPlate1 = 0
        nNodeCountInPlate2 = 0
        'Firslt plate
        Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_EDGE")
        If Not oEdgeNodeList1 Is Nothing Then
            For nCurNode = 0 To oEdgeNodeList1.Length - 1
                Set oElem = oEdgeNodeList1.Item(nCurNode)
                Set oAttribute1 = oElem.getAttributeNode("TYPE")
                ' Count no. of openings
                If oAttribute1.Value = "Hole" Then
                    nNodeCountInPlate1 = nNodeCountInPlate1 + 1
                End If
                Set oElem = Nothing
                Set oAttribute1 = Nothing
            Next
        End If
        Set oEdgeNodeList1 = Nothing

        'Second Plate
        Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_EDGE")
        If Not oEdgeNodeList2 Is Nothing Then
            For nCurNode = 0 To oEdgeNodeList2.Length - 1
                Set oElem = oEdgeNodeList2.Item(nCurNode)
                Set oAttribute1 = oElem.getAttributeNode("TYPE")
                ' Count no. of openings
                If oAttribute1.Value = "Hole" Then
                    nNodeCountInPlate2 = nNodeCountInPlate2 + 1
                End If
                Set oElem = Nothing
                Set oAttribute1 = Nothing
            Next
        End If
        Set oEdgeNodeList2 = Nothing

        'Check the opening(hole) count
        If Not nNodeCountInPlate1 = nNodeCountInPlate2 Then
            IJMfgComparisonRule_CompareMfgPlate = False
            strRC = "No. of openings differs"
            GoTo SetToNothing

        End If

        'Check for number of markings
        '*****************************
        Dim nNoOfMarkingsInPlate1 As Integer
        Dim nNoOfMarkingsInPlate2 As Integer
        nNoOfMarkingsInPlate1 = 0
        nNoOfMarkingsInPlate2 = 0

        ' Count marking in plate1
        Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_MARKING")
        Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_MARKING")
        If Not oEdgeNodeList1 Is Nothing Or Not oEdgeNodeList2 Is Nothing Then
            nNoOfMarkingsInPlate1 = oEdgeNodeList1.Length
            Set oEdgeNodeList1 = Nothing

            ' Count marking in plate2
            nNoOfMarkingsInPlate2 = oEdgeNodeList2.Length
            Set oEdgeNodeList2 = Nothing
            ' Check number of markings are same in plate1 and plate2
            If Not nNoOfMarkingsInPlate1 = nNoOfMarkingsInPlate2 Then
                IJMfgComparisonRule_CompareMfgPlate = False
                strRC = "No. of markings differs"
                GoTo SetToNothing
            End If

            Set oEdgeNodeList1 = Nothing
            Set oEdgeNodeList2 = Nothing

            'Check for each marking name in both plates, if number of markings are same
            '**************************************************************************

            Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_MARKING")
            Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_MARKING")

            For nCurNode = 0 To oEdgeNodeList1.Length - 1
                Set oElem = oEdgeNodeList1.Item(nCurNode)
                Set oAttribute1 = oElem.getAttributeNode("MARKING_NAME")
                Set oElem = Nothing
                Set oElem = oEdgeNodeList2.Item(nCurNode)
                Set oAttribute2 = oElem.getAttributeNode("MARKING_NAME")
                Set oElem = Nothing
                If Not oAttribute1 Is Nothing And Not oAttribute2 Is Nothing Then
                    If Not Trim(oAttribute1.Value) = "" And Not Trim(oAttribute2.Value) = "" Then
                        If Not oAttribute1.Value = oAttribute2.Value Then
                            IJMfgComparisonRule_CompareMfgPlate = False
                            strRC = "Marking name differs"
                           GoTo SetToNothing
                        End If
                    End If
                End If
                Set oAttribute1 = Nothing
                Set oAttribute2 = Nothing
            Next

        End If

        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing

      'Check for Material type
        '***********************
        Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_PART_INFO")
        Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_PART_INFO")
        If Not oEdgeNodeList1 Is Nothing Or Not oEdgeNodeList2 Is Nothing Then
            Set oElem = oEdgeNodeList1.Item(0)

            Set oAttribute1 = oElem.getAttributeNode("MATERIAL_TYPE")
            Set oElem = Nothing
            Set oElem = oEdgeNodeList2.Item(0)
            Set oAttribute2 = oElem.getAttributeNode("MATERIAL_TYPE")
            Set oElem = Nothing

            If Not oAttribute1.Value = oAttribute2.Value Then
                IJMfgComparisonRule_CompareMfgPlate = False
                strRC = "Material type differs"
                GoTo SetToNothing
            End If
            Set oAttribute1 = Nothing
            Set oAttribute2 = Nothing
        End If
        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing

        'Check for material grade
        '------------------------
        Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_PART_INFO")
        Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_PART_INFO")
        Set oElem = oEdgeNodeList1.Item(0)
        Set oAttribute1 = oElem.getAttributeNode("MATERIAL_GRADE")
        Set oElem = Nothing
        Set oElem = oEdgeNodeList2.Item(0)
        Set oAttribute2 = oElem.getAttributeNode("MATERIAL_GRADE")
        Set oElem = Nothing

        If Not oAttribute1.Value = oAttribute2.Value Then
            IJMfgComparisonRule_CompareMfgPlate = False
            strRC = "Material grade differs"
            GoTo SetToNothing
        End If
        Set oAttribute1 = Nothing
        Set oAttribute2 = Nothing
        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing
        'Check for thickness of the plates
        '---------------------------------
        Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_PLATE")
        Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_PLATE")
        Set oElem = oEdgeNodeList1.Item(0)
        Set oAttribute1 = oElem.getAttributeNode("THICKNESS")
        Set oElem = Nothing
        Set oElem = oEdgeNodeList2.Item(0)
        Set oAttribute2 = oElem.getAttributeNode("THICKNESS")
        Set oElem = Nothing

        If Not oAttribute1.Value = oAttribute2.Value Then
            IJMfgComparisonRule_CompareMfgPlate = False
            strRC = "Thickness differs"
            GoTo SetToNothing
        End If
        Set oAttribute1 = Nothing
        Set oAttribute2 = Nothing
        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing

        'Before checking no .of bevels, & no. of margins - first compare no. of edges
        '----------------------------------------------------------------------------
        Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_EDGE")
        Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_EDGE")
        
        If Not oEdgeNodeList1.Length = oEdgeNodeList2.Length Then
            IJMfgComparisonRule_CompareMfgPlate = False
            strRC = "No. of edges differs"
            GoTo SetToNothing
        Else
            'Check for Number of bevel symbols in the plate
            '------------------------------------------
            For nCurNode = 0 To oEdgeNodeList1.Length - 1
                Set oElem = oEdgeNodeList1.Item(nCurNode)
                Set oAttribute1 = oElem.getAttributeNode("NO_BEVEL")
                Set oElem = Nothing
                Set oElem = oEdgeNodeList2.Item(nCurNode)
                Set oAttribute2 = oElem.getAttributeNode("NO_BEVEL")
    
                If Not oAttribute1 Is Nothing And Not oAttribute2 Is Nothing Then
                    ' check bevel count
                    If Not oAttribute1.Value = oAttribute2.Value Then
                        IJMfgComparisonRule_CompareMfgPlate = False
                        strRC = "No. of bevel symbols differs"
                        GoTo SetToNothing
                    End If
                ElseIf oAttribute1 Is Nothing And oAttribute2 Is Nothing Then
                    ' Both plate may have no bevels in this case, but no need to exit from the method, can proceed next comparison
                    IJMfgComparisonRule_CompareMfgPlate = False
                Else
                    IJMfgComparisonRule_CompareMfgPlate = False
                    strRC = "No. of bevel symbols differs"
                    GoTo SetToNothing
                End If
    
                Set oAttribute1 = Nothing
                Set oAttribute2 = Nothing
            Next
    
            'Compare Number of Margins
            '---------------------------
            For nCurNode = 0 To oEdgeNodeList1.Length - 1
                Set oElem = oEdgeNodeList1.Item(nCurNode)
                Set oAttribute1 = oElem.getAttributeNode("NO_MARGIN")
                Set oElem = Nothing
                Set oElem = oEdgeNodeList2.Item(nCurNode)
                Set oAttribute2 = oElem.getAttributeNode("NO_MARGIN")
                Set oElem = Nothing
                If Not oAttribute1 Is Nothing And Not oAttribute2 Is Nothing Then
                    ' check margin count
                    If Not oAttribute1.Value = oAttribute2.Value Then
                        IJMfgComparisonRule_CompareMfgPlate = False
                        strRC = "No. of margins differs"
                        GoTo SetToNothing
                    End If
                ElseIf oAttribute1 Is Nothing And oAttribute2 Is Nothing Then
                    ' Both plates may have no margins in this case, but no need to exit from the method, can proceed next comparison
                    IJMfgComparisonRule_CompareMfgPlate = False
                Else
                    IJMfgComparisonRule_CompareMfgPlate = False
                    strRC = "No. of margins differs"
                    GoTo SetToNothing
                End If
    
                Set oAttribute1 = Nothing
                Set oAttribute2 = Nothing
            Next
        End If
        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing
        
        ' Confirm no. of bevel are same are not
        Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_BEVEL")
        Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_BEVEL")
    
        If Not oEdgeNodeList1.Length = oEdgeNodeList2.Length Then
            IJMfgComparisonRule_CompareMfgPlate = False
            strRC = "No. of bevel differs"
            GoTo SetToNothing
        End If
    
        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing
        ' Confirm no. of margins are same are not
        Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_PART_MARGIN_INFO")
        Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_PART_MARGIN_INFO")
    
        If Not oEdgeNodeList1.Length = oEdgeNodeList2.Length Then
            IJMfgComparisonRule_CompareMfgPlate = False
            strRC = "No. of margins differs"
            GoTo SetToNothing
        End If
    
        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing
        
        
        ' Compare the given contours are identical or not
        '*************************************************
        Dim oContourInPlate1 As IJComplexString
        Dim oContourInPlate2 As IJComplexString
        Dim oContourColInPlate1 As IJMfgGeomCol2d
        Dim oContourColInPlate2 As IJMfgGeomCol2d
        
        Dim oGeom As IJMfgGeom2d
        Dim oMathGeom As GSCADMathGeom.MfgMathGeom

        Dim index As Integer, nCurContour As Integer
        ' Assign tolerance for checking contours
        Dim dContourTolerance As Double
        dContourTolerance = 0.001

        Dim oMfgParent As IJMfgGeomParent
        Dim oChildrenColl1 As IJDTargetObjectCol, oChildrenColl2 As IJDTargetObjectCol
        
        Set oMfgParent = oMfgPlatePart1
        Set oChildrenColl1 = oMfgParent.GetChildren
        Set oContourColInPlate1 = oChildrenColl1.Item(1)
        Set oChildrenColl1 = Nothing
        
        Set oMfgParent = oMfgPlatePart2
        Set oChildrenColl2 = oMfgParent.GetChildren
        Set oContourColInPlate2 = oChildrenColl2.Item(1)
        Set oChildrenColl2 = Nothing
        
        Dim oGeomHelper As MfgGeomHelper
        Set oGeomHelper = New MfgGeomHelper
        Dim lNoOfContoursInP1 As Long
        Dim lNoOfContoursInP2 As Long
        Dim oMfgPlate1Contours As IJElements
        Dim oMfgPlate2Contours As IJElements
        
        Set oMathGeom = New GSCADMathGeom.MfgMathGeom
        
        oGeomHelper.GetContours2d oContourColInPlate1, lNoOfContoursInP1, oMfgPlate1Contours
        oGeomHelper.GetContours2d oContourColInPlate2, lNoOfContoursInP2, oMfgPlate2Contours
       
        If Not lNoOfContoursInP1 = lNoOfContoursInP2 Then
           IJMfgComparisonRule_CompareMfgPlate = False
           strRC = "No.of contours differs"
           GoTo SetToNothing
        Else
            For nCurContour = 1 To lNoOfContoursInP1
                Set oContourInPlate1 = oMfgPlate1Contours.Item(nCurContour)
                Set oContourInPlate2 = oMfgPlate2Contours.Item(nCurContour)
                IJMfgComparisonRule_CompareMfgPlate = oMathGeom.AreContoursIdentical2d(oContourInPlate1, oContourInPlate2, dContourTolerance)
                If IJMfgComparisonRule_CompareMfgPlate = False Then
                    strRC = "Contours are not identical"
                    GoTo SetToNothing
                End If
                Set oContourInPlate1 = Nothing
                Set oContourInPlate2 = Nothing
            Next
        End If

        

'        ' If all Comparison succeed the IJMfgComparisonRule_CompareMfgPlate will be True and return successfully.
        strRC = "Comparison rule succeed"

SetToNothing:
        Set oContourColInPlate1 = Nothing
        Set oContourColInPlate2 = Nothing
        Set oPartInfo = Nothing
        Set oMathGeom = Nothing
        Set oAttribute1 = Nothing
        Set oAttribute2 = Nothing
        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing
        Set oXmlDom1 = Nothing
        Set oXmlDom2 = Nothing
        Set oMfgPlatePart1 = Nothing
        Set oMfgPlatePart2 = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, Method, , "SMCustomWarningMessages", 8002, , "RULES")
    GoTo SetToNothing:
End Function
'*********************************************************************************************************
'Name       :   CompareMfgProfile
'
'Class      :   CMfgComparisonRule      ;Interface  :   IJMfgComparisonRule
'
'Purpose    :   Comparison rule for Structural Manufacturing Profiles
'               - Two input profile part is given for comparison
'
'
'
'Rule ID    :   Default Rule
'
'Etc. .     :   !
'*********************************************************************************************************
Private Function IJMfgComparisonRule_CompareMfgProfile(ByVal pMfgProfile1 As Object, ByVal pMfgProfile2 As Object, strRC As String, Optional ByVal eCompareMode As GSCADStructMfgGlobals.StrMfgIsIdenticalCompareMode = CompareAll) As Boolean

Const Method = "_CompareMfgProfile"
On Error GoTo ErrorHandler
        'initialise with false
        IJMfgComparisonRule_CompareMfgProfile = False
        Dim oMfgProfilePart1 As IJMfgProfilePart
        Dim oMfgProfilePart2 As IJMfgProfilePart
        Dim oProfilePart As IJProfilePart
        Dim oProfile1CS As IJCrossSection
        Dim oProfile2CS As IJCrossSection


        Set oMfgProfilePart1 = pMfgProfile1
        Set oMfgProfilePart2 = pMfgProfile2

        'Check Profiles are straight Or Not using SD support
        '**********************************
        Dim oSdProfileSupport As IJProfilePartSupport
        Dim oSdPartSupport As IJPartSupport
        Dim oDisp As Object

        Set oProfilePart = oMfgProfilePart1.GetDetailedPart
        Set oDisp = oProfilePart
        Set oProfilePart = Nothing
        Set oSdPartSupport = New ProfilePartSupport
        Set oSdPartSupport.Part = oDisp
        Set oSdProfileSupport = oSdPartSupport
        ' Get the cross section of the first profile
        oSdProfileSupport.GetCrossSection oProfile1CS

        If oSdProfileSupport.IsTwisted = True Then
            IJMfgComparisonRule_CompareMfgProfile = False
            strRC = "First selected profile is twisted"
            GoTo SetToNothing
        End If

        If oSdProfileSupport.IsLinear = False Then
            IJMfgComparisonRule_CompareMfgProfile = False
            strRC = "Profile is not linear"
            GoTo SetToNothing
        End If
        
        Set oSdProfileSupport = Nothing
        Set oProfilePart = oMfgProfilePart2.GetDetailedPart
        Set oDisp = oProfilePart
        Set oProfilePart = Nothing
        Set oSdPartSupport = New ProfilePartSupport
        Set oSdPartSupport.Part = oDisp
        Set oSdProfileSupport = oSdPartSupport

        'Check the second profile twisted or not
        If oSdProfileSupport.IsTwisted = True Then
            IJMfgComparisonRule_CompareMfgProfile = False
            strRC = "Profile is twisted"
            GoTo SetToNothing
        End If
        'Check the second profile isLinear or not
        If oSdProfileSupport.IsLinear = False Then
            IJMfgComparisonRule_CompareMfgProfile = False
            strRC = "Profile is not linear"
            GoTo SetToNothing
        End If

        oSdProfileSupport.GetCrossSection oProfile2CS
        ' Check the cross section of the profile are same or not
        ' ******************************************************
        If Not oProfile1CS.Type = oProfile2CS.Type Then
            IJMfgComparisonRule_CompareMfgProfile = False
            strRC = "Cross section type differs"
            GoTo SetToNothing
        End If
        If Not oProfile1CS.Attributes.Count = oProfile2CS.Attributes.Count Then
             strRC = "Cross section count(size) differs"
            IJMfgComparisonRule_CompareMfgProfile = False
            GoTo SetToNothing
        End If
        Set oProfile1CS = Nothing
        Set oProfile2CS = Nothing
        Set oSdProfileSupport = Nothing


        'Get xml service
        '****************
        Dim oEdgeNodeList1 As IXMLDOMNodeList
        Dim oEdgeNodeList2 As IXMLDOMNodeList
        Dim oElem As IXMLDOMElement
        Dim oAttribute1 As IXMLDOMAttribute
        Dim oAttribute2 As IXMLDOMAttribute
        Dim oXmlDom1 As DOMDocument
        Dim oXmlDom2 As DOMDocument

        Set oXmlDom1 = oMfgProfilePart1.GetXMLData
        Set oXmlDom2 = oMfgProfilePart2.GetXMLData
        
        If oXmlDom1 Is Nothing Or oXmlDom2 Is Nothing Then
            IJMfgComparisonRule_CompareMfgProfile = False
            strRC = "XML data is not available"
            GoTo SetToNothing
        End If

        If oXmlDom1.getElementsByTagName("SMS_PLATE").Length <> oXmlDom1.getElementsByTagName("SMS_PLATE").Length Then
            IJMfgComparisonRule_CompareMfgProfile = False
            strRC = "XML data is not available"
            GoTo SetToNothing
        End If

        Dim i As Integer
        For i = 1 To oXmlDom1.getElementsByTagName("SMS_PLATE").Length

            'Check for number of openings in the Profile
            '*****************************************
            Dim nCurNode As Integer
            Dim nNodeCountInProfile1 As Integer
            Dim nNodeCountInProfile2 As Integer
            nNodeCountInProfile1 = 0
            nNodeCountInProfile2 = 0
            'Firslt Profile
            Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_EDGE")
            If Not oEdgeNodeList1 Is Nothing Then
                For nCurNode = 0 To oEdgeNodeList1.Length - 1
                    Set oElem = oEdgeNodeList1.Item(nCurNode)
                    Set oAttribute1 = oElem.getAttributeNode("TYPE")
                    ' Count no. of openings
                    If oAttribute1.Value = "Hole" Then
                        nNodeCountInProfile1 = nNodeCountInProfile1 + 1
                    End If
                    Set oElem = Nothing
                    Set oAttribute1 = Nothing
                Next
            End If
            Set oEdgeNodeList1 = Nothing
    
            'Second Profile
            Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_EDGE")
            If Not oEdgeNodeList2 Is Nothing Then
                For nCurNode = 0 To oEdgeNodeList2.Length - 1
                    Set oElem = oEdgeNodeList2.Item(nCurNode)
                    Set oAttribute1 = oElem.getAttributeNode("TYPE")
                    ' Count no. of openings
                    If oAttribute1.Value = "Hole" Then
                        nNodeCountInProfile2 = nNodeCountInProfile2 + 1
                    End If
                    Set oElem = Nothing
                    Set oAttribute1 = Nothing
                Next
            End If
            Set oEdgeNodeList2 = Nothing
    
    
            'Check the opening(hole) count
            If Not nNodeCountInProfile1 = nNodeCountInProfile2 Then
                IJMfgComparisonRule_CompareMfgProfile = False
                strRC = "No. of openings differs"
                GoTo SetToNothing
            End If
    
            'Check for number of markings
            '*****************************
            Dim nNoOfMarkingsInProfile1 As Integer
            Dim nNoOfMarkingsInProfile2 As Integer
    
            ' Count marking in Profile1
            Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_MARKING")
            Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_MARKING")
            If Not oEdgeNodeList1 Is Nothing Or Not oEdgeNodeList2 Is Nothing Then
                nNoOfMarkingsInProfile1 = oEdgeNodeList1.Length
                Set oEdgeNodeList1 = Nothing
    
                ' Count marking in Profile2
                nNoOfMarkingsInProfile2 = oEdgeNodeList2.Length
                Set oEdgeNodeList2 = Nothing
                ' Check number of markings are same in Profile1 and Profile2
                If Not nNoOfMarkingsInProfile1 = nNoOfMarkingsInProfile2 Then
                    IJMfgComparisonRule_CompareMfgProfile = False
                    strRC = "No. of marking differs"
                    GoTo SetToNothing
                End If
    
                Set oEdgeNodeList1 = Nothing
                Set oEdgeNodeList2 = Nothing
    
                'Check for each marking name in both Profiles, if number of markings are same
                '**************************************************************************
    
                Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_MARKING")
                Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_MARKING")
    
                For nCurNode = 0 To oEdgeNodeList1.Length - 1
                    Set oElem = oEdgeNodeList1.Item(nCurNode)
                    Set oAttribute1 = oElem.getAttributeNode("MARKING_NAME")
                    Set oElem = Nothing
                    Set oElem = oEdgeNodeList2.Item(nCurNode)
                    Set oAttribute2 = oElem.getAttributeNode("MARKING_NAME")
                    Set oElem = Nothing
                    If Not oAttribute1 Is Nothing And Not oAttribute2 Is Nothing Then
                        If Not Trim(oAttribute1.Value) = "" Or Not Trim(oAttribute2.Value) = "" Then
                            If Not oAttribute1.Value = oAttribute2.Value Then
                                IJMfgComparisonRule_CompareMfgProfile = False
                                strRC = "Marking name differs"
                                GoTo SetToNothing
                            End If
                        End If
                    End If
                    Set oAttribute1 = Nothing
                    Set oAttribute2 = Nothing
                Next
    
            End If
    
            Set oEdgeNodeList1 = Nothing
            Set oEdgeNodeList2 = Nothing
    
          'Check for Material type
            '***********************
            Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_PART_INFO")
            Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_PART_INFO")
            If Not oEdgeNodeList1 Is Nothing Or Not oEdgeNodeList2 Is Nothing Then
                Set oElem = oEdgeNodeList1.Item(0)
    
                Set oAttribute1 = oElem.getAttributeNode("MATERIAL_TYPE")
                Set oElem = Nothing
                Set oElem = oEdgeNodeList2.Item(0)
                Set oAttribute2 = oElem.getAttributeNode("MATERIAL_TYPE")
                Set oElem = Nothing
    
                If Not oAttribute1.Value = oAttribute2.Value Then
                    IJMfgComparisonRule_CompareMfgProfile = False
                    strRC = "Material type differs"
                    GoTo SetToNothing
                End If
                Set oAttribute1 = Nothing
                Set oAttribute2 = Nothing
            End If
            Set oEdgeNodeList1 = Nothing
            Set oEdgeNodeList2 = Nothing
    
            'Check for material grade
            '------------------------
            Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_PART_INFO")
            Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_PART_INFO")
            Set oElem = oEdgeNodeList1.Item(0)
            Set oAttribute1 = oElem.getAttributeNode("MATERIAL_GRADE")
            Set oElem = Nothing
            Set oElem = oEdgeNodeList2.Item(0)
            Set oAttribute2 = oElem.getAttributeNode("MATERIAL_GRADE")
            Set oElem = Nothing
    
            If Not oAttribute1.Value = oAttribute2.Value Then
                IJMfgComparisonRule_CompareMfgProfile = False
                strRC = "Material grade differs"
                GoTo SetToNothing
            End If
            Set oAttribute1 = Nothing
            Set oAttribute2 = Nothing
            Set oEdgeNodeList1 = Nothing
            Set oEdgeNodeList2 = Nothing
    
             'Before checking no .of bevels, & no. of margins - first compare no. of edges
            '----------------------------------------------------------------------------
            Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_EDGE")
            Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_EDGE")
    
            If Not oEdgeNodeList1.Length = oEdgeNodeList2.Length Then
                IJMfgComparisonRule_CompareMfgProfile = False
                strRC = "No. of edges differs"
                GoTo SetToNothing
            Else
                'Check for Number of bevel symbols in the Profile
                '----------------------------------------------
                For nCurNode = 0 To oEdgeNodeList1.Length - 1
                    Set oElem = oEdgeNodeList1.Item(nCurNode)
                    Set oAttribute1 = oElem.getAttributeNode("NO_BEVEL")
                    Set oElem = Nothing
                    Set oElem = oEdgeNodeList2.Item(nCurNode)
                    Set oAttribute2 = oElem.getAttributeNode("NO_BEVEL")
                    Set oElem = Nothing
                    If Not oAttribute1 Is Nothing And Not oAttribute2 Is Nothing Then
                        ' check bevel count
                        If Not oAttribute1.Value = oAttribute2.Value Then
                            IJMfgComparisonRule_CompareMfgProfile = False
                            strRC = "No. of bevel symbol differs"
                            GoTo SetToNothing
                        End If
                    ElseIf oAttribute1 Is Nothing And oAttribute2 Is Nothing Then
                        ' In this case both profile may have no bevels, but no need to exit from the method, can proceed next comparison
                        IJMfgComparisonRule_CompareMfgProfile = False
                    Else
                        IJMfgComparisonRule_CompareMfgProfile = False
                        strRC = "No. of bevel symbol differs"
                        GoTo SetToNothing
                    End If
        
                    Set oAttribute1 = Nothing
                    Set oAttribute2 = Nothing
                Next
        
                'Compare Number of Margins
                '---------------------------
                For nCurNode = 0 To oEdgeNodeList1.Length - 1
                    Set oElem = oEdgeNodeList1.Item(nCurNode)
                    Set oAttribute1 = oElem.getAttributeNode("NO_MARGIN")
                    Set oElem = Nothing
                    Set oElem = oEdgeNodeList2.Item(nCurNode)
                    Set oAttribute2 = oElem.getAttributeNode("NO_MARGIN")
                    Set oElem = Nothing
                    If Not oAttribute1 Is Nothing And Not oAttribute2 Is Nothing Then
                        ' check no. of margins
                        If Not oAttribute1.Value = oAttribute2.Value Then
                            IJMfgComparisonRule_CompareMfgProfile = False
                            strRC = "No. of margins differs"
                            GoTo SetToNothing
                        End If
                    ElseIf oAttribute1 Is Nothing And oAttribute2 Is Nothing Then
                        ' Both profile may have no margins in this case, but no need to exit from the method, can proceed next comparison
                        IJMfgComparisonRule_CompareMfgProfile = False
                    Else
                        IJMfgComparisonRule_CompareMfgProfile = False
                        strRC = "No. of margins differs"
                        GoTo SetToNothing
                    End If
                    Set oAttribute1 = Nothing
                    Set oAttribute2 = Nothing
                Next
            End If
            Set oEdgeNodeList1 = Nothing
            Set oEdgeNodeList2 = Nothing
    
    
            ' Confirm no. of bevel are same are not
            Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_BEVEL")
            Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_BEVEL")
        
            If Not oEdgeNodeList1.Length = oEdgeNodeList2.Length Then
                IJMfgComparisonRule_CompareMfgProfile = False
                strRC = "No. of bevel differs"
                GoTo SetToNothing
            End If
        
            Set oEdgeNodeList1 = Nothing
            Set oEdgeNodeList2 = Nothing
            ' Confirm no. of margins are same are not
            Set oEdgeNodeList1 = oXmlDom1.getElementsByTagName("SMS_PART_MARGIN_INFO")
            Set oEdgeNodeList2 = oXmlDom2.getElementsByTagName("SMS_PART_MARGIN_INFO")
        
            If Not oEdgeNodeList1.Length = oEdgeNodeList2.Length Then
                IJMfgComparisonRule_CompareMfgProfile = False
                strRC = "No. of margins differs"
                GoTo SetToNothing
            End If
        
            Set oEdgeNodeList1 = Nothing
            Set oEdgeNodeList2 = Nothing

        Next
        
        
        ' Workaround code for checking the  profile length
        ' Comparison of Mfg profile part length
        If Not oMfgProfilePart1.Length = oMfgProfilePart2.Length Then
                IJMfgComparisonRule_CompareMfgProfile = False
                strRC = "Length of the profile differs"
                GoTo SetToNothing
        End If
        
        Dim dPart1Length As Double
        Dim dPart2Length As Double
        Set oProfilePart = oMfgProfilePart1.GetDetailedPart
        dPart1Length = oProfilePart.ProfileLength
        
        Set oProfilePart = Nothing
        Set oProfilePart = oMfgProfilePart2.GetDetailedPart
        dPart2Length = oProfilePart.ProfileLength
       
        'Compare the detailed profile part length
        If Not dPart1Length = dPart2Length Then
                IJMfgComparisonRule_CompareMfgProfile = False
                strRC = "Detailed profile part length differs"
                GoTo SetToNothing
        End If
        

        
'/////////////////////////////////////////////////////////////////////////////////////////////
' Code implemented to get the profile length between end cuts along with the landing curve
' Needs to be modified and verified
'/////////Profile length calculation based on landing curve end cuts///////////////////////

'            Dim oLandCrv As IJLandCurve
'            Dim oWireBody As IJWireBody
'            Dim oSGOWireBodyUtilities As SGOWireBodyUtilities
'            Dim oLandCrvdefPointsColl As Collection
'
'
'            Dim m_oProfileUtils As IJProfileDefinition ' Plate Utilities
'
'            ' Set the cross section
'            Dim pIJProfileAttributes As IJProfileAttributes
'            Set m_oProfileUtils = New ProfileUtils ' Create the Profile utilties
'
'            Dim oProfPart As IJProfilePart
'            Set oProfPart = oMfgProfilePart1.GetDetailedPart
'
'            Set pIJProfileAttributes = m_oProfileUtils
'
'            MsgBox "Landing curve 1 & 2"
'
'            ' Getting Landing Curve from Profile
'            pIJProfileAttributes.GetLandingCurveFromProfile oProfPart, oLandCrv
'
'            ' Getting Wire Body from Landing Curve
'            Set oWireBody1 = oLandCrv
'            oLandCrv = Nothing
'
'            Set oProfPart = oMfgProfilePart2.GetDetailedPart
'
'            MsgBox "oProfPart.ProfileLength : " & oProfPart.ProfileLength
'
'            If oProfPart Is Nothing Then
'            MsgBox "No prof"
'            End If
'            pIJProfileAttributes.GetLandingCurveFromProfile oProfPart, oLandCrv
'            Set oWireBody2 = oLandCrv
'            ' Create Object of Type SGOWireBodyUtilities
'
'            Set oSGOWireBodyUtilities = New SGOWireBodyUtilities
'
'                Dim dblDistance1 As Double
'                Dim dblDistance2 As Double
'                Dim intDefPointsCount As Integer
'                Dim oPositon1 As IJDPosition
'                Dim oPositon2 As IJDPosition
'
'            MsgBox "Curve def points"
'            ' Getting the LandCurve definition Point(s) collection
'            Set oLandCrvdefPointsColl = oSGOWireBodyUtilities.GetCurveDefinitionPoints(oWireBody1)
'
'             ' Get distance for profile part 1
'            intDefPointsCount = oLandCrvdefPointsColl.Count
'
'            If intDefPointsCount >= 2 Then
'                 For intDefPointsCount = 1 To 2
'                     If intDefPointsCount = 1 Then Set oPositon1 = oLandCrvdefPointsColl.Item(intDefPointsCount)
'                     If intDefPointsCount = 2 Then Set oPositon2 = oLandCrvdefPointsColl.Item(intDefPointsCount)
'                     If intDefPointsCount = 2 Then
'                         Exit For
'                     End If
'                 Next
'             End If
'
'            Call oSGOWireBodyUtilities.GetDistanceAlongCurve(oWireBody1, oPositon1, oPositon2, dblDistance1)
'
'             ' Get distance for profile part 2
'            Set oLandCrvdefPointsColl = oSGOWireBodyUtilities.GetCurveDefinitionPoints(oWireBody2)
'
'             intDefPointsCount = oLandCrvdefPointsColl.Count
'
'             If intDefPointsCount >= 2 Then
'                 For intDefPointsCount = 1 To 2
'                     If intDefPointsCount = 1 Then Set oPositon1 = oLandCrvdefPointsColl.Item(intDefPointsCount)
'                     If intDefPointsCount = 2 Then Set oPositon2 = oLandCrvdefPointsColl.Item(intDefPointsCount)
'                     If intDefPointsCount = 2 Then
'                         Exit For
'                     End If
'                 Next
'             End If
'
'             Call oSGOWireBodyUtilities.GetDistanceAlongCurve(oWireBody2, oPositon1, oPositon2, dblDistance2)
'
'
'
'             Set oLandCrvdefPointsColl = Nothing
'             Set oPositon1 = Nothing
'             Set oPositon2 = Nothing
'             Set pIJProfileAttributes = Nothing
'             Set oWireBody = Nothing
'             Set oLandCrv = Nothing
'             Set oLandCrvdefPointsColl = Nothing
'             Set oSGOWireBodyUtilities = Nothing
'
'
'
'                 ' Distance needs to be
'                 If Not dblDistance1 = dblDistance2 Then
'                         IJMfgComparisonRule_CompareMfgProfile = False
'                         strRC = "Length based on end-cuts differs"
'                         GoTo SetToNothing
'                 End If
'


'//////////////////////////////////////////////////////////////////////////////////////////
' Code implemented for comparing identical: getMfgGeomColFromAE() method needs to be
' implemented for IJMfgProfilePart in order to get the contours of the profile.
'//////////////////////////////////////////////////////////////////////////////////////////
'        ' Compare the given contours are identical or not
'        '*************************************************
'        Dim oContourInProfile1 As IJComplexString
'        Dim oContourInProfile2 As IJComplexString
'        Dim oContourColInProfile1 As IJMfgGeomCol2d
'        Dim oContourColInProfile2 As IJMfgGeomCol2d
'        Dim oGeom As IJMfgGeom2d
'        Dim oMathGeom As GSCADMathGeom.MfgMathGeom
'        Set oMathGeom = New GSCADMathGeom.MfgMathGeom
'
'        ' Assign tolerance for checking contours
'        Dim dContourTolerance As Double
'        dContourTolerance = 0.1
'
'        Dim oPart1Geom2d As IUnknown
'        Dim oPart2Geom2d As IUnknown
'
'MsgBox "getMfgGeomColFromAE is not available for profile"
'' getMfgGeomColFromAE is not available for profile
'        Set oContourColInProfile1 = oMfgProfilePart1.getMfgGeomColFromAE("MfgProfileDevelopmentOutput_ORIG")
'        Set oContourColInProfile2 = oMfgProfilePart2.getMfgGeomColFromAE("MfgProfileDevelopmentOutput_ORIG")
'
'
'
'        Dim index As Integer, nCurContour As Integer
'
'        If Not oContourColInProfile1.GetColcount = oContourColInProfile2.GetColcount Then
'            IJMfgComparisonRule_CompareMfgProfile = False
'               'm_oCPNDebugLog.LogMsg "Profile contours differs."
'            GoTo SetToNothing
'        End If
'
'        index = oContourColInProfile1.GetColcount
'        If Not index = 0 Then
'        For nCurContour = 0 To index
'            Set oGeom = oContourColInProfile1.GetGeometry(nCurContour)
'            Set oContourInProfile1 = oGeom
'            Set oGeom = Nothing
'            Set oGeom = oContourColInProfile2.GetGeometry(nCurContour)
'            Set oContourInProfile2 = oGeom
'            Set oGeom = Nothing
'            IJMfgComparisonRule_CompareMfgProfile = oMathGeom.AreContoursIdentical2d(oContourInProfile1, oContourInProfile2, dContourTolerance)
'            If IJMfgComparisonRule_CompareMfgProfile = False Then
'                'm_oCPNDebugLog.LogMsg "Profile " & m_nProfileCount & " contours are not identical with the first selected profile part"
'                GoTo SetToNothing
'            End If
'        Next nCurContour
'        End If
'
'/////////////////////////////////////////////////////////////////////////////////////
        
        ' If all Comparison succeed the IJMfgComparisonRule_CompareMfgProfile will be True and return successfully.
        IJMfgComparisonRule_CompareMfgProfile = True
        strRC = "Comparison rule succeed"



SetToNothing:
        Set oProfilePart = Nothing
        Set oDisp = Nothing
        Set oProfilePart = Nothing
        Set oSdPartSupport = Nothing
        Set oSdProfileSupport = Nothing

        Set oContourColInProfile1 = Nothing
        Set oContourColInProfile2 = Nothing
        Set oMathGeom = Nothing
        Set oAttribute1 = Nothing
        Set oAttribute2 = Nothing
        Set oEdgeNodeList1 = Nothing
        Set oEdgeNodeList2 = Nothing
        Set oXmlDom1 = Nothing
        Set oXmlDom2 = Nothing
        Set oMfgProfilePart1 = Nothing
        Set oMfgProfilePart2 = Nothing
        Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, Method, , "SMCustomWarningMessages", 8003, , "RULES")
    GoTo SetToNothing:
End Function

 
