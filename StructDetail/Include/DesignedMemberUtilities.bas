Attribute VB_Name = "DesignedMemberUtilities"
'*******************************************************************
'
'Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'File : DesignedMemberUtilities.bas
'
'Author : D.A. Trent
'
'Description :
'   Utilites for processing DesignedMember ("Built Up") objects
'
'History:
'*****************************************************************************
Option Explicit
Private Const MODULE = "StructDetail\Data\Include\DesignedMemberUtilities"
'
'


'*************************************************************************
'Function
'IsDesignedMemberDetailed
'
'Abstract
'   Given the Assembly Connection (IJAppConnection Interface)
'   Initialize the BoundedConnection and Bounding Connection Data structures
'
'input
'   oMemberPart - Designed Member to check
'
'Return
'       True  - all children of DesignedMember system are detailed Parts
'               Plate Parts (MemberPartPrismatic, Profile Systems, etc)
'       False - at least one child is NOT valid Detailed Part
'
'Exceptions
'
'***************************************************************************
Public Function IsDesignedMemberDetailed(oMemberPart As ISPSMemberPartCommon) As Boolean
Const METHOD = "::IsDesignedMemberDetailed"
    
    Dim sMsg As String
    On Error GoTo ErrorHandler
    
    Dim iIndex As Long
    Dim nObjects As Long
    
    Dim bDetailed As Boolean
    
    Dim oObject As Object
    Dim oMemberObjects As IJDMemberObjects
    Dim oDesignedMember As ISPSDesignedMember
    Dim oSmartOccurrence As IJSmartOccurrence
    
    nObjects = 0
    bDetailed = False
    IsDesignedMemberDetailed = False
    
    If TypeOf oMemberPart Is IJDMemberObjects Then
        Set oMemberObjects = oMemberPart
        nObjects = oMemberObjects.Count
    End If
    
    For iIndex = 1 To nObjects
        bDetailed = True
        
        If oMemberObjects.Item(iIndex) Is Nothing Then
        Else
            Set oObject = oMemberObjects.Item(iIndex)
            If TypeOf oObject Is IJPlate Then
                ' Check if root and all leaf systems are detailed
                bDetailed = IsDetailedPlateParts(oObject)
            
            ElseIf TypeOf oObject Is IJProfile Then
                ' Check if root and all leaf systems are detailed
                bDetailed = IsDetailedProfileParts(oObject)
                
            ElseIf TypeOf oObject Is ISPSDesignedMember Then
                ' Do we expect nested DesignedMembers
                
            ElseIf TypeOf oObject Is ISPSMemberPartPrismatic Then
            
            ElseIf TypeOf oObject Is ISPSMemberPartCommon Then
            
            Else
                ' Unknown type, assume it is valid
            End If
        End If
        
                   
        If Not bDetailed Then
            Exit For
        End If
        
    Next iIndex
    
    IsDesignedMemberDetailed = bDetailed

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'IsDetailedPlateParts
'
'Abstract
'   Check if given IJPlate and all of its hildren are Detailed
'
'input
'   oPlate - Plate System/Part to check
'
'Return
'       True  - all children of Plate system are detailed Parts
'               Plate Parts (MemberPartPrismatic, Profile Systems, etc)
'       False - at least one child is NOT valid Detailed Part
'
'Exceptions
'
'***************************************************************************
Public Function IsDetailedPlateParts(oPlate As IJPlate) As Boolean
Const METHOD = "::IsDetailedPlateParts"
    
    Dim sMsg As String
    On Error GoTo ErrorHandler
    
    Dim iIndex As Long
    Dim nObjects As Long
    Dim lResultType As Long
            
    Dim bDetailed As Boolean
    
    Dim oObject As Object
    Dim oChildren As IJDObjectCollection
    Dim oDesignParent As IJDesignParent
    Dim oSDNameRulesUtilHelper As SDNameRulesUtilHelper
    
    nObjects = 0
    bDetailed = False
    IsDetailedPlateParts = False
    
    If TypeOf oPlate Is IJDesignParent Then
        Set oDesignParent = oPlate
        oDesignParent.GetChildren oChildren
        If oChildren Is Nothing Then
            nObjects = 0
        Else
            nObjects = oChildren.Count
        End If
    
    End If
    
    If nObjects > 0 Then
        Set oSDNameRulesUtilHelper = New SDNameRulesUtilHelper
        
        iIndex = 0
        For Each oObject In oChildren
            bDetailed = True
            iIndex = iIndex + 1
            
            If TypeOf oObject Is IJPlatePart Then
                lResultType = oSDNameRulesUtilHelper.ResultType(oObject)
                'lResultType: 32768 = 0x8000 = STRUCT_RESULT_TYPE_LIGHT
                If (lResultType And 32768) = 32768 Then
                    bDetailed = False
                    Exit For
                End If
    
            ElseIf TypeOf oObject Is IJPlateSystem Then
                bDetailed = IsDetailedPlateParts(oObject)
                If Not bDetailed Then
                    Exit For
                End If
                
            ElseIf TypeOf oObject Is IJProfile Then
                bDetailed = IsDetailedProfileParts(oObject)
                If Not bDetailed Then
                    Exit For
                End If
            
            Else
                ' Unknown type, assume it is valid
            End If
        Next oObject
    
    End If
    
    IsDetailedPlateParts = bDetailed

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'*************************************************************************
'Function
'IsDetailedProfileParts
'
'Abstract
'   Check if given IJProfile and all of its hildren are Detailed
'
'input
'   oProfile - Profile System/Part to check
'
'Return
'       True  - all children of Profile system are detailed Parts
'               Profile Parts (MemberPartPrismatic, Profile Systems, etc)
'       False - at least one child is NOT valid Detailed Part
'
'Exceptions
'
'***************************************************************************
Public Function IsDetailedProfileParts(oProfile As IJProfile) As Boolean
Const METHOD = "::IsDetailedProfileParts"
    
    Dim sMsg As String
    On Error GoTo ErrorHandler
    
    Dim iIndex As Long
    Dim nObjects As Long
    Dim lResultType As Long
            
    Dim bDetailed As Boolean
    
    Dim oObject As Object
    Dim oChildren As IJDObjectCollection
    Dim oDesignParent As IJDesignParent
    Dim oSDNameRulesUtilHelper As SDNameRulesUtilHelper
    
    nObjects = 0
    bDetailed = False
    IsDetailedProfileParts = False
    
    If TypeOf oProfile Is IJDesignParent Then
        Set oDesignParent = oProfile
        oDesignParent.GetChildren oChildren
        If oChildren Is Nothing Then
            nObjects = 0
        Else
            nObjects = oChildren.Count
        End If
        
    End If
    
    If nObjects > 0 Then
        Set oSDNameRulesUtilHelper = New SDNameRulesUtilHelper
        
        iIndex = 0
        For Each oObject In oChildren
            bDetailed = True
            iIndex = iIndex + 1
            
            If TypeOf oObject Is IJProfilePart Then
                lResultType = oSDNameRulesUtilHelper.ResultType(oObject)
                'lResultType: 32768 = 0x8000 = STRUCT_RESULT_TYPE_LIGHT
                If (lResultType And 32768) = 32768 Then
                    bDetailed = False
                    Exit For
                End If

            ElseIf TypeOf oObject Is IJProfile Then
                bDetailed = IsDetailedProfileParts(oObject)
                If Not bDetailed Then
                    Exit For
                End If
                
            ElseIf TypeOf oObject Is IJPlateSystem Then
                bDetailed = IsDetailedPlateParts(oObject)
                If Not bDetailed Then
                    Exit For
                End If
            
            Else
                ' Unknown type, assume it is valid
            End If
        Next oObject
        
    End If
    
    IsDetailedProfileParts = bDetailed

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'***********************************************************************
' METHOD:  IsPortFromBuiltUpMember
'
' DESCRIPTION:  Compare the normal of the molded surface of two
'               connected plate parts. If they point to the same
'               direction, then return true. If not, return false.
'***********************************************************************
Public Sub IsFromBuiltUpMember(oObject As Object, _
                               bFromBuiltUp As Boolean, _
                               Optional oBuiltupMember As ISPSDesignedMember)
Const METHOD = "::IsFromBuiltUpMember"
On Error GoTo ErrorHandler
    
    bFromBuiltUp = False
    
    Dim oPort As IJPort
    Dim oConnectable As Object
    Dim oParentObject As Object
    Dim oDesignChild As IJDesignChild
    Dim oPlateSystem As IJPlateSystem
    Dim oDesignParent As IJDesignParent
    
    ' Check if type of Object passed in
    If oObject Is Nothing Then
        Exit Sub
        
    ElseIf TypeOf oObject Is IJPort Then
        ' a IJPort was passed in: get its Connectable
        Set oPort = oObject
        Set oConnectable = oPort.Connectable
        Set oPort = Nothing
        
        If Not TypeOf oConnectable Is IJPlatePart Then
            Exit Sub
        End If
        
    ElseIf TypeOf oObject Is IJPlatePart Then
        ' a IJPlatePart was passed in
        Set oConnectable = oObject
    
    ElseIf TypeOf oObject Is IJPlateSystem Then
        ' a IJPlateSystem was passed in (may be Root or Leaf System)
        Set oParentObject = oObject
    
    Else
        Exit Sub
    End If
    
    ' If Have a Connectable ( a IJPlatePart from Port or passed in)
    ' Get the Plate Part's Parent Object
    If Not oConnectable Is Nothing Then
        If TypeOf oConnectable Is IJDesignChild Then
            Set oDesignChild = oConnectable
            Set oParentObject = oDesignChild.GetParent
            Set oDesignChild = Nothing
        End If
    End If
    
    ' Verify have a valid IJPlateSystem either from the Port/IJPlatePate or passed in
    If oParentObject Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oParentObject Is IJPlateSystem Then
        Exit Sub
    
    ElseIf Not TypeOf oParentObject Is IJDesignChild Then
        Exit Sub
    End If
    
    ' Check if current IJPlateSystem is from a BuiltUP:
    ' (a Leaf Plate System is never)
    Set oPlateSystem = oParentObject
    bFromBuiltUp = oPlateSystem.IsBuiltupPlateSystem
    If bFromBuiltUp Then
        Set oBuiltupMember = oPlateSystem.ParentBuiltup
        Exit Sub
    End If
    
    ' The Above PlateSystem may be a Leaf Plate System
    ' Get the Above Plate Systems's Parent object
    Set oDesignChild = oParentObject
    Set oParentObject = oDesignChild.GetParent
    Set oDesignChild = Nothing
            
    ' Check if the Plate System's Parent object is IJPlateSystem
    If TypeOf oParentObject Is IJPlateSystem Then
        Set oPlateSystem = oParentObject
        bFromBuiltUp = oPlateSystem.IsBuiltupPlateSystem
                
        If bFromBuiltUp Then
            Set oBuiltupMember = oPlateSystem.ParentBuiltup
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub


