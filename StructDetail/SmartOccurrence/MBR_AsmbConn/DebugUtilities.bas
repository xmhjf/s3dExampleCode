Attribute VB_Name = "DebugUtilities"
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : DebugUtilities.bas
'
'Author : D.A. Trent
'
'Description :
'   Utilites to assist in Debugging MemberAssyConn Custom Rules
'
'History:
'*****************************************************************************
Option Explicit
Private Const MODULE = "StructDetail\Data\SmartOccurrence\MemberAssyConn\DebugUtilities"
Public Const bSM_trace = False

'

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub zvMsgBox(vText As Variant, _
                   Optional sDumArg1 As String, Optional sTitle As String)
On Error Resume Next
Dim sText As String
    sText = vText
    zMsgBox sText, sDumArg1, sTitle

End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub zMsgBox(sText As String, _
                   Optional sDumArg1 As String, Optional sTitle As String)
On Error Resume Next

Dim iFileNumber
Dim sFileName As String

'$$$Exit Sub
'$$$Debug $$$    MsgBox sText, , sTitle

    iFileNumber = FreeFile
    sFileName = "C:\Temp\TraceFile.txt"
    Open sFileName For Append Shared As #iFileNumber
    
    If Len(Trim(sTitle)) > 0 Then
        Write #iFileNumber, sTitle
    End If
    
    Write #iFileNumber, sText
    Close #iFileNumber
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Debug_Matrix(sText As String, oMatrix As IJDT4x4) As String
Const METHOD = "::Debug_Matrix"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    If oMatrix Is Nothing Then
        sMsg = sText & " ... Matrix is Nothing"
    Else
        sMsg = sText & _
                vbCrLf & _
                "oMatrix(u):" & Format(oMatrix.IndexValue(0), "0.0000") & _
                " , " & Format(oMatrix.IndexValue(1), "0.0000") & _
                " , " & Format(oMatrix.IndexValue(2), "0.0000") & _
                vbCrLf & _
                "oMatrix(v):" & Format(oMatrix.IndexValue(4), "0.0000") & _
                " , " & Format(oMatrix.IndexValue(5), "0.0000") & _
                " , " & Format(oMatrix.IndexValue(6), "0.0000") & _
                vbCrLf & _
                "oMatrix(w):" & Format(oMatrix.IndexValue(8), "0.0000") & _
                " , " & Format(oMatrix.IndexValue(9), "0.0000") & _
                " , " & Format(oMatrix.IndexValue(10), "0.0000") & _
                vbCrLf & _
                "oMatrix(r):" & Format(oMatrix.IndexValue(12), "0.0000") & _
                " , " & Format(oMatrix.IndexValue(13), "0.0000") & _
                " , " & Format(oMatrix.IndexValue(14), "0.0000")
    End If

    If Len(Trim(sText)) > 0 Then
        zMsgBox sMsg
    End If

    Debug_Matrix = sMsg
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Debug_Vector(sText As String, oVector As IJDVector) As String
Const METHOD = "::Debug_Vector"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = sText & _
            " :" & Format(oVector.x, "0.0000") & _
            " , " & Format(oVector.y, "0.0000") & _
            " , " & Format(oVector.z, "0.0000")
    
    If Len(Trim(sText)) > 0 Then
        zMsgBox sMsg
    End If
    
    Debug_Vector = sMsg

    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Debug_Point(sText As String, oPoint As IJPoint) As String
Const METHOD = "::Debug_Point"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    
    oPoint.GetPoint dx, dy, dz
    sMsg = sText & _
            " : " & Format(dx, "0.0000") & _
            " , " & Format(dy, "0.0000") & _
            " , " & Format(dz, "0.0000")

    If Len(Trim(sText)) > 0 Then
        zMsgBox sMsg
    End If
    
    Debug_Point = sMsg
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Debug_Position(sText As String, oPosition As IJDPosition) As String
Const METHOD = "::Debug_Vector"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = sText & _
            " : " & Format(oPosition.x, "0.0000") & _
            " , " & Format(oPosition.y, "0.0000") & _
            " , " & Format(oPosition.z, "0.0000")

    If Len(Trim(sText)) > 0 Then
        zMsgBox sMsg
    End If
    
    Debug_Position = sMsg
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Debug_MemberConnectionData(sText As String, oMemberData As MemberConnectionData) As String
Const METHOD = "::Debug_MemberConnectionData"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oNamedItem As IJNamedItem
    
    sMsg = sText & vbCrLf & "   ...MemberPart:"
    If oMemberData.MemberPart Is Nothing Then
        sMsg = sMsg & "oMemberData.MemberPart Is Nothing"
    Else
        If TypeOf oMemberData.MemberPart Is IJNamedItem Then
            Set oNamedItem = oMemberData.MemberPart
            sMsg = sMsg & oNamedItem.Name
        Else
            sMsg = sMsg & "??"
        End If
        
        sMsg = sMsg & vbCrLf & "   ...ePortId = "
        If oMemberData.ePortId = SPSMemberAxisAlong Then
            sMsg = sMsg & "SPSMemberAxisAlong"
        ElseIf oMemberData.ePortId = SPSMemberAxisEnd Then
            sMsg = sMsg & " SPSMemberAxisEnd"
        ElseIf oMemberData.ePortId = SPSMemberAxisStart Then
            sMsg = sMsg & "SPSMemberAxisStart"
        Else
            sMsg = sMsg & "???"
        End If
    End If
    
    sMsg = sMsg & vbCrLf & Debug_Matrix("", oMemberData.Matrix)
    If Len(Trim(sText)) > 0 Then
        zMsgBox sMsg
    End If
    
    Debug_MemberConnectionData = sMsg

    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

' =========================================================================
' =========================================================================
' =========================================================================
'Public Sub Debug_InputConnPorts(oAppConnection As IJAppConnection)
'Const METHOD = "::Debug_Matrix"
'    On Error GoTo ErrorHandler
'    Dim sMsg As String
'
'    Dim iIndex As Long
'    Dim lCount As Long
'
'    Dim sText As String
'
'    Dim dx As Double
'    Dim dy As Double
'    Dim dz As Double
'
'    Dim bColinear As Boolean
'    Dim bEndToEnd As Boolean
'    Dim bRightAngle As Boolean
'
'    Dim oPoint As IJPoint
'    Dim oVector As IJDVector
'    Dim oMatrix As IJDT4x4
'    Dim oAxisCurve As IJCurve
'    Dim oNamedItem As IJNamedItem
'    Dim oElements_Ports As IJElements
'
'    Dim oPortObj As Object
'
'    Dim oBounded_Point As IJPoint
'    Dim oBounded_AxisPort As ISPSSplitAxisPort
'    Dim oBounded_AxisCurve As IJCurve
'    Dim oBounded_MemberPart As ISPSMemberPartPrismatic
'
'    Dim oBounding_Point As IJPoint
'    Dim oBounding_AxisPort As ISPSSplitAxisPort
'    Dim oBounding_AxisCurve As IJCurve
'    Dim oBounding_MemberPart As ISPSMemberPartPrismatic
'
'    Dim ePortId As SPSMemberAxisPortIndex
'    Dim eBounded_PortId As SPSMemberAxisPortIndex
'    Dim eBounding_PortId As SPSMemberAxisPortIndex
'
'    Dim oAxisEndPort As ISPSSplitAxisEndPort
'    Dim oAxisAlongPort As ISPSSplitAxisAlongPort
'    Dim oSplitAxisPort As ISPSSplitAxisPort
'
'    Dim oPort As IJPort
'    Dim oMemberPart As ISPSMemberPartPrismatic
'
'    On Error GoTo ErrorHandler
'    sMsg = "Unknown Error"
'    sText = "FrmAsmConnSel::SelectorLogic"
'    sText = sText & vbCrLf & Debug_ObjectName(oAppConnection, True)
'
'    oAppConnection.enumPorts oElements_Ports
'    lCount = oElements_Ports.Count
'    sText = sText & vbCrLf & "oElements_Ports.Count = " & Str(lCount)
'
'    If lCount = 2 Then
'        For iIndex = 1 To lCount
'            Set oPortObj = oElements_Ports.Item(iIndex)
'            If TypeOf oPortObj Is ISPSSplitAxisPort Then
'                Set oSplitAxisPort = oPortObj
'                ePortId = oSplitAxisPort.PortIndex
'                Set oMemberPart = oSplitAxisPort.Part
'                If TypeOf oPortObj Is IJPort Then
'                    Set oPort = oSplitAxisPort.Port
'                    If oPort Is Nothing Then
'                    ElseIf TypeOf oPort.Geometry Is IJPoint Then
'                        eBounded_PortId = ePortId
'                        Set oBounded_MemberPart = oMemberPart
'                    ElseIf TypeOf oPort.Geometry Is IJCurve Then
'                        eBounding_PortId = ePortId
'                        Set oBounding_MemberPart = oMemberPart
'                    End If
'                End If
'            End If
'        Next iIndex
'
'        For iIndex = 1 To lCount
'            Set oPortObj = oElements_Ports.Item(iIndex)
'            sText = sText & vbCrLf & _
'                    "(" & Trim(Str(iIndex)) & ") Typename: " & TypeName(oPortObj)
'
'            If TypeOf oPortObj Is ISPSSplitAxisPort Then
'                Set oSplitAxisPort = oPortObj
'
'                Set oPort = oSplitAxisPort.Port
'                Set oMemberPart = oSplitAxisPort.Part
'
'                ePortId = oSplitAxisPort.PortIndex
'                If ePortId = SPSMemberAxisAlong Then
'                    sText = sText & vbCrLf & "   ...ePortId = SPSMemberAxisAlong"
'                    If TypeOf oPortObj Is ISPSSplitAxisAlongPort Then
'                    Else
'                        sText = sText & vbCrLf & "   .*.oPortObj is NOT ISPSSplitAxisAlongPort"
'                    End If
'
'                ElseIf ePortId = SPSMemberAxisEnd Then
'                    sText = sText & vbCrLf & "   ...ePortId = SPSMemberAxisEnd"
'                    If TypeOf oPortObj Is ISPSSplitAxisEndPort Then
'                    Else
'                        sText = sText & vbCrLf & "   .*.oPortObj is NOT ISPSSplitAxisEndPort"
'                    End If
'
'                ElseIf ePortId = SPSMemberAxisStart Then
'                    sText = sText & vbCrLf & "   ...ePortId = SPSMemberAxisStart"
'                    If TypeOf oPortObj Is ISPSSplitAxisEndPort Then
'                    Else
'                        sText = sText & vbCrLf & "   .*.oPortObj is NOT ISPSSplitAxisEndPort"
'                    End If
'
'                Else
'                    sText = sText & vbCrLf & "   .*.ePortId = ???"
'                End If
'
'                If TypeOf oMemberPart Is IJNamedItem Then
'                    Set oNamedItem = oMemberPart
'                    sText = sText & vbCrLf & "   ...PartName:" & oNamedItem.Name
'                End If
'
'                If oPort Is Nothing Then
'                ElseIf TypeOf oPort.Geometry Is IJPoint Then
'                    eBounded_PortId = ePortId
'                    Set oBounded_AxisPort = oSplitAxisPort
'                    Set oBounded_MemberPart = oMemberPart
'                    Set oPoint = oPort.Geometry
'                    oPoint.GetPoint dx, dy, dz
'
'                ElseIf TypeOf oPort.Geometry Is IJCurve Then
'                    eBounding_PortId = ePortId
'                    Set oBounding_AxisPort = oSplitAxisPort
'                    Set oBounding_MemberPart = oMemberPart
'                    Dim oPosition As IJDPosition
'                    Set oPosition = GetConnectionPositionOnSupping(oBounding_MemberPart, _
'                                                                   oBounded_MemberPart, _
'                                                                   eBounded_PortId)
'                    oPosition.Get dx, dy, dz
'                Else
'                    Set oPort = Nothing
'                End If
'
'                If Not oPort Is Nothing Then
'                    'get the curve segment at this position.
'                    'The axis may be made up of many curve segments
'                    sText = sText & vbCrLf & "   ...Pnt:" & Format(dx, "0.000") & _
'                            " , " & Format(dy, "0.000") & _
'                            " , " & Format(dz, "0.000")
'
'                    Set oAxisCurve = GetAxisCurveAtPosition(dx, dy, dz, oMemberPart)
'                    ' Matrix: U is direction Along Axis
'                    ' Matrix: V is direction normal to Web (Web Right to Web Left)
'                    ' Matrix: W is direction normal to Flange (Flange Bottom to Flange Top)
'                    oMemberPart.Rotation.GetTransformAtPosition dx, dy, dz, oMatrix, Nothing
'                    Set oVector = New DVector
'                    oVector.Set oMatrix.IndexValue(0), oMatrix.IndexValue(1), oMatrix.IndexValue(2)
'                    sText = sText & vbCrLf & "   ...Dir:" & Format(oVector.x, "0.000") & _
'                            " , " & Format(oVector.y, "0.000") & _
'                            " , " & Format(oVector.z, "0.000")
'
'                    If ePortId = SPSMemberAxisEnd Then
'                        Set oBounded_Point = oPoint
'                        Set oBounded_AxisCurve = oAxisCurve
'                    ElseIf ePortId = SPSMemberAxisStart Then
'                        Set oBounded_Point = oPoint
'                        Set oBounded_AxisCurve = oAxisCurve
'                    ElseIf ePortId = SPSMemberAxisAlong Then
'                        Set oBounding_Point = oPoint
'                        Set oBounding_AxisCurve = oAxisCurve
'                    End If
'
'                End If
'
'            Else
'                sText = sText & vbCrLf & "   .*.Else oPortObj Is ISPSSplitAxisPort "
'
'            End If
'
'
'        Next iIndex
'
'        If Not oBounded_AxisCurve Is Nothing Then
'            If Not oBounding_AxisCurve Is Nothing Then
'                If (TypeOf oBounded_AxisCurve Is IJLine) And _
'                   (TypeOf oBounding_AxisCurve Is IJLine) Then
'                    bColinear = IsMemberAxesColinear(oBounded_AxisCurve, oBounding_AxisCurve)
'                    bEndToEnd = IsMemberAxesEndToEnd(oBounded_AxisCurve, oBounding_AxisCurve)
'                    bRightAngle = IsMemberAxesAtRightAngles(oBounded_AxisCurve, oBounding_AxisCurve)
'                    sText = sText & vbCrLf & "...bEndToEnd:" & bEndToEnd & _
'                            "   ...bColinear:" & bColinear & _
'                            "   ...bRightAngle:" & bRightAngle
'                End If
'            End If
'        End If
'
'    End If
'
'    zMsgBox sText
'    Exit Sub
'
'ErrorHandler:
'    HandleError MODULE, METHOD, sMsg
'End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Debug_ObjectName(oObject As Object, _
                                 bIncludeType As Boolean) As String
Const METHOD = "::Debug_ObjectName"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim sText As String
    Dim oPortItem As IJPort
    Dim oConnectable As IJConnectable
    Dim oJNamedItem As IJNamedItem
    Dim EntityNaming As IJDStructEntityNaming
    Dim oSPS_SplitAxisPort As SPSMembers.ISPSSplitAxisPort
        
    On Error Resume Next
    Debug_ObjectName = ""
    
    sText = ""
    If oObject Is Nothing Then
        sText = " ...?oObject Is Nothing?"
        Debug_ObjectName = sText
        Exit Function

    ElseIf TypeOf oObject Is IJPort Then
        Set oPortItem = oObject
        Set oConnectable = oPortItem.Connectable
        If oConnectable Is Nothing Then
            sText = sText & "??IJPort.Connectable Is Nothing??"
        ElseIf TypeOf oConnectable Is IJDStructEntityNaming Then
            Set EntityNaming = oConnectable
        ElseIf TypeOf oConnectable Is IJNamedItem Then
            Set oJNamedItem = oConnectable
        Else
            sText = sText & "??TypeOf IJPort.Connectable Is ??"
        End If
    Else
        If TypeOf oObject Is IJNamedItem Then
            Set oJNamedItem = oObject
        ElseIf TypeOf oObject Is IJDStructEntityNaming Then
            Set EntityNaming = oObject
        ElseIf TypeOf oObject Is IJNamedItem Then
            Set oJNamedItem = oObject
        Else
            sText = sText & "??TypeOf oObject Is ??"
        End If
    End If
    
    If EntityNaming Is Nothing Then
        If oJNamedItem Is Nothing Then
            sText = sText & "??Name??"
        Else
            sText = sText & oJNamedItem.Name
        End If
    Else
        sText = sText & EntityNaming.Name
    End If
            
    If Len(Trim(sText)) < 1 Then
        Dim pSystemParent As IJSystem
        Dim pSystemChild As IJSystemChild
        Set pSystemChild = oObject
        If Not pSystemChild Is Nothing Then
            Set pSystemParent = pSystemChild.GetParent
            If Not pSystemParent Is Nothing Then
                sText = "(" & Debug_ObjectName(pSystemParent, False) & ")"
                Set pSystemParent = Nothing
            End If
            Set pSystemChild = Nothing
        End If
    End If
    
    If bIncludeType Then
        
        If oObject Is Nothing Then
            sText = sText & " ... oObject Is Nothing"
        
        ElseIf TypeOf oObject Is ISPSSplitAxisEndPort Then
            Set oSPS_SplitAxisPort = oObject
            If oSPS_SplitAxisPort.PortIndex = SPSMemberAxisAlong Then
                sText = sText & " ... SPSMemberAxisAlong"
            ElseIf oSPS_SplitAxisPort.PortIndex = SPSMemberAxisEnd Then
                sText = sText & " ... SPSMemberAxisEnd"
            ElseIf oSPS_SplitAxisPort.PortIndex = SPSMemberAxisStart Then
                sText = sText & " ... SPSMemberAxisStart"
            Else
                sText = sText & " ... ?PortIndex? " & _
                        Trim(Str(oSPS_SplitAxisPort.PortIndex))
            End If
            Set oSPS_SplitAxisPort = Nothing
            
        ElseIf TypeOf oObject Is IJPort Then
            sText = sText & " ... IJPort"
        
        ElseIf TypeOf oObject Is IJWireBody Then
            sText = sText & " ... IJWireBody"
        
        ElseIf TypeOf oObject Is IJSurfaceBody Then
            sText = sText & " ... IJSurfaceBody"
        
        ElseIf TypeOf oObject Is IJSolidBody Then
            sText = sText & " ... IJSolidBody"
        
        ElseIf TypeOf oObject Is IJPointsGraphBody Then
            sText = sText & " ... IJPointsGraphBody"
        
        ElseIf TypeOf oObject Is IJModelBody Then
            sText = sText & " ... IJModelBody"
        
        ElseIf TypeOf oObject Is IJSmartOccurrence Then
            Dim oSmartItem As IJSmartItem
            Dim oSmartClass As IJSmartClass
            Dim oSmartOccurrence As IJSmartOccurrence
            Set oSmartOccurrence = oObject
            Set oSmartItem = oSmartOccurrence.ItemObject
            If Not oSmartItem Is Nothing Then
                sText = sText & " ... IJSmartItem:" & oSmartItem.Name
            Else
                Set oSmartClass = oSmartOccurrence.ItemObject
                If Not oSmartClass Is Nothing Then
                    sText = sText & " ... IJSmartClass:" & oSmartClass.SCName
                Else
                    sText = sText & " ... IJSmartOccurrence:"
                End If
            End If
            sText = sText & " ... TypeOf " & TypeName(oObject)
        
        Else
            sText = sText & " ... TypeOf " & TypeName(oObject)
        End If
        
    End If
    
    Debug_ObjectName = sText
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
            
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Debug_PortData(oObject As Object, _
                               bIncludeName As Boolean) As String
Const METHOD = "::Debug_PortData"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim sText As String
    
    Dim lXId As Long
    Dim lCtxId As Long
    Dim lOptId As Long
    Dim lOprId As Long
    Dim lPortType As Long

    Dim oPortItem As IJPort
    Dim oPortMoniker As IUnknown
    
    Dim oStructPort As IJStructPort
    Dim oSP3D_StructPort As SP3DStructPorts.IJStructPort
    Dim oStructGeomBasicPort As StructGeomBasicPort
    
    Dim oPortHelper As PORTHELPERLib.PortHelper
        
    On Error Resume Next
    Debug_PortData = ""
    
    sText = ""
    If oObject Is Nothing Then
        sText = " ...?oObject Is Nothing?"
        Debug_PortData = sText
        Exit Function

    ElseIf TypeOf oObject Is ISPSSplitAxisEndPort Then
        sText = "ISPSSplitAxisEndPort:"

    ElseIf TypeOf oObject Is StructGeomBasicPort Then
        Set oStructGeomBasicPort = oObject
        Set oSP3D_StructPort = oStructGeomBasicPort
        Set oPortMoniker = oSP3D_StructPort.PortMoniker
        If oPortMoniker Is Nothing Then
            sText = "PortMoniker Is Nothing: Type: ?" & "  Ctx: ?" & "  Opt: ?" & "  Opr: ?" & "  Xid: ?"
        Else
            Set oPortHelper = New PORTHELPERLib.PortHelper
            oPortHelper.DecodeTopologyProxyMoniker oPortMoniker, lPortType, lCtxId, lOptId, lOprId, lXId
            sText = "StructGeomBasicPort: Type: " & Trim(Str(lPortType)) & "  Ctx: " & Trim(Str(lCtxId)) & _
                    "  Opt: " & Trim(Str(lOptId)) & "  Opr: " & Trim(Str(lOprId)) & _
                    "  Xid: " & Trim(Str(lXId))
        End If
        
    ElseIf TypeOf oObject Is IJStructPort Then
        Set oStructPort = oObject
        sText = "IJStructPort: Type: " & Trim(Str(oStructPort.ProxyType)) & "  Ctx: " & Trim(Str(oStructPort.ContextID)) & _
                "  Opt: " & Trim(Str(oStructPort.OperationID)) & "  Opr: " & Trim(Str(oStructPort.OperatorID)) & _
                "  Xid: " & Trim(Str(oStructPort.SectionID))
    Else
        sText = "Type: ?" & "  Ctx: ?" & "  Opt: ?" & "  Opr: ?" & "  Xid: ?"
    End If
    
    If bIncludeName Then
        If oObject Is Nothing Then
            sText = sText & " ... oObject Is Nothing"
        Else
            sText = sText & " ... " & Debug_ObjectName(oObject, True)
        End If
    End If
    
    Debug_PortData = sText
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
            
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub zSM_trace(sText As String, _
                   Optional sDumArg1 As String, Optional sTitle As String)
On Error Resume Next
    If bSM_trace Then
        zMsgBox sText, sDumArg1, sTitle
    End If
    Exit Sub
End Sub


