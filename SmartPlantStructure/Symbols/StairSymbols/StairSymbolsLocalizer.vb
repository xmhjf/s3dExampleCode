'=====================================================================================================
'
'Copyright 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'Stair symbols localizer
'
'File
'  StairSymbolsLocalizer.vb
'
'=====================================================================================================
Imports Ingr.SP3D.Common.Middle.Services

Namespace Ingr.SP3D.Content.Structure

    NotInheritable Class StairSymbolsLocalizer
        Public Shared Function GetString(ByVal iID As Integer, ByVal defMsgStr As String) As String
            Return CmnLocalizer.GetString(iID, defMsgStr, StairSymbolsResourceIDs.DEFAULT_RESOURCE, StairSymbolsResourceIDs.DEFAULT_ASSEMBLY)
        End Function
    End Class

End Namespace
