VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CustomMethods.cls
'   ProgID:         HgrSupportSymbolUtilities.CustomMethods
'   Author:         Caroline
'   Creation Date:  14.Jan.2002
'   Description:
'
'   Change History:
'       14.Jan.2002     Caroline            Creation Date
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Following method will be removed very soon. This has been
'included to improve the performance. These lines will be removed as soon as
'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.

' CMCache custom method to cache the input argument into a parameter contend and the reverse conversion
'It is up to you to find a way to convert your reference data object to a parameter content
Public Sub CMCacheForHgrPart(pInputCM As Object, bArgToCache As Boolean, pToConvert As Object, ByRef pOutput As Object)

 If bArgToCache Then
       
        'Need to convert the graphic input pToConvert into a Parameter ( pOutput)
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
        
        Dim oPart As IJDPart
        Set oPart = pToConvert
        'MsgBox "Partnumber :" & oPart.PartNumber
        
        ' Create a PC whose value is an identifier of the input
        ' Raju,
        ' the property Part_Number must be retrieved form from the pToConvert argument.
        ' I am hardcoding it
        '
        oPC.Type = igString
        oPC.String = oPart.PartNumber

        Set oPart = Nothing

        ' Always return this PC
        Set pOutput = oPC
        Set oPC = Nothing
 Else
        'Need to convert the cached Parameter pToConvert into your reference data object pOutput
        Dim oPCout As IJDParameterContent
        Dim oPCCache As IJDParameterContent
        Set oPCout = New DParameterContent
        
        ' Here there is three options
        ' o Return a parameter contents, that containts the part_number stored by the pToConvert argument
        '   In this case the edit command will have to retrieve the Part object when needed.
        '   Note : It is better to return a copy of the cached object instead the cached object itself.
        '          This allow to avoid the edition of the cached object.
        ' o Retrieve the Part_number and retrieve your part object with it.
        '   With this solution you can have assoc assertion while when this method is called you are in
        '   in a compute process.
        ' o Get the symbol (or equipment) from the pInputCM, query for the IJDReferencesArg interface,
        '   then get the argument at index 1 it is the SiteProxy.
        '   With this last method, the part has to be passed by reference to the symbol,
        '   but that is what you doing with your design.
        
        ' returning NULL means that the cache method is unable to resolve the cache.
        ' as of now anyway the part is passed as an input argument. this issue of caching
        ' has to be resolved in cycle2.
        
        Set pOutput = Nothing
        
        ' Here is the implementation of the option 1
'        Set oPCCache = pToConvert
'        oPCout.Type = oPCCache.Type
'        oPCout.String = oPCCache.String
'        Set pOutput = oPCout
'        Set oPCout = Nothing
        
        ' Here is the implementation of the option 2
'        Dim oRefDB As GSCADRefDataServices.IJDRefDBGlobal
'        Dim oPart As IJDPart

'        Set oRefDB = New GSCADRefDataServices.RefDBGlobal
'        Set oPart = oRefDB.GetCatalogPart("Storage Tanks", oPCCache.String)
'        Set pOutput = oRefDB
'        Set oRefDB = Nothing
'        Set oPart = Nothing
End If

End Sub

