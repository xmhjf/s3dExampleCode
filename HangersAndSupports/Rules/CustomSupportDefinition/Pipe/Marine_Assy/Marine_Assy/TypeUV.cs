//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   TypeUV.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.TypeUV
//   Author       :BS
//   Creation Date:12.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  12.Jun.2013     BS       CR-CP-224470 Convert HS_Marine_Assy to C# .Net
//  04.Dec.2013   Rajeswari  DI-CP-241804 Modified the code as part of hardening
//  11.Dec.2014     PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Linq;
using Ingr.SP3D.Content.Support.Symbols;
namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class TypeUV : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        private const string HORSECTION1 = "HORSECTION1";
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
        private double PI = 0;
        string sectionSize = string.Empty, flatBarSize = string.Empty;
        bool includeRightPad, includeLeftPad, isVerticalRoute, isVerticalStruct;
        int routeCount = 0;
        double uBoltOffset = 0.0, rotAngle = 0.0;
        string[] uBoltPart, uBoltPart1, ubolts;
        bool[] isUboltExist;
        double[] routeAngle;
        MarineAssemblyServices.PADProperties padProperties;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    PI = (Math.Atan(1) * 4);
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //Get the attributes from assembly
                    bool flatBarFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnFBSection", "FlatBarFromRule")).PropValue;
                    includeRightPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnRightPad", "IncludeRightPad")).PropValue;
                    includeLeftPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnLeftPad", "IncludeLeftPad")).PropValue;
                    bool uBoltOption = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOpt", "PipeAttachOption")).PropValue;
                    uBoltOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnPAttachOff", "Offset")).PropValue;
                    bool uboltOffsetFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOff", "OffsetFrmRule")).PropValue;
                    rotAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnRotAngle", "RotAngle")).PropValue;

                    routeCount = SupportHelper.SupportedObjects.Count;
                    if (routeCount > 10)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Number of Pipes should be less than or equal to 10.", "", "TypeUV.cs", 77);
                        return null;
                    }

                    //Give warning message when multiple pipes selected and Rotation angle given
                    if (routeCount > 1)
                    {
                        if (rotAngle > 0)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Rotation around Pipe is applicable only for Single Pipe. Re-setting the rotation angle to 0 degrees.", "", "TypeUV.cs", 86);
                            rotAngle = 0;
                        }
                    }
                    //Get the Section Size
                    sectionSize = MarineAssemblyServices.GetSectionSize(this);

                    //Get the FlatBar Size
                    if (flatBarFromRule) //1 means true
                        GenericHelper.GetDataByRule("hsMrnRLFlatBarSize", (BusinessObject)support, out flatBarSize);
                    else
                    {
                        PropertyValueCodelist flatBarSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFBSection", "FlatBarSize");
                        flatBarSize = flatBarSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(flatBarSizeCodeList.PropValue).DisplayName;
                        if (flatBarSize.Equals("NONE") || String.IsNullOrEmpty(flatBarSize))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Flat Bar size is not available.", "", "TypeUV.cs", 102);
                            return null;
                        }
                    }
                    //Get the UBolt
                    uBoltPart = MarineAssemblyServices.GetUboltPart(this);

                    string TempStr = string.Empty;
                    int[] pipeAttachment, uBolt;
                    pipeAttachment = new int[routeCount];
                    uBolt = new int[routeCount];
                    isUboltExist = new bool[routeCount];
                    uBoltPart1 = new string[routeCount];

                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        TempStr = "U" + routeIndex;
                        pipeAttachment[routeIndex - 1] = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPAttach", TempStr)).PropValue;
                        if (uBoltOption == false)      //User defined
                        {
                            if (pipeAttachment[routeIndex - 1] == 3)       //None
                                isUboltExist[routeIndex - 1] = false;
                            else
                                isUboltExist[routeIndex - 1] = true;
                        }
                        else        //Default Ubolt is True
                            isUboltExist[routeIndex - 1] = true;
                    }

                    string[] temp = new string[1];
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        if (isUboltExist[routeIndex - 1] == true)
                        {
                            temp = uBoltPart[routeIndex - 1].Split('-');
                            uBoltPart1[routeIndex - 1] = temp[0] + "-";
                        }
                    }

                    routeAngle = MarineAssemblyServices.GetRoutAngleAndIsRouteVertical(this, PI, routeCount, out isVerticalRoute);
                    if (routeCount > 1)
                    {
                        for (int routeIndex = 1; routeIndex <= routeCount - 1; routeIndex++)
                        {
                            if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[routeIndex - 1]), Math.Round(routeAngle[(routeIndex - 1) + 1])) == false)
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe slopes are not equal. Please check", "", "TypeUV.cs", 148);
                                return null;
                            }
                        }
                    }

                    //Get Pad part and Dimensions
                    string sectionCode, steelStd;
                    padProperties = MarineAssemblyServices.GetPadPartAndDimensions(this, sectionSize, out sectionCode, out steelStd);
                    //GetSteelStandard    
                    string steelStandard = MarineAssemblyServices.GetSteelStandard(this, steelStd);
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    //Get the UboltOffset
                    if (uboltOffsetFromRule == true)
                    {
                        PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachOffset");
                        IEnumerable<BusinessObject> hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvPAttachOff", "SectionSize")).PropValue == sectionCode));
                        if (hsMarineServiceDimPart.Count() > 0)
                            uBoltOffset = (double)((PropertyValueDouble)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachOff", "PipeAttachOffset")).PropValue;
                    }

                    parts.Add(new PartInfo(HORSECTION1, sectionSize + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION1, flatBarSize + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION2, flatBarSize + " " + steelStandard));

                    ubolts = new string[routeCount];
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        if (isUboltExist[routeIndex - 1] == true)
                        {
                            ubolts[routeIndex - 1] = "sUbolt_" + routeIndex;
                            parts.Add(new PartInfo(ubolts[routeIndex - 1], uBoltPart[routeIndex - 1]));
                        }
                    }

                    if (includeLeftPad == true)
                        parts.Add(new PartInfo(LEFTPAD, padProperties.padPart));

                    if (includeRightPad == true)
                        parts.Add(new PartInfo(RIGHTPAD, padProperties.padPart));

                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 2;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                string supportingType = MarineAssemblyServices.GetSupportingTypes(this);
                double boundingBoxWidth = 0, boundingBoxHeight = 0, leftPipeDiameter = 0, rightPipeDiameter = 0;
                MarineAssemblyServices.GetBoundingBoxDimensionsAndPipeDiameter(this, routeCount, ref boundingBoxWidth, ref boundingBoxHeight, ref leftPipeDiameter, ref rightPipeDiameter);

                //Auto Dimensioning of Supports                
                int textIndex = 0;
                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        textIndex = textIndex + 1;
                        MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[ubolts[routeIndex - 1]], "Route");
                    }
                }

                textIndex++;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[HORSECTION1], "BeginCap");

                textIndex++;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[HORSECTION1], "EndCap");

                PropertyValueCodelist cardinalCodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJOAhsSteelCP", "CP2");
                cardinalCodeList.PropValue = 1;
                componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP2");
                cardinalCodeList.PropValue = 3;
                componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP4");

                string strVerticalSection1Port = string.Empty, strVerticalSection2Port = string.Empty;
                double leftStructAngle = 0.0, rightStructAngle = 0.0;
                MarineAssemblyServices.GetLeftRightStructAngleAndIsVerticalStructure(this, PI, out isVerticalStruct, out leftStructAngle, out rightStructAngle);
                if (leftStructAngle <= 0)
                    strVerticalSection1Port = "EndCap";
                else
                    strVerticalSection1Port = "EndFace";

                cardinalCodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJOAhsSteelCP", "CP2");
                cardinalCodeList.PropValue = 1;
                componentDictionary[VERTSECTION2].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP2");
                cardinalCodeList.PropValue = 3;
                componentDictionary[VERTSECTION2].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP4");

                if (rightStructAngle <= 0)
                    strVerticalSection2Port = "BeginCap";
                else
                    strVerticalSection2Port = "BeginFace";

                textIndex++;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[VERTSECTION1], strVerticalSection1Port);

                textIndex++;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[VERTSECTION2], strVerticalSection2Port);

                //Get Section Structure dimensions
                BusinessObject sectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThick = 0.0, sectionWidth1 = 0.0, sectionDepth1 = 0.0;
                int sectionType = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionType")).PropValue;
                if (sectionType == 4)
                {
                    sectionWidth = crossSection.Width;
                    sectionDepth = crossSection.Depth;
                }
                else
                {
                    sectionWidth = crossSection.Width;
                    sectionDepth = crossSection.Depth;
                    sectionThick = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                }
                sectionPart = componentDictionary[VERTSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                sectionWidth1 = crossSection.Width;
                sectionDepth1 = crossSection.Depth;

                //get Codelisted Property Values for Hgr Beam
                MarineAssemblyServices.BeamCLProperties beamCLProperties = MarineAssemblyServices.GetBeamCLProperties(componentDictionary);

                string[] partNames = new string[3];
                partNames[0] = HORSECTION1;
                partNames[1] = VERTSECTION1;
                partNames[2] = VERTSECTION2;

                //Intialize the Hgr beam properties
                MarineAssemblyServices.IntializeBeamProperties(componentDictionary, support, beamCLProperties, partNames);

                double verticalSection1EndOverLength = 0, verticalSection2EndOverLength = 0;
                //set the default cut back angle for the vertical section
                if (sectionType == 4 && Configuration == 2)//For RS Section set the cutback angle
                {
                    componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                    componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                    componentDictionary[VERTSECTION1].SetPropertyValue(-(Math.Atan(1) * 4) / 4, "IJOAhsCutback", "CutbackBeginAngle");
                    componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                    componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                    componentDictionary[VERTSECTION2].SetPropertyValue(-(Math.Atan(1) * 4) / 4, "IJOAhsCutback", "CutbackEndAngle");
                }
                else
                {
                    componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                    componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                    componentDictionary[VERTSECTION1].SetPropertyValue((Math.Atan(1) * 4) / 4, "IJOAhsCutback", "CutbackBeginAngle");
                    componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                    componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                    componentDictionary[VERTSECTION2].SetPropertyValue((Math.Atan(1) * 4) / 4, "IJOAhsCutback", "CutbackEndAngle");
                }
                //Get the largest Pipe Dia
                double[] pipeDiameter = new double[routeCount];

                for (int index = 1; index <= routeCount; index++)
                {
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(index);
                    pipeDiameter[index - 1] = pipeInfo.OutsideDiameter;
                }
                //For UBolt
                string uBoltType = string.Empty;
                double[] H = new double[routeCount];

                //Set attribute values for the U-Bolt
                int[] uBoltTypes = new int[routeCount];
                string tempUBoltType = string.Empty;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        tempUBoltType = "U" + routeIndex;

                        PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachSel");
                        IEnumerable<BusinessObject> hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachPart")).PropValue == uBoltPart1[routeIndex - 1]));
                        if (hsMarineServiceDimPart.Count() > 0)
                            uBoltTypes[routeIndex - 1] = (int)((PropertyValueCodelist)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachType")).PropValue;

                        //Get the Ublot Dimensions
                        componentDictionary[ubolts[routeIndex - 1]].SetPropertyValue(pipeDiameter[routeIndex - 1], "IJOAhsPipeOD", "PipeOD");
                        if (uBoltTypes[routeIndex - 1] == 1)
                        {
                            if (sectionType == 4)
                                componentDictionary[ubolts[routeIndex - 1]].SetPropertyValue(sectionWidth, "IJOAhsSteelThickness", "SteelThickness");
                            else
                                componentDictionary[ubolts[routeIndex - 1]].SetPropertyValue(sectionThick, "IJOAhsSteelThickness", "SteelThickness");
                        }
                        support.SetPropertyValue(uBoltTypes[routeIndex - 1], "IJOAhsMrnPAttach", tempUBoltType);
                    }
                }
                //Set attribute values for Pads
                string[] arrayOfPadKeys = new string[2];

                if (includeLeftPad == false && includeRightPad == true)
                {
                    verticalSection1EndOverLength = verticalSection1EndOverLength - padProperties.padThickness;
                    arrayOfPadKeys[0] = RIGHTPAD;
                }
                else if (includeLeftPad == true && includeRightPad == false)
                {
                    verticalSection2EndOverLength = verticalSection2EndOverLength - padProperties.padThickness;
                    arrayOfPadKeys[0] = LEFTPAD;
                }
                else if (includeLeftPad == true && includeRightPad == true)
                {
                    verticalSection2EndOverLength = verticalSection2EndOverLength - padProperties.padThickness;
                    verticalSection1EndOverLength = verticalSection1EndOverLength - padProperties.padThickness;
                    arrayOfPadKeys[0] = RIGHTPAD;
                    arrayOfPadKeys[1] = LEFTPAD;
                }
                //Set the Pad properties
                MarineAssemblyServices.SetPadProperties(componentDictionary, support, padProperties, arrayOfPadKeys);
                //=======================
                // Set Assembly Attributes
                //=======================
                //Set UboltOffset value                
                support.SetPropertyValue(uBoltOffset, "IJOAhsMrnPAttachOff", "Offset");

                //Set Assembly Attributes
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;

                //Set Frame Orientation attribute
                PropertyValueCodelist orientCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmOrient", "FrmOrient");
                string fromOrient = orientCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(orientCodeList.PropValue).DisplayName;
                CodelistItem codeList = metadataManager.GetCodelistInfo("hsMrnCLFrmOrient", "UDP").GetCodelistItem(fromOrient);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnFrmOrient", "FrmOrient");

                //Set Section Size properties
                codeList = metadataManager.GetCodelistInfo("hsMrnCLSecSize", "UDP").GetCodelistItem(sectionSize);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnSection", "SectionSize");
                codeList = metadataManager.GetCodelistInfo("hsMrnCLFBSecSize", "UDP").GetCodelistItem(flatBarSize);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnFBSection", "FlatBarSize");

                double overhangLeft, overhangRight;
                //Set Overhang attributes
                double hangerOverHangLeft = 0, hangerOverHangRight = 0;
                int overhangOption = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnOHOption", "OverhangOpt")).PropValue;
                if (overhangOption == 1) //By Catalog Rule
                {
                    Collection<object> hangerOverHangCollection;

                    bool value = GenericHelper.GetDataByRule("hsMrnRLFrmOH", (BusinessObject)support, out hangerOverHangCollection);
                    if (hangerOverHangCollection[0] == null)
                    {
                        hangerOverHangLeft = (double)hangerOverHangCollection[1];
                        hangerOverHangRight = (double)hangerOverHangCollection[2];
                    }
                    else
                    {
                        hangerOverHangLeft = (double)hangerOverHangCollection[0];
                        hangerOverHangRight = (double)hangerOverHangCollection[1];
                    }
                    support.SetPropertyValue(hangerOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                    support.SetPropertyValue(hangerOverHangRight, "IJOAhsMrnOHRight", "OverhangRight");
                }
                else if (overhangOption == 3) //User Defined
                {
                    hangerOverHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                    hangerOverHangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHRight", "OverhangRight")).PropValue;
                }
                overhangLeft = hangerOverHangLeft - rightPipeDiameter / 2;
                overhangRight = hangerOverHangRight - leftPipeDiameter / 2;

                //=======================================
                //Do Something if more than one Structure
                //=======================================
                //Get the structure surface

                //Set Length of horzontal member
                double routeLowStruct1HorDist, routeLowStruct2HorDist, routeHighStruct1HorDist, routeHighStruct2HorDist, horizontalSectionLength = 0, distanceBetweenStruct;
                //For Sloped Structure
                Boolean[] bIsOffsetApplied = MarineAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] idxStructPort = MarineAssemblyServices.GetIndexedStructPortName(this, bIsOffsetApplied);
                string rightStructPort = idxStructPort[0];
                string leftStructPort = idxStructPort[1];

                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                if ((Math.Abs(routeAngle[0]) < (0 + 0.01) && Math.Abs(routeAngle[0]) > (0 - 0.01)) || (Math.Abs(routeAngle[0]) < ((Math.Atan(1) * 4) + 0.01) && Math.Abs(routeAngle[0]) > ((Math.Atan(1) * 4) - 0.01)))
                {
                    routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);
                    routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                    routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal);
                    routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal);
                    distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                }
                else
                {
                    if ((Math.Abs(routeStructAngle) < (0 + 0.01) && Math.Abs(routeStructAngle) > (0 - 0.01)) || (Math.Abs(routeStructAngle) < ((Math.Atan(1) * 4) + 0.01) && Math.Abs(routeStructAngle) > ((Math.Atan(1) * 4) - 0.01)))
                    {
                        routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Vertical);
                        routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Vertical);
                        routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Vertical);
                        routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Vertical);
                        distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Vertical);
                    }
                    else if (Math.Abs(routeStructAngle) < ((Math.Atan(1) * 4) / 2 + 0.01) && Math.Abs(routeStructAngle) > ((Math.Atan(1) * 4) / 2 - 0.01))
                    {
                        routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);
                        routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                        routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal);
                        routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal);
                        distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                    }
                    else
                    {
                        routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Vertical);
                        routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Vertical);
                        routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Vertical);
                        routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Vertical);
                        distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Vertical);
                    }
                }

                if (supportingType.ToUpper() == "SLAB")
                {
                    if (overhangOption == 1 || overhangOption == 3)          //By catalog OR User Defined
                        horizontalSectionLength = boundingBoxHeight + overhangLeft + overhangRight;
                    else if (overhangOption == 2)    //adjust to structure
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Joints" + ": " + "WARNING: " + "Adjust to Supporting steel is applicable only when the support placement is By Point.", "", "TypeUV.cs", 495);
                        return;
                    }
                }
                else if (supportingType.ToUpper() == "STEEL")
                {
                    if (overhangOption == 1 || overhangOption == 3)          //By catalog
                        horizontalSectionLength = boundingBoxHeight + overhangLeft + overhangRight;
                    else if (overhangOption == 2)                             //Adjust to structure
                    {
                        overhangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                        overhangLeft = Math.Min(routeLowStruct1HorDist, routeHighStruct1HorDist) + sectionWidth / 2;
                        horizontalSectionLength = distanceBetweenStruct + sectionWidth;
                        support.SetPropertyValue(overhangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                        support.SetPropertyValue(overhangLeft, "IJOAhsMrnOHRight", "OverhangRight");
                    }
                }
                else if (supportingType.ToUpper() == "STEEL-SLAB" || supportingType.ToUpper() == "SLAB-STEEL")
                {
                    if (overhangOption == 1 || overhangOption == 3)         // by catalog
                        horizontalSectionLength = boundingBoxHeight + overhangLeft + overhangRight;
                    else if (overhangOption == 2)    //adjust to structure
                    {
                        overhangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                        overhangLeft = Math.Min(routeLowStruct1HorDist, routeHighStruct1HorDist) + sectionWidth / 2;
                        horizontalSectionLength = boundingBoxHeight + overhangLeft + overhangRight;
                    }
                }

                componentDictionary[HORSECTION1].SetPropertyValue(horizontalSectionLength, "IJUAHgrOccLength", "Length");

                //=============
                //Create Joints
                //=============
                //Create a collection to hold the joints
                MarineAssemblyServices.ConfigIndex[] sectionUboltConfigIndex = new MarineAssemblyServices.ConfigIndex[routeCount];
                string horizontalSectionPort1 = string.Empty, horizontalSectionPort2 = string.Empty;
                if (Configuration == 1)
                {
                    horizontalSectionPort1 = "EndCap";
                    horizontalSectionPort2 = "BeginCap";
                }
                else if (Configuration == 2)
                {
                    horizontalSectionPort1 = "BeginCap";
                    horizontalSectionPort2 = "EndCap";
                }
                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        if (sectionType != 4)
                            sectionUboltConfigIndex[routeIndex - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeYZ, Axis.Y, Axis.Z);//11574;
                        else
                            sectionUboltConfigIndex[routeIndex - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.NegativeX);//1454;
                    }
                }

                MarineAssemblyServices.ConfigIndex horizontalVertConfigIndex1, horizontalVertConfigIndex2, routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex();
                double routeSectionPlanOffset, horRoutePlaneOffset = 0, horRouteOriginOffset = 0, horRouteAxisOffset = 0, horVertPlaneOffset1 = 0, horVertPlaneOffset2 = 0;
                double supRotAngle = 0, horVertAxisOffset1 = 0, horVertOriginOffset1 = 0, horVertAxisOffset2 = 0, horVertOriginOffset2 = 0;
                string horizontalSectionPort = string.Empty;
                double shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnShoeHeight", "ShoeHeight")).PropValue;
                if (routeCount == 1)          //Implementation for Angular rigid joint
                {
                    if (Configuration == 1)
                        supRotAngle = rotAngle;
                    else if (Configuration == 2)
                    {
                        if (rotAngle > (Math.Atan(1) * 4))
                            supRotAngle = rotAngle - (Math.Atan(1) * 4);
                        else
                            supRotAngle = (Math.Atan(1) * 4) + rotAngle;
                    }

                    if (sectionType == 4)
                    {
                        horizontalVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX);
                        horizontalVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeX);
                        routeSectionPlanOffset = -sectionDepth / 2 + uBoltOffset;

                        horRoutePlaneOffset = sectionWidth + shoeHeight;
                        horRouteOriginOffset = -overhangLeft - boundingBoxHeight;
                        horRouteAxisOffset = 0;

                        horVertPlaneOffset1 = sectionDepth + 0.01;
                        horVertAxisOffset1 = -sectionDepth1 + sectionWidth + 0.005;
                        horVertOriginOffset1 = sectionWidth1;

                        horVertPlaneOffset2 = sectionDepth + 0.01;
                        horVertAxisOffset2 = sectionWidth + 0.005;
                        horVertOriginOffset2 = 0;
                    }
                    else
                    {
                        horizontalVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeY);
                        horizontalVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.Y);
                        routeSectionPlanOffset = -sectionWidth / 2 + uBoltOffset;

                        horRoutePlaneOffset = -overhangLeft - boundingBoxHeight;
                        horRouteOriginOffset = sectionWidth / 2;
                        horRouteAxisOffset = -shoeHeight;

                        horVertPlaneOffset1 = sectionWidth + 0.01;
                        horVertAxisOffset1 = sectionWidth1;
                        horVertOriginOffset1 = sectionDepth1 - 0.005;

                        horVertPlaneOffset2 = sectionWidth + 0.01;
                        horVertAxisOffset2 = 0;
                        horVertOriginOffset2 = -0.005;
                    }
                }
                else            // for Multi Pipes
                {
                    if (sectionType == 4)
                    {
                        routeSectionPlanOffset = -sectionDepth / 2 + uBoltOffset;
                        if (Configuration == 1)
                        {
                            rotAngle = 0;
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.Z); 
                            horizontalVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX);
                            horizontalVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeX); 
                            horizontalSectionPort = "EndCap";
                            horRoutePlaneOffset = sectionWidth + shoeHeight;
                            horRouteOriginOffset = -overhangRight - boundingBoxHeight;
                            horRouteAxisOffset = sectionDepth / 2;
                            horVertPlaneOffset1 = sectionDepth + 0.01;
                            horVertAxisOffset1 = -sectionDepth1 + sectionWidth + 0.005;
                            horVertOriginOffset1 = sectionWidth1;
                            horVertPlaneOffset2 = sectionDepth + 0.01;
                            horVertAxisOffset2 = sectionWidth + 0.005;
                            horVertOriginOffset2 = 0;
                        }
                        else
                        {
                            rotAngle = (Math.Atan(1) * 4);
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.ZX, Axis.Z, Axis.NegativeZ);
                            horizontalVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX);
                            horizontalVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeX); 
                            horizontalSectionPort = "BeginCap";
                            horRoutePlaneOffset = boundingBoxWidth + sectionWidth + shoeHeight;
                            horRouteOriginOffset = overhangRight;
                            horRouteAxisOffset = sectionDepth / 2;
                            horVertPlaneOffset1 = -0.01;
                            horVertAxisOffset1 = sectionWidth + 0.005;
                            horVertOriginOffset1 = 0;
                            horVertPlaneOffset2 = -0.01;
                            horVertAxisOffset2 = -sectionDepth1 + sectionWidth + 0.005;
                            horVertOriginOffset2 = sectionWidth1;
                        }
                    }
                    else
                    {
                        horizontalVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeY); 
                        horizontalVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.Y); 
                        routeSectionPlanOffset = -sectionWidth / 2 + uBoltOffset;

                        if (Configuration == 1)
                        {
                            rotAngle = 0;
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX); 
                            horizontalSectionPort = "EndCap";

                            horRoutePlaneOffset = -overhangRight - boundingBoxHeight;
                            horRouteOriginOffset = sectionWidth / 2;
                            horRouteAxisOffset = -shoeHeight;

                            horVertPlaneOffset1 = sectionWidth + 0.01;
                            horVertAxisOffset1 = sectionWidth1;
                            horVertOriginOffset1 = sectionDepth1 - 0.005;

                            horVertPlaneOffset2 = sectionWidth + 0.01;
                            horVertAxisOffset2 = 0;
                            horVertOriginOffset2 = -0.005;
                        }

                        else if (Configuration == 2)
                        {
                            rotAngle = (Math.Atan(1) * 4);
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX); 
                            horizontalSectionPort = "BeginCap";

                            horRoutePlaneOffset = overhangRight + boundingBoxHeight;
                            horRouteOriginOffset = sectionWidth / 2;
                            horRouteAxisOffset = -boundingBoxWidth - shoeHeight;

                            horVertPlaneOffset1 = sectionWidth + 0.01;
                            horVertAxisOffset1 = 0;
                            horVertOriginOffset1 = sectionDepth1 - 0.005;

                            horVertPlaneOffset2 = sectionWidth + 0.01;
                            horVertAxisOffset2 = sectionWidth1;
                            horVertOriginOffset2 = -0.005;
                        }
                    }
                }
                if (routeCount == 1)//For Single Pipe use Angular Rigid Joint to Rotate around the Pipe
                {
                    //Add joint Between Horizontal Section and BoundingBox
                    JointHelper.CreateAngularRigidJoint("-1", "Route", ubolts[0], "Route", new Vector(0, 0, 0), new Vector(0, -(Math.Atan(1) * 4) / 2, supRotAngle));
                    //Add Joint Between the UBolt and Horizontal Section
                    if (sectionType == 4)
                        JointHelper.CreateRigidJoint(ubolts[0], "Steel", HORSECTION1, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Z, -horRoutePlaneOffset, routeSectionPlanOffset + sectionDepth / 2, horRouteOriginOffset + boundingBoxHeight / 2);
                    else
                        JointHelper.CreateRigidJoint(ubolts[0], "Steel", HORSECTION1, "BeginCap", Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX, horRoutePlaneOffset + boundingBoxHeight / 2, horRouteAxisOffset, uBoltOffset);

                }
                else //For Multi Pipes
                {
                    //Add joint Between Horizontal Section and BoundingBox
                    JointHelper.CreateRigidJoint(HORSECTION1, horizontalSectionPort, "-1", "BBR_High", routeHorConfigIndex.A, routeHorConfigIndex.B, routeHorConfigIndex.C, routeHorConfigIndex.D, horRoutePlaneOffset, horRouteAxisOffset, horRouteOriginOffset);

                    //Joint For Ubolt and Horizontal Section
                    string routePortName;
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        if (isUboltExist[routeIndex - 1] == true)
                        {
                            if (routeIndex == 1)
                                routePortName = "Route";
                            else
                                routePortName = "Route_" + routeIndex;

                            //Add Joint Between the UBolt and Horizontal Section
                            JointHelper.CreateTranslationalJoint(ubolts[routeIndex - 1], "Route", HORSECTION1, "Neutral", sectionUboltConfigIndex[routeIndex - 1].A, sectionUboltConfigIndex[routeIndex - 1].B, sectionUboltConfigIndex[routeIndex - 1].C, sectionUboltConfigIndex[routeIndex - 1].D, routeSectionPlanOffset);

                            //Add Joint Between the UBolt and Route
                            JointHelper.CreatePointOnAxisJoint(ubolts[routeIndex - 1], "Route", "-1", routePortName, Axis.Z);

                        }
                    }
                }
                //Add Joint Between the Horizontal and Vertical Beams
                JointHelper.CreateRigidJoint(HORSECTION1, horizontalSectionPort1, VERTSECTION1, "BeginCap", horizontalVertConfigIndex1.A, horizontalVertConfigIndex1.B, horizontalVertConfigIndex1.C, horizontalVertConfigIndex1.D, horVertPlaneOffset1, horVertAxisOffset1, horVertOriginOffset1);

                //Add Joint Between the Horizontal and Vertical Beams
                JointHelper.CreateRigidJoint(HORSECTION1, horizontalSectionPort2, VERTSECTION2, "EndCap", horizontalVertConfigIndex2.A, horizontalVertConfigIndex2.B, horizontalVertConfigIndex2.C, horizontalVertConfigIndex2.D, horVertPlaneOffset2, horVertAxisOffset2, horVertOriginOffset2);

                //Add Joint Between the Ports of Vertical Beam 1
                JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint Between the Ports of Vertical Beam 2
                JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint Between the Ports of Vertical Beam 1
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", "-1", rightStructPort, Plane.XY);

                //Add Joint Between the Ports of Vertical Beam 2
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, "BeginCap", "-1", leftStructPort, Plane.XY);

                if (includeLeftPad == false && includeRightPad == true)
                    //Add Joint Between the Left and the Vertical Section 1
                    JointHelper.CreateRigidJoint(RIGHTPAD, "Port2", VERTSECTION1, "EndFace", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY, 0, 0, 0);
                else if (includeLeftPad == true && includeRightPad == false)
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "Port1", VERTSECTION2, "BeginFace", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);
                else if (includeRightPad == true && includeLeftPad == true)
                {
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "Port1", VERTSECTION1, "EndFace", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY, 0, 0, 0);

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(RIGHTPAD, "Port2", VERTSECTION2, "BeginFace", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);
                }

                componentDictionary[VERTSECTION1].SetPropertyValue(verticalSection1EndOverLength, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(verticalSection2EndOverLength, "IJUAHgrOccOverLength", "BeginOverLength");
                support.SetPropertyValue(rotAngle, "IJOAhsMrnRotAngle", "RotAngle");
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(ubolts[0], 1)); // partindex, routeindex
                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections.Add(new ConnectionInfo(VERTSECTION1, 1)); // partindex, routeindex
                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                string bomDescriptionValue = (string)((PropertyValueString)support.GetPropertyValue("IJOAhsMrnBOMDesc", "BOM_DESC")).PropValue;

                if (string.IsNullOrEmpty(bomDescriptionValue))
                    bomDescription = (string)((PropertyValueString)support.GetPropertyValue("IJOAhsMrnSupType", "SupType")).PropValue + "-" + MarineAssemblyServices.GetLargePipeDiameter(this);
                else
                    bomDescription = bomDescriptionValue;

                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - TypeUV" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion

    }
}

