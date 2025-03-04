//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Guide.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.Guide
//   Author       :Vijaya
//   Creation Date:23.July.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  23.July.2013     Vijaya   CR-CP-224470 Convert HS_Marine_Assy to C# .Net
//  04.Dec.2013    Rajeswari  DI-CP-241804 Modified the code as part of hardening
//  31.Oct.2014      PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
//  11.Dec.2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//  28-Apr-2015      PVK	  Resolve Coverity issues found in April
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    public class Guide : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
        private double PI = 0;
        string verticalSectionSize = string.Empty, horizontalSectionSize = string.Empty;
        bool includeRightPad, includeLeftPad, includeHorSection, isVerticalRoute, isVerticalSruct;
        int routeCount = 0;
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
                    PI = Math.Atan(1) * 4;
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //Get the attributes from assembly
                    bool sectionFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnGuideSec", "SectionFromRule")).PropValue;
                    PropertyValueCodelist verSectionTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnVerSection", "VerSectionType");
                    PropertyValueCodelist verSectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnVerSection", "VerSectionSize");
                   
                    PropertyValueCodelist horSectionTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnHorSection", "HorSectionType");
                    PropertyValueCodelist horSectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnHorSection", "HorSectionSize");

                    includeRightPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnRightPad", "IncludeRightPad")).PropValue;
                    includeLeftPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnLeftPad", "IncludeLeftPad")).PropValue;
                    includeHorSection = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnIncHorSec", "IncludeHorSec")).PropValue;

                    //Error message for flatbar
                    if (verSectionTypeCodeList.PropValue == 4 || horSectionTypeCodeList.PropValue == 4)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "This assembly is not applicable for Flatbar Section", "", "Guide.cs", 78);
                        return null;
                    }

                    Collection<object> sectionSizes = new Collection<object>();
                    //Get the section size
                    if (sectionFromRule == true)
                    {
                        GenericHelper.GetDataByRule("hsMrnRLGuideSectSize", (BusinessObject)support, out sectionSizes);
                        if (sectionSizes != null)
                        {
                            if (sectionSizes[0] == null)
                                verticalSectionSize = sectionSizes[1].ToString();
                            else
                                verticalSectionSize = sectionSizes[0].ToString();

                            if (includeHorSection == true)
                            {
                                if (sectionSizes[0] == null)
                                    horizontalSectionSize = sectionSizes[2].ToString();
                                else
                                    horizontalSectionSize = sectionSizes[1].ToString();
                            }
                        }
                    }
                    else
                    {
                        verticalSectionSize = verSectionSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(verSectionSizeCodeList.PropValue).DisplayName;
                        if (string.IsNullOrEmpty(verticalSectionSize))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Vertical Section size is not available.", "", "Guide.cs", 105);
                            return null;
                        }
                        if (includeHorSection == true)
                        {
                            horizontalSectionSize = horSectionSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(horSectionSizeCodeList.PropValue).DisplayName;
                            if (string.IsNullOrEmpty(horizontalSectionSize))
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Horizontal Section size is not available.", "", "Guide.cs", 113);
                                return null;
                            }
                        }
                    }
                    routeCount = SupportHelper.SupportedObjects.Count;

                    routeAngle = MarineAssemblyServices.GetRoutAngleAndIsRouteVertical(this, PI, routeCount, out isVerticalRoute);
                    if (routeCount > 1)
                    {
                        for (int index = 0; index > routeCount; index++)
                        {
                            if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[index]), Math.Round(routeAngle[index + 1])) == true)
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe slopes are not equal. Please check", "", "Guide.cs", 127);
                                return null;
                            }
                        }
                    }

                    //Get Pad part and Dimensions
                    string sectionCode = string.Empty, steelStd = string.Empty;
                    padProperties = MarineAssemblyServices.GetPadPartAndDimensions(this, verticalSectionSize, out sectionCode, out steelStd);
                    //GetSteelStandard    
                    string steelStandard = MarineAssemblyServices.GetSteelStandard(this, steelStd);

                    parts.Add(new PartInfo(VERTSECTION1, verticalSectionSize + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION2, verticalSectionSize + " " + steelStandard));

                    if (includeHorSection == true)
                    {
                        parts.Add(new PartInfo(HORSECTION1, horizontalSectionSize + " " + steelStandard));
                        parts.Add(new PartInfo(HORSECTION2, horizontalSectionSize + " " + steelStandard));
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
                //load standard bounding box.
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;
                string refPlane = string.Empty, verSec1Port = string.Empty, verSec2Port = string.Empty;

                PropertyValueCodelist overHangCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnGuideOH", "OverhangOpt");

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    refPlane = "BBSR";
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                }
                else
                {
                    refPlane = "BBR";
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }
                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height, leftPipeDiameter = 0.0, rightPipeDiameter = 0.0, leftStructAngle = 0.0, rightStructAngle = 0.0, structAngle = 0.0;;

                MarineAssemblyServices.GetLeftRightStructAngleAndIsVerticalStructure(this, PI, out isVerticalSruct, out leftStructAngle, out rightStructAngle);

                //Check to see what they are connecting to              
                string supportingType = MarineAssemblyServices.GetSupportingTypes(this);

                bool slopedSteel = false, slopedRoute = false;
                if (HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[0], 3) , 0)== false)
                    slopedRoute = true;
                else if (HgrCompareDoubleService.cmpdbl(Math.Round(leftStructAngle, 3) , 0)== false || HgrCompareDoubleService.cmpdbl(Math.Round(rightStructAngle, 3) , 0)== false)    //for sloped steel
                    slopedSteel = true;

                //For Chamfered Plate
                structAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, OrientationAlong.Global_Z);

                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct && supportingType.ToUpper() == "SLAB")
                    if (HgrCompareDoubleService.cmpdbl(Math.Round(structAngle, 3) , 0)== false)
                        slopedSteel = true;

                PropertyValueCodelist orientCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmOrient", "FrmOrient");
                string fromOrient = orientCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(orientCodeList.PropValue).DisplayName;
                //Get dimension of the BBX
                if (slopedSteel == true && SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    if (supportingType.ToUpper() == "SLAB" && fromOrient.ToUpper() == "PERPENDICULAR TO PIPE" && isVerticalRoute == false)
                    {
                        boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedVertical);
                        boundingBoxWidth = boundingBox.Width;
                        boundingBoxHeight = boundingBox.Height;
                    }
                    else
                    {
                        boundingBox = MarineAssemblyServices.GetBoundingBoxDimensions(this, routeCount);
                        boundingBoxWidth = boundingBox.Width;
                        boundingBoxHeight = boundingBox.Height;
                    }
                }
                else
                {
                    boundingBox = MarineAssemblyServices.GetBoundingBoxDimensions(this, routeCount);
                    boundingBoxWidth = boundingBox.Width;
                    boundingBoxHeight = boundingBox.Height;
                }

                //get route object collection on the reference plane
                if (SupportHelper.SupportedObjects.Count == 1)
                {                    
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    leftPipeDiameter = pipeInfo.OutsideDiameter;
                    rightPipeDiameter = leftPipeDiameter;
                }
                else
                {
                    int routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(refPlane, BoundingBoxEdge.Right);
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                    leftPipeDiameter = pipeInfo.OutsideDiameter;

                    routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(refPlane, BoundingBoxEdge.Left);
                    pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                    rightPipeDiameter = pipeInfo.OutsideDiameter;
                }

                //Auto Dimensioning of Supports
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                if (includeHorSection == true)
                {
                    MarineAssemblyServices.CreateDimensionNote(this, "Dim 1", componentDictionary[HORSECTION1], "BeginCap");
                    MarineAssemblyServices.CreateDimensionNote(this, "Dim 2", componentDictionary[HORSECTION1], "EndCap"); 
                }

                PropertyValueCodelist cardinalCP2CodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJOAhsSteelCP", "CP2");
                componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCP2CodeList.PropValue = 1, "IJOAhsSteelCP", "CP2");
                if (leftStructAngle <= 0.0)
                    verSec1Port = "EndCap";
                else
                    verSec1Port = "EndFace";

                PropertyValueCodelist cardinalCP1CodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJOAhsSteelCP", "CP1");
                componentDictionary[VERTSECTION2].SetPropertyValue(cardinalCP1CodeList.PropValue = 1, "IJOAhsSteelCP", "CP1");
                if (rightStructAngle <= 0.0)
                    verSec2Port = "BeginCap";
                else
                    verSec2Port = "BeginFace";

                MarineAssemblyServices.CreateDimensionNote(this, "Dim 3", componentDictionary[VERTSECTION1], verSec1Port);
                MarineAssemblyServices.CreateDimensionNote(this, "Dim 4", componentDictionary[VERTSECTION2], verSec2Port);
                MarineAssemblyServices.CreateDimensionNote(this, "Dim 5", componentDictionary[VERTSECTION1], "BeginCap");
                MarineAssemblyServices.CreateDimensionNote(this, "Dim 6", componentDictionary[VERTSECTION2], "EndCap");

                //Get Section Structure dimensions
                BusinessObject sectionPart;
                CrossSection crossSection;
                //Get SteelWidth and SteelDepth 
                double sectionWidth = 0.0, sectionDepth = 0.0, steelThickness = 0.0, vertSectionWidth = 0.0, vertSectionDepth = 0.0, vertSteelThickness = 0.0;
                if (includeHorSection == true)
                {
                    sectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                    crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                    sectionWidth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                    sectionDepth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                    steelThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                }

                sectionPart = componentDictionary[VERTSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                vertSectionWidth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                vertSectionDepth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                vertSteelThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                //Get the Current Location in the Route Connection Cycle
                string[] partNames = new string[4];
                partNames[0] = VERTSECTION1;
                partNames[1] = VERTSECTION2;
                if (includeHorSection == true)
                {
                    partNames[2] = HORSECTION1;
                    partNames[3] = HORSECTION2;
                }
                //get Codelisted Property Values for Hgr Beam
                MarineAssemblyServices.BeamCLProperties beamCLProperties = MarineAssemblyServices.GetBeamCLProperties(componentDictionary);
                //set HGR Beam Properties.
                MarineAssemblyServices.IntializeBeamProperties(componentDictionary, support, beamCLProperties, partNames);

                // Set Values of Part Occurance Attributes
                //set the Cutback for the vertcal section
                if (slopedSteel == true)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (includeHorSection == true)
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlFlg.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlFlg.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlFlg.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        }
                        else
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlFlg.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlFlg.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlFlg.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        }

                        if (Configuration == 2)
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(leftStructAngle, "IJOAhsCutback", "CutbackEndAngle");
                            componentDictionary[VERTSECTION2].SetPropertyValue(rightStructAngle, "IJOAhsCutback", "CutbackBeginAngle");
                        }
                        else if (Configuration == 1)
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(-(leftStructAngle), "IJOAhsCutback", "CutbackEndAngle");
                            componentDictionary[VERTSECTION2].SetPropertyValue(-(rightStructAngle), "IJOAhsCutback", "CutbackBeginAngle");
                        }
                    }
                }

                double[] pipeDiameter = new double[routeCount];
                for (int index = 0; index < routeCount; index++)
                {
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(index + 1);
                    pipeDiameter[index] = pipeInfo.OutsideDiameter;
                }
                //Set attribute values for Pads
                double vertSec1EndOverLength = 0.0, vertSec2BeginOverLength = 0.0;

                string[] padPartNames = new string[2];
                if (includeLeftPad == false && includeRightPad == true)
                {
                    vertSec1EndOverLength = vertSec1EndOverLength - padProperties.padThickness;
                    padPartNames[0] = RIGHTPAD;
                }
                else if (includeLeftPad == true && includeRightPad == false)
                {
                    vertSec2BeginOverLength = vertSec2BeginOverLength - padProperties.padThickness;
                    padPartNames[0] = LEFTPAD;
                }
                else if (includeLeftPad == true && includeRightPad == true)
                {
                    vertSec2BeginOverLength = vertSec2BeginOverLength - padProperties.padThickness;
                    vertSec1EndOverLength = vertSec1EndOverLength - padProperties.padThickness;
                    padPartNames[0] = RIGHTPAD;
                    padPartNames[1] = LEFTPAD;
                }
                if (includeLeftPad == true || includeRightPad == true)
                MarineAssemblyServices.SetPadProperties(componentDictionary, support, padProperties, padPartNames);
                //Set Assembly Attributes
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;

                //Set Frame Orientation attribute
                CodelistItem codeList = metadataManager.GetCodelistInfo("hsMrnCLFrmOrient", "UDP").GetCodelistItem(fromOrient);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnFrmOrient", "FrmOrient");

                //Set Section Size properties
                codeList = metadataManager.GetCodelistInfo("hsMrnCLSecSize", "UDP").GetCodelistItem(verticalSectionSize);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnVerSection", "VerSectionSize");

                if (includeHorSection == true)
                {
                    codeList = metadataManager.GetCodelistInfo("hsMrnCLSecSize", "UDP").GetCodelistItem(horizontalSectionSize);
                    support.SetPropertyValue(codeList.Value, "IJOAhsMrnHorSection", "HorSectionSize");
                }
                //Set Overhang attributes
                double hgrOverHangLeft = 0.0, hgrOverHangRight = 0.0, userDefOverHangLeft = 0.0, userDefOverHangRight = 0.0, overHangLeft = 0.0, overHangRight = 0.0, cutbackAngle1 = 0.0, cutbackAngle2 = 0.0;
                bool isPlatesloped = false;
                Collection<object> overHang = new Collection<object>();
                GenericHelper.GetDataByRule("hsMrnRLFrmOH", (BusinessObject)support, out overHang);
                if (overHang != null)
                {
                    if (overHang[0] == null)
                    {
                        hgrOverHangLeft = (double)overHang[1];
                        hgrOverHangRight = (double)overHang[2];
                    }
                    else
                    {
                        hgrOverHangLeft = (double)overHang[0];
                        hgrOverHangRight = (double)overHang[1];
                    }
                }
                if (includeHorSection == true)
                {
                    if (overHangCodeList.PropValue == 1)    //By Catalog Rule
                    {
                        overHangLeft = hgrOverHangLeft - rightPipeDiameter / 2;
                        overHangRight = hgrOverHangRight - leftPipeDiameter / 2;
                        support.SetPropertyValue(hgrOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                        support.SetPropertyValue(hgrOverHangRight, "IJOAhsMrnOHRight", "OverhangRight");
                    }
                    else if (overHangCodeList.PropValue == 2) //User Defined
                    {
                        userDefOverHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                        userDefOverHangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHRight", "OverhangRight")).PropValue;

                        support.SetPropertyValue(userDefOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                        support.SetPropertyValue(userDefOverHangRight, "IJOAhsMrnOHRight", "OverhangRight");

                        overHangLeft = userDefOverHangLeft - rightPipeDiameter / 2;
                        overHangRight = userDefOverHangRight - leftPipeDiameter / 2;
                    }
                }
                else
                {
                    if (overHangCodeList.PropValue == 1)
                    {
                        support.SetPropertyValue(rightPipeDiameter / 2, "IJOAhsMrnOHLeft", "OverhangLeft");
                        support.SetPropertyValue(leftPipeDiameter / 2, "IJOAhsMrnOHRight", "OverhangRight");
                    }
                    else if (overHangCodeList.PropValue == 2)
                    {
                        userDefOverHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                        userDefOverHangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHRight", "OverhangRight")).PropValue;

                        if (HgrCompareDoubleService.cmpdbl(userDefOverHangLeft , rightPipeDiameter / 2)==false || HgrCompareDoubleService.cmpdbl(userDefOverHangRight , leftPipeDiameter / 2)==false)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING:" + "Overhang Option not applicable when Horizontal section not available. Resetting the values.", "", "Guide.cs", 451);

                        support.SetPropertyValue(rightPipeDiameter / 2, "IJOAhsMrnOHLeft", "OverhangLeft");
                        support.SetPropertyValue(leftPipeDiameter / 2, "IJOAhsMrnOHRight", "OverhangRight");
                    }
                }

                //Get the structure surface
                if (supportingType.ToUpper() == "STEEL")
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        if (SupportHelper.SupportingObjects.Count < 2)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Select more than one Structure", "", "Guide.cs", 464);
                            return;
                        }
                }

                //Set the cutback angle for Place By Point while placing on Knuckle Plate
                string slabType = string.Empty, routePort = string.Empty, pipePosition = string.Empty; ;
                if (supportingType.ToUpper() == "SLAB")
                {
                    Matrix4X4 orientation = RefPortHelper.PortLCS("Structure");
                    Vector globalZ = new Vector(0, 0, 1);
                    double angle = Math.Acos(globalZ.Dot(orientation.ZAxis));
                    if ((angle < PI / 2 + 0.0001 && angle > PI / 2 - 0.0001) || (angle < -PI / 2 + 0.0001 && angle > -PI / 2 - 0.0001))
                        slabType = "WALL";
                    else
                        slabType = "SLAB";
                }
                //Get the Projection Length on surface
                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct && slabType.ToUpper() == "SLAB") //Place by point
                {
                    double sectionLength1, sectionLength2, routeStructDistance;
                    routeStructDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct);
                    sectionLength1 = routeStructDistance + pipeDiameter[0] / 2 + sectionDepth;
                    sectionLength2 = routeStructDistance + pipeDiameter[0] / 2 + sectionDepth;
                    //Initialize
                    for (int index = 0; index < partNames.Length; index++)
                        componentDictionary[partNames[index]].SetPropertyValue(0.0001, "IJUAHgrOccLength", "Length");

                    Position projectedVector = new Position(0, 0, 0), leftPosition = new Position(0, 0, 0), rightPosition = new Position(0, 0, 0), oPosOnStruct = new Position(0, 0, 0);
                    Vector normalprojectedVector = new Vector(0, 0, 0);
                    double leftOffset, rightOffset;

                    BusinessObject structObject = (BusinessObject)SupportHelper.SupportingObjects.First();
                    Matrix4X4 portPosition1 = new Matrix4X4(), portPosition2 = new Matrix4X4();
                    try
                    {
                        portPosition1 = RefPortHelper.PortLCS("BBR_High");
                        portPosition2 = RefPortHelper.PortLCS("BBRV_High");
                        if (includeHorSection == false)
                        {
                            overHangLeft = 0;
                            overHangRight = 0;
                        }
                        if (slopedSteel == true)
                        {
                            if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                            {
                                leftOffset = overHangLeft;
                                rightOffset = overHangRight + boundingBoxWidth;

                                Vector localVetor = portPosition2.YAxis;
                                localVetor.Length = leftOffset;
                                leftPosition = portPosition2.Origin.Offset(localVetor);

                                localVetor.Length = -rightOffset;
                                rightPosition = portPosition2.Origin.Offset(localVetor);

                                SupportingHelper.GetProjectedPointOnSurface(leftPosition, portPosition2.ZAxis, structObject, out projectedVector, out normalprojectedVector);

                                if (projectedVector != null)
                                    sectionLength2 = projectedVector.DistanceToPoint(leftPosition) + boundingBoxHeight + sectionDepth;

                                cutbackAngle1 = PI - normalprojectedVector.Angle(portPosition2.ZAxis, portPosition2.XAxis);

                                SupportingHelper.GetProjectedPointOnSurface(rightPosition, portPosition2.ZAxis, structObject, out projectedVector, out normalprojectedVector);

                                if (projectedVector != null)
                                    sectionLength1 = projectedVector.DistanceToPoint(rightPosition) + boundingBoxHeight + sectionDepth;

                                cutbackAngle2 = PI - normalprojectedVector.Angle(portPosition2.ZAxis, portPosition2.XAxis);
                            }
                            else
                            {
                                leftOffset = overHangLeft - vertSectionWidth;
                                rightOffset = overHangRight + boundingBoxWidth - vertSectionWidth;

                                Vector localVetor = portPosition1.YAxis;
                                localVetor.Length = leftOffset;
                                leftPosition = portPosition1.Origin.Offset(localVetor);

                                localVetor.Length = -rightOffset;
                                rightPosition = portPosition1.Origin.Offset(localVetor);

                                SupportingHelper.GetProjectedPointOnSurface(leftPosition, portPosition1.ZAxis, structObject, out projectedVector, out normalprojectedVector);

                                if (projectedVector != null)
                                    sectionLength2 = projectedVector.DistanceToPoint(leftPosition) + boundingBoxHeight + sectionDepth;

                                cutbackAngle1 = PI - normalprojectedVector.Angle(portPosition1.ZAxis, portPosition1.XAxis);

                                SupportingHelper.GetProjectedPointOnSurface(rightPosition, portPosition1.ZAxis, structObject, out projectedVector, out normalprojectedVector);

                                if (projectedVector != null)
                                    sectionLength1 = projectedVector.DistanceToPoint(rightPosition) + boundingBoxHeight + sectionDepth;

                                cutbackAngle2 = PI - normalprojectedVector.Angle(portPosition1.ZAxis, portPosition1.XAxis);
                            }
                        }
                        else
                        {
                            if (includeHorSection == true)
                            {
                                leftOffset = overHangLeft - vertSectionWidth;
                                rightOffset = overHangRight + boundingBoxWidth - vertSectionWidth;
                            }
                            else
                            {
                                leftOffset = overHangLeft;
                                rightOffset = overHangRight + boundingBoxWidth;
                            }

                            Vector localVetor = portPosition1.YAxis;
                            localVetor.Length = leftOffset;
                            leftPosition = portPosition1.Origin.Offset(localVetor);

                            localVetor.Length = -rightOffset;
                            rightPosition = portPosition1.Origin.Offset(localVetor);

                            SupportingHelper.GetProjectedPointOnSurface(leftPosition, portPosition1.ZAxis, structObject, out projectedVector, out normalprojectedVector);

                            if (projectedVector != null)
                                sectionLength2 = projectedVector.DistanceToPoint(leftPosition) + boundingBoxHeight + sectionDepth;

                            cutbackAngle1 = PI - normalprojectedVector.Angle(portPosition1.ZAxis, portPosition1.XAxis);
                            SupportingHelper.GetProjectedPointOnSurface(rightPosition, portPosition1.ZAxis, structObject, out projectedVector, out normalprojectedVector);

                            if (projectedVector != null)
                                sectionLength1 = projectedVector.DistanceToPoint(rightPosition) + boundingBoxHeight + sectionDepth;

                            cutbackAngle2 = PI - normalprojectedVector.Angle(portPosition1.ZAxis, portPosition1.XAxis);
                        }
                    }
                    catch
                    {
                    }

                    if (structAngle < 0.0001 && structAngle > -0.0001 || structAngle < PI + 0.0001 && structAngle > PI - 0.0001)
                        if (HgrCompareDoubleService.cmpdbl(cutbackAngle1, 0) == false && HgrCompareDoubleService.cmpdbl(cutbackAngle2, 0) == false)
                            isPlatesloped = true;
                        else
                            isPlatesloped = false;
                    else
                        isPlatesloped = false;
                    //Set the cutback type for BBR or BBRV
                    int cutbackType = 0;
                    if (isPlatesloped == true)
                        cutbackType = 3;
                    else
                    {
                        if (HgrCompareDoubleService.cmpdbl(cutbackAngle1, 0) == true && HgrCompareDoubleService.cmpdbl(cutbackAngle2, 0) == false)
                            cutbackType = 1;
                        else if (HgrCompareDoubleService.cmpdbl(cutbackAngle1, 0) == false && HgrCompareDoubleService.cmpdbl(cutbackAngle2, 0) == true)
                            cutbackType = 1;
                        else if (HgrCompareDoubleService.cmpdbl(cutbackAngle1, 0) == false && HgrCompareDoubleService.cmpdbl(cutbackAngle2, 0) == false)
                            if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                                if (isVerticalRoute == true)
                                    cutbackType = 1;
                                else
                                    cutbackType = 2;
                            else
                                cutbackType = 1;
                        else if (HgrCompareDoubleService.cmpdbl(Math.Round(cutbackAngle1, 3) , 0)==true || HgrCompareDoubleService.cmpdbl(Math.Round(cutbackAngle2, 3) , 0)==true)
                            if (HgrCompareDoubleService.cmpdbl(Math.Round(structAngle, 3) , 0) == false && Math.Round(structAngle, 3) < PI)      //For Sloped Deck
                                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                                    if (isVerticalRoute == true)
                                        cutbackType = 1;
                                    else
                                        cutbackType = 2;
                                else
                                    cutbackType = 1;
                            else
                                cutbackType = 1;
                        else
                            cutbackType = 1;
                    }

                    switch (cutbackType)
                    {
                        case 1:
                            routePort = "BBR_High";
                            if (Configuration == 1)
                            {
                                componentDictionary[VERTSECTION1].SetPropertyValue(sectionLength1, "IJUAHgrOccLength", "Length");
                                componentDictionary[VERTSECTION2].SetPropertyValue(sectionLength2, "IJUAHgrOccLength", "Length");
                                componentDictionary[VERTSECTION2].SetPropertyValue(-cutbackAngle1, "IJOAhsCutback", "CutbackBeginAngle");
                                componentDictionary[VERTSECTION1].SetPropertyValue(-cutbackAngle2, "IJOAhsCutback", "CutbackEndAngle");
                            }
                            else
                            {
                                componentDictionary[VERTSECTION1].SetPropertyValue(sectionLength2, "IJUAHgrOccLength", "Length");
                                componentDictionary[VERTSECTION2].SetPropertyValue(sectionLength1, "IJUAHgrOccLength", "Length");
                                componentDictionary[VERTSECTION2].SetPropertyValue(cutbackAngle2, "IJOAhsCutback", "CutbackBeginAngle");
                                componentDictionary[VERTSECTION1].SetPropertyValue(cutbackAngle1, "IJOAhsCutback", "CutbackEndAngle");
                            }
                            break;

                        case 2:
                            routePort = "BBRV_High";
                            double dTempCutbackAngle;
                            //For both cutback angle are different swap the angle
                            if (HgrCompareDoubleService.cmpdbl(cutbackAngle1 , cutbackAngle2) == false)
                            {
                                dTempCutbackAngle = cutbackAngle1;
                                cutbackAngle1 = cutbackAngle2;
                                cutbackAngle2 = dTempCutbackAngle;
                            }

                            if (Configuration == 1)
                            {
                                componentDictionary[VERTSECTION1].SetPropertyValue(sectionLength2, "IJUAHgrOccLength", "Length");
                                componentDictionary[VERTSECTION2].SetPropertyValue(sectionLength1, "IJUAHgrOccLength", "Length");
                                componentDictionary[VERTSECTION2].SetPropertyValue(cutbackAngle1, "IJOAhsCutback", "CutbackBeginAngle");
                                componentDictionary[VERTSECTION1].SetPropertyValue(cutbackAngle2, "IJOAhsCutback", "CutbackEndAngle");
                            }
                            else
                            {
                                componentDictionary[VERTSECTION1].SetPropertyValue(sectionLength1, "IJUAHgrOccLength", "Length");
                                componentDictionary[VERTSECTION2].SetPropertyValue(sectionLength2, "IJUAHgrOccLength", "Length");
                                componentDictionary[VERTSECTION2].SetPropertyValue(-cutbackAngle2, "IJOAhsCutback", "CutbackBeginAngle");
                                componentDictionary[VERTSECTION1].SetPropertyValue(-cutbackAngle1, "IJOAhsCutback", "CutbackEndAngle");

                            }
                            break;

                        case 3:      // If the Plate sloped and Pipe direction is on the slope not along slope
                            routePort = "BBR_High";
                            double tempLength1 = 0.0, tempLength2 = 0.0;
                            Vector localVetor = portPosition1.XAxis;
                            localVetor.Length = vertSectionDepth / 2;
                            leftPosition = portPosition1.Origin.Offset(localVetor);

                            localVetor.Length = -vertSectionDepth / 2;
                            rightPosition = portPosition1.Origin.Offset(localVetor);

                            SupportingHelper.GetProjectedPointOnSurface(leftPosition, portPosition1.ZAxis, structObject, out projectedVector, out normalprojectedVector);

                            if (projectedVector != null)
                                tempLength1 = projectedVector.DistanceToPoint(leftPosition);

                            cutbackAngle1 = PI - normalprojectedVector.Angle(portPosition1.ZAxis, portPosition1.XAxis);
                            SupportingHelper.GetProjectedPointOnSurface(rightPosition, portPosition1.ZAxis, structObject, out projectedVector, out normalprojectedVector);

                            if (projectedVector != null)
                                tempLength2 = projectedVector.DistanceToPoint(rightPosition) + pipeDiameter[0] / 2 + sectionDepth;

                            cutbackAngle2 = PI - normalprojectedVector.Angle(portPosition1.ZAxis, portPosition1.XAxis);
                            componentDictionary[VERTSECTION1].SetPropertyValue(sectionLength1 - (vertSectionDepth / 2) * Math.Tan(cutbackAngle2), "IJUAHgrOccLength", "Length");
                            componentDictionary[VERTSECTION2].SetPropertyValue(sectionLength2 - (vertSectionDepth / 2) * Math.Tan(cutbackAngle1), "IJUAHgrOccLength", "Length");

                            if (tempLength1 > tempLength2)
                                if (Configuration == 1)
                                {
                                    componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    componentDictionary[VERTSECTION2].SetPropertyValue(cutbackAngle1, "IJOAhsCutback", "CutbackBeginAngle");
                                    componentDictionary[VERTSECTION1].SetPropertyValue(-cutbackAngle2, "IJOAhsCutback", "CutbackEndAngle");


                                }
                                else
                                {
                                    componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    componentDictionary[VERTSECTION2].SetPropertyValue(-cutbackAngle2, "IJOAhsCutback", "CutbackBeginAngle");
                                    componentDictionary[VERTSECTION1].SetPropertyValue(cutbackAngle1, "IJOAhsCutback", "CutbackEndAngle");


                                }
                            else
                                if (Configuration == 1)
                                {
                                    componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    componentDictionary[VERTSECTION2].SetPropertyValue(-cutbackAngle1, "IJOAhsCutback", "CutbackBeginAngle");
                                    componentDictionary[VERTSECTION1].SetPropertyValue(cutbackAngle2, "IJOAhsCutback", "CutbackEndAngle");

                                }
                                else
                                {
                                    componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    componentDictionary[VERTSECTION2].SetPropertyValue(cutbackAngle2, "IJOAhsCutback", "CutbackBeginAngle");
                                    componentDictionary[VERTSECTION1].SetPropertyValue(-cutbackAngle1, "IJOAhsCutback", "CutbackEndAngle");

                                }
                            break;
                    }
                }
                else
                {
                    routePort = "BBR_High";
                }
                //Set Length of horzontal member
                double horSectionLength = 0.0;
                if (includeHorSection == true)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        double tempDistance = 0.0;
                        if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 2)      //By Catalog Rule
                            tempDistance = boundingBoxWidth;

                        horSectionLength = tempDistance + overHangLeft + overHangRight;
                        componentDictionary[HORSECTION1].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                        componentDictionary[HORSECTION2].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                    }
                    else                        //for Place by point
                    {
                        if (supportingType.ToUpper() == "SLAB")
                            if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 2)          //By catalog OR User Defined
                                horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                            else if (supportingType.ToUpper() == "STEEL")
                                if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 2)         //By catalog
                                    horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                                else if (supportingType.ToUpper() == "STEEL-SLAB" || supportingType.ToUpper() == "SLAB-STEEL")
                                    if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 2)         //by catalog
                                        horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                        componentDictionary[HORSECTION1].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                        componentDictionary[HORSECTION2].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                    }
                }
                //set the properties for wrapping connections
                if (includeHorSection == true)
                {
                    cardinalCP1CodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJOAhsSteelCP", "CP1");
                    cardinalCP2CodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJOAhsSteelCP", "CP2");
                    componentDictionary[VERTSECTION1].SetPropertyValue(-PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[VERTSECTION1].SetPropertyValue(-PI / 2, "IJOAhsEndCap", "EndCapRotZ");
                    componentDictionary[VERTSECTION2].SetPropertyValue(-PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[VERTSECTION2].SetPropertyValue(-PI / 2, "IJOAhsEndCap", "EndCapRotZ");
                    componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCP1CodeList.PropValue = 3, "IJOAhsSteelCP", "CP1");
                    componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCP2CodeList.PropValue = 3, "IJOAhsSteelCP", "CP2");
                    componentDictionary[VERTSECTION2].SetPropertyValue(cardinalCP2CodeList.PropValue = 3, "IJOAhsSteelCP", "CP2");
                    componentDictionary[VERTSECTION2].SetPropertyValue(cardinalCP1CodeList.PropValue = 3, "IJOAhsSteelCP", "CP1");
                }
                //Set the cutback angle for Sloped Pipe
                Matrix4X4 routePortOrientation = new Matrix4X4(), structPortOrientation = new Matrix4X4();
                routePortOrientation = RefPortHelper.PortLCS("Route");
                structPortOrientation = RefPortHelper.PortLCS("Structure");

                if (HgrCompareDoubleService.cmpdbl(structPortOrientation.Origin.Z , routePortOrientation.Origin.Z) == false || HgrCompareDoubleService.cmpdbl(structPortOrientation.Origin.X , routePortOrientation.Origin.X) == false || HgrCompareDoubleService.cmpdbl(structPortOrientation.Origin.Y , routePortOrientation.Origin.Y) == false)
                {
                    if (structPortOrientation.Origin.Z > routePortOrientation.Origin.Z || structPortOrientation.Origin.X > routePortOrientation.Origin.X || structPortOrientation.Origin.Y > routePortOrientation.Origin.Y)
                        pipePosition = "BELOW";
                    else
                        pipePosition = "ABOVE";
                }

                if (slopedRoute == true)
                {
                    if ((SupportHelper.PlacementType == PlacementType.PlaceByStruct && fromOrient.ToUpper() == "PERPENDICULAR TO PIPE") || SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {

                        double angle = 0.0;
                        if (pipePosition == "ABOVE")
                            if (Configuration == 1)
                                angle = -routeAngle[0];
                            else if (Configuration == 2)
                                angle = routeAngle[0];

                            else if (pipePosition == "BELOW")
                                if (Configuration == 1)
                                    angle = routeAngle[0];
                                else if (Configuration == 2)
                                    angle = -routeAngle[0];

                        if (Configuration == 1)
                        {
                            componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[VERTSECTION2].SetPropertyValue(-angle, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[VERTSECTION1].SetPropertyValue(angle, "IJOAhsCutback", "CutbackEndAngle");
                        }
                        else
                        {
                            componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlWeb, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlWeb, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[VERTSECTION2].SetPropertyValue(-angle, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[VERTSECTION1].SetPropertyValue(angle, "IJOAhsCutback", "CutbackEndAngle");
                        }
                    }
                }
                //Create Joints
                //For Sloped Structure
                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = MarineAssemblyServices.GetIsLugEndOffsetApplied(this);

                string[] structPort = new string[2];
                structPort = MarineAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0], rightStructPort = structPort[1];

                //Get Route Port               
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routePort = MarineAssemblyServices.GetRoutePort(this, support, slopedRoute, slopedSteel, isVerticalRoute, isVerticalSruct);

                if (includeHorSection == false)
                    if (routePort == "BBR_High" || routePort == "BBSR_High")
                        if (Configuration == 1)
                            componentDictionary[VERTSECTION1].SetPropertyValue(PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                        else
                            componentDictionary[VERTSECTION1].SetPropertyValue(-PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                    else
                        if (Configuration == 1)
                            componentDictionary[VERTSECTION1].SetPropertyValue(-PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                        else
                            componentDictionary[VERTSECTION1].SetPropertyValue(PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");

                Plane routeHorPlaneA = new Plane(), routeHorPlaneB = new Plane(), sectionPlaneA = new Plane(), sectionPlaneB = new Plane();
                Axis routeHorAxisA = new Axis(), routeHorAxisB = new Axis(), sectionAxisA = new Axis(), sectionAxisB = new Axis();
                routeHorPlaneA = Plane.ZX;
                routeHorPlaneB = Plane.NegativeXY;
                routeHorAxisA = Axis.X;

                sectionPlaneA = Plane.XY;
                sectionPlaneB = Plane.NegativeXY;
                sectionAxisA = Axis.X;
                string horSectionPort = string.Empty;
                double horRoutePlaneOffset = 0.0, horRouteAxisOffset = 0.0, verRouteOriginOffset = 0.0, verRouteAxisOffset = 0.0;
                horRoutePlaneOffset = -boundingBoxHeight;
                if (includeHorSection == true)
                {
                    if (Configuration == 1)
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            routeHorAxisB = Axis.NegativeX;
                            horSectionPort = "BeginCap";
                            horRouteAxisOffset = overHangLeft;
                        }
                        else
                        {
                            routeHorAxisB = Axis.X;
                            horSectionPort = "EndCap";
                            horRouteAxisOffset = -overHangLeft;
                        }
                    else if (Configuration == 2)

                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            routeHorAxisB = Axis.X;
                            horSectionPort = "EndCap";
                            horRouteAxisOffset = -overHangLeft;
                        }
                        else
                        {
                            routeHorAxisB = Axis.NegativeX;
                            horSectionPort = "BeginCap";
                            horRouteAxisOffset = overHangLeft;
                        }
                }
                else        // Include Horizontal section =false
                {
                    if (Configuration == 1)
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            sectionAxisB = Axis.Y;
                            verRouteAxisOffset = boundingBoxWidth;
                            verRouteOriginOffset = vertSectionDepth / 2;
                        }
                        else
                        {
                            sectionAxisB = Axis.NegativeY;
                            verRouteAxisOffset = 0;
                            verRouteOriginOffset = -vertSectionDepth / 2;
                        }
                    else if (Configuration == 2)
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            sectionAxisB = Axis.NegativeY;
                            verRouteAxisOffset = 0;
                            verRouteOriginOffset = -vertSectionDepth / 2;
                        }
                        else
                        {
                            sectionAxisB = Axis.Y;
                            verRouteAxisOffset = boundingBoxWidth;
                            verRouteOriginOffset = vertSectionDepth / 2;
                        }
                }
                if (includeHorSection == true)
                {
                    //Add joint Between Horizontal Section and BoundingBox                  
                    JointHelper.CreateRigidJoint(HORSECTION1, horSectionPort, "-1", routePort, routeHorPlaneA, routeHorPlaneB, routeHorAxisA, routeHorAxisB, horRoutePlaneOffset, horRouteAxisOffset, -vertSectionDepth / 2);

                    //Add Joint Between the Horizontal1 and Horizontal2 Beams
                    JointHelper.CreateRigidJoint(HORSECTION1, "Neutral", HORSECTION2, "Neutral", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -boundingBoxHeight - sectionDepth, 0);

                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HORSECTION1, "BeginCap", VERTSECTION1, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, sectionDepth, horSectionLength, 0);//9381

                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HORSECTION1, "EndCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.X, sectionDepth, -horSectionLength, 0);
                }
                else
                {
                    //Add Joint Between the Route and Vertical Beams
                    JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", "-1", routePort, Plane.XY, Plane.XY, Axis.X, Axis.X, boundingBoxHeight, verRouteAxisOffset, verRouteOriginOffset);

                    //Add Joint Between the Vertical 1 and Vertical 2 Beams
                    JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", VERTSECTION1, "BeginCap", sectionPlaneA, sectionPlaneB, sectionAxisA, sectionAxisB, 0, 0, -boundingBoxWidth);
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct || (SupportHelper.PlacementType != PlacementType.PlaceByStruct && slabType.ToUpper() != "SLAB"))
                {
                    //Add Joint Between the Ports of Vertical Beam 1
                    JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);//11757

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);//11757

                    //Add Joint Between the Ports of Vertical Beam 1
                    JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", "-1", leftStructPort, Plane.XY);

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, "BeginCap", "-1", rightStructPort, Plane.XY);
                }

                if (includeLeftPad == false && includeRightPad == true)
                    //Add Joint Between the Left and the Vertical Section 1
                    JointHelper.CreateRigidJoint(RIGHTPAD, "Port1", VERTSECTION1, "EndFace", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                else if (includeLeftPad == true && includeRightPad == false)
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "Port2", VERTSECTION2, "BeginFace", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                else if (includeRightPad == true && includeLeftPad == true)
                {
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(RIGHTPAD, "Port1", VERTSECTION1, "EndFace", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "Port2", VERTSECTION2, "BeginFace", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                }
                componentDictionary[VERTSECTION1].SetPropertyValue(vertSec1EndOverLength, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(vertSec2BeginOverLength, "IJUAHgrOccOverLength", "BeginOverLength");
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
                    if (includeHorSection == true)
                        routeConnections.Add(new ConnectionInfo(HORSECTION1, 1)); // partindex, routeindex
                    else
                        routeConnections.Add(new ConnectionInfo(VERTSECTION1, 1));

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

                if (string.IsNullOrEmpty(bomDescription))
                    bomDescription = (string)((PropertyValueString)support.GetPropertyValue("IJOAhsMrnSupType", "SupType")).PropValue + "-" + MarineAssemblyServices.GetLargePipeDiameter(this);
                else
                    bomDescription = bomDescriptionValue;

                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - RectHorTypeB" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

