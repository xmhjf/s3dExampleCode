//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   FieldSupports.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.FieldSupports
//   Author       : Rajeswari
//   Creation Date: 02-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 02-Sep-2013  Rajeswari   CR-CP-224478 Convert FlSample_Supports to C# .Net 
// 28-Apr-2015  PVK			Resolve Coverity issues found in April
// 17-Dec-2015  Ramya       TR 284319	Multiple Record exception dumps are created on copy pasting supports
// 29-Apr-2016  PVK         TR-CP-292883	Resolve the issues found in FL Sample Assemblies
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class FieldSupports : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        public string HORIZONTALSECTION = "HORIZONTALSECTION";
        public string ENDPLATE1 = "ENDPLATE1";
        public string ENDPLATE2 = "ENDPLATE2";
        public string BASEPLATE = "BASEPLATE";
        public string CONNECTION1 = "CONNECTION1";
        public string CONNECTION2 = "CONNECTION2";

        Double basePlateWidth, basePlateDepth, basePlateThickness, endPlateDepth, endPlateWidth, endPlateThickness, outsidePipeDiaMeter, horizontalLength, boundingBoxWidth, boundingBoxHeight;
        string basePlate, endPlates, reverseSec, supportName, outsidePipePort;
        object[] outSideRouteParams;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    string sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_SecSize", "SecSize")).PropValue;
                    basePlate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_Plate", "Plate")).PropValue;
                    basePlateWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_BP_Width", "BP_Width")).PropValue;
                    basePlateDepth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_BP_Depth", "BP_Depth")).PropValue;
                    basePlateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_BP_Thickness", "BP_Thickness")).PropValue;

                    int endPlate = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrFl_EndPlate", "EndPlate")).PropValue;
                    MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                    endPlates = metadataManager.GetCodelistInfo("FlSampleYesNo", "UDP").GetCodelistItem(endPlate).DisplayName;
                    endPlateWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_EPWidth", "EPWidth")).PropValue;
                    endPlateDepth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_EPDepth", "EPDepth")).PropValue;
                    endPlateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_EPThickness", "EPThickness")).PropValue;

                    reverseSec = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_ReverseSec", "ReverseSec")).PropValue;
                    supportName = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_SupportName", "SupportName")).PropValue;

                    // Each assembly has at least these two parts
                    parts.Add(new PartInfo(HORIZONTALSECTION, sectionSize));

                    // add in a baseplate if we need one
                    if (basePlate.ToUpper().Equals("WITH"))
                    {
                        parts.Add(new PartInfo(BASEPLATE, "Utility_USER_FIXED_BOX_1"));

                        if (HgrCompareDoubleService.cmpdbl(basePlateWidth, 0) == true || HgrCompareDoubleService.cmpdbl(basePlateDepth, 0) == true || HgrCompareDoubleService.cmpdbl(basePlateThickness, 0) == true)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Dimensions for the base plate are not properly defined.", "", "FieldSupports.cs", 84);
                            basePlate = "No";
                            support.SetPropertyValue(basePlate, "IJUAHgrFL_Plate", "Plate");
                        }
                    }
                    // ==========================
                    // 1. Load standard bounding box definition
                    // ==========================
                    BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                    BoundingBox boundingBox;
                    // ==========================
                    // 2. Get bounding box boundary objects dimension information
                    // ==========================
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                    else
                        boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                    // ==========================
                    // 3. retrieve dimension of the bounding box
                    // ==========================
                    // Get route box geometry
                    //  ____________________
                    // |                    |
                    // |  ROUTE BOX BOUND   | dHeight
                    // |____________________|
                    //        dWidth
                    boundingBoxWidth = boundingBox.Width;
                    boundingBoxHeight = boundingBox.Height;

                    outSideRouteParams = FLSampleSupportServices.GetOutsideRouteProps(this, 2);
                    outsidePipeDiaMeter = Convert.ToDouble(outSideRouteParams[1]);
                    outsidePipePort = (outSideRouteParams[3]).ToString();
                    horizontalLength = RefPortHelper.DistanceBetweenPorts(outsidePipePort, "Structure", PortDistanceType.Horizontal);

                    int endPlatesNeeded = 2; // To Set No to EndPlate Codelist
                    // add in a endplates if we need one
                    if (endPlates.ToUpper().Equals("YES"))
                    {
                        if (supportName.ToUpper().Equals("FS4T") || supportName.ToUpper().Equals("USS"))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Endplates are not used for this support.", "", "FieldSupports.cs", 127);
                             support.SetPropertyValue(endPlatesNeeded, "IJOAHgrFL_EndPlate", "EndPlate");
                        }
                        else
                        {
                            if (HgrCompareDoubleService.cmpdbl(endPlateDepth, 0) == true || HgrCompareDoubleService.cmpdbl(endPlateDepth, 0) == true || HgrCompareDoubleService.cmpdbl(endPlateThickness, 0) == true)
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Dimensions for the end plate are not properly defined.", "", "FieldSupports.cs", 134);
                                support.SetPropertyValue(endPlatesNeeded, "IJOAHgrFL_EndPlate", "EndPlate");
                            }
                            else
                            {
                                parts.Add(new PartInfo(ENDPLATE1, "Utility_USER_FIXED_BOX_1"));
                                if ((boundingBoxHeight > horizontalLength + outsidePipeDiaMeter / 2) || (boundingBoxWidth > horizontalLength) || (FLSampleSupportServices.GetWidestDistance(this) > horizontalLength))
                                    parts.Add(new PartInfo(ENDPLATE2, "Utility_USER_FIXED_BOX_1"));
                            }
                        }
                    }

                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {
                        parts.Add(new PartInfo(CONNECTION1, "Log_Conn_Part_1"));

                        if (reverseSec.ToUpper().Equals("YES"))
                            parts.Add(new PartInfo(CONNECTION2, "Log_Conn_Part_1"));
                    }
                    // Return the collection of Catalog Parts
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)  //It wont place by reference
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Can not place the support in Place-By-Reference.", "", "FieldSupports.cs", 172);
                    return;
                }
                
                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Wall))
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Slab"))   //It wont place by point with wall or slab
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Can not place the support with slab and wall.", "", "FieldSupports.cs", 178);
                    return;
                }

                string sectionType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_SecType", "SecType")).PropValue;
                string connectTo = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_ConnectTo", "ConnectTo")).PropValue;
                double shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrFl_ShoeH", "ShoeH")).PropValue;
                double overhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_OverHang", "OverHang")).PropValue;
                double freeEndOverLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_OverLength", "OverLength")).PropValue;
                double maxSpan = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_MaxSpan", "MaxSpan")).PropValue;
                double maxAssemblyLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_MaxAssyLength", "MaxAssyLength")).PropValue;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Double pipeOD = pipeInfo.OutsideDiameter;

                // Get interface for accessing items on the collection of Part Occurences
                BusinessObject horizontalSectionPart = componentDictionary[HORIZONTALSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                double steelWidth = crossSection.Width;
                double steelDepth = crossSection.Depth;
                outsidePipePort = "Route";
                string partSheet = supportName; string partNumber = string.Empty;

                if (connectTo.ToUpper().Equals("FLANGE FACE"))
                    partNumber = supportName + "_Flange";
                else
                    partNumber = supportName + "_Web";

                double ndFrom = FLSampleSupportServices.GetDataByCondition(partSheet, "IJHgrSupportDefinition", "NDFrom", "IJDPart", "PartNumber", partNumber);
                double ndTo = FLSampleSupportServices.GetDataByCondition(partSheet, "IJHgrSupportDefinition", "NDTo", "IJDPart", "PartNumber", partNumber);
                outSideRouteParams = FLSampleSupportServices.GetOutsideRouteProps(this, 2);
                outsidePipeDiaMeter = Convert.ToDouble(outSideRouteParams[1]);
                outsidePipePort = Convert.ToString(outSideRouteParams[3]);
                double outPipeOrientation = (RefPortHelper.AngleBetweenPorts(outsidePipePort, PortAxisType.X, OrientationAlong.Global_Z) * 180 / Math.PI);

                if (((HgrCompareDoubleService.cmpdbl(outPipeOrientation , 0)==true && outPipeOrientation < 0.00001) || (outPipeOrientation > 179.999999 && outPipeOrientation < 180.00001)) && SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    horizontalLength = FLSampleSupportServices.GetHorizontalDistanceBetweenPorts(this, outsidePipePort, "Structure");
                else
                    horizontalLength = RefPortHelper.DistanceBetweenPorts(outsidePipePort, "Structure", PortDistanceType.Horizontal_Perpendicular);

                if (SupportHelper.PlacementType == PlacementType.PlaceByPoint && Convert.ToInt32(FLSampleSupportServices.CheckPipeOrientation(this)[1]) == 1)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Use Place-by-Structure for horizontal pipes.", "", "FieldSupports.cs", 220);
                    return;
                }

                // Check for valid pipe size
                double pipeND = 0;
                for (int i = 1; i <= SupportHelper.SupportedObjects.Count; i++)
                {
                    pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    pipeND = pipeInfo.NominalDiameter.Size;

                    if ((supportName.Equals("FS4B") || supportName.Equals("FS5B")) && i > 1)
                    {
                        if (pipeND < (0.075 - 0.000001) || pipeND > (ndTo + 0.000001))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "FieldSupports.cs", 235);
                            return;
                        }
                    }
                    else
                    {
                        if (pipeND < (ndFrom - 0.000001) || pipeND > (ndTo + 0.000001))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "FieldSupports.cs", 243);
                            return;
                        }
                    }
                }

                if (horizontalLength > (maxSpan + 0.000001))
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Distance between structure and farthest route is more than maximum span.", "", "FieldSupports.cs", 250);

                // If there are no horizonal routes, check if there are vertical routes on both sides of strucutre.
                // If there are vertical routes on both sides, then get their distance from strucuture
                double distStructToVerticalRoute = 0;
                string verticalRoutePort = FLSampleSupportServices.GetDistanceAndPortForVerticalRoute(this, outsidePipePort)[1].ToString();

                if (boundingBoxWidth > horizontalLength)
                    distStructToVerticalRoute = Convert.ToDouble(FLSampleSupportServices.GetDistanceAndPortForVerticalRoute(this, outsidePipePort)[0]);

                // If dDistStructToVerticalRoute is not zero then there is vertical route on other sides of sturcutre
                double lengthOutsideBBX = Convert.ToDouble(FLSampleSupportServices.GetLengthAndDiameterOutsideBBX(this)[0]);
                double distFromBBXHigh = Convert.ToDouble(FLSampleSupportServices.GetLengthAndDiameterOutsideBBX(this)[1]);
                double distFromBBXLow = Convert.ToDouble(FLSampleSupportServices.GetLengthAndDiameterOutsideBBX(this)[2]);

                double distBBXLowToStruct = RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Structure", PortDistanceType.Horizontal_Perpendicular);
                double distBBXHighToStruct = RefPortHelper.DistanceBetweenPorts("BBSR_High", "Structure", PortDistanceType.Horizontal_Perpendicular);

                Boolean isTwoSides = false;
                if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false || HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                {
                    if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false)
                    {
                        if (distBBXHighToStruct < distFromBBXHigh && (lengthOutsideBBX + FLSampleSupportServices.GetGreaterValue(boundingBoxWidth, boundingBoxHeight) > distFromBBXHigh + distBBXHighToStruct))
                            isTwoSides = true;
                        else if (distBBXHighToStruct < boundingBoxHeight)
                            isTwoSides = true;
                    }
                    else
                    {
                        if (distBBXLowToStruct < distFromBBXLow)
                        {
                            if (lengthOutsideBBX + FLSampleSupportServices.GetGreaterValue(boundingBoxWidth, boundingBoxHeight) > distFromBBXLow + distBBXLowToStruct)
                                isTwoSides = true;
                            else if (distBBXLowToStruct < boundingBoxHeight)
                                isTwoSides = true;
                        }
                    }
                }
                else if (boundingBoxHeight > horizontalLength + outsidePipeDiaMeter / 2)
                    isTwoSides = true;
                else if (distStructToVerticalRoute > 0)
                    isTwoSides = true;

                double outSideDiaNearBBXLow = 0, outSideDiaNearBBXHigh = 0;
                if (isTwoSides == true)
                {
                    outSideDiaNearBBXHigh = Convert.ToDouble(FLSampleSupportServices.GetOutSideRouteDiameterOnOtherSideStruct(this, outsidePipePort, outsidePipeDiaMeter)[0]);
                    outSideDiaNearBBXLow = Convert.ToDouble(FLSampleSupportServices.GetOutSideRouteDiameterOnOtherSideStruct(this, outsidePipePort, outsidePipeDiaMeter)[1]);
                }

                double structSteelDepth = 0, structSteelWidth = 0, structFlangeThickness = 0, structWebThickness = 0;
                if (supportingType == "Steel")
                {
                    structSteelWidth = SupportingHelper.SupportingObjectInfo(1).Width;
                    structSteelDepth = SupportingHelper.SupportingObjectInfo(1).Depth;
                    structFlangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                    structWebThickness = SupportingHelper.SupportingObjectInfo(1).WebThickness;
                }

                // NOTE - we need to recompile this in SP4 and use the new HORZ_PERP distance to get the correct distance
                double offsetfromBBXtoVertFarPipe = RefPortHelper.DistanceBetweenPorts(outsidePipePort, "BBSR_Low", PortDistanceType.Horizontal);

                // we only need to use this offset if we have a vertical pipe that farther out than any other pipe
                if (offsetfromBBXtoVertFarPipe < outsidePipeDiaMeter / 2 + 0.0001 && offsetfromBBXtoVertFarPipe > outsidePipeDiaMeter / 2 - 0.0001)
                    offsetfromBBXtoVertFarPipe = 0;

                double beamLength = 0, structureOffset = 0, endPlateOffset1 = 0, endPlateOffset2 = 0, basePlateOffset = 0, offsetFromMiddle = 0, structWidthOffset = 0, structuralDim = 0;

                if (connectTo.ToUpper().Equals("FLANGE FACE"))
                {
                    structuralDim = structSteelWidth;
                    structWidthOffset = 0;
                }
                else if (connectTo.ToUpper().Equals("ACROSS WEB"))
                {
                    structuralDim = structSteelDepth;
                    offsetFromMiddle = structSteelWidth / 2;
                    basePlateThickness = offsetFromMiddle - structWebThickness / 2;
                    basePlateOffset = offsetFromMiddle - structWebThickness / 2;
                    structWidthOffset = offsetFromMiddle - structWebThickness / 2;
                }
                int facenumber;
                if (basePlate.ToUpper().Equals("WITH"))
                {
                    string baseplateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, basePlateThickness, UnitName.DISTANCE_MILLIMETER) + "Plate Steel," + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, basePlateWidth, UnitName.DISTANCE_MILLIMETER) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, basePlateDepth, UnitName.DISTANCE_MILLIMETER);
                    double basePlateWeight = 7900 * basePlateWidth * basePlateDepth * basePlateThickness;

                    // NOTE: Change this method from ConfigHlpr to HH
                    componentDictionary[BASEPLATE].SetPropertyValue(basePlateDepth, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                    componentDictionary[BASEPLATE].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                    componentDictionary[BASEPLATE].SetPropertyValue(basePlateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                    componentDictionary[BASEPLATE].SetPropertyValue(baseplateBOM, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");
                    componentDictionary[BASEPLATE].SetPropertyValue(basePlateWeight, "IJOAHgrUtility_USER_FIXED_BOX", "DryWt");

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        basePlateOffset = basePlateThickness;
                        if (!isTwoSides)
                            beamLength = horizontalLength + overhang + basePlateWidth / 2 + freeEndOverLength;
                        else
                        {
                            if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                                beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX;
                            else if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) == true && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                                beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX - outSideDiaNearBBXHigh / 2;
                            else if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) == true)
                                beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX - outSideDiaNearBBXLow / 2;
                            else
                                beamLength = boundingBoxHeight + 2 * overhang - outSideDiaNearBBXHigh / 2 - outSideDiaNearBBXLow / 2;
                        }

                        if (beamLength > (maxAssemblyLength + 0.00001))
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Assembly length is more than maximum assembly length.", "", "FieldSupports.cs", 363);

                        structureOffset = basePlateWidth / 2 + freeEndOverLength;

                        if (sectionType.Equals("L"))
                        {
                            if (!isTwoSides)
                            {
                                if (Configuration == 1 || Configuration == 3)
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, 0, -basePlateWidth / 2 - freeEndOverLength, steelDepth / 2);
                                else
                                    // Add Connection for the end of the Angled Beam
                                    JointHelper.CreateRigidJoint(BASEPLATE, "StartOther", HORIZONTALSECTION, "EndCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0, -steelDepth / 2, basePlateWidth / 2 + freeEndOverLength);
                            }
                            else
                            {
                                // do this if placing on w section
                                if (HgrCompareDoubleService.cmpdbl(structFlangeThickness , 0) == false)
                                {
                                    if (Configuration == 1 || Configuration == 3)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, boundingBoxWidth + steelDepth / 2, basePlateThickness, distBBXLowToStruct);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, -steelDepth / 2, basePlateThickness, distBBXLowToStruct);
                                }
                                else
                                {
                                    if (Configuration == 1)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, boundingBoxWidth + steelDepth / 2, basePlateThickness, distBBXLowToStruct);
                                    else if (Configuration == 2)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, basePlateThickness, -distBBXHighToStruct);
                                    else if (Configuration == 3)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, -boundingBoxWidth - steelDepth / 2, basePlateThickness, -distBBXHighToStruct);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, -steelDepth / 2, basePlateThickness, distBBXLowToStruct);
                                }
                            }
                        }
                        else
                        {
                            facenumber = 513;

                            if ((SupportHelper.SupportingObjects.Count != 0))
                                facenumber = SupportingHelper.SupportingObjectInfo(1).FaceNumber;

                            if (reverseSec.ToUpper().Equals("NO"))
                            {
                                if (!isTwoSides)
                                {
                                    // If the support is placed on concreate colimn
                                    if (HgrCompareDoubleService.cmpdbl(structFlangeThickness , 0) == true && (facenumber == 513 || facenumber == 514))
                                    {
                                        if (Configuration == 1 || Configuration == 2)
                                            JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, 0.0, -basePlateWidth / 2);
                                        else
                                            JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, 0.0, basePlateWidth / 2);
                                    }
                                    else
                                    {
                                        if (Configuration == 1 || Configuration == 2)
                                            JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, 0.0, -basePlateWidth / 2);
                                        else
                                            JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, 0.0, basePlateWidth / 2);
                                    }
                                }
                                else
                                {
                                    if (Configuration == 1 || Configuration == 4)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", BASEPLATE, "EndOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, boundingBoxWidth + steelDepth / 2, 0.0, distBBXLowToStruct);
                                    else if (Configuration == 2 || Configuration == 3)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", BASEPLATE, "EndOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, -boundingBoxWidth - steelDepth / 2, 0.0, -distBBXHighToStruct);
                                }
                            }
                            else
                            {
                                if (!isTwoSides)
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, structWidthOffset + steelWidth + basePlateThickness, freeEndOverLength - beamLength + structSteelWidth / 2);
                                else
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, structWidthOffset + steelWidth + basePlateThickness, -horizontalLength - basePlateWidth / 2);
                            }

                        }
                    }
                    else
                    {
                        // We needed to add the code for By Point Base Plates in below
                        basePlateOffset = basePlateThickness;
                        if (!isTwoSides)
                            beamLength = horizontalLength + overhang + basePlateWidth / 2 + freeEndOverLength;
                        else
                        {
                            if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                            {
                                if (boundingBoxHeight > boundingBoxWidth)
                                    beamLength = boundingBoxHeight + 2 * overhang - outSideDiaNearBBXHigh / 2 - outSideDiaNearBBXLow / 2;
                                else
                                    beamLength = boundingBoxWidth + 2 * overhang - outSideDiaNearBBXHigh / 2 - outSideDiaNearBBXLow / 2;
                            }
                            else if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                                beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX;
                            else if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) == true && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                                beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX - outSideDiaNearBBXHigh / 2;
                            else if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) == true)
                                beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX - outSideDiaNearBBXLow / 2;
                            else
                                beamLength = boundingBoxHeight + 2 * overhang - outSideDiaNearBBXHigh / 2 + outSideDiaNearBBXLow / 2;
                        }

                        if (beamLength > (maxAssemblyLength + 0.00001))
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Assembly length is more than maximum assembly length.", "", "FieldSupports.cs", 468);

                        structureOffset = basePlateWidth / 2 + freeEndOverLength;
                    }
                }
                else  // bp with
                {
                    if (!isTwoSides)
                        beamLength = horizontalLength + overhang + structuralDim / 2 + freeEndOverLength;
                    else
                    {
                        // First - Calculate beam length if there are no horizontal routes
                        if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        {
                            if (boundingBoxHeight > boundingBoxWidth)
                                beamLength = boundingBoxHeight + 2 * overhang - outSideDiaNearBBXHigh / 2 - outSideDiaNearBBXLow / 2;
                            else
                                beamLength = boundingBoxWidth + 2 * overhang - outSideDiaNearBBXHigh / 2 - outSideDiaNearBBXLow / 2;
                        }
                        else if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                            beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX;
                        else if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) == true && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                            beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX - outSideDiaNearBBXHigh / 2;
                        else if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) == true)
                            beamLength = boundingBoxHeight + 2 * overhang + lengthOutsideBBX - outSideDiaNearBBXLow / 2;
                        else
                            beamLength = boundingBoxHeight + 2 * overhang - outSideDiaNearBBXHigh / 2 - outSideDiaNearBBXLow / 2;

                        if (beamLength > (maxAssemblyLength + 0.00001))
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Assembly length is more than maximum assembly length.", "", "FieldSupports.cs", 497);

                        structureOffset = structuralDim / 2 + freeEndOverLength;
                    }
                }

                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[HORIZONTALSECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[HORIZONTALSECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;
                // Change this methods from ConfigHlpr to HH
                componentDictionary[HORIZONTALSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORIZONTALSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORIZONTALSECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[HORIZONTALSECTION].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[HORIZONTALSECTION].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                // ====== ======
                // Create Joints
                // ====== ======
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (sectionType.Equals("L"))
                    {
                        // NOTE : Change this methods from ConfigHlpr to HH
                        componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength, "IJUAHgrOccLength", "Length");
                        // When there are no outermost vertical pipes on either sides of structure
                        if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) == true && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) == true)
                        {
                            if (HgrCompareDoubleService.cmpdbl(structFlangeThickness , 0) == false)
                            {
                                if (!isTwoSides)
                                {
                                    endPlateOffset1 = structWidthOffset + steelWidth / 2;
                                    endPlateOffset2 = 0;
                                    if (Configuration == 1 || Configuration == 3)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -overhang - offsetfromBBXtoVertFarPipe + outsidePipeDiaMeter / 2, basePlateThickness, -shoeHeight);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, -overhang - offsetfromBBXtoVertFarPipe + outsidePipeDiaMeter / 2, boundingBoxWidth + shoeHeight, basePlateOffset);
                                }
                                else // do this if there are pipes on both sides
                                {
                                    if (Configuration == 1 || Configuration == 3)
                                    {
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", "-1", "BBSR_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, overhang - outsidePipeDiaMeter / 2, -boundingBoxWidth - shoeHeight, -basePlateThickness);
                                        endPlateOffset1 = structWidthOffset + steelWidth / 2;
                                        endPlateOffset2 = 0;
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", "-1", "BBSR_Low", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, overhang - outsidePipeDiaMeter / 2, -basePlateThickness, -shoeHeight);
                                        endPlateOffset1 = 0;
                                        endPlateOffset2 = steelDepth / 2;
                                    }
                                }
                            }
                            else
                            {
                                // Previous code , still might be used for the concrete
                                if (Configuration == 1)
                                {
                                    // Add Joint Between Route and Horizontal Beam
                                    JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -overhang - offsetfromBBXtoVertFarPipe + outsidePipeDiaMeter / 2, basePlateThickness, -shoeHeight);
                                    endPlateOffset1 = structWidthOffset + steelWidth / 2;
                                    endPlateOffset2 = 0;
                                }
                                else if (Configuration == 2)
                                {
                                    offsetfromBBXtoVertFarPipe = RefPortHelper.DistanceBetweenPorts(outsidePipePort, "BBSR_High", PortDistanceType.Horizontal);
                                    // Add Joint Between Route and Horizontal Beam
                                    JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, -beamLength - offsetfromBBXtoVertFarPipe + overhang, shoeHeight, basePlateThickness);
                                    endPlateOffset1 = 0;
                                    endPlateOffset2 = steelDepth / 2;
                                }
                                else if (Configuration == 3)
                                {
                                    offsetfromBBXtoVertFarPipe = RefPortHelper.DistanceBetweenPorts(outsidePipePort, "BBSR_High", PortDistanceType.Horizontal);
                                    // Add Joint Between Route and Horizontal Beam
                                    JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -beamLength - offsetfromBBXtoVertFarPipe + overhang, basePlateThickness, -shoeHeight - boundingBoxWidth);
                                    endPlateOffset1 = 0;
                                    endPlateOffset2 = steelDepth / 2;
                                }
                                else if (Configuration == 4)
                                {
                                    // Add Joint Between Route and Horizontal Beam
                                    JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, -overhang - offsetfromBBXtoVertFarPipe + outsidePipeDiaMeter / 2, boundingBoxWidth + shoeHeight, basePlateOffset);
                                    endPlateOffset1 = structWidthOffset + steelDepth / 2;
                                    endPlateOffset2 = 0;
                                }
                            }
                        }
                        else
                        {
                            endPlateOffset1 = structWidthOffset + steelWidth / 2;
                            endPlateOffset2 = 0;
                            if (Configuration == 1 || Configuration == 3)
                            {
                                if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                                    JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, -distFromBBXLow - overhang, boundingBoxWidth + shoeHeight, basePlateOffset);
                                else
                                {
                                    if (HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, -distFromBBXLow - overhang, boundingBoxWidth + shoeHeight, basePlateOffset);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, outsidePipeDiaMeter / 2 - overhang, boundingBoxWidth + shoeHeight, basePlateOffset);
                                }
                            }
                            else if (Configuration == 2 || Configuration == 4)
                            {
                                if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                                    JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -boundingBoxHeight - distFromBBXLow - overhang, basePlateThickness, -boundingBoxWidth - shoeHeight);
                                else
                                {
                                    if (HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -boundingBoxHeight - distFromBBXLow - overhang, basePlateThickness, -boundingBoxWidth - shoeHeight);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -boundingBoxHeight - overhang, basePlateThickness, -boundingBoxWidth - shoeHeight);
                                }
                            }
                        }
                    }
                    else // Top is C Section, Brace is L Section
                    {
                        facenumber = 513;

                        if ((SupportHelper.SupportingObjects.Count != 0))
                            facenumber = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                        // NOTE : Change this methods from ConfigHlpr to HH
                        componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength, "IJUAHgrOccLength", "Length");
                        if (reverseSec.ToUpper().Equals("NO"))
                        {
                            if (!isTwoSides)
                            {
                                double basePlateLengthOffset = 0;
                                if (basePlate.ToUpper().Equals("WITH"))
                                    basePlateLengthOffset = basePlateWidth / 2 - structSteelWidth / 2;
                                else
                                    basePlateLengthOffset = 0;

                                if (HgrCompareDoubleService.cmpdbl(structFlangeThickness , 0) == true && facenumber == 513)
                                {
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength + structSteelDepth / 2, "IJUAHgrOccLength", "Length");
                                    if (Configuration == 1)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, boundingBoxWidth + shoeHeight, -overhang + pipeOD / 2, structSteelWidth / 2 + basePlateThickness);
                                    else if (Configuration == 2)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, -steelDepth - shoeHeight, -overhang + pipeOD / 2, structSteelWidth / 2 + basePlateThickness);
                                    else if (Configuration == 3)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, boundingBoxWidth + shoeHeight, pipeOD / 2 - overhang + beamLength + structSteelDepth / 2, -structSteelWidth / 2 - basePlateThickness);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, -steelDepth - shoeHeight, pipeOD / 2 - overhang + beamLength + structSteelDepth / 2, -structSteelWidth / 2 - basePlateThickness);
                                }
                                else if (HgrCompareDoubleService.cmpdbl(structFlangeThickness , 0) == true && facenumber == 514)
                                {
                                    if (Configuration == 1)
                                        /// Add Joint Between Route and Horizontal Beam
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, boundingBoxWidth + shoeHeight, -beamLength + distBBXLowToStruct + structuralDim / 2 + basePlateLengthOffset, basePlateOffset);
                                    else if (Configuration == 2)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, -shoeHeight - steelDepth, -beamLength + distBBXLowToStruct + structuralDim / 2 + basePlateLengthOffset, basePlateOffset);
                                    else if (Configuration == 3)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, boundingBoxWidth + shoeHeight, -beamLength + distBBXLowToStruct + structuralDim / 2 + basePlateLengthOffset, -structSteelDepth - basePlateThickness);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, -shoeHeight - steelDepth, -beamLength + distBBXLowToStruct + structuralDim / 2 + basePlateLengthOffset, -structSteelDepth - basePlateThickness);
                                }
                                else
                                {
                                    if (Configuration == 1 || Configuration == 3)
                                        // Add Joint Between Route and Horizontal Beam
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, boundingBoxWidth + shoeHeight, -beamLength + distBBXLowToStruct + structuralDim / 2 + basePlateLengthOffset, basePlateOffset);
                                    else
                                        // Add Joint Between Route and Horizontal Beam
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, -shoeHeight - steelDepth, -beamLength + distBBXLowToStruct + structuralDim / 2 + basePlateLengthOffset, basePlateOffset);
                                }
                            }
                            else
                            {
                                if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) == true && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) == true)
                                {
                                    if (Configuration == 1 || Configuration == 3)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, shoeHeight, -boundingBoxHeight + outSideDiaNearBBXLow / 2 - overhang, basePlateOffset);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, -boundingBoxWidth - shoeHeight - steelDepth, -boundingBoxHeight + outSideDiaNearBBXLow / 2 - overhang, basePlateOffset);
                                }
                                else
                                {
                                    if (Configuration == 1)
                                    {
                                         JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, boundingBoxWidth + shoeHeight, -distFromBBXLow - overhang, basePlateOffset);
                                    }
                                    else if (Configuration == 2)
                                    {
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, -boundingBoxWidth - shoeHeight - steelDepth, -boundingBoxHeight - distFromBBXLow - overhang, basePlateOffset);
                                    }
                                    else if (Configuration == 3)
                                    {
                                        if (HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false)
                                            JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, boundingBoxWidth + shoeHeight, -distFromBBXLow - overhang, basePlateOffset);
                                        else
                                            JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, 0, -boundingBoxHeight + outSideDiaNearBBXLow / 2 - overhang, basePlateOffset);
                                    }
                                    else
                                    {
                                        if (HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false)
                                            JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, -boundingBoxWidth - shoeHeight - steelDepth, -boundingBoxHeight - distFromBBXLow - overhang, basePlateOffset);
                                        else
                                            JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.X, -boundingBoxWidth - shoeHeight - steelDepth, -boundingBoxHeight + outSideDiaNearBBXLow / 2 - overhang, basePlateOffset);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (!isTwoSides)
                            {
                                if (Configuration == 1 || Configuration == 3)
                                    // Add Joint Between Route and Horizontal Beam
                                    JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, boundingBoxWidth + shoeHeight, distBBXLowToStruct + structSteelWidth / 2, basePlateOffset + structWidthOffset + steelWidth);
                                else
                                    // Add Joint Between Route and Horizontal Beam
                                    JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, -boundingBoxWidth - shoeHeight - steelDepth, distBBXHighToStruct + structSteelWidth / 2, basePlateOffset + structWidthOffset + steelWidth);
                            }
                            else
                            {
                                if (Configuration == 1 || Configuration == 3)
                                {
                                    if (HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, boundingBoxWidth + shoeHeight, boundingBoxHeight + distFromBBXHigh + overhang, basePlateOffset + structWidthOffset + steelWidth);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_Low", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, boundingBoxWidth + shoeHeight, boundingBoxHeight - outSideDiaNearBBXHigh / 2 + overhang, basePlateOffset + structWidthOffset + steelWidth);
                                }
                                else if (Configuration == 2 || Configuration == 4)
                                {
                                    if (HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) ==false && HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) ==false)
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, -boundingBoxWidth - shoeHeight - steelDepth, boundingBoxHeight + distFromBBXHigh + overhang, basePlateOffset + structWidthOffset + steelWidth);
                                    else
                                        JointHelper.CreateRigidJoint("-1", "BBSR_High", HORIZONTALSECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, -boundingBoxWidth - shoeHeight - steelDepth, -outSideDiaNearBBXHigh / 2 + overhang, basePlateOffset + structWidthOffset + steelWidth);
                                }
                            }
                        }
                    }
                }
                else
                {
                    double byPointAngle1 = 0;
                    double distStructToFirstRoute = FLSampleSupportServices.GetNearestOuterMostRouteDistance(this, outsidePipePort, verticalRoutePort);
                    double beamLengthOffset = beamLength - distStructToFirstRoute - overhang;
                    if (sectionType.Equals("L"))
                    {
                        componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength, "IJUAHgrOccLength", "Length");
                        JointHelper.CreatePrismaticJoint(HORIZONTALSECTION, "EndCap", CONNECTION1, "Connection", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        if (basePlate.ToUpper().Equals("WITH"))
                        {
                            if (Configuration == 1 || Configuration == 2)
                                JointHelper.CreateRevoluteJoint("-1", "Structure", BASEPLATE, "EndOther", Axis.X, Axis.NegativeY);
                            else
                                JointHelper.CreateRevoluteJoint("-1", "Structure", BASEPLATE, "EndOther", Axis.X, Axis.Y);

                            if (!isTwoSides)
                            {
                                if (Configuration == 1 || Configuration == 3)
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, 0, -basePlateWidth / 2 - freeEndOverLength);
                                else
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.X, steelDepth / 2, 0, -basePlateWidth / 2 - freeEndOverLength);
                            }
                            else
                            {
                                if (Configuration == 1 || Configuration == 3)
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, 0, -beamLengthOffset);
                                else
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.X, steelDepth / 2, 0, -beamLengthOffset);
                            }
                        }
                        else // bp with
                        {
                            if (!isTwoSides)
                            {
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength - structureOffset, "IJUAHgrOccLength", "Length");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(structureOffset, "IJUAHgrOccOverLength", "EndOverLength");
                            }
                            else
                            {
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength - beamLengthOffset, "IJUAHgrOccLength", "Length");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLengthOffset, "IJUAHgrOccOverLength", "EndOverLength");
                            }

                            if (Configuration == 1)
                                JointHelper.CreateRevoluteJoint("-1", "Structure", HORIZONTALSECTION, "EndCap", Axis.X, Axis.NegativeY);
                            else if (Configuration == 2)
                                JointHelper.CreateRevoluteJoint("-1", "Structure", HORIZONTALSECTION, "EndCap", Axis.X, Axis.Y);
                            else if (Configuration == 3)
                                JointHelper.CreateRevoluteJoint("-1", "Structure", HORIZONTALSECTION, "EndCap", Axis.X, Axis.NegativeX);
                            else if (Configuration == 4)
                                JointHelper.CreateRevoluteJoint("-1", "Structure", HORIZONTALSECTION, "EndCap", Axis.X, Axis.X);
                        }

                        byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);

                        if ((byPointAngle1 * 180 / Math.PI) < 180 && (byPointAngle1 * 180 / Math.PI) > 89.999999)
                        {
                            if ((byPointAngle1 * 180 / Math.PI) > 89.999999 && (byPointAngle1 * 180 / Math.PI) < 90.00001)
                                JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.XY, 0);
                            else
                                JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.NegativeZX, 0);
                        }
                        else
                            JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.ZX, 0);
                    }
                    else // Top is C Section
                    {
                        if (reverseSec.ToUpper().Equals("NO"))
                        {
                            componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength, "IJUAHgrOccLength", "Length");
                            JointHelper.CreatePrismaticJoint(HORIZONTALSECTION, "EndCap", CONNECTION1, "Connection", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                            if (basePlate.ToUpper().Equals("WITH"))
                            {
                                if (Configuration == 1 || Configuration == 3)
                                    JointHelper.CreateRevoluteJoint("-1", "Structure", BASEPLATE, "EndOther", Axis.X, Axis.NegativeY);
                                else
                                    JointHelper.CreateRevoluteJoint("-1", "Structure", BASEPLATE, "EndOther", Axis.X, Axis.Y);

                                if (!isTwoSides)
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, 0, -basePlateWidth / 2 - freeEndOverLength);
                                else
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.X, steelDepth / 2, 0, -beamLengthOffset);
                            }
                            else // bp with
                            {
                                if (!isTwoSides)
                                {
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength - structureOffset, "IJUAHgrOccLength", "Length");
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(structureOffset, "IJUAHgrOccOverLength", "EndOverLength");
                                }
                                else
                                {
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength - beamLengthOffset, "IJUAHgrOccLength", "Length");
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLengthOffset, "IJUAHgrOccOverLength", "EndOverLength");
                                }

                                if (Configuration == 1 || Configuration == 3)
                                    JointHelper.CreateRevoluteJoint("-1", "Structure", HORIZONTALSECTION, "EndCap", Axis.X, Axis.NegativeY);
                                else if (Configuration == 2 || Configuration == 4)
                                    JointHelper.CreateRevoluteJoint("-1", "Structure", HORIZONTALSECTION, "EndCap", Axis.X, Axis.Y);
                            }

                            byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);

                            if ((byPointAngle1 * 180 / Math.PI) < 180 && (byPointAngle1 * 180 / Math.PI) > 89.999999)
                            {
                                if ((byPointAngle1 * 180 / Math.PI) > 89.999999 && (byPointAngle1 * 180 / Math.PI) < 90.00001)
                                    JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.XY, 0);
                                else
                                    JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.NegativeZX, 0);
                            }
                            else
                                JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.ZX, 0);
                        }
                        else
                        {
                            componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength, "IJUAHgrOccLength", "Length");
                            JointHelper.CreatePrismaticJoint(HORIZONTALSECTION, "EndCap", CONNECTION1, "Connection", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                            if (basePlate.ToUpper().Equals("WITH"))
                            {
                                if (Configuration == 1 || Configuration == 3)
                                    JointHelper.CreateRevoluteJoint("-1", "Structure", BASEPLATE, "EndOther", Axis.X, Axis.NegativeY);
                                else
                                    JointHelper.CreateRevoluteJoint("-1", "Structure", BASEPLATE, "EndOther", Axis.X, Axis.Y);

                                if (!isTwoSides)
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.X, steelDepth / 2, structWidthOffset + steelWidth, -basePlateWidth / 2 - freeEndOverLength);
                                else
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", BASEPLATE, "StartOther", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.X, steelDepth / 2, structWidthOffset + steelWidth, -beamLengthOffset);
                            }
                            else // bp with
                            {
                                if (!isTwoSides)
                                {
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength - structureOffset, "IJUAHgrOccLength", "Length");
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(structureOffset, "IJUAHgrOccOverLength", "EndOverLength");
                                }
                                else
                                {
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLength - beamLengthOffset, "IJUAHgrOccLength", "Length");
                                    componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamLengthOffset, "IJUAHgrOccOverLength", "EndOverLength");
                                }

                                JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", CONNECTION2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, structWidthOffset + steelWidth);
                                if (Configuration == 1 || Configuration == 3)
                                    JointHelper.CreateRevoluteJoint("-1", "Structure", CONNECTION2, "Connection", Axis.X, Axis.Y);
                                else if (Configuration == 2 || Configuration == 4)
                                    JointHelper.CreateRevoluteJoint("-1", "Structure", CONNECTION2, "Connection", Axis.X, Axis.NegativeY);
                            }

                            byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);

                            if ((byPointAngle1 * 180 / Math.PI) < 180 && (byPointAngle1 * 180 / Math.PI) > 89.999999)
                            {
                                if ((byPointAngle1 * 180 / Math.PI) > 89.999999 && (byPointAngle1 * 180 / Math.PI) < 90.00001)
                                    JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.XY, 0);
                                else
                                    JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.NegativeZX, 0);
                            }
                            else
                                JointHelper.CreatePlanarJoint(CONNECTION1, "Connection", "-1", "Route", Plane.XY, Plane.ZX, 0);
                        }
                    }
                }

                if (endPlates.ToUpper().Equals("YES"))
                {
                    string endPlateBom = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, endPlateThickness, UnitName.DISTANCE_MILLIMETER) + "Plate Steel," + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, endPlateWidth, UnitName.DISTANCE_MILLIMETER) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, endPlateDepth, UnitName.DISTANCE_MILLIMETER);
                    double endPlateWeight = 7900 * endPlateThickness * endPlateWidth * endPlateDepth;

                    componentDictionary[ENDPLATE1].SetPropertyValue(endPlateDepth, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                    componentDictionary[ENDPLATE1].SetPropertyValue(endPlateWidth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                    componentDictionary[ENDPLATE1].SetPropertyValue(endPlateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                    componentDictionary[ENDPLATE1].SetPropertyValue(endPlateBom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");
                    componentDictionary[ENDPLATE1].SetPropertyValue(endPlateWeight, "IJOAHgrUtility_USER_FIXED_BOX", "DryWt");

                    if (!string.IsNullOrEmpty(ENDPLATE2))
                    {
                        componentDictionary[ENDPLATE2].SetPropertyValue(endPlateDepth, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                        componentDictionary[ENDPLATE2].SetPropertyValue(endPlateWidth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                        componentDictionary[ENDPLATE2].SetPropertyValue(endPlateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                        componentDictionary[ENDPLATE2].SetPropertyValue(endPlateBom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");
                        componentDictionary[ENDPLATE2].SetPropertyValue(endPlateWeight, "IJOAHgrUtility_USER_FIXED_BOX", "DryWt");
                    }

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (sectionType.Equals("L"))
                        {
                            if (isTwoSides)
                            {
                                JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, endPlateOffset2, endPlateOffset1);
                                if (!string.IsNullOrEmpty(ENDPLATE2))
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, endPlateOffset2, endPlateOffset1);
                            }
                            else
                            {
                                if (HgrCompareDoubleService.cmpdbl(distFromBBXHigh , 0) == true && HgrCompareDoubleService.cmpdbl(distFromBBXLow , 0) == true)
                                {
                                    if (Configuration == 1 || Configuration == 3)
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -endPlateThickness, endPlateOffset2, endPlateOffset1);
                                    else
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, endPlateOffset2, endPlateOffset1);
                                }
                                else
                                {
                                    if (Configuration == 1 || Configuration == 3)
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, endPlateOffset2, endPlateOffset1);
                                    else
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -endPlateThickness, endPlateOffset2, endPlateOffset1);
                                }
                            }
                        }
                        else
                        {
                            if (reverseSec.ToUpper().Equals("NO"))
                            {
                                if (Configuration == 1 || Configuration == 3)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, 0, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, 0, structWidthOffset + steelWidth / 2);
                                }
                                else
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, steelDepth, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, steelDepth, structWidthOffset + steelWidth / 2);
                                }
                            }
                            else
                            {
                                if (Configuration == 1 || Configuration == 3)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, 0, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, 0, structWidthOffset + steelWidth / 2);
                                }
                                else
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, steelDepth, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, steelDepth, structWidthOffset + steelWidth / 2);
                                }
                            }
                        }
                    }
                    else // his is for BY-Point
                    {
                        if (basePlate.ToUpper().Equals("WITH"))
                        {
                            if (sectionType.ToUpper().Equals("L"))
                            {
                                if (Configuration == 1 || Configuration == 3)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, 0, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, 0, structWidthOffset + steelWidth / 2);
                                }
                                else if (Configuration == 2 || Configuration == 4)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -endPlateThickness, 0, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, beamLength, 0, structWidthOffset + steelWidth / 2);
                                }
                            }
                            else
                            {
                                if (reverseSec.ToUpper().Equals("NO"))
                                {
                                    endPlateOffset1 = steelDepth;
                                    endPlateOffset2 = 0;
                                }
                                else
                                {
                                    endPlateOffset1 = 0;
                                    endPlateOffset2 = steelDepth;
                                }

                                if (Configuration == 1 || Configuration == 3)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, endPlateOffset1, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, endPlateOffset1, structWidthOffset + steelWidth / 2);
                                }
                                else if (Configuration == 2 || Configuration == 4)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, endPlateOffset2, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, endPlateOffset2, structWidthOffset + steelWidth / 2);
                                }
                            }
                        }
                        else
                        {
                            if (sectionType.ToUpper().Equals("L"))
                            {
                                if (Configuration == 1)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, 0, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, 0, structWidthOffset + steelWidth / 2);
                                }
                                else if (Configuration == 2)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -endPlateThickness, structWidthOffset + steelWidth / 2, 0);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, beamLength, structWidthOffset + steelWidth / 2, 0);
                                }
                                else if (Configuration == 3)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, structWidthOffset + steelWidth / 2, 0);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, structWidthOffset + steelWidth / 2, 0);
                                }
                                else if (Configuration == 4)
                                {
                                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -endPlateThickness, 0, structWidthOffset + steelWidth / 2);
                                    if (!string.IsNullOrEmpty(ENDPLATE2))
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, beamLength, 0, structWidthOffset + steelWidth / 2);
                                }

                            }
                            else
                            {
                                if (reverseSec.ToUpper().Equals("NO"))
                                {
                                    if (Configuration == 1 || Configuration == 3)
                                    {
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, steelDepth, structWidthOffset + steelWidth / 2);
                                        if (!string.IsNullOrEmpty(ENDPLATE2))
                                            JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, steelDepth, structWidthOffset + steelWidth / 2);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -endPlateThickness, structWidthOffset + steelWidth / 2, 0);
                                        if (!string.IsNullOrEmpty(ENDPLATE2))
                                            JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, beamLength, structWidthOffset + steelWidth / 2, 0);
                                    }
                                }
                                else
                                {
                                    if (Configuration == 1 || Configuration == 3)
                                    {
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -endPlateThickness, 0, structWidthOffset + steelWidth / 2);
                                        if (!string.IsNullOrEmpty(ENDPLATE2))
                                            JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, beamLength, 0, structWidthOffset + steelWidth / 2);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE1, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -endPlateThickness, structWidthOffset + steelWidth / 2, steelDepth);
                                        if (!string.IsNullOrEmpty(ENDPLATE2))
                                            JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", ENDPLATE2, "StartOther", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, beamLength, structWidthOffset + steelWidth / 2, steelDepth);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 4;
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
                    // Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    for (int index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                    {
                        routeConnections.Add(new ConnectionInfo(HORIZONTALSECTION, index)); // partindex, routeindex
                    }

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
                    // Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    for (int index = 1; index <= SupportHelper.SupportingObjects.Count; index++)
                    {
                        if (basePlate == "With")
                            structConnections.Add(new ConnectionInfo(BASEPLATE, index)); // partindex, routeindex
                        else
                            structConnections.Add(new ConnectionInfo(HORIZONTALSECTION, index)); // partindex, routeindex
                    }

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

        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                string supportName = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJUAHgrFl_SupportName", "SupportName")).PropValue;
                int endPlate = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrFl_EndPlate", "EndPlate")).PropValue;

                if (endPlate == 2)
                    bomDescription = supportName;
                else
                    bomDescription = supportName + "- with End Plate";

                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOMDescription." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
        /// <summary>
        /// This method returns the Mirrored configuration Value
        /// </summary>
        /// <param name="CurrentMirrorToggleValue">int - Toggle Value.</param>
        /// <param name="eMirrorPlane">MirrorPlane - eMirrorPlane.</param>
        /// <returns>int</returns>        
        /// <code>
        ///     MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane);
        ///</code>
        public override int MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane)
        {
            try
            {
                if (CurrentMirrorToggleValue == 1)
                    return 2;
                else if (CurrentMirrorToggleValue == 2)
                    return 1;
                else if (CurrentMirrorToggleValue == 3)
                    return 4;
                else if (CurrentMirrorToggleValue == 4)
                    return 3;

                return CurrentMirrorToggleValue;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Mirrored Configuration." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
