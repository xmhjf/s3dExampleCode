//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   LFieldSupports.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.LFieldSupports
//   Author       : Rajeswari
//   Creation Date: 03-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 03-Sep-2013  Rajeswari CR-CP-224478 Convert FlSample_Supports to C# .Net 
// 28-Apr-2015      PVK	     Resolve Coverity issues found in April
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
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

    public class LFieldSupports : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        public string HORIZONTALSECTION = "HORIZONTALSECTION";
        public string VERTICALSECTION = "VERTICALSECTION";
        public string BASEPLATEHORIZONTAL = "BASEPLATEHORIZONTAL";
        public string BASEPLATEVERTICAL = "BASEPLATEVERTICAL";

        string basePlate = string.Empty;
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

                    if (basePlate.ToUpper().Equals("WITH"))
                    {
                        parts.Add(new PartInfo(HORIZONTALSECTION, sectionSize));
                        parts.Add(new PartInfo(VERTICALSECTION, sectionSize));
                        parts.Add(new PartInfo(BASEPLATEHORIZONTAL, "Utility_USER_FIXED_BOX_1"));
                        parts.Add(new PartInfo(BASEPLATEVERTICAL, "Utility_USER_FIXED_BOX_1"));
                    }
                    else
                    {
                        parts.Add(new PartInfo(HORIZONTALSECTION, sectionSize));
                        parts.Add(new PartInfo(VERTICALSECTION, sectionSize));
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

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Can not place the support in Place-By-Reference.", "", "LFieldSupports.cs", 91);
                    return;
                }

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Wall))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Can not place the support with slab and wall.", "", "LFieldSupports.cs", 97);
                        return;
                    }
                }

                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                double shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrFl_ShoeH", "ShoeH")).PropValue;
                double maxSpan = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_MaxSpan", "MaxSpan")).PropValue;
                double basePlateWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_BP_Width", "BP_Width")).PropValue;
                double basePlateDepth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_BP_Depth", "BP_Depth")).PropValue;
                double basePlateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_BP_Thickness", "BP_Thickness")).PropValue;
                double maxAssemblyLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_MaxAssyLength", "MaxAssyLength")).PropValue;
                string supportName = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_SupportName", "SupportName")).PropValue;

                int jointType = 0;
                double basePlateOffset = 0, overlap = 0;
                if (supportName.Equals("5FSX3B"))
                {
                    jointType = ((int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrFl_JointType", "JointType")).PropValue);
                    basePlateOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_BP_Offset", "BP_Offset")).PropValue;
                    overlap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrFl_Overlap", "Overlap")).PropValue;
                }
                double structOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrFl_StructOffset", "StructOffset")).PropValue;
                double clearance = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrFl_Clearance", "Clearance")).PropValue;

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
                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                // Get interface for accessing items on the collection of Part Occurences
                BusinessObject horizontalSectionPart = componentDictionary[HORIZONTALSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                double steelWidth = crossSection.Width;
                double steelDepth = crossSection.Depth;

                double structSteelDepth = 0, structSteelWidth = 0, structFlangeThickness = 0, structWebThickness = 0;
                if (supportingType == "Steel")
                {
                    structSteelWidth = SupportingHelper.SupportingObjectInfo(1).Width;
                    structFlangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                    structWebThickness = SupportingHelper.SupportingObjectInfo(1).WebThickness;
                }

                string partSheet = supportName; string partNumber = string.Empty;
                if (supportName.ToUpper().Equals("FS6B"))
                    partNumber = "FS6B_1";
                else
                {
                    if (basePlate.ToUpper().Equals("WITH"))
                        partNumber = "5FSX3B_2";
                    else
                        partNumber = "5FSX3B_1";
                }

                double ndFrom = FLSampleSupportServices.GetDataByCondition(partSheet, "IJHgrSupportDefinition", "NDFrom", "IJDPart", "PartNumber", partNumber);
                double ndTo = FLSampleSupportServices.GetDataByCondition(partSheet, "IJHgrSupportDefinition", "NDTo", "IJDPart", "PartNumber", partNumber);
                //  Check pipe diameter
                double pipeND = 0;
                // Check for valid pipe size
                for (int i = 1; i <= SupportHelper.SupportedObjects.Count; i++)
                {
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    pipeND = pipeInfo.NominalDiameter.Size;

                    if (pipeND < (ndFrom - 0.000001) || pipeND > (ndTo + 0.000001))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "LFieldSupports.cs", 182);
                        return;
                    }
                }

                Boolean[] isOffsetApplied = FLSampleSupportServices.GetIsLugEndOffsetApplied(this);
                string[] idxStructPort = FLSampleSupportServices.GetIndexedStructPortName(this, isOffsetApplied);
                string horizontalStructPort = idxStructPort[0];
                string verticalStructPort = idxStructPort[1];

                // Check the First Structure is Vertical or Horizontal
                double tempAngle = (RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_Z) * 180 / Math.PI);
                bool isPrimHorStr;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(tempAngle, 1)) , 0)==true || HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(tempAngle, 1)) , 180) == true)
                    isPrimHorStr = true;
                else
                    isPrimHorStr = false;

                if (isPrimHorStr == true)
                {
                    horizontalStructPort = idxStructPort[1];
                    verticalStructPort = idxStructPort[0];
                }

                string connectionPortHorizontal = string.Empty, connectionPortVertical = string.Empty;
                double distHoriStructToBBXLow = RefPortHelper.DistanceBetweenPorts("BBSR_Low", horizontalStructPort, PortDistanceType.Vertical);
                double distHoriStructToBBXHigh = RefPortHelper.DistanceBetweenPorts("BBSR_High", horizontalStructPort, PortDistanceType.Vertical);

                if (distHoriStructToBBXHigh > distHoriStructToBBXLow)
                    connectionPortHorizontal = "BBSR_High";
                else
                    connectionPortHorizontal = "BBSR_Low";

                double distVerticalStructToBBXPort = RefPortHelper.DistanceBetweenPorts(connectionPortHorizontal, verticalStructPort, PortDistanceType.Horizontal_Perpendicular);
                // Set horizonal and vertical offset and connection port for vertical section depending on the position of vertical structure
                double distVertStructToBBXLow = RefPortHelper.DistanceBetweenPorts("BBSR_Low", verticalStructPort, PortDistanceType.Horizontal_Perpendicular);
                double distVertStructToBBXHigh = RefPortHelper.DistanceBetweenPorts("BBSR_High", verticalStructPort, PortDistanceType.Horizontal_Perpendicular);

                double vertSectionHoriOffset = 0, vertSectionVertOffset = 0, vertSectionDepthOffset = 0, horizontalOffset = 0;
                Plane[] planeA = new Plane[2]; Plane[] planeB = new Plane[2]; Axis[] axisA = new Axis[2]; Axis[] axisB = new Axis[2];
                double horizontalSpan = boundingBoxWidth + distVerticalStructToBBXPort + clearance;
                structSteelDepth = 0;
                if (distVertStructToBBXHigh > distVertStructToBBXLow)
                {
                    if (Configuration == 1)
                    {
                        planeB[0] = Plane.YZ;
                        planeB[1] = Plane.NegativeZX;
                        axisB[0] = Axis.Y;
                        axisB[1] = Axis.Z;
                        if (connectionPortHorizontal.Equals("BBSR_High"))
                        {
                            horizontalOffset = -(horizontalSpan - distVerticalStructToBBXPort + structSteelDepth);
                            connectionPortVertical = "BeginCap";
                            vertSectionHoriOffset = 0;

                            if (supportName.ToUpper().Equals("5FSX3B"))
                            {
                                planeA[0] = Plane.ZX;
                                planeA[1] = Plane.YZ;
                                axisA[0] = Axis.Z;
                                axisA[1] = Axis.NegativeY;
                                if (jointType == 3)
                                    vertSectionHoriOffset = overlap;
                                else if (jointType == 2)
                                    vertSectionHoriOffset = steelDepth;
                                else
                                    vertSectionHoriOffset = 0;
                            }
                        }
                        else
                        {
                            horizontalSpan = horizontalSpan + steelWidth - structWebThickness / 2;
                            horizontalOffset = -distVerticalStructToBBXPort + structSteelDepth + basePlateThickness + structWebThickness / 2;

                            connectionPortVertical = "EndCap";
                            vertSectionHoriOffset = -steelWidth;
                            if (supportName.ToUpper().Equals("5FSX3B"))
                            {
                                horizontalSpan = horizontalSpan - steelWidth - basePlateThickness;
                                planeA[0] = Plane.ZX;
                                planeA[1] = Plane.ZX;
                                axisA[0] = Axis.Z;
                                axisA[1] = Axis.X;
                                if (jointType == 3)
                                    vertSectionHoriOffset = -steelWidth - overlap;
                                else if (jointType == 2)
                                    vertSectionHoriOffset = -steelDepth;
                                else
                                    vertSectionHoriOffset = 0;
                            }
                        }
                    }
                    else
                    {
                        planeB[0] = Plane.YZ;
                        planeB[1] = Plane.ZX;
                        axisB[0] = Axis.Y;
                        axisB[1] = Axis.NegativeZ;
                        if (connectionPortHorizontal.Equals("BBSR_Low"))
                        {
                            horizontalSpan = horizontalSpan + steelWidth - structWebThickness / 2;
                            horizontalOffset = horizontalSpan - distVerticalStructToBBXPort + structSteelDepth;
                            connectionPortVertical = "BeginCap";
                            vertSectionHoriOffset = 0;

                            if (supportName.ToUpper().Equals("5FSX3B"))
                            {
                                planeA[0] = Plane.ZX;
                                planeA[1] = Plane.YZ;
                                axisA[0] = Axis.Z;
                                axisA[1] = Axis.NegativeY;
                                if (jointType == 3)
                                    vertSectionHoriOffset = overlap;
                                else if (jointType == 2)
                                    vertSectionHoriOffset = steelDepth;
                                else
                                    vertSectionHoriOffset = 0;
                            }
                        }
                        else
                        {
                            connectionPortVertical = "EndCap";
                            vertSectionHoriOffset = -steelWidth;
                            if (supportName.ToUpper().Equals("5FSX3B"))
                            {
                                horizontalSpan = horizontalSpan - steelWidth - basePlateThickness;
                                planeA[0] = Plane.ZX;
                                planeA[1] = Plane.ZX;
                                axisA[0] = Axis.Z;
                                axisA[1] = Axis.X;
                                if (jointType == 3)
                                    vertSectionHoriOffset = -steelWidth - overlap;
                                else if (jointType == 2)
                                    vertSectionHoriOffset = -steelDepth;
                                else
                                    vertSectionHoriOffset = 0;
                            }
                        }
                    }
                }
                else
                {
                    if (isPrimHorStr == true)
                    {
                        if (Configuration == 1)
                        {
                            if (connectionPortHorizontal.Equals("BBSR_High"))
                            {
                                horizontalOffset = -(horizontalSpan - distVerticalStructToBBXPort + structSteelDepth);
                                connectionPortVertical = "BeginCap";
                                vertSectionHoriOffset = 0;
                                planeB[0] = Plane.YZ;
                                planeB[1] = Plane.NegativeZX;
                                axisB[0] = Axis.Y;
                                axisB[1] = Axis.X;
                                if (supportName.ToUpper().Equals("5FSX3B"))
                                {
                                    planeA[0] = Plane.ZX;
                                    planeA[1] = Plane.YZ;
                                    axisA[0] = Axis.Z;
                                    axisA[1] = Axis.NegativeY;
                                    if (jointType == 3)
                                        vertSectionHoriOffset = overlap;
                                    else if (jointType == 2)
                                        vertSectionHoriOffset = steelDepth;
                                    else
                                        vertSectionHoriOffset = 0;
                                }
                            }
                            else
                            {
                                horizontalSpan = horizontalSpan - boundingBoxWidth + steelWidth;
                                horizontalOffset = -(-distVerticalStructToBBXPort + structSteelDepth + basePlateThickness);
                                planeB[0] = Plane.YZ;
                                planeB[1] = Plane.NegativeZX;
                                axisB[0] = Axis.Y;
                                axisB[1] = Axis.NegativeX;
                                connectionPortVertical = "EndCap";
                                vertSectionHoriOffset = -steelWidth;
                                if (supportName.ToUpper().Equals("5FSX3B"))
                                {
                                    horizontalSpan = horizontalSpan - steelWidth - basePlateThickness;
                                    planeA[0] = Plane.ZX;
                                    planeA[1] = Plane.ZX;
                                    axisA[0] = Axis.Z;
                                    axisA[1] = Axis.X;
                                    if (jointType == 3)
                                        vertSectionHoriOffset = -steelWidth - overlap;
                                    else if (jointType == 2)
                                        vertSectionHoriOffset = -steelDepth;
                                    else
                                        vertSectionHoriOffset = 0;
                                }
                            }
                        }
                        else
                        {
                            if (connectionPortHorizontal.Equals("BBSR_High"))
                            {
                                horizontalSpan = horizontalSpan + steelWidth;
                                horizontalOffset = -(-distVerticalStructToBBXPort + structSteelDepth + basePlateThickness);
                                planeB[0] = Plane.YZ;
                                planeB[1] = Plane.ZX;
                                axisB[0] = Axis.Y;
                                axisB[1] = Axis.X;
                                connectionPortVertical = "EndCap";
                                vertSectionHoriOffset = -steelWidth;
                                if (supportName.ToUpper().Equals("5FSX3B"))
                                {
                                    horizontalSpan = horizontalSpan - steelWidth - basePlateThickness;
                                    planeA[0] = Plane.ZX;
                                    planeA[1] = Plane.ZX;
                                    axisA[0] = Axis.Z;
                                    axisA[1] = Axis.X;
                                    if (jointType == 3)
                                        vertSectionHoriOffset = -steelWidth - overlap;
                                    else if (jointType == 2)
                                        vertSectionHoriOffset = -steelDepth;
                                    else
                                        vertSectionHoriOffset = 0;
                                }
                            }
                            else
                            {
                                horizontalSpan = horizontalSpan - boundingBoxWidth;
                                horizontalOffset = -(-horizontalSpan - distVerticalStructToBBXPort + structSteelDepth);
                                connectionPortVertical = "BeginCap";
                                vertSectionHoriOffset = 0;
                                planeB[0] = Plane.YZ;
                                planeB[1] = Plane.ZX;
                                axisB[0] = Axis.Y;
                                axisB[1] = Axis.NegativeX;
                                if (supportName.ToUpper().Equals("5FSX3B"))
                                {
                                    planeA[0] = Plane.ZX;
                                    planeA[1] = Plane.YZ;
                                    axisA[0] = Axis.Z;
                                    axisA[1] = Axis.NegativeY;
                                    if (jointType == 3)
                                        vertSectionHoriOffset = overlap;
                                    else if (jointType == 2)
                                        vertSectionHoriOffset = steelDepth;
                                    else
                                        vertSectionHoriOffset = 0;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (Configuration == 1)
                        {
                            planeB[0] = Plane.YZ;
                            planeB[1] = Plane.NegativeZX;
                            axisB[0] = Axis.Y;
                            axisB[1] = Axis.Z;
                            if (connectionPortHorizontal.Equals("BBSR_Low"))
                            {
                                horizontalSpan = horizontalSpan - boundingBoxWidth + steelWidth;
                                horizontalOffset = -(horizontalSpan - distVerticalStructToBBXPort + structSteelDepth);
                                connectionPortVertical = "BeginCap";
                                vertSectionHoriOffset = 0;

                                if (supportName.ToUpper().Equals("5FSX3B"))
                                {
                                    horizontalSpan = horizontalSpan - steelWidth; // because horizontal menber is shorter
                                    horizontalOffset = horizontalOffset + steelWidth;
                                    planeA[0] = Plane.ZX;
                                    planeA[1] = Plane.YZ;
                                    axisA[0] = Axis.Z;
                                    axisA[1] = Axis.NegativeY;
                                    if (jointType == 3)
                                        vertSectionHoriOffset = overlap;
                                    else if (jointType == 2)
                                        vertSectionHoriOffset = steelDepth;
                                    else
                                        vertSectionHoriOffset = 0;
                                }
                            }
                            else
                            {
                                horizontalOffset = -distVerticalStructToBBXPort + structSteelDepth + basePlateThickness;
                                connectionPortVertical = "EndCap";
                                vertSectionHoriOffset = -steelWidth;
                                if (supportName.ToUpper().Equals("5FSX3B"))
                                {
                                    planeA[0] = Plane.ZX;
                                    planeA[1] = Plane.ZX;
                                    axisA[0] = Axis.Z;
                                    axisA[1] = Axis.X;
                                    if (jointType == 3)
                                        vertSectionHoriOffset = -steelWidth - overlap;
                                    else if (jointType == 2)
                                        vertSectionHoriOffset = -steelDepth;
                                    else
                                        vertSectionHoriOffset = 0;
                                }
                            }
                        }
                        else if (Configuration == 2)
                        {
                            planeB[0] = Plane.YZ;
                            planeB[1] = Plane.ZX;
                            axisB[0] = Axis.Y;
                            axisB[1] = Axis.NegativeZ;
                            if (connectionPortHorizontal.Equals("BBSR_Low"))
                            {
                                horizontalSpan = horizontalSpan - boundingBoxWidth;
                                horizontalOffset = -(-distVerticalStructToBBXPort + structSteelDepth + basePlateThickness);
                                connectionPortVertical = "EndCap";
                                vertSectionHoriOffset = -steelWidth;

                                if (supportName.ToUpper().Equals("5FSX3B"))
                                {
                                    planeA[0] = Plane.ZX;
                                    planeA[1] = Plane.ZX;
                                    axisA[0] = Axis.Z;
                                    axisA[1] = Axis.X;
                                    if (jointType == 3)
                                        vertSectionHoriOffset = -steelWidth - overlap;
                                    else if (jointType == 2)
                                        vertSectionHoriOffset = -steelDepth;
                                    else
                                        vertSectionHoriOffset = 0;
                                }
                            }
                            else
                            {
                                horizontalSpan = horizontalSpan - boundingBoxWidth + steelWidth;
                                horizontalOffset = (horizontalSpan - distVerticalStructToBBXPort + structSteelDepth);
                                connectionPortVertical = "BeginCap";
                                vertSectionHoriOffset = 0;
                                if (supportName.ToUpper().Equals("5FSX3B"))
                                {
                                    horizontalSpan = horizontalSpan - steelWidth; // because horizontal menber is shorter
                                    horizontalOffset = horizontalOffset + steelWidth;
                                    planeA[0] = Plane.ZX;
                                    planeA[1] = Plane.YZ;
                                    axisA[0] = Axis.Z;
                                    axisA[1] = Axis.NegativeY;
                                    if (jointType == 3)
                                        vertSectionHoriOffset = overlap;
                                    else if (jointType == 2)
                                        vertSectionHoriOffset = steelDepth;
                                    else
                                        vertSectionHoriOffset = 0;
                                }
                            }
                        }
                    }
                }

                string basePlateConnectionPort = string.Empty, basePlatePort = string.Empty;
                double horizontalLength = 0, verticalLength = 0;

                if (basePlate.ToUpper().Equals("WITH"))
                {
                    if (planeA[0] == Plane.ZX && planeA[1] == Plane.YZ && axisA[0] == Axis.Z && axisA[1] == Axis.NegativeY)
                    {
                        basePlateConnectionPort = "EndCap";
                        horizontalLength = horizontalSpan - basePlateThickness;
                        basePlatePort = "StartOther";
                    }
                    else
                    {
                        basePlateConnectionPort = "BeginCap";
                        horizontalLength = horizontalSpan;
                        basePlatePort = "EndOther";
                    }
                }
                else
                    horizontalLength = horizontalSpan;

                verticalLength = RefPortHelper.DistanceBetweenPorts(connectionPortHorizontal, horizontalStructPort, PortDistanceType.Vertical);
                if (supportName.ToUpper().Equals("FS6B"))
                {
                    planeA[0] = Plane.ZX;
                    planeA[1] = Plane.ZX;
                    axisA[0] = Axis.Z;
                    axisA[1] = Axis.X;
                }
                else // Support is 5FSX3B
                {
                    if (jointType == 1)
                    {
                        verticalLength = verticalLength + steelDepth;
                        vertSectionVertOffset = steelDepth;
                    }
                    else if (jointType == 2)
                    {
                        vertSectionHoriOffset = steelDepth;
                        vertSectionVertOffset = basePlateThickness;
                    }
                    else
                    {
                        planeA[0] = Plane.ZX;
                        planeA[1] = Plane.NegativeYZ;
                        axisA[0] = Axis.Z;
                        axisA[1] = Axis.Y;
                        vertSectionDepthOffset = -steelDepth;
                        verticalLength = verticalLength + steelDepth + overlap;
                        vertSectionVertOffset = overlap + steelDepth;
                    }
                }

                // Check if Assyembly span is greater than dist between vertical strucutre and port
                if (horizontalSpan < distVerticalStructToBBXPort)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Assembly length is greater than maximum assembly length.", "", "LFieldSupports.cs", 590);

                if ((supportName.ToUpper().Equals("FS6B") || supportName.ToUpper().Equals("FS6B")) && jointType == 1)
                {
                    if ((horizontalSpan - steelDepth) > maxSpan)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Assembly span specified by user is greater than maximum span.", "", "LFieldSupports.cs", 595);
                }
                else if (supportName.ToUpper() == "" && jointType == 2)
                {
                    if (horizontalSpan > maxSpan)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Assembly span specified by user is greater than maximum span.", "", "LFieldSupports.cs", 600);
                }

                if (verticalLength + shoeHeight > maxAssemblyLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Assembly length is greater than maximum assembly length.", "", "LFieldSupports.cs", 604);

                // ====== ======
                // Set Values of Part Occurance Attributes
                // ====== ======
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

                componentDictionary[VERTICALSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTICALSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTICALSECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTICALSECTION].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTICALSECTION].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                // ====== ======
                // Create Joints
                // ====== ======
                componentDictionary[HORIZONTALSECTION].SetPropertyValue(horizontalLength, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTICALSECTION].SetPropertyValue(verticalLength - basePlateThickness + shoeHeight, "IJUAHgrOccLength", "Length");

                if (isPrimHorStr == false)
                {
                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint("-1", connectionPortHorizontal, HORIZONTALSECTION, "BeginCap", planeB[0], planeB[1], axisB[0], axisB[1], (steelDepth / 2 + structOffset), -shoeHeight, horizontalOffset);
                    else
                        JointHelper.CreateRigidJoint("-1", connectionPortHorizontal, HORIZONTALSECTION, "BeginCap", planeB[0], planeB[1], axisB[0], axisB[1], -(steelDepth / 2 + structOffset), -shoeHeight, horizontalOffset);

                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, connectionPortVertical, VERTICALSECTION, "BeginCap", planeA[0], planeA[1], axisA[0], axisA[1], 0, vertSectionVertOffset, vertSectionHoriOffset);
                }
                else if (isPrimHorStr == true)
                {
                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint("-1", connectionPortHorizontal, HORIZONTALSECTION, "BeginCap", planeB[0], planeB[1], axisB[0], axisB[1], (steelDepth / 2 + structOffset), horizontalOffset, shoeHeight);
                    else
                        JointHelper.CreateRigidJoint("-1", connectionPortHorizontal, HORIZONTALSECTION, "BeginCap", planeB[0], planeB[1], axisB[0], axisB[1], -(steelDepth / 2 + structOffset), horizontalOffset, shoeHeight);

                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, connectionPortVertical, VERTICALSECTION, "BeginCap", planeA[0], planeA[1], axisA[0], axisA[1], 0, vertSectionVertOffset, vertSectionHoriOffset);
                }

                if (basePlate.ToUpper().Equals("WITH"))
                {
                    string baseplateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, basePlateThickness, UnitName.DISTANCE_MILLIMETER) + "Plate Steel," + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, basePlateWidth, UnitName.DISTANCE_MILLIMETER) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, basePlateDepth, UnitName.DISTANCE_MILLIMETER);
                    double basePlateWeight = 7900 * basePlateWidth * basePlateDepth * basePlateThickness;

                    // NOTE: Change this method from ConfigHlpr to HH
                    componentDictionary[BASEPLATEHORIZONTAL].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                    componentDictionary[BASEPLATEHORIZONTAL].SetPropertyValue(basePlateDepth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                    componentDictionary[BASEPLATEHORIZONTAL].SetPropertyValue(basePlateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                    componentDictionary[BASEPLATEHORIZONTAL].SetPropertyValue(baseplateBOM, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");
                    componentDictionary[BASEPLATEHORIZONTAL].SetPropertyValue(basePlateWeight, "IJOAHgrUtility_USER_FIXED_BOX", "DryWt");

                    componentDictionary[BASEPLATEVERTICAL].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                    componentDictionary[BASEPLATEVERTICAL].SetPropertyValue(basePlateDepth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                    componentDictionary[BASEPLATEVERTICAL].SetPropertyValue(basePlateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                    componentDictionary[BASEPLATEVERTICAL].SetPropertyValue(baseplateBOM, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");
                    componentDictionary[BASEPLATEVERTICAL].SetPropertyValue(basePlateWeight, "IJOAHgrUtility_USER_FIXED_BOX", "DryWt");

                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, basePlateConnectionPort, BASEPLATEHORIZONTAL, basePlatePort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, steelDepth / 2, basePlateOffset);
                    JointHelper.CreateRigidJoint(VERTICALSECTION, "EndCap", BASEPLATEVERTICAL, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, steelDepth / 2, basePlateOffset);
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
                return 2;
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

                    if (basePlate.ToUpper().Equals("WITH"))
                    {
                        structConnections.Add(new ConnectionInfo(BASEPLATEVERTICAL, 1)); // partindex, routeindex
                        structConnections.Add(new ConnectionInfo(BASEPLATEHORIZONTAL, 2)); // partindex, routeindex
                    }
                    else
                    {
                        structConnections.Add(new ConnectionInfo(VERTICALSECTION, 1)); // partindex, routeindex
                        structConnections.Add(new ConnectionInfo(HORIZONTALSECTION, 2)); // partindex, routeindex
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
                bomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJUAHgrFl_SupportName", "SupportName")).PropValue;

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
                if (eMirrorPlane == MirrorPlane.YZPlane)
                {
                    if (CurrentMirrorToggleValue == 1)
                        return 2;
                    else if (CurrentMirrorToggleValue == 2)
                        return 1;
                }

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
