//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_BC_CBW2.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_BC_CBW2
//   Author       :  Vijay
//   Creation Date:  19-04-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19-04-2013     Vijay   CR-CP-224484  Convert HS_Assembly to C# .Net  
//   26-03-2014  B.Chethan  DM 252050  Assy_FR_BC_CS support symbol is not shown on all segments of the pipeline in ISO
//   22-02-2015       PVK   TR-CP-264951  Resolve coverity issues found in November 2014 report
//   27-04-2015       PVK   TR-CP-253033  Elevation CP not shown by default
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
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

    public class Assy_BC_CBW2 : CustomSupportDefinition
    {
        private const string TOP_SECTION = "TOP_SECTION";
        private const string ANGLE_SECTION = "ANGLE_SECTION";
        private const string TOP_PLATE = "TOP_PLATE";
        private const string BOT_PLATE = "BOT_PLATE";
        private const string WELDPART1 = "WELDPART1";
        private const string WELDPART2 = "WELDPART2";
        private const string WELDPART3 = "WELDPART3";

        private double shoeH, beamOverHang, overHang, bpWidth, bpHoleSize, bpHoleInset, cutBackAngle;
        private string plateSize, basePlate, sectionSize;
        private Boolean includeWeld;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyShoeH", "SHOE_H")).PropValue;
                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWSize", "WSize")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyBC", "PLATE_SIZE")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyBC", "OVERHANG")).PropValue;
                    beamOverHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyBC", "BEAM_OVERHANG")).PropValue;
                    bpWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyBC", "BP_WIDTH")).PropValue;
                    basePlate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlates", "Plates")).PropValue;
                    bpHoleSize = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyBC", "BP_HOLE_SIZE")).PropValue;
                    bpHoleInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyBC", "BP_HOLE_INSET")).PropValue;
                    cutBackAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyCutback", "CUTBACK_ANGLE")).PropValue;
                    string weld1 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWeld", "Weld1")).PropValue;
                    string weld2 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWeld", "Weld2")).PropValue;

                    //on error goto errorHandler
                    if (weld1 == "" && weld2 == "")
                    {
                        includeWeld = false;
                    }
                    else
                        includeWeld = true;
                    if (basePlate == "With")
                    {
                        parts.Add(new PartInfo(TOP_SECTION, sectionSize));
                        parts.Add(new PartInfo(ANGLE_SECTION, "Utility_CUTBACK_W2_1"));
                        parts.Add(new PartInfo(TOP_PLATE, "Utility_FOUR_HOLE_PLATE_" + plateSize));
                        parts.Add(new PartInfo(BOT_PLATE, "Utility_FOUR_HOLE_PLATE_" + plateSize));

                        if (includeWeld == true)
                        {
                            parts.Add(new PartInfo(WELDPART1, weld1));
                            parts.Add(new PartInfo(WELDPART2, weld2));
                            parts.Add(new PartInfo(WELDPART3, weld1));
                        }
                    }
                    else
                    {
                        parts.Add(new PartInfo(TOP_SECTION, sectionSize));
                        parts.Add(new PartInfo(ANGLE_SECTION, "Utility_CUTBACK_W2_1"));
                        if (includeWeld == true)
                        {
                            parts.Add(new PartInfo(WELDPART1, weld1));
                            parts.Add(new PartInfo(WELDPART2, weld2));
                        }
                    }
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
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                //Load standard bounding box definition
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                //==========================
                //1. Load standard bounding box definition
                //==========================

                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                //=========================
                //2. Get bounding box boundary objects dimension information
                //=========================

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                //====== ======
                //3. retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry

                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | dHeight
                // |____________________|
                //        dWidth

                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                Note note1 = CreateNote("Dim1", componentDictionary[TOP_SECTION], "BeginCap");
                note1.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList1 = note1.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note1.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");
                note1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note2 = CreateNote("Dim2", componentDictionary[TOP_SECTION], "EndCap");
                note2.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList2 = note2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note2.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");
                note2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note3 = CreateNote("Dim3", componentDictionary[ANGLE_SECTION], "EndStructure");
                note3.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList3 = note3.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note3.SetPropertyValue(codeList3, "IJGeneralNote", "Purpose");
                note3.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                BusinessObject horizontalSectionPart = componentDictionary[TOP_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = crosssection.Width;
                double steelDepth = crosssection.Depth;
                double steelFlange = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                double steelWeb = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tw")).PropValue;

                PropertyValueCodelist topbeginMiterCodelist = (PropertyValueCodelist)componentDictionary[TOP_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (topbeginMiterCodelist.PropValue == -1)
                    topbeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist topendMiterCodelist = (PropertyValueCodelist)componentDictionary[TOP_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (topendMiterCodelist.PropValue == -1)
                    topendMiterCodelist.PropValue = 1;

                //====== ======
                // Set Values of Part Occurance Attributes
                //====== ======

                componentDictionary[TOP_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[TOP_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[TOP_SECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[TOP_SECTION].SetPropertyValue(topbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[TOP_SECTION].SetPropertyValue(topendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                //====== ======
                //Create Joints
                //====== ======

                //Create a collection to hold the joints

                Plane[] confStructPlane = new Plane[2];
                Axis[] confStructAxis = new Axis[2];

                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter;

                int index2 = 0;
                double length = 0, tempLength = 0, offset = 0, structOffset = 0, originBraceOffset = 0;
                for (int index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                {
                    if (index == 1)
                    {
                        tempLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                        if (HgrCompareDoubleService.cmpdbl(tempLength, 0) == true)
                            tempLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                    }
                    else
                    {
                        tempLength = RefPortHelper.DistanceBetweenPorts("Route_" + index.ToString(), "Structure", PortDistanceType.Horizontal);
                        if (HgrCompareDoubleService.cmpdbl(tempLength, 0) == true)
                            tempLength = RefPortHelper.DistanceBetweenPorts("Route_" + index.ToString(), "Structure", PortDistanceType.Vertical);
                    }
                    if (length < tempLength)
                    {
                        length = tempLength;
                        index2 = index;
                    }
                    tempLength = 0;
                }

                PipeObjectInfo routeInfo1 = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(index2);
                double pipeDiaHeight = routeInfo1.OutsideDiameter;

                double angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double angle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Z, OrientationAlong.Global_Z);
                double angle3 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);
                if (Math.Abs(angle) < Math.Round(Math.PI, 2) / 2)
                {
                    if (Configuration == 1)
                    {
                        offset = -steelWidth / 2.0 - shoeH;
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            confStructPlane[0] = Plane.XY;
                            confStructPlane[1] = Plane.XY;
                            confStructAxis[0] = Axis.X;
                            confStructAxis[1] = Axis.NegativeY;

                            structOffset = -steelWidth / 2;
                            originBraceOffset = steelWidth / 2;
                        }
                        else
                        {
                            if ((SupportHelper.SupportingObjects.Count != 0))
                            {
                                if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && HgrCompareDoubleService.cmpdbl(Math.Round(angle3, 3) , Math.Round(Math.Round(Math.PI, 2) / 2, 3))==true)
                                {
                                    confStructPlane[0] = Plane.XY;
                                    confStructPlane[1] = Plane.XY;
                                    confStructAxis[0] = Axis.X;
                                    confStructAxis[1] = Axis.NegativeY;

                                    structOffset = -steelWidth / 2;
                                    originBraceOffset = steelWidth / 2;
                                }
                                else
                                {
                                    confStructPlane[0] = Plane.XY;
                                    confStructPlane[1] = Plane.XY;
                                    confStructAxis[0] = Axis.X;
                                    confStructAxis[1] = Axis.X;

                                    structOffset = steelWidth / 2;
                                    originBraceOffset = steelWidth / 2;
                                }
                            }
                            else
                            {
                                confStructPlane[0] = Plane.XY;
                                confStructPlane[1] = Plane.XY;
                                confStructAxis[0] = Axis.X;
                                confStructAxis[1] = Axis.X;

                                structOffset = steelWidth / 2;
                                originBraceOffset = steelWidth / 2;
                            }
                        }
                    }
                    else
                    {
                        offset = boundingBoxWidth / 2.0 + shoeH - 2 * pipeDiameter;
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            confStructPlane[0] = Plane.XY;
                            confStructPlane[1] = Plane.XY;
                            confStructAxis[0] = Axis.X;
                            confStructAxis[1] = Axis.Y;

                            structOffset = steelWidth / 2;
                            originBraceOffset = steelWidth / 2;
                        }
                        else
                        {
                            if ((SupportHelper.SupportingObjects.Count != 0))
                            {
                                if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && HgrCompareDoubleService.cmpdbl(Math.Round(angle3, 3) , Math.Round(Math.Round(Math.PI, 2) / 2, 3))==true)
                                {
                                    confStructPlane[0] = Plane.XY;
                                    confStructPlane[1] = Plane.XY;
                                    confStructAxis[0] = Axis.X;
                                    confStructAxis[1] = Axis.Y;

                                    structOffset = steelWidth / 2;
                                    originBraceOffset = steelWidth / 2;
                                }
                                else
                                {
                                    confStructPlane[0] = Plane.XY;
                                    confStructPlane[0] = Plane.XY;
                                    confStructAxis[0] = Axis.X;
                                    confStructAxis[1] = Axis.NegativeX;

                                    structOffset = steelWidth / 2;
                                    originBraceOffset = steelWidth / 2;
                                }
                            }
                            else
                            {
                                confStructPlane[0] = Plane.XY;
                                confStructPlane[0] = Plane.XY;
                                confStructAxis[0] = Axis.X;
                                confStructAxis[1] = Axis.NegativeX;

                                structOffset = steelWidth / 2;
                                originBraceOffset = steelWidth / 2;
                            }
                        }
                    }
                }
                else
                {
                    originBraceOffset = steelWidth / 2;
                    if (Configuration == 1)
                    {
                        offset = boundingBoxWidth / 2.0 + shoeH;
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            confStructPlane[0] = Plane.XY;
                            confStructPlane[1] = Plane.XY;
                            confStructAxis[0] = Axis.X;
                            confStructAxis[1] = Axis.Y;

                            structOffset = steelWidth / 2;
                        }
                        else
                        {
                            if ((SupportHelper.SupportingObjects.Count != 0))
                            {
                                if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && HgrCompareDoubleService.cmpdbl(Math.Round(angle3, 3) , Math.Round(Math.Round(Math.PI, 2) / 2, 3))==true)
                                {
                                    confStructPlane[0] = Plane.XY;
                                    confStructPlane[1] = Plane.XY;
                                    confStructAxis[0] = Axis.X;
                                    confStructAxis[1] = Axis.Y;

                                    structOffset = steelWidth / 2;
                                }
                                else
                                {
                                    confStructPlane[0] = Plane.XY;
                                    confStructPlane[1] = Plane.XY;
                                    confStructAxis[0] = Axis.X;
                                    confStructAxis[1] = Axis.X;

                                    structOffset = -steelWidth / 2;
                                }
                            }
                            else
                            {
                                confStructPlane[0] = Plane.XY;
                                confStructPlane[1] = Plane.XY;
                                confStructAxis[0] = Axis.X;
                                confStructAxis[1] = Axis.X;

                                structOffset = -steelWidth / 2;
                            }
                        }
                    }
                    else
                    {
                        offset = -boundingBoxWidth / 2.0 - shoeH;
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            confStructPlane[0] = Plane.XY;
                            confStructPlane[1] = Plane.XY;
                            confStructAxis[0] = Axis.X;
                            confStructAxis[1] = Axis.NegativeY;

                            structOffset = -steelWidth / 2;
                        }
                        else
                        {
                            if ((SupportHelper.SupportingObjects.Count != 0))
                            {
                                if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && HgrCompareDoubleService.cmpdbl(Math.Round(angle3, 3) , Math.Round(Math.Round(Math.PI, 2) / 2, 3))==true)
                                {
                                    confStructPlane[0] = Plane.XY;
                                    confStructPlane[1] = Plane.XY;
                                    confStructAxis[0] = Axis.X;
                                    confStructAxis[1] = Axis.NegativeY;

                                    structOffset = -steelWidth / 2;
                                }
                                else
                                {
                                    confStructPlane[0] = Plane.XY;
                                    confStructPlane[0] = Plane.XY;
                                    confStructAxis[0] = Axis.X;
                                    confStructAxis[1] = Axis.NegativeX;

                                    structOffset = steelWidth / 2;
                                }
                            }
                            else
                            {
                                confStructPlane[0] = Plane.XY;
                                confStructPlane[0] = Plane.XY;
                                confStructAxis[0] = Axis.X;
                                confStructAxis[1] = Axis.NegativeX;

                                structOffset = steelWidth / 2;
                            }
                        }
                    }
                }

                double triLeg1 = 0, trileg2 = 0, trileg3 = 0, calc1, z1, z4, orderLength, axisBraceOffSet, plateThickness;
                string cutBackBom;
                double cutBackAngle2 = Math.PI / 2 - cutBackAngle;
                axisBraceOffSet = (steelDepth / Math.Cos(cutBackAngle2)) / 2.0;

                if (basePlate == "With")
                {
                    BusinessObject topPlateConnectorPart = componentDictionary[TOP_PLATE].GetRelationship("madeFrom", "part").TargetObjects[0];

                    plateThickness = (double)((PropertyValueDouble)topPlateConnectorPart.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;
                    calc1 = steelDepth / Math.Cos(cutBackAngle);
                    triLeg1 = length + overHang - beamOverHang - plateThickness;
                    trileg2 = triLeg1 / Math.Tan(cutBackAngle2);
                    trileg3 = (length - plateThickness) / Math.Cos(cutBackAngle);

                    z1 = (steelDepth / 2.0) / Math.Tan(cutBackAngle);
                    z4 = (steelDepth / 2.0) / (Math.Tan(cutBackAngle2));
                    orderLength = trileg3 + z1 + z4;

                    cutBackBom = sectionSize.Substring(10, sectionSize.Length - 10) + ", Starting Angle = " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, cutBackAngle, UnitName.ANGLE_DEGREE) + ", Ending Angle = " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, (cutBackAngle2), UnitName.ANGLE_DEGREE) + ", Order Length = " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, orderLength, UnitName.DISTANCE_INCH);

                    componentDictionary[TOP_PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                    componentDictionary[TOP_PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                    componentDictionary[TOP_PLATE].SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                    componentDictionary[TOP_PLATE].SetPropertyValue(bpHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                    componentDictionary[BOT_PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                    componentDictionary[BOT_PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                    componentDictionary[BOT_PLATE].SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                    componentDictionary[BOT_PLATE].SetPropertyValue(bpHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                    componentDictionary[TOP_SECTION].SetPropertyValue(triLeg1, "IJUAHgrOccLength", "Length");
                    componentDictionary[TOP_SECTION].SetPropertyValue((pipeDiaHeight / 2) + overHang, "IJUAHgrOccOverLength", "EndOverLength");

                    componentDictionary[ANGLE_SECTION].SetPropertyValue(trileg3, "IJOAHgrUtility_GENERIC_W", "L");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(steelWidth, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(steelDepth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(steelFlange, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(steelWeb, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(cutBackAngle, "IJOAHgrUtility_CUTBACK", "ANGLE");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(cutBackAngle2, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(cutBackBom, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");


                    //'Add Connection for the end of the Angled Beam
                    JointHelper.CreateRigidJoint(TOP_SECTION, "BeginCap", TOP_PLATE, "TopStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, steelDepth / 2, steelWidth / 2);
                    JointHelper.CreateRigidJoint(TOP_PLATE, "TopStructure", BOT_PLATE, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, steelDepth / 2 + trileg2, 0);

                    if (includeWeld == true)
                        //Add Joint Between the Top Plate and the Weld Part 3
                        JointHelper.CreateRigidJoint(TOP_PLATE, "BotStructure", WELDPART3, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -bpWidth / 2, 0);
                }
                else
                {
                    plateThickness = 0;
                    triLeg1 = length + overHang - beamOverHang;
                    trileg2 = triLeg1 / Math.Tan(cutBackAngle);
                    trileg3 = length / Math.Cos(cutBackAngle);
                    z1 = (steelDepth / 2.0) / Math.Tan(cutBackAngle);
                    z4 = (steelDepth / 2.0) / (Math.Tan(cutBackAngle2));
                    orderLength = trileg3 + z1 + z4;

                    cutBackBom = sectionSize.Substring(10, sectionSize.Length - 10) + ", Starting Angle = " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, cutBackAngle, UnitName.ANGLE_DEGREE) + ", Ending Angle = " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, (cutBackAngle2), UnitName.ANGLE_DEGREE) + ", Order Length = " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, orderLength, UnitName.DISTANCE_INCH);

                    componentDictionary[TOP_SECTION].SetPropertyValue(triLeg1, "IJUAHgrOccLength", "Length");
                    componentDictionary[TOP_SECTION].SetPropertyValue((pipeDiaHeight / 2) + overHang, "IJUAHgrOccOverLength", "EndOverLength");

                    componentDictionary[ANGLE_SECTION].SetPropertyValue(trileg3, "IJOAHgrUtility_GENERIC_W", "L");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(steelWidth, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(steelDepth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(steelFlange, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(steelWeb, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(cutBackAngle, "IJOAHgrUtility_CUTBACK", "ANGLE");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(cutBackAngle2, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[ANGLE_SECTION].SetPropertyValue(cutBackBom, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");

                }
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreateRigidJoint("-1", "Structure", TOP_SECTION, "BeginCap", confStructPlane[0], confStructPlane[1], confStructAxis[0], confStructAxis[1], plateThickness, structOffset, offset);

                else
                {
                    if ((SupportHelper.SupportingObjects.Count != 0))
                    {
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && HgrCompareDoubleService.cmpdbl(Math.Round(angle3, 3) , Math.Round(Math.Round(Math.PI, 2) / 2, 3))==true || HgrCompareDoubleService.cmpdbl(Math.Round(angle3, 3) , Math.Round(Math.Round(Math.PI, 2), 3))==true)
                        {
                            JointHelper.CreateRigidJoint("-1", "Structure", TOP_SECTION, "BeginCap", confStructPlane[0], confStructPlane[1], confStructAxis[0], confStructAxis[1], plateThickness, structOffset, offset);
                        }
                        else
                        {
                            if (HgrCompareDoubleService.cmpdbl(angle, 0) == true && HgrCompareDoubleService.cmpdbl(Math.Round(angle2, 3) , Math.Round(Math.Round(Math.PI, 2) / 2, 3))==true)
                                JointHelper.CreateRigidJoint("-1", "Structure", TOP_SECTION, "BeginCap", confStructPlane[0], confStructPlane[1], confStructAxis[0], confStructAxis[1], plateThickness, offset + pipeDiameter, structOffset);
                            else
                                JointHelper.CreateRigidJoint("-1", "Structure", TOP_SECTION, "BeginCap", confStructPlane[0], confStructPlane[1], confStructAxis[0], confStructAxis[1], plateThickness, offset, structOffset);
                        }
                    }
                    else
                    {
                        if (HgrCompareDoubleService.cmpdbl(angle, 0) == true && HgrCompareDoubleService.cmpdbl(Math.Round(angle2, 3) , Math.Round(Math.Round(Math.PI, 2) / 2, 3))==true)
                            JointHelper.CreateRigidJoint("-1", "Structure", TOP_SECTION, "BeginCap", confStructPlane[0], confStructPlane[1], confStructAxis[0], confStructAxis[1], plateThickness, offset + pipeDiameter, structOffset);
                        else
                            JointHelper.CreateRigidJoint("-1", "Structure", TOP_SECTION, "BeginCap", confStructPlane[0], confStructPlane[1], confStructAxis[0], confStructAxis[1], plateThickness, offset, structOffset);
                    }
                }

                JointHelper.CreateRigidJoint(TOP_SECTION, "EndCap", ANGLE_SECTION, "StartStructure", Plane.ZX, Plane.YZ, Axis.X, Axis.NegativeY, steelDepth, axisBraceOffSet, originBraceOffset);

                if (includeWeld == true)
                {
                    //'Add Joint Between the Angle Section and the Weld Part 1
                    JointHelper.CreateRigidJoint(ANGLE_SECTION, "EndStructure", WELDPART1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    // 'Add Joint Between the Angle Section and the Weld Part 2
                    JointHelper.CreateRigidJoint(ANGLE_SECTION, "StartStructure", WELDPART2, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                
                Note note;
                ControlPoint controlPoint;
                bool excludeNote;
                if (support.SupportsInterface("IJUAhsExcludeNotes"))
                    excludeNote = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                else
                    excludeNote = false;

                if (excludeNote == false)
                {
                    //This object will tell the drawing task where the cad detail sht is and what it should point to
                    if (basePlate == "With")
                        note = CreateNote("CAD Detail", TOP_PLATE, "BotStructure", new Position(0.0, 0.0, 0.0), @"\HangersAndSupports\CAD Details\Utility four hole plate.sym", false, 2, 53, out controlPoint);

                    //This method creates a CP first and then uses the location of the CP to define the location of the note.
                    //This example creates the CP 100mm from the end of the HgrBeam
                    note = CreateNote("Ad Hoc Note", TOP_SECTION, "BeginCap", new Position(0.0, 0.0, 0.1), "Paint member yellow", false, 2, 52, out controlPoint);

                    //This object provides the location in the support that we want an elevation label.
                    note = CreateNote("Elevation Callout", TOP_SECTION, "EndCap", new Position(0.0, 0.0, 0.0), "EL", false, 2, 51, out controlPoint);
                }
                else
                {
                    if (basePlate == "With")
                        DeleteNoteIfExists("CAD Detail");
                    DeleteNoteIfExists("Ad Hoc Note");
                    DeleteNoteIfExists("Elevation Callout");
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

                    for (int index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                    {
                        //Steel section included in the support that connects to the Route
                        //Value representing the route we are connecting to   
                        routeConnections.Add(new ConnectionInfo(TOP_SECTION, index));
                    }
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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(TOP_SECTION, 1));      //partindex, routeindex

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

        public override int MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane)
        {
            try
            {
                //Find out the plane about which the mirroring is being done
                if (eMirrorPlane == MirrorPlane.XYPlane)
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


