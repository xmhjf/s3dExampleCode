//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_FR_TS_LS.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_FR_TS_LS
//   Author       :  Manikanth
//   Creation Date:  11.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   11-04-2013     Manikanth   CR-CP-224484  Convert HS_Assembly to C# .Net  
//   22-02-2015       PVK       TR-CP-264951  Resolve coverity issues found in November 2014 report
//   27-04-2015       PVK       TR-CP-253033  Elevation CP not shown by default
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

    public class Assy_FR_TS_LS : CustomSupportDefinition
    {
        //Constants
        private const string VERT_SECTION = "VERT_SECTION";
        private const string HOR_SECTION = "HOR_SECTION";
        private const string PLATE = "PLATE";

        double shoeHeight, width1, width2, gap, bpWidth, basePlateHoleSize, bpHoleInset;
        string plateSize, basePlate, sectionSize, weld1, weld2, weldPart1, weldPart2, weldPart3;
        bool includeWeld;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyShoeH", "SHOE_H")).PropValue;
                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrLSize", "LSize")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyTS", "PLATE_SIZE")).PropValue;
                    width1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyTS", "WIDTH1")).PropValue;
                    width2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyTS", "WIDTH2")).PropValue;
                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyFR_TS_LS", "GAP")).PropValue;
                    bpWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyTS", "BP_WIDTH")).PropValue;
                    basePlate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlates", "Plates")).PropValue;
                    basePlateHoleSize = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyTS", "BP_HOLE_SIZE")).PropValue;
                    bpHoleInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyTS", "BP_HOLE_INSET")).PropValue;
                    weld1 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWeld", "Weld1")).PropValue;
                    weld2 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWeld", "Weld2")).PropValue;

                    if (weld1 == "" && weld2 == "")
                        includeWeld = false;
                    else
                        includeWeld = true;

                    parts.Add(new PartInfo(VERT_SECTION, sectionSize));
                    parts.Add(new PartInfo(HOR_SECTION, sectionSize));

                    if (basePlate == "With")
                    {
                        parts.Add(new PartInfo(PLATE, "Utility_FOUR_HOLE_PLATE_" + plateSize));
                        if (includeWeld == true)
                        {
                            weldPart1 = "4";
                            weldPart2 = "5";
                            weldPart3 = "6";
                            parts.Add(new PartInfo(weldPart1, weld2));
                            parts.Add(new PartInfo(weldPart2, weld1));
                            parts.Add(new PartInfo(weldPart3, weld1));
                        }
                    }
                    else
                    {
                        if (includeWeld == true)
                        {
                            weldPart1 = "3";
                            weldPart2 = "4";
                            parts.Add(new PartInfo(weldPart1, weld1));
                            parts.Add(new PartInfo(weldPart2, weld1));
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
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========

                //==========================
                //1. Load standard bounding box definition
                //==========================
                BoundingBox boundingBox;

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
                //
                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | BBXHeight
                // |____________________|
                //        BBXWidth
                //
                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject horizontalSectionPart = (componentDictionary[HOR_SECTION]).GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = crosssection.Width;                                                                                             //Width, Depth, Flange, or Web as 2nd argument
                double steelDepth = crosssection.Depth;                                                                                             //Width, Depth, Flange, or Web as 2nd argument
                double steelThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;   //Width, Depth, Flange, or Web as 2nd argument

                Note note1 = CreateNote("Dim1", (componentDictionary[HOR_SECTION]), "BeginCap");
                note1.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note1PropertyValueCL = (PropertyValueCodelist)note1.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList1 = note1PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note1.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");
                note1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note2 = CreateNote("Dim2", (componentDictionary[HOR_SECTION]), "EndCap");
                note2.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note2PropertyValueCL = (PropertyValueCodelist)note2.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList2 = note2PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note2.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");
                note2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note3 = CreateNote("Dim3", (componentDictionary[VERT_SECTION]), "BeginCap");
                note3.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note3PropertyValueCL = (PropertyValueCodelist)note3.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList3 = note3PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);  //value 3 means fabrication
                note3.SetPropertyValue(codeList3, "IJGeneralNote", "Purpose");
                note3.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================

                PropertyValueCodelist topbeginMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERT_SECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (topbeginMiterCodelist.PropValue == -1)
                    topbeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist topendMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERT_SECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (topendMiterCodelist.PropValue == -1)
                    topendMiterCodelist.PropValue = 1;

                PropertyValueCodelist anglebeginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HOR_SECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (anglebeginMiterCodelist.PropValue == -1)
                    anglebeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist angleendMiterCodelist = (PropertyValueCodelist)(componentDictionary[HOR_SECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (angleendMiterCodelist.PropValue == -1)
                    angleendMiterCodelist.PropValue = 1;

                //====== ======
                //Set Values of Part Occurance Attributes
                //====== ======
                (componentDictionary[VERT_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[VERT_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[VERT_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[VERT_SECTION]).SetPropertyValue(topbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[VERT_SECTION]).SetPropertyValue(topendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[HOR_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HOR_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HOR_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HOR_SECTION]).SetPropertyValue(topbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HOR_SECTION]).SetPropertyValue(topendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                double plateThickness = 0, byPointStructOffset = 0, byPointRouteOffset = 0, length, horizontalLength, byPointAngle1, byPointAngle2;
                Plane byPointStructPlane1 = new Plane(); Plane byPointStructPlane2 = new Plane(); Plane byPointRoutePlane1 = new Plane(); Plane byPointRoutePlane2 = new Plane();
                Axis byPointStructAxis1 = new Axis(); Axis byPointStructAxis2 = new Axis(); Axis byPointRouteAxis1 = new Axis(); Axis byPointRouteAxis2 = new Axis();

                length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                horizontalLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);   //figure out the orientation of the structure port

                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI / 2, 2))   //The structure is oriented in the standard direction
                {
                    if (Configuration == 1)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI / 2, 2))
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointStructOffset = -steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = -width1;
                        }
                        else
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.X;
                            byPointStructOffset = steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.Z;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                    }
                    else if (Configuration == 2)
                    {
                        if (Math.Abs(byPointAngle1) > Math.Round(Math.PI / 2, 2))
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointStructOffset = -steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                        else
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.X;
                            byPointStructOffset = steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                    }
                    else if (Configuration == 3)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI / 2, 2))
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.XY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.X;
                            byPointStructOffset = -steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.Z;
                            byPointRouteOffset = -width1;
                        }
                        else
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.XY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointStructOffset = steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                    }
                    else if (Configuration == 4)
                    {
                        if (Math.Abs(byPointAngle1) > Math.Round(Math.PI / 2, 2))
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.XY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.X;
                            byPointStructOffset = -steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.Z;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                        else
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.XY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointStructOffset = steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = -width1;
                        }
                    }
                }
                else                                     //The structure is oriented in the opposite direction
                {
                    if (Configuration == 1)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI / 2, 2))
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.X;
                            byPointStructOffset = steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = -width1;
                        }
                        else
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointStructOffset = -steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.Z;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                    }
                    else if (Configuration == 2)
                    {
                        if (Math.Abs(byPointAngle1) > Math.Round(Math.PI / 2, 2))
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.X;
                            byPointStructOffset = steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                        else
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointStructOffset = -steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.Z;
                            byPointRouteOffset = -width1;
                        }
                    }
                    else if (Configuration == 3)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI / 2, 2))
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.XY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointStructOffset = steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.Z;
                            byPointRouteOffset = -width1;
                        }
                        else
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.XY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.X;
                            byPointStructOffset = -steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                    }
                    else if (Configuration == 4)
                    {
                        if (Math.Abs(byPointAngle1) > Math.Round(Math.PI / 2, 2))
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.XY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointStructOffset = steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.Z;
                            byPointRouteOffset = width1 + boundingBoxWidth;
                        }
                        else
                        {
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.XY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.X;
                            byPointStructOffset = -steelDepth / 2.0;
                            byPointRoutePlane1 = Plane.XY;
                            byPointRoutePlane2 = Plane.ZX;
                            byPointRouteAxis1 = Axis.Y;
                            byPointRouteAxis2 = Axis.NegativeZ;
                            byPointRouteOffset = -width1;
                        }
                    }
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    (componentDictionary[HOR_SECTION]).SetPropertyValue(boundingBoxWidth + width1 + width2, "IJUAHgrOccLength", "Length");

                    if (basePlate == "With")
                    {
                        BusinessObject plate = (componentDictionary[PLATE]).GetRelationship("madeFrom", "part").TargetObjects[0];
                        plateThickness = (double)((PropertyValueDouble)plate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;
                        (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                        (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                        (componentDictionary[PLATE]).SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                        (componentDictionary[PLATE]).SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                        //Add Joint Between Structure and Vertical Beam
                        if (Configuration == 1)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, plateThickness, -steelWidth / 2.0);
                        else if (Configuration == 2)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, plateThickness, steelWidth / 2.0);
                        else if (Configuration == 3)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.Y, plateThickness, steelWidth / 2.0);
                        else if (Configuration == 4)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, plateThickness, -steelWidth / 2.0);

                        //Add Joint Between the Plate and the Vertical Beam
                        if (Configuration == 1 || Configuration == 2)
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                        else if (Configuration == 3 || Configuration == 4)
                            JointHelper.CreateRigidJoint(PLATE, "BotStructure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);

                        if (includeWeld)
                        {
                            //Add Joint Between the Plate and the Weld Part 1
                            JointHelper.CreateRigidJoint(PLATE, "BotStructure", weldPart1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -bpWidth / 2, 0);
                            //Add Joint Between the Vertical Beam and the Weld Part 3
                            JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", weldPart3, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }
                    }
                    else
                    {
                        //Add Joint Between Structure and Vertical Beam
                        if (Configuration == 1)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -steelWidth / 2.0);
                        else if (Configuration == 2)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0.0, steelWidth / 2.0);
                        else if (Configuration == 3)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0.0, steelWidth / 2.0);
                        else if (Configuration == 3)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0.0, -steelWidth / 2.0);

                        if (includeWeld)
                            //Add Joint Between the Vertical Beam and the Weld Part 1
                            JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", weldPart1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }

                    //Add joints between Route and Beam
                    if (Configuration == 1 || Configuration == 4)
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "EndCap", Plane.XY, Plane.ZX, Axis.X, Axis.X, boundingBoxHeight + shoeHeight, -width1);
                    else if (Configuration == 2 || Configuration == 3)
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX, boundingBoxHeight + shoeHeight, -width1);

                    if (includeWeld)
                        //Add Joint Between the Horizontal Beam and the Weld Part 2
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", weldPart2, "Other", Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeY, 0, 0, boundingBoxWidth + width1 + width2);

                    //Add Joint Between the Horizontal and Vertical Beams
                    if (Configuration == 1 || Configuration == 2)
                        JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.X, gap, -boundingBoxWidth / 2.0 - width1 + steelDepth / 2.0, steelThickness);
                    else if (Configuration == 3 || Configuration == 4)
                        JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, gap, -boundingBoxWidth / 2.0 - width1 - steelDepth / 2.0, steelThickness);

                    //Flexible Member
                    JointHelper.CreatePrismaticJoint(VERT_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                else
                {
                    bool bSlopedRoute = false, IsVerticalRoute;
                    double dRouteAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);

                    if (dRouteAngle < Math.PI / 4)
                        IsVerticalRoute = true;
                    else if (dRouteAngle > 3 * Math.PI / 4)
                        IsVerticalRoute = true;
                    else
                        IsVerticalRoute = false;

                    if (HgrCompareDoubleService.cmpdbl(Math.Round(dRouteAngle, 3) , 0) ==false)            //for sloped pipe
                        bSlopedRoute = true;
                    string sRoutePort, sStructuePort;
                    if (bSlopedRoute == false)
                    {
                        sRoutePort = "BBR_Low";
                        sStructuePort = "Structure";
                    }
                    else
                    {
                        if (IsVerticalRoute)
                        {
                            Vector BB_X = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);
                            Vector BB_Z = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X);
                            Vector GlobalX = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalX, BB_X);
                            Vector GlobalY = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalY, BB_X);
                            //AngleBetweenVectors = ACos(Vector1.Dot(Vector2) / (Vector1.Length * Vector2.Length))

                            if (((Math.Acos(GlobalX.Dot(BB_Z)) / (GlobalX.Length * BB_Z.Length)) < Math.PI / 4))
                                BB_Z = GlobalX;
                            else if (((Math.Acos(GlobalX.Dot(BB_Z)) / (GlobalX.Length * BB_Z.Length)) < 3 * Math.PI / 4))
                                BB_Z.Set(-GlobalX.X, -GlobalX.Y, GlobalX.Z);
                            else if (((Math.Acos(GlobalY.Dot(BB_Z)) / (GlobalY.Length * BB_Z.Length)) < Math.PI / 4))
                                BB_Z = GlobalY;
                            else
                                BB_Z.Set(-GlobalY.X, -GlobalY.Y, -GlobalY.Z);

                            BoundingBox BBX = BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, "HSFrameBBX", false, true, false);
                            sRoutePort = "HSFrameBBX_Low";
                            sStructuePort = "StructAlt";

                        }
                        else
                        {
                            sRoutePort = "BBRV_Low";
                            sStructuePort = "StructAlt";
                        }
                    }

                    if (horizontalLength < boundingBoxWidth)
                        (componentDictionary[HOR_SECTION]).SetPropertyValue(boundingBoxWidth + width1 + width2, "IJUAHgrOccLength", "Length");
                    else
                        (componentDictionary[HOR_SECTION]).SetPropertyValue(horizontalLength + width1 + width2, "IJUAHgrOccLength", "Length");

                    if (basePlate == "With")
                    {
                        BusinessObject plate = (componentDictionary[PLATE]).GetRelationship("madeFrom", "part").TargetObjects[0];
                        plateThickness = (double)((PropertyValueDouble)plate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                        (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                        (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                        (componentDictionary[PLATE]).SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                        (componentDictionary[PLATE]).SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                        //Add Joint Between Structure and Vertical Beam
                        if (Configuration == 1 || Configuration == 2)
                            JointHelper.CreatePrismaticJoint("-1", sStructuePort, VERT_SECTION, "EndCap", byPointStructPlane1, byPointStructPlane2, byPointStructAxis1, byPointStructAxis2, plateThickness, byPointStructOffset);
                        else if (Configuration == 3 || Configuration == 4)
                            JointHelper.CreatePrismaticJoint("-1", sStructuePort, VERT_SECTION, "BeginCap", byPointStructPlane1, byPointStructPlane2, byPointStructAxis1, byPointStructAxis2, plateThickness, byPointStructOffset);

                        //Add Joint Between the Plate and the Vertical Beam
                        if (Configuration == 1 || Configuration == 2)
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                        else if (Configuration == 3 || Configuration == 4)
                            JointHelper.CreateRigidJoint(PLATE, "BotStructure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);

                        if (includeWeld)
                        {
                            //Add Joint Between the Plate and the Weld Part 1
                            JointHelper.CreateRigidJoint(PLATE, "BotStructure", weldPart1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -bpWidth / 2.0, 0);
                            //Add Joint Between the Vertical Beam and the Weld Part 3
                            JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", weldPart3, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }
                    }
                    else
                    {
                        //Add Joint Between Structure and Vertical Beam
                        if (Configuration == 1 || Configuration == 2)
                            JointHelper.CreatePrismaticJoint("-1", sStructuePort, VERT_SECTION, "EndCap", byPointStructPlane1, byPointStructPlane2, byPointStructAxis1, byPointStructAxis2, 0.0, byPointStructOffset);
                        else if (Configuration == 3 || Configuration == 4)
                            JointHelper.CreatePrismaticJoint("-1", sStructuePort, VERT_SECTION, "BeginCap", byPointStructPlane1, byPointStructPlane2, byPointStructAxis1, byPointStructAxis2, 0.0, byPointStructOffset);

                        if (includeWeld)
                            //Add Joint Between the Vertical Beam and the Weld Part 1
                            JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", weldPart1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                    //Add Joint Between the Horizontal and Vertical Beams
                    if (Configuration == 1 || Configuration == 2)
                        JointHelper.CreatePrismaticJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeY, gap, steelThickness);
                    else if (Configuration == 3 || Configuration == 4)
                        JointHelper.CreatePrismaticJoint(HOR_SECTION, "EndCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.Y, gap, steelThickness);

                    //Add joints between Route and Beam
                    if (Configuration == 1 || Configuration == 4)
                        JointHelper.CreateRigidJoint("-1", sRoutePort, HOR_SECTION, "EndCap", byPointRoutePlane1, byPointRoutePlane2, byPointRouteAxis1, byPointRouteAxis2, shoeHeight + boundingBoxHeight, 0.0, byPointRouteOffset);
                    else if (Configuration == 2 || Configuration == 3)
                        JointHelper.CreateRigidJoint("-1", sRoutePort, HOR_SECTION, "BeginCap", byPointRoutePlane1, byPointRoutePlane2, byPointRouteAxis1, byPointRouteAxis2, shoeHeight + boundingBoxHeight, 0.0, byPointRouteOffset);

                    if (horizontalLength < boundingBoxWidth)
                    {
                        if (includeWeld)
                            //Add Joint Between the Horizontal Beam and the Weld Part 2
                            JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", weldPart2, "Other", Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeY, 0, 0, (boundingBoxWidth + width1 + width2) / 2);
                    }
                    else
                    {
                        if (includeWeld)
                            //Add Joint Between the Horizontal Beam and the Weld Part 2
                            JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", weldPart2, "Other", Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeY, 0, 0, (horizontalLength + width1 + width2) / 2);
                    }

                    JointHelper.CreatePrismaticJoint(VERT_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                
                Note note;
                ControlPoint controlPoint;
                bool excludeNote;
                if (support.SupportsInterface("IJUAhsExcludeNotes"))
                    excludeNote = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                else
                    excludeNote = false;

                if (basePlate == "With")
                {
                    //This object will tell the drawing task where the cad detail sht is and what it should point to
                    if (excludeNote == false)
                        note = CreateNote("CAD Detail", PLATE, "BotStructure", new Position(0.0, 0.0, 0.0), @"\HangersAndSupports\CAD Details\Utility four hole plate.sym", false, 2, 53, out controlPoint);
                    else
                        DeleteNoteIfExists("CAD Detail");
                    // By default the location of the support will refer to the primary pipe centerline.  This can be changed by adding a keyplan object to the support. Here we add one to the support where it attaches to the building structure.
                    note = CreateNote("KeyPlan Callout", PLATE, "BotStructure", new Position(0.0, 0.0, 0.0), "", false, 2, 54, out controlPoint);
                }
                else
                    note = CreateNote("KeyPlan Callout", VERT_SECTION, "EndCap", new Position(0.0, 0.0, 0.0), "", false, 2, 54, out controlPoint);

                //This method creates a CP first and then uses the location of the CP to define the location of the note. This example creates the CP 100mm from the end of the HgrBeam
                if (!excludeNote)
                    note = CreateNote("Ad Hoc Note", VERT_SECTION, "BeginCap", new Position(0.0, 0.0, 0.1), "Paint member yellow", false, 2, 52, out controlPoint);
                else
                    DeleteNoteIfExists("Ad Hoc Note");

                //This object provides the location in the support that we want an elevation label.
                if (!excludeNote)
                    note = CreateNote("Elevation Callout", HOR_SECTION, "BeginCap", new Position(0.0, 0.0, 0.0), "EL", false, 2, 51, out controlPoint);
                else
                    DeleteNoteIfExists("Elevation Callout");
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Assy_FR_TS_LS." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        public override int MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane)
        {
            try
            {

                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                width1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyTS", "WIDTH1")).PropValue;
                width2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyTS", "WIDTH2")).PropValue;

                support.SetPropertyValue(width2, "IJOAHgrAssyTS", "WIDTH1");
                support.SetPropertyValue(width1, "IJOAHgrAssyTS", "WIDTH2");
                if (eMirrorPlane == MirrorPlane.XZPlane)
                {
                    if (CurrentMirrorToggleValue == 1)
                        return 4;
                    else if (CurrentMirrorToggleValue == 2)
                        return 3;
                    if (CurrentMirrorToggleValue == 3)
                        return 2;
                    else if (CurrentMirrorToggleValue == 4)
                        return 1;
                }
                else if (eMirrorPlane == MirrorPlane.YZPlane)
                {
                    if (CurrentMirrorToggleValue == 1)
                        return 3;
                    else if (CurrentMirrorToggleValue == 2)
                        return 4;
                    if (CurrentMirrorToggleValue == 3)
                        return 1;
                    else if (CurrentMirrorToggleValue == 4)
                        return 2;
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
        public override int ConfigurationCount
        {
            get
            {
                return 4;
            }
        }
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold ALL the Route Connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    for (int index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                    {
                        //Steel section included in the support that connects to the Route
                        //Value representing the route we are connecting to   
                        routeConnections.Add(new ConnectionInfo(HOR_SECTION, index));
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
       

        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

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
    }
}