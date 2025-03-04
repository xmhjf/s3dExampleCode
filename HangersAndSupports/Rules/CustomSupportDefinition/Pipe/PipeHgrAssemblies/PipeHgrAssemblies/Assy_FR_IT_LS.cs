//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_FR_IT_LS.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_FR_IT_LS
//   Author       :  Vijay
//   Creation Date:  04-04-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04-04-2013     Vijay   CR-CP-224484  Convert HS_Assembly to C# .Net  
//   07-01-2014     Ramya   DM-CP-246092	 [TR] Failure of the test 1.6.1.1 of QTP HangersDevTestSet
//   06-Jan-2015 Chethan    TR-CP-262663  Certain attributes doesn’t get modified for the support assembly “Assy_FR_UC_CS”
//   22-02-2015       PVK   TR-CP-264951  Resolve coverity issues found in November 2014 report
//   27-04-2015       PVK   TR-CP-253033  Elevation CP not shown by default
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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

    public class Assy_FR_IT_LS : CustomSupportDefinition
    {
        private const string VERT_SECTION = "VERT_SECTION";
        private const string HOR_SECTION = "HOR_SECTION";
        private const string PLATE = "PLATE";
        private string WELDPART1 = "WELDPART1";
        private string WELDPART2 = "WELDPART2";
        private string WELDPART3 = "WELDPART3";

        private double shoeH, bpWidth, bpHoleSize, bpHoleInset, width1, width2, overLap;
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
                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrLSize", "LSize")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyTS", "PLATE_SIZE")).PropValue;
                    width1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyTS", "WIDTH1")).PropValue;
                    width2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyTS", "WIDTH2")).PropValue;
                    overLap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyFR_IT_LS", "OVERLAP")).PropValue;
                    bpWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyTS", "BP_WIDTH")).PropValue;
                    basePlate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlates", "Plates")).PropValue;
                    bpHoleSize = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyTS", "BP_HOLE_SIZE")).PropValue;
                    bpHoleInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyTS", "BP_HOLE_INSET")).PropValue;
                    string weld1 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWeld", "Weld1")).PropValue;
                    string weld2 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWeld", "Weld2")).PropValue;

                    if (weld1 == "" && weld2 == "")
                        includeWeld = false;
                    else
                        includeWeld = true;

                    if (basePlate == "With")
                    {
                        parts.Add(new PartInfo(VERT_SECTION, sectionSize));
                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(PLATE, "Utility_FOUR_HOLE_PLATE_" + plateSize));

                        if (includeWeld == true)
                        {
                            parts.Add(new PartInfo(WELDPART1, weld2));
                            parts.Add(new PartInfo(WELDPART2, weld1));
                            parts.Add(new PartInfo(WELDPART3, weld1));
                        }
                    }
                    else
                    {
                        parts.Add(new PartInfo(VERT_SECTION, sectionSize));
                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        if (includeWeld == true)
                        {
                            parts.Add(new PartInfo(WELDPART1, weld1));
                            parts.Add(new PartInfo(WELDPART2, weld1));
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
                return 8;
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

                Note note1 = CreateNote("Dim1", componentDictionary[HOR_SECTION], "BeginCap");
                note1.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList1 = note1.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note1.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");
                note1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note2 = CreateNote("Dim2", componentDictionary[HOR_SECTION], "EndCap");
                note2.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList2 = note2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note2.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");
                note2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note3 = CreateNote("Dim3", componentDictionary[VERT_SECTION], "BeginCap");
                note3.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList3 = note3.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note3.SetPropertyValue(codeList3, "IJGeneralNote", "Purpose");
                note3.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                BusinessObject horizontalSectionPart = componentDictionary[HOR_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = crosssection.Width;
                double steelDepth = crosssection.Depth;

                PropertyValueCodelist topbeginMiterCodelist = (PropertyValueCodelist)componentDictionary[VERT_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (topbeginMiterCodelist.PropValue == -1)
                    topbeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist topendMiterCodelist = (PropertyValueCodelist)componentDictionary[VERT_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (topendMiterCodelist.PropValue == -1)
                    topendMiterCodelist.PropValue = 1;

                PropertyValueCodelist anglebeginMiterCodelist = (PropertyValueCodelist)componentDictionary[HOR_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (anglebeginMiterCodelist.PropValue == -1)
                    anglebeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist angleendMiterCodelist = (PropertyValueCodelist)componentDictionary[HOR_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (angleendMiterCodelist.PropValue == -1)
                    angleendMiterCodelist.PropValue = 1;

                //====== ======
                // Set Values of Part Occurance Attributes
                //====== ======

                componentDictionary[VERT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERT_SECTION].SetPropertyValue(topbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERT_SECTION].SetPropertyValue(topendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[HOR_SECTION].SetPropertyValue(anglebeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[HOR_SECTION].SetPropertyValue(angleendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                string connection, keyPlanPart = string.Empty, keyPlanPort = string.Empty;
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                        connection = "Steel";
                    else
                        connection = "Slab";
                }
                else
                    connection = "Slab";

                double length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                if (Ingr.SP3D.Content.Support.Symbols.HgrCompareDoubleService.cmpdbl(length, 0) == true)
                    length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter;

                double byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                //====== ======
                //Create Joints
                //====== ======

                //Create a collection to hold the joints

                Plane[] confRoutePlane = new Plane[2];
                Axis[] confRouteAxis = new Axis[2];
                Plane[] structRoutePlane = new Plane[2];
                Axis[] structRouteAxis = new Axis[2];
                Plane[] pointStructPlane = new Plane[2];
                Axis[] pointStructAxis = new Axis[2];
                Plane[] pointRoutePlane = new Plane[2];

                double plateThickness = 0, vertOffset = 0, routeOffset = 0, pointStructOffset = 0;

                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI, 2) / 2.0)
                {
                    if (Configuration == 1 || Configuration == 5)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = -width1;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = width1 + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                    else if (Configuration == 2 || Configuration == 6)
                    {
                        if (Math.Abs(byPointAngle1) > Math.Round(Math.PI, 2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = width1 + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = -width1;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                    else if (Configuration == 3 || Configuration == 7)
                    {
                        if (Math.Abs(byPointAngle1) > Math.Round(Math.PI, 2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = width1 + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.XY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = -width1;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.XY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                    else if (Configuration == 4 || Configuration == 8)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = -width1;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.XY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = width1 + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.XY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                }
                else      //The structure is oriented in the opposite direction
                {
                    if (Configuration == 1 || Configuration == 5)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = -width1;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = width1 + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                    else if (Configuration == 2 || Configuration == 6)
                    {
                        if (Math.Abs(byPointAngle1) > Math.Round(Math.PI, 2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = width1 + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = -width1;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                    else if (Configuration == 3 || Configuration == 7)
                    {
                        if (Math.Abs(byPointAngle1) > Math.Round(Math.PI, 2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = width1 + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.XY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = -width1;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.XY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                    else if (Configuration == 4 || Configuration == 8)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = -width1;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.XY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = width1 + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.XY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (Configuration == 1)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.NegativeZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.NegativeX;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.NegativeX;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = boundingBoxWidth / 2.0 - width1 / 2.0;
                    }
                    else if (Configuration == 2)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.NegativeZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.X;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.X;

                        vertOffset = steelWidth / 2.0;
                        routeOffset = boundingBoxWidth / 2.0 + width1 / 2.0;
                    }
                    else if (Configuration == 3)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.NegativeZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.X;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.XY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.NegativeX;

                        vertOffset = steelWidth / 2.0;
                        routeOffset = boundingBoxWidth / 2.0 + width1 / 2.0;
                    }
                    else if (Configuration == 4)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.NegativeZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.NegativeX;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.XY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.X;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = boundingBoxWidth / 2.0 - width1 / 2.0;
                    }
                    else if (Configuration == 5)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.NegativeZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.NegativeX;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.NegativeX;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = -width1 - width2 / 2.0 + boundingBoxWidth / 2.0;
                    }
                    else if (Configuration == 6)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.NegativeZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.X;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.X;

                        vertOffset = steelWidth / 2.0;
                        routeOffset = width1 + width2 / 2.0 + boundingBoxWidth / 2.0;
                    }
                    else if (Configuration == 7)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.NegativeZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.X;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.XY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.NegativeX;

                        vertOffset = steelWidth / 2.0;
                        routeOffset = width1 + width2 / 2.0 + boundingBoxWidth / 2.0;
                    }
                    else if (Configuration == 8)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.NegativeZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.NegativeX;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.XY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.X;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = -width1 - width2 / 2.0 + boundingBoxWidth / 2.0;
                    }

                    componentDictionary[HOR_SECTION].SetPropertyValue(width1 + width2, "IJUAHgrOccLength", "Length");

                    if (basePlate == "With")
                    {
                        BusinessObject businessObjectPlate = componentDictionary[PLATE].GetRelationship("madeFrom", "part").TargetObjects[0];
                        plateThickness = (double)((PropertyValueDouble)businessObjectPlate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                        componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                        componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                        componentDictionary[PLATE].SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                        componentDictionary[PLATE].SetPropertyValue(bpHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                        //Add Joint Between Structure and Vertical Beam
                        if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], plateThickness, vertOffset);
                        else if (Configuration == 3 || Configuration == 4 || Configuration == 7 || Configuration == 8)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "BeginCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], plateThickness, vertOffset);

                        //Add Joint Between the Plate and the Vertical Beam
                        if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                        {
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                            keyPlanPart = PLATE;    //set the part/port for the KP object
                            keyPlanPort = "BotStructure";   //set the part/port for the KP object
                        }
                        else if (Configuration == 3 || Configuration == 4 || Configuration == 7 || Configuration == 8)
                        {
                            JointHelper.CreateRigidJoint(PLATE, "BotStructure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                            keyPlanPart = PLATE;    //set the part/port for the KP object
                            keyPlanPort = "TopStructure";   //set the part/port for the KP object
                        }
                        if (includeWeld == true)
                        {
                            //Add Joint Between the Plate and the Weld Part 1
                            JointHelper.CreateRigidJoint(PLATE, "BotStructure", WELDPART1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -bpWidth / 2, 0);//**************************************
                            //Add Joint Between the Vertical Beam and the Weld Part 3
                            JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", WELDPART3, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }
                    }
                    else
                    {
                        //Add Joint Between Structure and Vertical Beam
                        if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                        {
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], 0.0, vertOffset);
                            keyPlanPart = VERT_SECTION;  //set the part/port for the KP object
                            keyPlanPort = "EndCap";     //set the part/port for the KP object
                        }
                        else if (Configuration == 3 || Configuration == 4 || Configuration == 7 || Configuration == 8)
                        {
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "BeginCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], 0.0, vertOffset);
                            keyPlanPart = VERT_SECTION;      //set the part/port for the KP object
                            keyPlanPort = "BeginCap";       //set the part/port for the KP object
                        }

                        if (includeWeld == true)
                            //Add Joint Between the Vertical Beam and the Weld Part 1
                            JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", WELDPART1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                    //Add joints between Route and Beam
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "EndCap", confRoutePlane[0], confRoutePlane[1], confRouteAxis[0], confRouteAxis[1], -shoeH, routeOffset);

                    if (includeWeld == true)
                        //Add Joint Between the Horizontal Beam and the Weld Part 2
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", WELDPART2, "Other", Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeY, 0, 0, width2);

                    //Add Joint Between the Horizontal and Vertical Beams
                    if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeY, steelDepth - overLap, width2 - steelDepth / 2.0, 0);
                    else if (Configuration == 3 || Configuration == 4 || Configuration == 7 || Configuration == 8)
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, steelDepth - overLap, width2 + steelDepth / 2.0, 0);

                    //Flexible Member
                    JointHelper.CreatePrismaticJoint(VERT_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                else
                {
                    if (connection == "Steel")
                    {
                        if (basePlate == "With")
                        {
                            BusinessObject businessObjectPlate = componentDictionary[PLATE].GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)businessObjectPlate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE].SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE].SetPropertyValue(bpHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            componentDictionary[VERT_SECTION].SetPropertyValue(length + pipeDiameter / 2.0 + steelDepth - overLap + shoeH - plateThickness, "IJUAHgrOccLength", "Length");

                            //Add Joint Between Structure and Vertical Beam
                            if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                                JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], plateThickness, pointStructOffset, 0);
                            else if (Configuration == 3 || Configuration == 4 || Configuration == 7 || Configuration == 8)
                                JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "BeginCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], plateThickness, pointStructOffset, 0);

                            //Add Joint Between the Plate and the Vertical Beam
                            if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                            {
                                JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                                keyPlanPart = PLATE;    //set the part/port for the KP object
                                keyPlanPort = "BotStructure";   //set the part/port for the KP object
                            }
                            else if (Configuration == 3 || Configuration == 4 || Configuration == 7 || Configuration == 8)
                            {
                                JointHelper.CreateRigidJoint(PLATE, "BotStructure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                                keyPlanPart = PLATE;    //set the part/port for the KP object
                                keyPlanPort = "TopStructure";   //set the part/port for the KP object
                            }
                            if (includeWeld == true)
                            {
                                //Add Joint Between the Plate and the Weld Part 1
                                JointHelper.CreateRigidJoint(PLATE, "BotStructure", WELDPART1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -bpWidth / 2, 0);

                                //Add Joint Between the Vertical Beam and the Weld Part 3
                                JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", WELDPART3, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            }
                        }

                        else
                        {
                            componentDictionary[VERT_SECTION].SetPropertyValue(length + pipeDiameter / 2.0 + steelDepth - overLap + shoeH, "IJUAHgrOccLength", "Length");

                            //Add Joint Between Structure and Vertical Beam
                            if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                            {
                                JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], 0.0, pointStructOffset, 0.0);
                                keyPlanPart = VERT_SECTION;      //set the part/port for the KP object
                                keyPlanPort = "EndCap";     //set the part/port for the KP object
                            }
                            else if (Configuration == 3 || Configuration == 4 || Configuration == 7 || Configuration == 8)
                            {
                                JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "BeginCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], 0.0, pointStructOffset, 0.0);
                                keyPlanPart = VERT_SECTION;      //set the part/port for the KP object
                                keyPlanPort = "BeginCap";       //set the part/port for the KP object
                            }
                            if (includeWeld == true)
                                //Add Joint Between the Vertical Beam and the Weld Part 1
                                JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", WELDPART1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }

                        //Add Joint Between the Horizontal and Vertical Beams
                        if (Configuration == 1 || Configuration == 5)
                            JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeY, steelDepth - overLap, width2 - steelDepth / 2.0, 0);
                        else if (Configuration == 2 || Configuration == 6)
                            JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeY, steelDepth - overLap, -(width2 + steelDepth / 2.0), 0);
                        else if (Configuration == 3 || Configuration == 7)
                            JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, steelDepth - overLap, -(width2 - steelDepth / 2.0), 0);
                        else if (Configuration == 4 || Configuration == 8)
                            JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, steelDepth - overLap, width2 + steelDepth / 2.0, 0);

                        if (includeWeld == true)
                            //Add Joint Between the Horizontal Beam and the Weld Part 2
                            JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", WELDPART2, "Other", Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeY, 0, 0, width2);

                        //Add joints between Route and Beam
                        if (Configuration == 1 || Configuration == 4 || Configuration == 5 || Configuration == 8)
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "EndCap", pointRoutePlane[0], pointRoutePlane[1], routeOffset);
                        else if (Configuration == 2 || Configuration == 3 || Configuration == 6 || Configuration == 7)
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "BeginCap", pointRoutePlane[0], pointRoutePlane[1], routeOffset);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(HOR_SECTION, "BeginCap", HOR_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                    double byOffSet = 0;
                    if (connection == "Slab")
                    {
                        componentDictionary[HOR_SECTION].SetPropertyValue(width1 + width2, "IJUAHgrOccLength", "Length");

                        if (Configuration == 1)
                        {
                            vertOffset = (width2 - steelWidth / 2.0) / 2.0;

                            structRoutePlane[0] = Plane.ZX;
                            structRoutePlane[1] = Plane.NegativeXY;
                            structRouteAxis[0] = Axis.X;
                            structRouteAxis[1] = Axis.NegativeX;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeX;

                            routeOffset = (width2 - steelWidth / 2.0) / 2.0 + boundingBoxWidth / 2.0;
                            pointStructOffset = width1 + steelWidth / 2.0;
                            byOffSet = steelDepth - overLap;
                        }
                        if (Configuration == 2)
                        {
                            vertOffset = -(width1 - steelWidth / 2.0) / 2.0;

                            structRoutePlane[0] = Plane.XY;
                            structRoutePlane[1] = Plane.NegativeYZ;
                            structRouteAxis[0] = Axis.Y;
                            structRouteAxis[1] = Axis.NegativeY;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            routeOffset = -(width1 - steelWidth / 2.0) / 2.0 + boundingBoxWidth / 2.0;
                            pointStructOffset = steelDepth - overLap;
                            byOffSet = width2 + steelWidth / 2.0;
                        }
                        if (Configuration == 3)
                        {
                            vertOffset = -width1 / 2.0 - 3 * steelWidth / 4;

                            structRoutePlane[0] = Plane.XY;
                            structRoutePlane[1] = Plane.NegativeYZ;
                            structRouteAxis[0] = Axis.Y;
                            structRouteAxis[1] = Axis.NegativeY;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            routeOffset = (width1 + steelWidth / 2.0) / 2.0 + boundingBoxWidth / 2.0 + width2;
                            pointStructOffset = steelDepth - overLap;
                            byOffSet = width2 + steelWidth / 2.0;
                        }
                        if (Configuration == 4)
                        {
                            vertOffset = width2 / 2.0 + 3 * steelWidth / 4;

                            structRoutePlane[0] = Plane.ZX;
                            structRoutePlane[1] = Plane.NegativeXY;
                            structRouteAxis[0] = Axis.X;
                            structRouteAxis[1] = Axis.NegativeX;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.X;

                            routeOffset = -(width2 - steelWidth / 2.0) / 2.0 + boundingBoxWidth / 2.0 - width1;
                            pointStructOffset = width1 + steelWidth / 2.0;
                            byOffSet = steelDepth - overLap;
                        }
                        if (Configuration == 5)
                        {
                            vertOffset = -(width2 + 3 * steelWidth / 2.0) / 2.0;

                            structRoutePlane[0] = Plane.ZX;
                            structRoutePlane[1] = Plane.NegativeXY;
                            structRouteAxis[0] = Axis.X;
                            structRouteAxis[1] = Axis.NegativeX;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeX;

                            routeOffset = (width2 + steelWidth / 2.0) / 2.0 + boundingBoxWidth / 2.0 + width1;
                            pointStructOffset = width1 + steelWidth / 2.0;
                            byOffSet = steelDepth - overLap;
                        }
                        if (Configuration == 6)
                        {
                            vertOffset = (width1 + 3 * steelWidth / 2.0) / 2.0;

                            structRoutePlane[0] = Plane.XY;
                            structRoutePlane[1] = Plane.NegativeYZ;
                            structRouteAxis[0] = Axis.Y;
                            structRouteAxis[1] = Axis.NegativeY;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeY;

                            routeOffset = (width2 + steelWidth / 2.0) / 2.0 + boundingBoxWidth / 2.0 + width1;
                            pointStructOffset = width1 + steelWidth / 2.0;
                            byOffSet = steelDepth - overLap;
                        }
                        if (Configuration == 7)
                        {
                            vertOffset = (width1 - steelWidth / 2.0) / 2.0;

                            structRoutePlane[0] = Plane.XY;
                            structRoutePlane[1] = Plane.NegativeYZ;
                            structRouteAxis[0] = Axis.Y;
                            structRouteAxis[1] = Axis.NegativeY;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.Y;

                            routeOffset = (width1 - steelWidth / 2.0) / 2.0 + boundingBoxWidth / 2.0;
                            pointStructOffset = steelDepth - overLap;
                            byOffSet = (width2 + steelWidth / 2.0);
                        }
                        if (Configuration == 8)
                        {
                            vertOffset = -(width2 - steelWidth / 2.0) / 2.0;

                            structRoutePlane[0] = Plane.ZX;
                            structRoutePlane[1] = Plane.NegativeXY;
                            structRouteAxis[0] = Axis.X;
                            structRouteAxis[1] = Axis.NegativeX;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.X;

                            routeOffset = -(width2 - steelWidth / 2.0) / 2.0 + boundingBoxWidth / 2.0;
                            pointStructOffset = (width1 + steelWidth / 2.0);
                            byOffSet = (steelDepth - overLap);
                        }

                        if (basePlate == "With")
                        {
                            BusinessObject businessObjectPlate = componentDictionary[PLATE].GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)businessObjectPlate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE].SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE].SetPropertyValue(bpHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");
                            componentDictionary[VERT_SECTION].SetPropertyValue(length + pipeDiameter / 2.0 + steelDepth - overLap + shoeH - plateThickness, "IJUAHgrOccLength", "Length");

                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], plateThickness, vertOffset, 0);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);

                            if (includeWeld == true)
                            {
                                //Add Joint Between the Plate and the Weld Part 1
                                JointHelper.CreateRigidJoint(PLATE, "BotStructure", WELDPART1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -bpWidth / 2, 0);

                                //Add Joint Between the Vertical Beam and the Weld Part 3
                                JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", WELDPART3, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            }
                            //set the part/port for the KP object
                            keyPlanPart = PLATE;
                            keyPlanPort = "BotStructure";
                        }
                        else
                        {
                            componentDictionary[VERT_SECTION].SetPropertyValue(length + pipeDiameter / 2.0 + steelDepth - overLap + shoeH, "IJUAHgrOccLength", "Length");

                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], 0, vertOffset, 0);

                            if (includeWeld == true)
                                //Add Joint Between the Vertical Beam and the Weld Part 1
                                JointHelper.CreateRigidJoint(VERT_SECTION, "EndCap", WELDPART1, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                            //set the part/port for the KP object
                            keyPlanPart = VERT_SECTION;
                            keyPlanPort = "EndCap";
                        }
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "BeginCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], byOffSet, pointStructOffset, 0);

                        if (includeWeld == true)
                            //Add Joint Between the Horizontal Beam and the Weld Part 2
                            JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", WELDPART2, "Other", Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeY, 0, 0, width2);

                        //Add joints between Route and Beam
                        if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "EndCap", pointRoutePlane[0], pointRoutePlane[1], routeOffset);
                        else if (Configuration == 3 || Configuration == 4 || Configuration == 7 || Configuration == 8)
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "EndCap", pointRoutePlane[0], pointRoutePlane[1], routeOffset);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(HOR_SECTION, "BeginCap", HOR_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                }
               
                //New for V9.1, the ability to place notes anywhere, not just at ports.  Use the CreateCPAtPoint method.  This
                //method creates a CP first and then uses the location of the CP to define the location of the note.
                //This example creates the CP 100mm from the end of the HgrBeam
                Note note;
                ControlPoint controlPoint;
                bool excludeNote;
                if (support.SupportsInterface("IJUAhsExcludeNotes"))
                    excludeNote = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                else
                    excludeNote = false;


                 note = CreateNote("KeyPlan Callout", keyPlanPart, keyPlanPort, new Position(0.0, 0.0, 0.0), "", false, 2, 54, out controlPoint);
                 if (excludeNote == false)
                 {

                     //New in V9.1 is the ability to show CAD Details on fab drawings of H&S.  This object will tell the drawing task where the
                     //cad detail sht is and what it should point to
                     if (basePlate == "With")
                         note = CreateNote("CAD Detail", keyPlanPart, keyPlanPort, new Position(0.0, 0.0, 0.0), @"\HangersAndSupports\CAD Details\Utility four hole plate.sym", false, 2, 53, out controlPoint);
                 }
                 else
                 {
                     if (basePlate == "With")
                         DeleteNoteIfExists("CAD Detail");
                 }

                string verSecPort = string.Empty;
                double zOffset = 0;
                if (keyPlanPort == "BotStructure" || keyPlanPort == "EndCap")
                {
                    verSecPort = "BeginCap";
                    zOffset = 0.1;
                }
                else
                {
                    verSecPort = "EndCap";
                    zOffset = -0.1;
                }

                if (excludeNote == false)
                    note = CreateNote("Ad Hoc Note", VERT_SECTION, verSecPort, new Position(0.0, 0.0, zOffset), "Paint member yellow", false, 2, 52, out controlPoint);
                else
                    DeleteNoteIfExists("Ad Hoc Note");

                //New for V9.1 is the ability for drawings to display elevations on fabrication drawings of H&S.  This object provides the
                //location in the support that we want an elevation label.

                string horSecPort = string.Empty;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (Configuration == 5 || Configuration == 6 || Configuration == 7 || Configuration == 8)
                        horSecPort = "EndCap";
                    else
                        horSecPort = "BeginCap";
                }
                else   //Place By Point
                {
                    if (Configuration == 2 || Configuration == 3 || Configuration == 6 || Configuration == 7)
                        horSecPort = "EndCap";
                    else
                        horSecPort = "BeginCap";
                }
                if (excludeNote == false)
                    note = CreateNote("Elevation Callout", HOR_SECTION, horSecPort, new Position(0.0, 0.0, 0.0), "EL", false, 2, 51, out controlPoint);
                else
                    DeleteNoteIfExists("Elevation Callout");
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

                    //For Clamp
                    for (int index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                        routeConnections.Add(new ConnectionInfo("HOR_SECTION", index));     //partindex, routeindex

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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    int numStruct = 0;
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        numStruct = 1;
                    else
                        numStruct = SupportHelper.SupportingObjects.Count;

                    for (int index = 1; index <= numStruct; index++)
                    {
                        if (basePlate == "With")
                            structConnections.Add(new ConnectionInfo("PLATE", index));  //partindex, routeindex
                        else
                            structConnections.Add(new ConnectionInfo("VERT_SECTION", index));   //partindex, routeindex
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

        public override int MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane)
        {
            try
            {

                string connection = string.Empty;
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                        connection = "Steel";
                    else
                        connection = "Slab";
                }
                else
                    connection = "Slab";

                //Find out the plane about which the mirroring is being done
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (eMirrorPlane == MirrorPlane.YZPlane)
                    {
                        if (CurrentMirrorToggleValue == 1)
                            return 7;
                        else if (CurrentMirrorToggleValue == 2)
                            return 8;
                        else if (CurrentMirrorToggleValue == 3)
                            return 5;
                        else if (CurrentMirrorToggleValue == 4)
                            return 6;
                        else if (CurrentMirrorToggleValue == 5)
                            return 3;
                        else if (CurrentMirrorToggleValue == 6)
                            return 4;
                        else if (CurrentMirrorToggleValue == 7)
                            return 1;
                        else if (CurrentMirrorToggleValue == 8)
                            return 2;
                    }
                    else if (eMirrorPlane == MirrorPlane.XZPlane)
                    {
                        if (CurrentMirrorToggleValue == 1)
                            return 8;
                        else if (CurrentMirrorToggleValue == 2)
                            return 7;
                        else if (CurrentMirrorToggleValue == 3)
                            return 6;
                        else if (CurrentMirrorToggleValue == 4)
                            return 5;
                        else if (CurrentMirrorToggleValue == 5)
                            return 4;
                        else if (CurrentMirrorToggleValue == 6)
                            return 3;
                        else if (CurrentMirrorToggleValue == 7)
                            return 2;
                        else if (CurrentMirrorToggleValue == 8)
                            return 1;
                    }
                }
                else if (connection == "Slab")
                {
                    if (eMirrorPlane == MirrorPlane.YZPlane)
                    {
                        if (CurrentMirrorToggleValue == 1)
                            return 7;
                        else if (CurrentMirrorToggleValue == 2)
                            return 8;
                        else if (CurrentMirrorToggleValue == 3)
                            return 5;
                        else if (CurrentMirrorToggleValue == 4)
                            return 6;
                        else if (CurrentMirrorToggleValue == 5)
                            return 3;
                        else if (CurrentMirrorToggleValue == 6)
                            return 4;
                        else if (CurrentMirrorToggleValue == 7)
                            return 1;
                        else if (CurrentMirrorToggleValue == 8)
                            return 2;
                    }
                    else if (eMirrorPlane == MirrorPlane.XZPlane)
                    {
                        if (CurrentMirrorToggleValue == 1)
                            return 2;
                        else if (CurrentMirrorToggleValue == 2)
                            return 1;
                        else if (CurrentMirrorToggleValue == 3)
                            return 4;
                        else if (CurrentMirrorToggleValue == 4)
                            return 3;
                        else if (CurrentMirrorToggleValue == 5)
                            return 6;
                        else if (CurrentMirrorToggleValue == 6)
                            return 5;
                        else if (CurrentMirrorToggleValue == 7)
                            return 8;
                        else if (CurrentMirrorToggleValue == 8)
                            return 7;
                    }
                }
                else
                {
                    if (CurrentMirrorToggleValue == 1)
                        return 3;
                    else if (CurrentMirrorToggleValue == 2)
                        return 4;
                    else if (CurrentMirrorToggleValue == 3)
                        return 1;
                    else if (CurrentMirrorToggleValue == 4)
                        return 2;
                    else if (CurrentMirrorToggleValue == 5)
                        return 7;
                    else if (CurrentMirrorToggleValue == 6)
                        return 8;
                    else if (CurrentMirrorToggleValue == 7)
                        return 5;
                    else if (CurrentMirrorToggleValue == 8)
                        return 6;
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
