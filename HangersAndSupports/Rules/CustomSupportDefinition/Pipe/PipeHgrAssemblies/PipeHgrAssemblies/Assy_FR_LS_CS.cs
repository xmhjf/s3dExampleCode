//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_FR_LS_CS.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_FR_LS_CS
//   Author       :  Vijay
//   Creation Date:  08-04-2013   
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   08-04-2013     Vijay   CR-CP-224484  Convert HS_Assembly to C# .Net  
//   07-01-2014     Ramya   DM-CP-246092	 [TR] Failure of the test 1.6.1.1 of QTP HangersDevTestSet
//   06-Jan-2015 Chethan    TR-CP-262663  Certain attributes doesn’t get modified for the support assembly “Assy_FR_UC_CS”
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
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
   
    public class Assy_FR_LS_CS : CustomSupportDefinition
    {
        private const string VERT_SECTION = "VERT_SECTION";
        private const string HOR_SECTION = "HOR_SECTION";
        private const string PLATE = "PLATE";
        private double shoeH, overHang, bpWidth, bpHoleSize, bpHoleInset, extension, overLap;
        private string plateSize, basePlate, sectionSize;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyShoeH", "SHOE_H")).PropValue;
                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrCSize", "CSize")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyLS", "PLATE_SIZE")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS", "OVERHANG")).PropValue;
                    overLap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS", "OVERLAP")).PropValue;
                    bpWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS", "BP_WIDTH")).PropValue;
                    basePlate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlates", "Plates")).PropValue;
                    bpHoleSize = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS", "BP_HOLE_SIZE")).PropValue;
                    bpHoleInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS", "BP_HOLE_INSET")).PropValue;
                    extension = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyLS", "EXTENSION")).PropValue;

                    if (basePlate == "With")
                    {
                        parts.Add(new PartInfo(VERT_SECTION, sectionSize));
                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(PLATE, "Utility_FOUR_HOLE_PLATE_" + plateSize));
                    }
                    else
                    {
                        parts.Add(new PartInfo(VERT_SECTION, sectionSize));
                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
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

                string connection;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    connection = "Slab";

                else
                {
                    if ((SupportHelper.SupportingObjects.Count != 0))
                    {
                        if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                            connection = "Steel";
                        else
                            connection = "Slab";
                    }
                    else
                        connection = "Slab";
                }

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

                double vertOffset = 0, routeOffset = 0, pointStructOffset = 0, horOffset = 0;
                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI,2) / 2.0)
                {
                    if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI,2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = -overHang;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeX;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = boundingBoxWidth + overHang;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.X;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                    }
                    else
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI,2) / 2.0)
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = -overHang;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.X;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = boundingBoxWidth + overHang;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeX;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                }
                else      //The structure is oriented in the opposite direction
                {
                    if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI,2) / 2.0)
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = -overHang;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.X;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                        else
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = overHang + boundingBoxWidth;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeX;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                    }
                    else
                    {
                        if (Math.Abs(byPointAngle1) < Math.Round(Math.PI,2) / 2.0)
                        {
                            vertOffset = -steelWidth / 2.0;
                            routeOffset = -overHang;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeX;

                            pointStructOffset = -steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                        }
                        else
                        {
                            vertOffset = steelWidth / 2.0;
                            routeOffset = boundingBoxWidth + overHang;

                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;
                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.X;

                            pointStructOffset = steelDepth / 2.0;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                        }
                    }
                }

                double pointLength, plateThickness;
                if (Configuration == 1 || Configuration == 3 || Configuration == 5 || Configuration == 7)
                    pointLength = length + pipeDiameter / 2.0 + steelDepth + overLap + shoeH;
                else
                    pointLength = length + pipeDiameter / 2.0 + overLap - shoeH - boundingBoxHeight;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (Configuration == 1)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.ZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.NegativeX;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.Y;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = -overHang;
                        horOffset = -steelDepth - shoeH;
                    }
                    else if (Configuration == 2)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.ZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.X;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.NegativeY;

                        vertOffset = steelWidth / 2.0;
                        routeOffset = boundingBoxWidth + overHang;
                        horOffset = -steelDepth - shoeH;
                    }
                    else if (Configuration == 3)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.ZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.NegativeX;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.Y;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = -overHang;
                        horOffset = boundingBoxHeight + shoeH;
                    }
                    else if (Configuration == 4)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.ZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.X;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.NegativeY;

                        vertOffset = steelWidth / 2.0;
                        routeOffset = boundingBoxWidth + overHang;
                        horOffset = boundingBoxHeight + shoeH;
                    }
                    else if (Configuration == 5)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.ZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.NegativeX;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.Y;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = boundingBoxWidth + overHang;
                        horOffset = -steelDepth - shoeH;
                    }
                    else if (Configuration == 6)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.ZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.X;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.NegativeY;

                        vertOffset = steelWidth / 2.0;
                        routeOffset = -overHang;
                        horOffset = -steelDepth - shoeH;
                    }
                    else if (Configuration == 7)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.ZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.NegativeX;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.Y;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = boundingBoxWidth + overHang;
                        horOffset = boundingBoxHeight + shoeH;
                    }
                    else if (Configuration == 8)
                    {
                        confRoutePlane[0] = Plane.XY;
                        confRoutePlane[1] = Plane.ZX;
                        confRouteAxis[0] = Axis.X;
                        confRouteAxis[1] = Axis.X;

                        structRoutePlane[0] = Plane.XY;
                        structRoutePlane[1] = Plane.NegativeXY;
                        structRouteAxis[0] = Axis.X;
                        structRouteAxis[1] = Axis.NegativeY;

                        vertOffset = -steelWidth / 2.0;
                        routeOffset = -overHang;
                        horOffset = boundingBoxHeight + shoeH;

                    }

                    componentDictionary[HOR_SECTION].SetPropertyValue(boundingBoxWidth + extension + steelDepth / 2.0 + overLap + overHang, "IJUAHgrOccLength", "Length");

                    if (basePlate == "With")
                    {
                        BusinessObject businessObjectPlate = componentDictionary[PLATE].GetRelationship("madeFrom", "part").TargetObjects[0];
                        plateThickness = (double)((PropertyValueDouble)businessObjectPlate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                        componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                        componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                        componentDictionary[PLATE].SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                        componentDictionary[PLATE].SetPropertyValue(bpHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                        //Add Joint Between Structure and Vertical Beam
                        JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], plateThickness, -vertOffset);

                        //Add Joint Between the Plate and the Vertical Beam
                        JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                    }
                    else
                        //Add Joint Between Structure and Vertical Beam
                        JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], 0.0, -vertOffset);

                    //Add joints between Route and Beam
                    if (Configuration == 1 || Configuration == 2 || Configuration == 3 || Configuration == 4)
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", confRoutePlane[0], confRoutePlane[1], confRouteAxis[0], confRouteAxis[1], horOffset, routeOffset);
                    else
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "EndCap", confRoutePlane[0], confRoutePlane[1], confRouteAxis[0], confRouteAxis[1], horOffset, routeOffset);

                    //Add Joint Between the Horizontal and Vertical Beams
                    if (Configuration == 1 || Configuration == 2 || Configuration == 3 || Configuration == 4)
                        JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, -overLap, -overLap - steelDepth, 0);
                    else
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, -overLap, overLap, 0);

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

                            componentDictionary[VERT_SECTION].SetPropertyValue(pointLength - plateThickness, "IJUAHgrOccLength", "Length");

                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], plateThickness, pointStructOffset, vertOffset);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                        }
                        else
                        {
                            componentDictionary[VERT_SECTION].SetPropertyValue(pointLength, "IJUAHgrOccLength", "Length");
                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], 0.0, pointStructOffset, vertOffset);
                        }

                        //Add Joint Between the Horizontal and Vertical Beams
                        if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                            JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, -overLap, -overLap - steelDepth, 0);
                        else
                            JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, -overLap, overLap, 0);

                        //Add joints between Route and Beam
                        if (Configuration == 1 || Configuration == 2 || Configuration == 5 || Configuration == 6)
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "BeginCap", pointRoutePlane[0], pointRoutePlane[1], routeOffset);
                        else
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "EndCap", pointRoutePlane[0], pointRoutePlane[1], routeOffset);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(HOR_SECTION, "BeginCap", HOR_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }

                    if (connection == "Slab")
                    {
                        componentDictionary[HOR_SECTION].SetPropertyValue(boundingBoxWidth + extension + steelDepth / 2.0 + overLap + overHang, "IJUAHgrOccLength", "Length");

                        if (Configuration == 1 || Configuration == 2)
                        {
                            vertOffset = extension;

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

                            routeOffset = boundingBoxWidth + overHang;
                            pointStructOffset = overLap + steelWidth;
                            horOffset = overLap + steelDepth;
                        }
                        if (Configuration == 3 || Configuration == 4)
                        {
                            vertOffset = -extension - steelDepth;

                            structRoutePlane[0] = Plane.XY;
                            structRoutePlane[1] = Plane.ZX;
                            structRouteAxis[0] = Axis.X;
                            structRouteAxis[1] = Axis.NegativeX;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.NegativeXY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.NegativeX;

                            routeOffset = -overHang;
                            pointStructOffset = -overLap;
                            horOffset = overLap;
                        }
                        if (Configuration == 5 || Configuration == 6)
                        {
                            vertOffset = extension + steelDepth;

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

                            routeOffset = boundingBoxWidth + overHang;
                            pointStructOffset = -overLap;
                            horOffset = overLap + steelDepth;
                        }
                        if (Configuration == 7 || Configuration == 8)
                        {
                            vertOffset = -extension;

                            structRoutePlane[0] = Plane.XY;
                            structRoutePlane[1] = Plane.ZX;
                            structRouteAxis[0] = Axis.X;
                            structRouteAxis[1] = Axis.NegativeX;

                            pointRoutePlane[0] = Plane.ZX;
                            pointRoutePlane[1] = Plane.XY;
                            pointStructPlane[0] = Plane.XY;
                            pointStructPlane[1] = Plane.NegativeXY;

                            pointStructAxis[0] = Axis.X;
                            pointStructAxis[1] = Axis.X;

                            routeOffset = -overHang;
                            pointStructOffset = -overLap;
                            horOffset = -overLap - steelDepth;
                        }

                        if (basePlate == "With")
                        {
                            BusinessObject businessObjectPlate = componentDictionary[PLATE].GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)businessObjectPlate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE].SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE].SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE].SetPropertyValue(bpHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");
                            componentDictionary[VERT_SECTION].SetPropertyValue(pointLength - plateThickness, "IJUAHgrOccLength", "Length");

                            //Add Joint Between Structure and Vertical Beam
                            if (Configuration == 1 || Configuration == 2 || Configuration == 3 || Configuration == 4)
                                JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], plateThickness, vertOffset, -steelWidth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], plateThickness, vertOffset, steelWidth / 2);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -steelDepth / 2.0, -steelWidth / 2.0);
                        }
                        else
                        {
                            componentDictionary[VERT_SECTION].SetPropertyValue(pointLength, "IJUAHgrOccLength", "Length");

                            //Add Joint Between Structure and Vertical Beam
                            if (Configuration == 1 || Configuration == 2 || Configuration == 3 || Configuration == 4)
                                JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], 0, vertOffset, -steelWidth / 2.0);
                            else
                                JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", pointStructPlane[0], pointStructPlane[1], pointStructAxis[0], pointStructAxis[1], 0, vertOffset, steelWidth / 2.0);
                        }
                        //Add Joint Between the Horizontal and Vertical Beams
                        if (Configuration == 1 || Configuration == 2 || Configuration == 3 || Configuration == 4)
                            JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "BeginCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], horOffset, pointStructOffset, 0);
                        else
                            JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", structRoutePlane[0], structRoutePlane[1], structRouteAxis[0], structRouteAxis[1], horOffset, pointStructOffset, 0);

                        //Add joints between Route and Beam
                        if (Configuration == 1 || Configuration == 2 || Configuration == 3 || Configuration == 4)
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "EndCap", pointRoutePlane[0], pointRoutePlane[1], routeOffset);
                        else
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "BeginCap", pointRoutePlane[0], pointRoutePlane[1], routeOffset);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(HOR_SECTION, "BeginCap", HOR_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
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
