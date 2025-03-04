//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_FR_LS_WS.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_FR_LS_WS
//   Author       :  Manikanth
//   Creation Date:  10.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10-04-2013     Manikanth   CR-CP-224484  Convert HS_Assembly to C# .Net 
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
    public class Assy_FR_LS_WS : CustomSupportDefinition
    {
        private const string VERT_SECTION = "VERT_SECTION";
        private const string HOR_SECTION = "HOR_SECTION";
        private const string PLATE = "PLATE";
        private const double CONST_INCH = 25.4 / 1000;

        double shoeHeight, overHang = 0, bpWidth, basePlateHoleSize, bpHoleInset, extension;
        string plate_Size, sectionSize, basePlate;

        public override Collection<PartInfo> Parts
        {
            get
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Collection<PartInfo> parts = new Collection<PartInfo>();

                shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyShoeH", "SHOE_H")).PropValue;
                sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWSize", "WSize")).PropValue;
                plate_Size = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyLS_HSS_WS", "PLATE_SIZE")).PropValue;
                overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS_HSS_WS", "OVERHANG")).PropValue;
                bpWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS_HSS_WS", "BP_WIDTH")).PropValue;
                basePlate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlates", "Plates")).PropValue;
                basePlateHoleSize = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS_HSS_WS", "BP_HOLE_SIZE")).PropValue;
                bpHoleInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyLS_HSS_WS", "BP_HOLE_INSET")).PropValue;
                extension = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyLS", "EXTENSION")).PropValue;

                parts.Add(new PartInfo(VERT_SECTION, sectionSize));
                parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                if (basePlate == "With")
                {
                    parts.Add(new PartInfo(PLATE, "Utility_FOUR_HOLE_PLATE_" + plate_Size));
                }

                return parts;  //Return the collection of Catalog Parts
            }
        }
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                //==========================
                //1. Load standard bounding box definition
                //==========================
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

                double steelWidth = crosssection.Width; //Width, Depth, Flange, or Web as 2nd argument
                double steelDepth = crosssection.Depth; //'Width, Depth, Flange, or Web as 2nd argument

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

                //Determine whether connecting to Steel or a Slab
                string connection;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    connection = "Slab";
                else
                {
                    if ((SupportHelper.SupportingObjects.Count != 0))
                    {
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                            connection = "Steel";
                        else
                            connection = "Slab";
                    }
                    else
                        connection = "Slab";                          
                }

                double verticalOffset, routeOffset, overLap, plateThickness=0, byPointStructOffset, length, byPointAngle1, byPointAngle2, horizontalOffset = 0;

                Plane confRoutePlane1=new Plane(); Plane confRoutePlane2=new Plane(); Plane structIndexPlane1=new Plane(); Plane structIndexPlane2=new Plane();
                Plane byPointStructPlane1=new Plane();Plane byPointStructPlane2=new Plane();Plane byPointRoutePlane1=new Plane();Plane byPointRoutePlane2=new Plane();
                Axis confRouteAxis1=new Axis();Axis confRouteAxis2=new Axis();Axis structIndexAxis1=new Axis();Axis structIndexAxis2 =new Axis();Axis byPointStructAxis1=new Axis();Axis byPointStructAxis2=new Axis();

               length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

               if (Ingr.SP3D.Content.Support.Symbols.HgrCompareDoubleService.cmpdbl(length, 0) == true)               
               length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);

                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter;

                overLap = 0.5 * CONST_INCH;

                byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI / 2, 2))  //The structure is oriented in the standard direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI / 2, 2))
                    {
                        verticalOffset = -steelWidth / 2;
                        routeOffset = -overHang;
                        byPointStructPlane1 = Plane.XY;
                        byPointStructPlane2 = Plane.NegativeXY;
                        byPointStructAxis1 = Axis.X;
                        byPointStructAxis2 = Axis.NegativeX;
                        byPointStructOffset = -steelDepth / 2;
                        byPointRoutePlane1 = Plane.ZX;
                        byPointRoutePlane2 = Plane.XY;
                    }
                    else                                                  //The structure is oriented in the opposite direction
                    {
                        verticalOffset = steelWidth / 2;
                        routeOffset = boundingBoxWidth + overHang;
                        byPointStructPlane1 = Plane.XY;
                        byPointStructPlane2 = Plane.NegativeXY;
                        byPointStructAxis1 = Axis.X;
                        byPointStructAxis2 = Axis.X;
                        byPointStructOffset = steelDepth / 2;
                        byPointRoutePlane1 = Plane.ZX;
                        byPointRoutePlane2 = Plane.NegativeXY;
                    }
                }
                else
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI / 2, 2))
                    {
                        verticalOffset = -steelWidth / 2;
                        routeOffset = -overHang;
                        byPointStructPlane1 = Plane.XY;
                        byPointStructPlane2 = Plane.NegativeXY;
                        byPointStructAxis1 = Axis.X;
                        byPointStructAxis2 = Axis.X;
                        byPointStructOffset = steelDepth / 2;
                        byPointRoutePlane1 = Plane.ZX;
                        byPointRoutePlane2 = Plane.XY;
                    }
                    else
                    {
                        verticalOffset = steelWidth / 2;
                        routeOffset = boundingBoxWidth + overHang;
                        byPointStructPlane1 = Plane.XY;
                        byPointStructPlane2 = Plane.NegativeXY;
                        byPointStructAxis1 = Axis.X;
                        byPointStructAxis2 = Axis.NegativeX;
                        byPointStructOffset = -steelDepth / 2;
                        byPointRoutePlane1 = Plane.ZX;
                        byPointRoutePlane2 = Plane.NegativeXY;
                    }
                }

                double byPointLength;

                if (Configuration == 1 || Configuration == 3)
                    byPointLength = length + pipeDiameter / 2 + steelDepth + overLap + shoeHeight;
                else
                    byPointLength = length + pipeDiameter / 2 + overLap - shoeHeight - boundingBoxHeight;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (Configuration == 1)
                    {
                        confRoutePlane1 = Plane.XY;
                        confRoutePlane2 = Plane.ZX;
                        confRouteAxis1 = Axis.X;
                        confRouteAxis2 = Axis.NegativeX;
                        structIndexPlane1 = Plane.XY;
                        structIndexPlane2 = Plane.NegativeXY;
                        structIndexAxis1 = Axis.X;
                        structIndexAxis2 = Axis.Y;
                        verticalOffset = -steelWidth / 2;
                        routeOffset = -overHang;
                        horizontalOffset = -steelDepth - shoeHeight;
                    }
                    else if (Configuration == 2)
                    {
                        confRoutePlane1 = Plane.XY;
                        confRoutePlane2 = Plane.ZX;
                        confRouteAxis1 = Axis.X;
                        confRouteAxis2 = Axis.X;
                        structIndexPlane1 = Plane.XY;
                        structIndexPlane2 = Plane.NegativeXY;
                        structIndexAxis1 = Axis.X;
                        structIndexAxis2 = Axis.NegativeY;
                        verticalOffset = steelWidth / 2;
                        routeOffset = boundingBoxWidth + overHang;
                        horizontalOffset = -steelDepth - shoeHeight;
                    }
                    else if (Configuration == 3)
                    {
                        confRoutePlane1 = Plane.XY;
                        confRoutePlane2 = Plane.ZX;
                        confRouteAxis1 = Axis.X;
                        confRouteAxis2 = Axis.NegativeX;
                        structIndexPlane1 = Plane.XY;
                        structIndexPlane2 = Plane.NegativeXY;
                        structIndexAxis1 = Axis.X;
                        structIndexAxis2 = Axis.Y;
                        verticalOffset = -steelWidth / 2;
                        routeOffset = -overHang;
                        horizontalOffset = boundingBoxHeight + shoeHeight;
                    }
                    else if (Configuration == 4)
                    {
                        confRoutePlane1 = Plane.XY;
                        confRoutePlane2 = Plane.ZX;
                        confRouteAxis1 = Axis.X;
                        confRouteAxis2 = Axis.X;
                        structIndexPlane1 = Plane.XY;
                        structIndexPlane2 = Plane.NegativeXY;
                        structIndexAxis1 = Axis.X;
                        structIndexAxis2 = Axis.NegativeY;
                        verticalOffset = steelWidth / 2;
                        routeOffset = boundingBoxWidth + overHang; ;
                        horizontalOffset = boundingBoxHeight + shoeHeight;
                    }

                    (componentDictionary[HOR_SECTION]).SetPropertyValue(boundingBoxWidth + extension + steelDepth / 2 + overLap + overHang, "IJUAHgrOccLength", "Length");

                    if (basePlate == "With")
                    {
                        BusinessObject plate = (componentDictionary[PLATE]).GetRelationship("madeFrom", "part").TargetObjects[0];
                        plateThickness = (double)((PropertyValueDouble)plate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                        (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                        (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                        (componentDictionary[PLATE]).SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                        (componentDictionary[PLATE]).SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                        //Add Joint Between Structure and Vertical Beam
                        JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", structIndexPlane1, structIndexPlane2, structIndexAxis1, structIndexAxis2, plateThickness, verticalOffset);
                        //Add Joint Between the Plate and the Vertical Beam
                        JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                    }
                    else
                        //Add Joint Between Structure and Vertical Beam
                        JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, "EndCap", structIndexPlane1, structIndexPlane2, structIndexAxis1, structIndexAxis2, 0, verticalOffset);
                    //Add joints between Route and Beam
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "BeginCap", confRoutePlane1, confRoutePlane2, confRouteAxis1, confRouteAxis2, horizontalOffset, routeOffset);
                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, -overLap, 0, steelWidth);
                    //Flexible Member
                    JointHelper.CreatePrismaticJoint(VERT_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                else
                {
                    if (connection == "Steel")
                    {
                        if (basePlate == "With")
                        {
                            BusinessObject plate = (componentDictionary[PLATE]).GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)plate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            (componentDictionary[PLATE]).SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            (componentDictionary[PLATE]).SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            (componentDictionary[VERT_SECTION]).SetPropertyValue(byPointLength - plateThickness, "IJUAHgrOccLength", "Length");
                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", byPointStructPlane1, byPointStructPlane2, byPointStructAxis1, byPointStructAxis2, plateThickness, byPointStructOffset, 0);
                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                        }
                        else
                        {
                            (componentDictionary[VERT_SECTION]).SetPropertyValue(byPointLength, "IJUAHgrOccLength", "Length");
                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", byPointStructPlane1, byPointStructPlane2, byPointStructAxis1, byPointStructAxis2, 0, byPointStructOffset, 0);
                        }
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, -overLap, 0, steelWidth);
                        //Add joints between Route and Beam
                        JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "BeginCap", byPointRoutePlane1, byPointRoutePlane2, routeOffset);
                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(HOR_SECTION, "BeginCap", HOR_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }

                    if (connection == "Slab")
                    {
                        (componentDictionary[HOR_SECTION]).SetPropertyValue(boundingBoxWidth + extension + steelDepth / 2 + overLap + overHang, "IJUAHgrOccLength", "Length");

                        if (Configuration == 1 || Configuration == 2)
                        {
                            verticalOffset = extension;
                            structIndexPlane1 = Plane.ZX;
                            structIndexPlane2 = Plane.NegativeXY;
                            structIndexAxis1 = Axis.X;
                            structIndexAxis2 = Axis.NegativeX;
                            byPointRoutePlane1 = Plane.ZX;
                            byPointRoutePlane2 = Plane.XY;
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            routeOffset = boundingBoxWidth + overHang;
                            horizontalOffset = overLap + steelDepth;
                            byPointStructOffset = 0;
                        }

                        else if (Configuration == 3 || Configuration == 4)
                        {
                            verticalOffset = -extension - steelDepth;
                            structIndexPlane1 = Plane.XY;
                            structIndexPlane2 = Plane.ZX;
                            structIndexAxis1 = Axis.X;
                            structIndexAxis2 = Axis.NegativeX;
                            byPointStructPlane1 = Plane.XY;
                            byPointStructPlane2 = Plane.NegativeXY;
                            byPointStructAxis1 = Axis.X;
                            byPointStructAxis2 = Axis.NegativeX;
                            byPointRoutePlane1 = Plane.ZX;
                            byPointRoutePlane2 = Plane.NegativeXY;
                            routeOffset = -overHang;
                            horizontalOffset = -steelDepth;
                            byPointStructOffset = -overLap;
                        }

                        if (basePlate == "With")
                        {
                            BusinessObject plate = (componentDictionary[PLATE]).GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)plate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            (componentDictionary[PLATE]).SetPropertyValue(bpWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            (componentDictionary[PLATE]).SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            (componentDictionary[PLATE]).SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            (componentDictionary[VERT_SECTION]).SetPropertyValue(byPointLength - plateThickness, "IJUAHgrOccLength", "Length");
                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", byPointStructPlane1, byPointStructPlane2, byPointStructAxis1, byPointStructAxis2, plateThickness, verticalOffset, 0);
                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                        }
                        else
                        {
                            (componentDictionary[VERT_SECTION]).SetPropertyValue(byPointLength, "IJUAHgrOccLength", "Length");
                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", byPointStructPlane1, byPointStructPlane2, byPointStructAxis1, byPointStructAxis2, 0, verticalOffset, 0);
                        }
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "BeginCap", structIndexPlane1, structIndexPlane2, structIndexAxis1, structIndexAxis2, horizontalOffset, byPointStructOffset, steelWidth);
                        //Add joints between Route and Beam
                        JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "EndCap", byPointRoutePlane1, byPointRoutePlane2, routeOffset);
                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(HOR_SECTION, "BeginCap", HOR_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Assy_FR_LS_WS" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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

                    int numStruct = 0;
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        numStruct = 1;
                    else
                        numStruct = SupportHelper.SupportingObjects.Count;

                    for (int index = 1; index <= numStruct; index++)
                    {
                        if (basePlate == "With")
                            structConnections.Add(new ConnectionInfo(PLATE, index));
                        else
                            structConnections.Add(new ConnectionInfo(VERT_SECTION, index));
                    }

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
                //Get IJHgrInputConfig Hlpr Interface off of passed Helper
                if (eMirrorPlane == MirrorPlane.XZPlane)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (CurrentMirrorToggleValue == 1)
                            return 2;
                        else if (CurrentMirrorToggleValue == 2)
                            return 1;
                        else if (CurrentMirrorToggleValue == 3)
                            return 4;
                        else if (CurrentMirrorToggleValue == 4)
                            return 3;
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
                    }
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