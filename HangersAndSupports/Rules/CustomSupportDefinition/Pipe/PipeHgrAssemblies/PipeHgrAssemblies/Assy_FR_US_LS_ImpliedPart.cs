//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_FR_US_LS_ImpliedPart.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_FR_US_LS_ImpliedPart
//   Author       :Vijaya
//   Creation Date:04.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04.Apr.2013     Vijaya   CR-CP-224484-Initial Creation           
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;

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
    public class Assy_FR_US_LS_ImpliedPart : CustomSupportDefinition
    {
        //Constants
        private const string HOR_SECTION = "HOR_SECTION"; 
        private const string VERT_SECTION1 = "VERT_SECTION1"; 
        private const string VERT_SECTION2 = "VERT_SECTION2"; 
        private const string PLATE1 = "PLATE1"; 
        private const string PLATE2 = "PLATE2"; 
        string[] boltPartKeys = new string[8];
       
        double shoeHeight, OverLap, basePlateWidth, basePlateHoleSize, basePlateHoleInset, span, gap;
        string sectionSize, plateSize, basePlate;
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

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyShoeH", "SHOE_H")).PropValue;
                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrLSize", "LSize")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyUS", "PLATE_SIZE")).PropValue;
                    OverLap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUS", "OVERLAP")).PropValue;
                    basePlateWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUS", "BP_WIDTH")).PropValue;
                    basePlate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlates", "Plates")).PropValue;
                    basePlateHoleSize = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUS", "BP_HOLE_SIZE")).PropValue;
                    basePlateHoleInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUS", "BP_HOLE_INSET")).PropValue;
                    span = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyUS", "SPAN")).PropValue;
                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUS", "GAP")).PropValue;

                    if (basePlate == "With")//ROD_CLEVIS_LUG
                    {
                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(VERT_SECTION1, sectionSize));
                        parts.Add(new PartInfo(VERT_SECTION2, sectionSize));
                        parts.Add(new PartInfo(PLATE1, "Utility_FOUR_HOLE_PLATE_" + plateSize));
                        parts.Add(new PartInfo(PLATE2, "Utility_FOUR_HOLE_PLATE_" + plateSize));
                    }
                    else
                    {
                        parts.Add(new PartInfo(HOR_SECTION, sectionSize));
                        parts.Add(new PartInfo(VERT_SECTION1, sectionSize));
                        parts.Add(new PartInfo(VERT_SECTION2, sectionSize));
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
        public override ReadOnlyCollection<PartInfo> ImpliedParts
        {
            get
            {
                try
                {
                    Collection<PartInfo> impliedParts = new Collection<PartInfo>();
                    //Add Bolts as Implied Parts      
                    for (int indexBolt = 1; indexBolt <= 8; indexBolt++)
                    {
                        boltPartKeys[indexBolt - 1] = "Bolt" + indexBolt.ToString();
                        impliedParts.Add(new PartInfo(boltPartKeys[indexBolt - 1], "S3Dhs_HexBolt-1/2x4"));
                    }
                    ReadOnlyCollection<PartInfo> rImpliedParts = new ReadOnlyCollection<PartInfo>(impliedParts);
                    return rImpliedParts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Implied Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                //Check the Route Vertical
                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                bool isVerticalRoute;

                if (Math.Round(Math.Abs(routeAngle), 3) < Math.Round(Math.PI / 4, 3))
                    isVerticalRoute = true;
                else if (Math.Round(Math.Abs(routeAngle), 3) > Math.Round(3 * Math.PI / 4, 3))
                    isVerticalRoute = true;
                else
                    isVerticalRoute = false;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;                
                BusinessObject horizontalSectionPart = componentDictionary[HOR_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                double steelThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                CommonAssembly commonAssembly = new CommonAssembly();

                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = commonAssembly.GetIsLugEndOffsetApplied(this);

                string[] structPort = new string[2];
                structPort = commonAssembly.GetIndexedStructPortName(this,isOffsetApplied);
                string leftStructPort = structPort[0];
                string rightStructPort = structPort[1];
                string connection;

                //Determine whether connecting to Steel or a Slab
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    connection = "Slab";
                else
                {
                    if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        connection = "Steel";
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                            connection = "Slab";          //Two Slabs

                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Member))
                            connection = "Slab-Steel";    //Slab then Steel

                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                            connection = "Steel-Slab";    //Steel then Slab
                    }
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                                connection = "Steel";    //Steel                      
                            else
                                connection = "Slab";
                        }
                        else
                            connection = "Slab";
                    }
                }
                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[HOR_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[HOR_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;

                // Set Values of Part Occurance Attributes
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[HOR_SECTION].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[HOR_SECTION].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                componentDictionary[HOR_SECTION].SetPropertyValue(0.001, "IJUAHgrOccLength", "Length");

                componentDictionary[VERT_SECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERT_SECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERT_SECTION1].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERT_SECTION1].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERT_SECTION1].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                componentDictionary[VERT_SECTION1].SetPropertyValue(0.001, "IJUAHgrOccLength", "Length");

                componentDictionary[VERT_SECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERT_SECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERT_SECTION2].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERT_SECTION2].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERT_SECTION2].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                componentDictionary[VERT_SECTION2].SetPropertyValue(0.001, "IJUAHgrOccLength", "Length");

                //Create Joints
                double length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                double horizontalLength1 = RefPortHelper.DistanceBetweenPorts("Route", rightStructPort, PortDistanceType.Horizontal);
                double horizontalLength1BBX = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal);

                double horizontalLength2 = RefPortHelper.DistanceBetweenPorts("Route", leftStructPort, PortDistanceType.Horizontal);

                double byPointAngle2 = commonAssembly.GetRouteStructConfigAngle(this, "Route", leftStructPort, PortAxisType.Y);
                double byPointAngle4 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, leftStructPort, PortAxisType.X, OrientationAlong.Direct);
                double routeOffset = 0;
                string swapPort;
                Plane routePlaneA = new Plane();
                Plane routePlaneB = new Plane();
                Axis routeAxisA = new Axis();
                Axis routeAxisB = new Axis();
                
                if (Math.Abs(byPointAngle4) > Math.PI / 2)//For Left Structure
                {
                        routeOffset = -(boundingBoxWidth / 2 - horizontalLength1 - OverLap - steelDepth / 2);
                        routePlaneA = Plane.XY;
                        routePlaneB = Plane.YZ;
                        routeAxisA = Axis.X;
                        routeAxisB = Axis.Y;
                }
                else //The structure is oriented in the opposite direction
                {
                    if (Configuration == 1)
                    {
                        if (Math.Abs(byPointAngle2) < Math.PI / 2)
                        {
                            swapPort = leftStructPort;
                            leftStructPort = rightStructPort;
                            rightStructPort = swapPort;
                            routeOffset = -boundingBoxWidth / 2 + horizontalLength2 + OverLap + steelDepth / 2;
                            routePlaneA = Plane.XY;
                            routePlaneB = Plane.YZ;
                            routeAxisA = Axis.X;
                            routeAxisB = Axis.Y;
                        }
                        else
                        {
                            routeOffset = -boundingBoxWidth / 2 + horizontalLength1 + OverLap + steelDepth / 2;
                            routePlaneA = Plane.XY;
                            routePlaneB = Plane.YZ;
                            routeAxisA = Axis.X;
                            routeAxisB = Axis.Y;
                        }
                    }
                    else if (Configuration == 2)
                    {
                        if (Math.Abs(byPointAngle2) < Math.PI / 2)
                        {
                            routeOffset = -(boundingBoxWidth / 2 + horizontalLength1 + OverLap + steelDepth / 2);
                            routePlaneA = Plane.XY;
                            routePlaneB = Plane.YZ;
                            routeAxisA = Axis.X;
                            routeAxisB = Axis.NegativeY;
                        }
                        else
                        {
                            swapPort = leftStructPort;
                            leftStructPort = rightStructPort;
                            rightStructPort = swapPort;
                            routeOffset = -(boundingBoxWidth / 2 + horizontalLength2 + OverLap + steelDepth / 2);
                            routePlaneA = Plane.XY;
                            routePlaneB = Plane.YZ;
                            routeAxisA = Axis.X;
                            routeAxisB = Axis.NegativeY;
                        }
                    }
                }

                double plateThickness = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    componentDictionary[HOR_SECTION].SetPropertyValue(span + steelDepth * 2 + OverLap * 2, "IJUAHgrOccLength", "Length");
                    if (basePlate == "With")
                    {
                        BusinessObject plate1Part = componentDictionary[PLATE1].GetRelationship("madeFrom", "part").TargetObjects[0];
                        plateThickness = (double)((PropertyValueDouble)plate1Part.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                        componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                        componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                        componentDictionary[PLATE1].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                        componentDictionary[PLATE1].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                        componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                        componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                        componentDictionary[PLATE2].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                        componentDictionary[PLATE2].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");
                        //Add Joint Between Structure and Vertical Beam
                        if (Configuration == 1)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, plateThickness, -steelWidth / 2);                  
                        else
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, plateThickness, steelWidth / 2);

                        //Add Joint Between the Plate and the Vertical Beam
                        JointHelper.CreateRigidJoint(PLATE1, "BotStructure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, steelDepth / 2, -steelWidth / 2);

                        //Add Joint Between Structure and Vertical Beam
                        JointHelper.CreatePlanarJoint("-1", "Structure", VERT_SECTION2, "BeginCap", Plane.XY, Plane.XY, plateThickness);

                        //Add Joint Between the Plate and the Vertical Beam
                        JointHelper.CreateRigidJoint(PLATE2, "BotStructure", VERT_SECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2); 
                    }
                    else
                    {
                        //Add Joint Between Structure and Vertical Beam
                        if (Configuration == 1)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, plateThickness, -steelWidth / 2);                   
                        else
                            JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, plateThickness, steelWidth / 2);

                        //Add Joint Between Structure and Vertical Beam
                        JointHelper.CreatePlanarJoint("-1", "Structure", VERT_SECTION2, "BeginCap", Plane.XY, Plane.XY, 0);   
                    }
                    // Add joints between Route and Beam
                    if (Configuration == 1)
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "EndCap", Plane.XY, Plane.ZX, Axis.X, Axis.X, boundingBoxHeight + shoeHeight, -((span + steelDepth * 2 + OverLap * 2) / 2 - boundingBoxWidth / 2));              
                    else
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "EndCap", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX, boundingBoxHeight + shoeHeight, ((span + steelDepth * 2 + OverLap * 2) / 2) + boundingBoxWidth / 2);

                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION1, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.X, gap, -OverLap, steelThickness);

                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, gap, OverLap, steelThickness);

                    //Flexible Member
                    JointHelper.CreatePrismaticJoint(VERT_SECTION1, "BeginCap", VERT_SECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Flexible Member
                    JointHelper.CreatePrismaticJoint(VERT_SECTION2, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                else
                {
                    if (connection == "Steel" && SupportHelper.SupportingObjects.Count > 1)
                    {
                        //Place by Point Geometry with Two Pieces of Steel
                        componentDictionary[VERT_SECTION1].SetPropertyValue(length, "IJUAHgrOccLength", "Length");
                        componentDictionary[VERT_SECTION2].SetPropertyValue(length, "IJUAHgrOccLength", "Length");
                        componentDictionary[HOR_SECTION].SetPropertyValue(horizontalLength1 + horizontalLength2 + OverLap * 2 + steelDepth, "IJUAHgrOccLength", "Length");

                        if (basePlate == "With")
                        {
                            BusinessObject plate1Part = componentDictionary[PLATE1].GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)plate1Part.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            //Add Joint Between Structure and Vertical Beam                          
                            JointHelper.CreatePointOnPlaneJoint(PLATE1, "TopStructure", "-1", rightStructPort, Plane.XY);

                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreatePointOnPlaneJoint(PLATE2, "TopStructure", "-1", leftStructPort, Plane.XY);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE1, "BotStructure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, steelDepth / 2, -steelWidth / 2);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE2, "BotStructure", VERT_SECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                        }
                        else
                        {
                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreatePointOnPlaneJoint(VERT_SECTION1, "EndCap", "-1", rightStructPort, Plane.XY);

                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreatePointOnPlaneJoint(VERT_SECTION2, "BeginCap", "-1", leftStructPort, Plane.XY);
                        }

                        //Add joints between Route and Beam
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HOR_SECTION, "EndCap", routePlaneA, routePlaneB, routeAxisA, routeAxisB, shoeHeight, routeOffset, -steelWidth / 2);
                        else
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HOR_SECTION, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, shoeHeight, routeOffset, steelWidth / 2);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION1, "BeginCap", Plane.ZX, Plane.YZ, Axis.X, Axis.Z, steelThickness, -OverLap - steelDepth, gap);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.YZ, Axis.X, Axis.NegativeZ, steelThickness, OverLap + steelDepth, gap);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(VERT_SECTION1, "BeginCap", VERT_SECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(VERT_SECTION2, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }

                    if (connection == "Slab" || connection == "Steel" && SupportHelper.SupportingObjects.Count < 2)
                    {
                        //Place by Point Geometry with Slab and with single piece steel
                        componentDictionary[HOR_SECTION].SetPropertyValue(span + steelDepth * 2 + OverLap * 2, "IJUAHgrOccLength", "Length");

                        String routePort;
                        if (isVerticalRoute == true)
                            routePort = "BBR_High";
                        else
                            routePort = "BBRV_High";

                        if (basePlate == "With")
                        {
                            BusinessObject plate1Part = componentDictionary[PLATE1].GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)plate1Part.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            //Add Joint Between Structure and Plate 1                            
                            JointHelper.CreatePointOnPlaneJoint(PLATE1, "TopStructure", "-1", "StructAlt", Plane.XY);

                            //Add Joint Between Structure and Plate 2
                            JointHelper.CreatePointOnPlaneJoint(PLATE2, "TopStructure", "-1", "StructAlt", Plane.XY);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE1, "BotStructure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, steelDepth / 2, -steelWidth / 2);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE2, "BotStructure", VERT_SECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                        }
                        else
                        {
                            //Add Joint Between Structure and Vertical Beam                            
                            JointHelper.CreatePointOnPlaneJoint(VERT_SECTION1, "EndCap", "-1", "StructAlt", Plane.XY);

                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreatePointOnPlaneJoint(VERT_SECTION2, "BeginCap", "-1", "StructAlt", Plane.XY);
                        }

                        //Add joints between Route and Beam
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", routePort, HOR_SECTION, "EndCap", routePlaneA, routePlaneB, routeAxisA, routeAxisB, shoeHeight, ((span + steelDepth * 2 + OverLap * 2) / 2 - boundingBoxWidth / 2), -steelWidth / 2);

                        else
                            JointHelper.CreateRigidJoint("-1", routePort, HOR_SECTION, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, shoeHeight, -((span + steelDepth * 2 + OverLap * 2) / 2 + boundingBoxWidth / 2), steelWidth / 2);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION1, "BeginCap", Plane.ZX, Plane.YZ, Axis.X, Axis.Z, steelThickness, -OverLap - steelDepth, gap);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.YZ, Axis.X, Axis.NegativeZ, steelThickness, OverLap + steelDepth, gap);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(VERT_SECTION1, "BeginCap", VERT_SECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(VERT_SECTION2, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }

                    if (connection == "Steel-Slab")
                    {
                        //Place by Point Geometry with Steel-Slab
                        componentDictionary[HOR_SECTION].SetPropertyValue(span + steelDepth * 2 + OverLap * 2, "IJUAHgrOccLength", "Length");

                        if (basePlate == "With")
                        {
                            BusinessObject plate1Part = componentDictionary[PLATE1].GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)plate1Part.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            //Steel Joint                            
                            JointHelper.CreatePointOnPlaneJoint(PLATE1, "TopStructure", "-1", rightStructPort, Plane.XY);

                            //Slab Joint
                            JointHelper.CreatePointOnPlaneJoint(PLATE2, "TopStructure", "-1", leftStructPort, Plane.XY);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE1, "BotStructure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, steelDepth / 2, -steelWidth / 2);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE2, "BotStructure", VERT_SECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                        }
                        else
                        {
                            //Add Joint Between Structure and Vertical Beam                            
                            JointHelper.CreatePointOnPlaneJoint(VERT_SECTION1, "EndCap", "-1", rightStructPort, Plane.XY);

                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreatePointOnPlaneJoint(VERT_SECTION2, "BeginCap", "-1", leftStructPort, Plane.XY);
                        }

                        //Add joints between Route and Beam
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HOR_SECTION, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.NegativeZ, shoeHeight, steelWidth / 2, horizontalLength1BBX + steelDepth / 2 + OverLap);                   
                        else if (Configuration == 2)
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HOR_SECTION, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, shoeHeight, -steelWidth / 2, horizontalLength1BBX - steelDepth - 2 * OverLap - span);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION1, "BeginCap", Plane.YZ, Plane.XY, Axis.Y, Axis.X, gap, OverLap, 0);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.YZ, Axis.X, Axis.NegativeZ, steelThickness, span + OverLap + steelDepth * 2, gap);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(VERT_SECTION1, "BeginCap", VERT_SECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(VERT_SECTION2, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }

                    if (connection == "Slab-Steel")
                    {
                        //Place by Point Geometry with Slab-Steel
                        componentDictionary[HOR_SECTION].SetPropertyValue(span + steelDepth * 2 + OverLap * 2, "IJUAHgrOccLength", "Length");

                        if (basePlate == "With")
                        {
                            BusinessObject plate1Part = componentDictionary[PLATE1].GetRelationship("madeFrom", "part").TargetObjects[0];
                            plateThickness = (double)((PropertyValueDouble)plate1Part.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                            componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE1].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                            componentDictionary[PLATE2].SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                            //Steel Joint                          
                            JointHelper.CreatePointOnPlaneJoint(PLATE1, "TopStructure", "-1", rightStructPort, Plane.XY);

                            //Slab Joint
                            JointHelper.CreatePointOnPlaneJoint(PLATE2, "TopStructure", "-1", leftStructPort, Plane.XY);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE1, "BotStructure", VERT_SECTION1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, steelDepth / 2, -steelWidth / 2);

                            //Add Joint Between the Plate and the Vertical Beam
                            JointHelper.CreateRigidJoint(PLATE2, "BotStructure", VERT_SECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                        }
                        else
                        {
                            //Add Joint Between Structure and Vertical Beam                            
                            JointHelper.CreatePointOnPlaneJoint(VERT_SECTION1, "EndCap", "-1", rightStructPort, Plane.XY);

                            //Add Joint Between Structure and Vertical Beam
                            JointHelper.CreatePointOnPlaneJoint(VERT_SECTION2, "BeginCap", "-1", leftStructPort, Plane.XY);
                        }

                        //Add joints between Route and Beam
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HOR_SECTION, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.NegativeZ, shoeHeight, steelWidth / 2, horizontalLength1BBX + steelDepth / 2 + OverLap);                   
                        else if (Configuration == 2)
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HOR_SECTION, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, shoeHeight, -steelWidth / 2, horizontalLength1BBX - steelDepth - 2 * OverLap - span);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION1, "BeginCap", Plane.YZ, Plane.XY, Axis.Y, Axis.X, gap, OverLap, 0);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.YZ, Axis.X, Axis.NegativeZ, steelThickness, span + OverLap + steelDepth * 2, gap);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(VERT_SECTION1, "BeginCap", VERT_SECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                        //Flexible Member
                        JointHelper.CreatePrismaticJoint(VERT_SECTION2, "BeginCap", VERT_SECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }//-----------------------------------------------------------------------------------
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

                    int numOfRoutes = SupportHelper.SupportedObjects.Count;
                    for (int index = 1; index <= numOfRoutes; index++)
                        routeConnections.Add(new ConnectionInfo(HOR_SECTION, index)); // partindex, routeindex

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

                    int stuctsCount = SupportHelper.SupportingObjects.Count;
                    for (int index = 1; index <= stuctsCount; index++)
                    {
                        if (basePlate == "With")
                        {
                            structConnections.Add(new ConnectionInfo(PLATE1, index)); // partindex, routeindex
                            structConnections.Add(new ConnectionInfo(PLATE2, index));
                        }
                        else
                        {
                            structConnections.Add(new ConnectionInfo(VERT_SECTION1, index));
                            structConnections.Add(new ConnectionInfo(VERT_SECTION2, index));
                        }
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
                if (eMirrorPlane != MirrorPlane.XZPlane)
                {
                    if (CurrentMirrorToggleValue == 1)
                        return 2;
                    else if (CurrentMirrorToggleValue == 2)
                        return 1;
                    else
                        return CurrentMirrorToggleValue;
                }
                else
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




