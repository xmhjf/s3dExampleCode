﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_FR_UC_CS.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_FR_UC_CS
//   Author       :  Manikanth
//   Creation Date:  12.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12-04-2013     Manikanth   CR-CP-224484  Convert HS_Assembly to C# .Net  
//   07-01-2014     Ramya   DM-CP-246092	 [TR] Failure of the test 1.6.1.1 of QTP HangersDevTestSet
//   06-Jan-2015 Chethan    TR-CP-262663  Certain attributes doesn’t get modified for the support assembly “Assy_FR_UC_CS”
//   22-02-2015       PVK   TR-CP-264951  Resolve coverity issues found in November 2014 report
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

    public class Assy_FR_UC_CS : CustomSupportDefinition
    {
        private const string TOP_SECTION = "TOP_SECTION";
        private const string HORCONNECTION = "HORCON";
        private const string TOP_PLATE = "TOP_PLATE";

        double shoeHeight, overHang, basePlateWidth, basePlateHoleSize, bpHoleInset;
        string plateSize, basePlate, sectionSize;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyShoeH", "SHOE_H")).PropValue;
                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrCSize", "CSize")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyUC", "PLATE_SIZE")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUC", "OVERHANG")).PropValue;
                    basePlateWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUC", "BP_WIDTH")).PropValue;
                    basePlate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlates", "Plates")).PropValue;
                    basePlateHoleSize = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUC", "BP_HOLE_SIZE")).PropValue;
                    bpHoleInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyUC", "BP_HOLE_INSET")).PropValue;

                    parts.Add(new PartInfo(TOP_SECTION, sectionSize));
                    parts.Add(new PartInfo(HORCONNECTION, "Log_Conn_Part_1"));
                    if (basePlate == "With")
                    {
                         parts.Add(new PartInfo(TOP_PLATE, "Utility_FOUR_HOLE_PLATE_" + plateSize));
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

                BusinessObject horizontalSectionPart = (componentDictionary[HORCONNECTION]).GetRelationship("madeFrom", "part").TargetObjects[0];
                BusinessObject topSectionPart = (componentDictionary[TOP_SECTION]).GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)topSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = crosssection.Width;
                double steelDepth = crosssection.Depth;

                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================

                PropertyValueCodelist topbeginMiterCodelist = (PropertyValueCodelist)(componentDictionary[TOP_SECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (topbeginMiterCodelist.PropValue == -1)
                    topbeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist topendMiterCodelist = (PropertyValueCodelist)(componentDictionary[TOP_SECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (topendMiterCodelist.PropValue == -1)
                    topendMiterCodelist.PropValue = 1;

                //====== ======
                //Set Values of Part Occurance Attributes
                //====== ======

                (componentDictionary[TOP_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[TOP_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[TOP_SECTION]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[TOP_SECTION]).SetPropertyValue(topbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[TOP_SECTION]).SetPropertyValue(topendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                double length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                if (Ingr.SP3D.Content.Support.Symbols.HgrCompareDoubleService.cmpdbl(length, 0) == true)
                    length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter ;
                double angle, angle1, angle2, structOffset = 0, offset = 0, steelOffset = 0, offset1=0;

                angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);
                angle1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                angle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Z, OrientationAlong.Global_Z);

                Plane routePlane1 = new Plane(); Plane routePlane2 = new Plane();
                Axis routeAxis1 = new Axis(); Axis routeAxis2 = new Axis();

                if (Math.Abs(angle) < Math.PI / 2)
                {
                    if (Configuration == 1)
                    {
                        routePlane1 = Plane.XY;
                        routePlane2 = Plane.NegativeXY;
                        routeAxis1 = Axis.X;
                        routeAxis2 = Axis.X;
                        structOffset = steelWidth / 2;
                        offset = -shoeHeight;
                        offset1 = shoeHeight;
                    }
                    else if (Configuration == 2)
                    {
                        routePlane1 = Plane.XY;
                        routePlane2 = Plane.NegativeXY;
                        routeAxis1 = Axis.X;
                        routeAxis2 = Axis.X;
                        offset = shoeHeight + pipeDiameter + steelDepth;
						offset1 = -shoeHeight;
                    }
                    else if (Configuration == 3)
                    {
                        routePlane1 = Plane.XY;
                        routePlane2 = Plane.NegativeXY;
                        routeAxis1 = Axis.X;
                        routeAxis2 = Axis.NegativeX;
                        offset = boundingBoxWidth - shoeHeight - pipeDiameter - steelDepth;
                        structOffset = -steelWidth / 2;
                        offset1 = shoeHeight;
                    }
                    else if (Configuration == 4)
                    {
                        routePlane1 = Plane.XY;
                        routePlane2 = Plane.NegativeXY;
                        routeAxis1 = Axis.X;
                        routeAxis2 = Axis.NegativeX;
                        offset = boundingBoxWidth + shoeHeight;
                        structOffset = -steelWidth / 2;
                        offset1 = -shoeHeight;                        
                    }
                }
                else
                {
                    if (Configuration == 1)
                    {
                        routePlane1 = Plane.XY;
                        routePlane2 = Plane.NegativeXY;
                        routeAxis1 = Axis.X;
                        routeAxis2 = Axis.NegativeX;
                        offset = boundingBoxWidth + shoeHeight;
                        structOffset = -steelWidth / 2;
                        offset1 = shoeHeight;
                    }
                    else if (Configuration == 2)
                    {
                        routePlane1 = Plane.XY;
                        routePlane2 = Plane.NegativeXY;
                        routeAxis1 = Axis.X;
                        routeAxis2 = Axis.NegativeX;
                        offset = boundingBoxWidth - shoeHeight - pipeDiameter - steelDepth;
                        structOffset = -steelWidth / 2;
                        offset1 = -shoeHeight;
                    }
                    else if (Configuration == 3)
                    {
                        routePlane1 = Plane.XY;
                        routePlane2 = Plane.NegativeXY;
                        routeAxis1 = Axis.X;
                        routeAxis2 = Axis.X;
                        offset = boundingBoxWidth + shoeHeight + steelDepth;
                        structOffset = steelWidth / 2;
                        offset1 = shoeHeight;
                    }
                    else if (Configuration == 4)
                    {
                        routePlane1 = Plane.XY;
                        routePlane2 = Plane.NegativeXY;
                        routeAxis1 = Axis.X;
                        routeAxis2 = Axis.X;
                        offset = -shoeHeight;
                        structOffset = steelWidth / 2;
                        offset1 = -shoeHeight;
                    }
                }

                if ((Math.Abs(angle) < Math.PI / 2) ||  HgrCompareDoubleService.cmpdbl(Math.Round(angle, 3) , Math.Round(Math.PI / 2, 3))==true)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (Configuration == 1)
                            steelOffset = 0;
                        else if (Configuration == 2)
                            steelOffset = -pipeDiameter - steelDepth;
                    }
                    else
                    {
                        if (Configuration == 1)
                        {
                            routePlane1 = Plane.XY;
                            routePlane2 = Plane.XY;
                            routeAxis1 = Axis.X;
                            routeAxis2 = Axis.NegativeX;
                            steelOffset = 0;
                            structOffset = steelWidth / 2;
                        }
                        else if (Configuration == 2)
                        {
                            routePlane1 = Plane.XY;
                            routePlane2 = Plane.XY;
                            routeAxis1 = Axis.X;
                            routeAxis2 = Axis.NegativeX;
                            steelOffset = -pipeDiameter - steelDepth;
                            structOffset = steelWidth / 2;
                        }
                        else if (Configuration == 3)
                        {
                            routePlane1 = Plane.XY;
                            routePlane2 = Plane.XY;
                            routeAxis1 = Axis.X;
                            routeAxis2 = Axis.X;
                            steelOffset = -steelDepth;
                            structOffset = -steelWidth / 2;
                        }
                        else if (Configuration == 4)
                        {
                            routePlane1 = Plane.XY;
                            routePlane2 = Plane.XY;
                            routeAxis1 = Axis.X;
                            routeAxis2 = Axis.X;
                            steelOffset = -pipeDiameter - 2 * steelDepth;
                            structOffset = -steelWidth / 2;
                        }
                    }
                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (Configuration == 1)
                            steelOffset = 0;
                        else if (Configuration == 2)
                            steelOffset = pipeDiameter + steelDepth;
                    }
                    else
                    {
                        if (Configuration == 1)
                        {
                            routePlane1 = Plane.XY;
                            routePlane2 = Plane.XY;
                            routeAxis1 = Axis.X;
                            routeAxis2 = Axis.NegativeX;
                            steelOffset = 0;
                            structOffset = steelWidth / 2;
                        }
                        else if (Configuration == 2)
                        {
                            routePlane1 = Plane.XY;
                            routePlane2 = Plane.XY;
                            routeAxis1 = Axis.X;
                            routeAxis2 = Axis.NegativeX;
                            if (SupportHelper.SupportingObjects.Count != 0)
                            {
                                if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                                    steelOffset = -pipeDiameter - steelDepth;
                                else
                                    steelOffset = pipeDiameter + steelDepth;
                            }
                            else
                                steelOffset = pipeDiameter + steelDepth;
                            structOffset = steelWidth / 2;
                        }
                        else if (Configuration == 3)
                        {
                            routePlane1 = Plane.XY;
                            routePlane2 = Plane.XY;
                            routeAxis1 = Axis.X;
                            routeAxis2 = Axis.X;
                            steelOffset = -steelDepth;
                            structOffset = -steelWidth / 2;
                        }
                        else if (Configuration == 4)
                        {
                            routePlane1 = Plane.XY;
                            routePlane2 = Plane.XY;
                            routeAxis1 = Axis.X;
                            routeAxis2 = Axis.X;
                            if (SupportHelper.SupportingObjects.Count != 0)
                            {
                                if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                                    steelOffset = -pipeDiameter - 2 * steelDepth;
                                else
                                    steelOffset = pipeDiameter;
                            }
                            else
                                steelOffset = pipeDiameter;
                            structOffset = -steelWidth / 2;
                        }
                    }
                }

                double beamLength;

                if (basePlate == "With")
                {
                    BusinessObject plate = (componentDictionary[TOP_PLATE]).GetRelationship("madeFrom", "part").TargetObjects[0];
                    double plate_T = (double)((PropertyValueDouble)plate.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")).PropValue;

                    (componentDictionary[TOP_PLATE]).SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH");
                    (componentDictionary[TOP_PLATE]).SetPropertyValue(basePlateWidth, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH");
                    (componentDictionary[TOP_PLATE]).SetPropertyValue(bpHoleInset, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C");
                    (componentDictionary[TOP_PLATE]).SetPropertyValue(basePlateHoleSize, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE");

                    //Add Connection for the end of the Angled Beam
                    JointHelper.CreateRigidJoint(TOP_SECTION, "BeginCap", TOP_PLATE, "TopStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, steelDepth / 2, steelWidth / 2);
                    //Add Connection for the end of the Angled Beam
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, plate_T, -structOffset);
                    else
                    {
                        beamLength = length + overHang + pipeDiameter / 2.0 - plate_T;

                        (componentDictionary[TOP_SECTION]).SetPropertyValue(beamLength, "IJUAHgrOccLength", "Length");

                        if (SupportHelper.SupportingObjects.Count != 0)
                        {
                            if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                                JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, plate_T, steelDepth + pipeDiameter / 2, offset1);
                            else
                            {
                                if (HgrCompareDoubleService.cmpdbl(Math.Round(angle2, 3) , Math.Round(Math.PI, 3))==true)
                                    JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, plate_T, -pipeDiameter / 2-offset1, 0);
                                else
                                    JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, plate_T, pipeDiameter/2 + steelDepth+offset1, 0);
                            }
                        }
                        else
                        {
                            if (HgrCompareDoubleService.cmpdbl(Math.Round(angle2, 3), Math.Round(Math.PI, 3)) == true)
                                JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, plate_T, -pipeDiameter / 2-offset1, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, plate_T, pipeDiameter / 2 + steelDepth+offset1, 0);
                        }
                    }
                }
                else
                {
                    //Add Connection for the end of the Angled Beam
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -structOffset);
                    else
                    {
                        beamLength = length + overHang + pipeDiameter / 2.0 + overHang;
                        (componentDictionary[TOP_SECTION]).SetPropertyValue(beamLength, "IJUAHgrOccLength", "Length");
                        if (SupportHelper.SupportingObjects.Count != 0)
                        {
                            if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                                JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0.0, steelDepth + pipeDiameter / 2 , offset);
                            else
                            {
                                if (HgrCompareDoubleService.cmpdbl(Math.Round(angle2, 3), Math.Round(Math.PI, 3)) == true)
                                    JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -pipeDiameter / 2-offset1, 0);
                                else
                                    JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, pipeDiameter / 2 + steelDepth + offset1, 0);
                            }
                        }
                        else
                        {
                            if (HgrCompareDoubleService.cmpdbl(Math.Round(angle2, 3), Math.Round(Math.PI, 3)) == true)
                                JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -pipeDiameter / 2-offset1, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "Structure", HORCONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, pipeDiameter / 2 + steelDepth + offset1, 0);
                        }
                    }
                }

                //Add a Spherical Joint between Top Beam and Horizontal Connection
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreateSphericalJoint(HORCONNECTION, "Connection", TOP_SECTION, "BeginCap");
                else
                    JointHelper.CreateRigidJoint(HORCONNECTION, "Connection", TOP_SECTION, "BeginCap", routePlane1, routePlane2, routeAxis1, routeAxis2, 0, steelOffset, structOffset);
                //Add joints between Route and Beam
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", TOP_SECTION, "EndCap", routePlane1, routePlane2, routeAxis1, routeAxis2, -overHang, offset);
                    //Flexible Member
                    JointHelper.CreatePrismaticJoint(TOP_SECTION, "BeginCap", TOP_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Assy_FR_UC_CS." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                            structConnections.Add(new ConnectionInfo(TOP_PLATE, index));
                        else
                            structConnections.Add(new ConnectionInfo(TOP_SECTION, index));
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
                if (eMirrorPlane == MirrorPlane.YZPlane)
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
                else if (eMirrorPlane == MirrorPlane.XYPlane)
                {
                    if ((SupportHelper.PlacementType == PlacementType.PlaceByStruct))
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
                    else
                    {
                        if (CurrentMirrorToggleValue == 1)
                            return 1;
                        else if (CurrentMirrorToggleValue == 2)
                            return 2;
                        else if (CurrentMirrorToggleValue == 3)
                            return 3;
                        else if (CurrentMirrorToggleValue == 4)
                            return 4;
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