//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   G4G_1400_CTA.cs
//   CabletrayAssy,Ingr.SP3D.Content.Support.Rules.G4G_1400_CTA
//   Author       : . MK
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013       MK      CR-CP-224477 - Converted CabletrayAssemblies to C# .Net
//  22-Jan-2015     PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Support.Middle.Hidden;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class G4G_1400_CTA : CustomSupportDefinition
    {
        double tierSpacing, trayOffset, beamOffset;
        int noOfParts, gbbrzLow, gbbrzZHigh, structure, tierNo;
        string[] part = new string[10];
        int leg1StructIdx = 1;
        int leg2StructIdx = 1;
        int leg3StructIdx = 1;
        int leg4StructIdx = 1;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    SupportComponentUtils supportComponentUtils = new SupportComponentUtils();

                    tierNo = (int)((PropertyValueInt)support.GetPropertyValue("IJUAHgrCWOcc", "TierNo")).PropValue;
                    tierSpacing = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrCWOcc", "TierSpacing")).PropValue;
                    trayOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrCWOcc", "TrayOffset")).PropValue;
                    beamOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrCWOcc", "BeamOffset")).PropValue;

                    if (tierSpacing == 0)
                        tierSpacing = 0.5;

                    if (tierNo == 0)
                        tierNo = 1;
                    //retrieve attributes from cableway support object
                    noOfParts = Convert.ToInt32(4.0 + tierNo);
                    gbbrzLow = noOfParts + 1;
                    gbbrzZHigh = gbbrzLow + 1;
                    structure = gbbrzZHigh + 1;

                    string[] partClass = new string[Convert.ToInt32(noOfParts) + 4];
                    partClass[2] = partClass[3] = partClass[1] = partClass[4] = "HgrBeam";

                    for (int i = 5; i <= noOfParts; i++)
                    {
                        partClass[i] = "HgrSupFlatPlate";
                    }
                    partClass[gbbrzLow] = "Connection";
                    partClass[gbbrzZHigh] = "Connection";
                    partClass[structure] = "Connection";

                    for (int i = 1; i <= partClass.Length - 1; i++)
                    {
                        Part FlatPlate;
                        part[i] = "part" + i;

                        if (partClass[i] == "HgrSupFlatPlate")
                        {
                            FlatPlate = supportComponentUtils.GetPartFromPartClass("HgrSupFlatPlate", "HgrPartByCW", support);
                            parts.Add(new PartInfo(part[i], FlatPlate.ToString()));
                        }
                        else if (partClass[i] == "HgrBeam")
                        {
                            FlatPlate = supportComponentUtils.GetPartFromPartClass("HgrBeam", "HgrCrossSectionByCW", support);
                            parts.Add(new PartInfo(part[i], FlatPlate.ToString()));
                        }
                        else
                        {
                            FlatPlate = supportComponentUtils.GetPartFromPartClass(partClass[i], "", support);
                            parts.Add(new PartInfo(part[i], FlatPlate.ToString()));
                        }
                    }
                    return parts;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
                }
            }
        }

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                string strBOverLength = "BeginOverLength";
                string strEOverLength = "EndOverLength";
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                for (int i = 1; i <= 4; i++)
                {
                    if (i <= 2)
                        (componentDictionary[part[i]]).SetPropertyValue(beamOffset, "IJUAHgrOccOverLength", strBOverLength);
                    else
                        (componentDictionary[part[i]]).SetPropertyValue(beamOffset, "IJUAHgrOccOverLength", strEOverLength);
                }
                int routeConnectionValue = 0;
                routeConnectionValue = Configuration;

               // If the number of structures is 1, only RouteConnectionValue from 1 to 4 is applied,
                //else RouteConnectionValue from 5 to 8 is applied.

                SupportingObjectInfo oObj = null;
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    oObj = SupportingHelper.SupportingObjectInfo(1);
                }

                if (oObj != null)
                {                    

                    if ((SupportHelper.SupportingObjects.Count < 2) && Configuration > 4)
                    {
                        routeConnectionValue = Configuration - 4;
                    }
                    else if ((SupportHelper.SupportingObjects.Count >= 2) && Configuration <= 4)
                    {
                        routeConnectionValue = Configuration + 4;
                    }

                }
                else
                {
                    if (Configuration > 4)
                        routeConnectionValue = Configuration - 4;
                }

                //Set default structure ports and structure connections
                string leg1StrPort = "Structure";
                string leg2StrPort = "Structure";
                string leg3StrPort = "Structure";
                string leg4StrPort = "Structure";


                double dWidth = 0;
                double dHeight = 0;
                long bbxLow = 0;                
                string strBBXLow = null;
                string rtePortName = null;
                string[] structPort = new string[6];

                if (routeConnectionValue == 1 || routeConnectionValue == 2) // BBR bounding box for 1 structure input.
                {

                    bbxLow = -1;                    
                    strBBXLow = "BBR_Low";
                    rtePortName = strBBXLow;

                    BoundingBoxHelper.CreateStandardBoundingBoxes(true);
                    BoundingBox boundingBox;

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                    else
                        boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                    dWidth = boundingBox.Width;
                    dHeight = boundingBox.Height;
                }
                else if (routeConnectionValue == 3 || routeConnectionValue == 4)    //Global bounding box for 1 structure input
                {
                    bbxLow = gbbrzLow;                   
                    strBBXLow = "Connection";
                    rtePortName = "GBBR_Z_Low";

                    AttachConnectionsToGBBR(part[gbbrzLow], part[gbbrzZHigh], out dWidth, out dHeight);

                }
                else if (routeConnectionValue == 5 || routeConnectionValue == 6)    //Global bounding box for 2 structure input, left and right about route
                {
                    bbxLow = gbbrzLow;                
                    strBBXLow = "Connection";
                    rtePortName = "GBBR_Z_Low";

                    AttachConnectionsToGBBR(part[gbbrzLow], part[gbbrzZHigh], out dWidth, out dHeight);
                    structPort = GetStructPortNamesInOrder(rtePortName, "LeftRightStructsAlongRoute");
                    if (structPort[1] == "Structure")
                    {
                        leg1StructIdx = 1;
                        leg2StructIdx = 1;
                        leg3StructIdx = 2;
                        leg4StructIdx = 2;
                    }
                    else
                    {
                        leg1StructIdx = 2;
                        leg2StructIdx = 2;
                        leg3StructIdx = 1;
                        leg4StructIdx = 1;
                    }
                    leg1StrPort = structPort[2];
                    leg2StrPort = structPort[2];
                    leg3StrPort = structPort[1];
                    leg4StrPort = structPort[1];

                }
                else if (routeConnectionValue == 7 || routeConnectionValue == 8)    //Global bounding box for 2 structure input, rear and front about route
                {
                    bbxLow = gbbrzLow;                   
                    strBBXLow = "Connection";
                    rtePortName = "GBBR_Z_Low";
                    AttachConnectionsToGBBR(part[gbbrzLow], part[gbbrzZHigh], out dWidth, out dHeight);
                    structPort = GetStructPortNamesInOrder(rtePortName, "RearFrontStructsAlongRoute");
                    if (structPort[1] == "Structure")
                    {
                        leg1StructIdx = 2;
                        leg2StructIdx = 1;
                        leg3StructIdx = 2;
                        leg4StructIdx = 1;
                    }
                    else
                    {
                        leg1StructIdx = 1;
                        leg2StructIdx = 2;
                        leg3StructIdx = 1;
                        leg4StructIdx = 2;
                    }
                    leg1StrPort = structPort[2];
                    leg2StrPort = structPort[1];
                    leg3StrPort = structPort[2];
                    leg4StrPort = structPort[1];
                }
                string strStructure;
                if (SupportHelper.SupportingObjects.Count > 1)
                {
                    Matrix4X4 structPortOrientation = new Matrix4X4();
                    Matrix4X4 struct2PortOrientation = new Matrix4X4();
                    structPortOrientation = RefPortHelper.PortLCS("Structure");
                    struct2PortOrientation = RefPortHelper.PortLCS("Struct_2");
                    strStructure = "Connection";
                    Position StructOrigin1 = new Position(structPortOrientation.Origin.X, structPortOrientation.Origin.Y, structPortOrientation.Origin.Z);
                    Vector StructXaxis1 = new Vector(structPortOrientation.XAxis.X, structPortOrientation.XAxis.Y, structPortOrientation.XAxis.Z);
                    Vector StructZaxis1 = new Vector(structPortOrientation.ZAxis.X, structPortOrientation.ZAxis.Y, structPortOrientation.ZAxis.Z);

                    Position StructOrigin2 = new Position(struct2PortOrientation.Origin.X, struct2PortOrientation.Origin.Y, struct2PortOrientation.Origin.Z);
                    Vector StructXaxis2 = new Vector(struct2PortOrientation.XAxis.X, struct2PortOrientation.XAxis.Y, struct2PortOrientation.XAxis.Z);
                    Vector StructZaxis2 = new Vector(struct2PortOrientation.ZAxis.X, struct2PortOrientation.ZAxis.Y, struct2PortOrientation.ZAxis.Z);

                    Vector StructVec = new Vector(StructOrigin2.X - StructOrigin1.X, StructOrigin2.Y - StructOrigin1.Y, StructOrigin2.Z - StructOrigin1.Z);

                    Vector StructYaxis1 = StructZaxis1.Cross(StructXaxis1);
                    double offset = 0.5 * StructYaxis1.Dot(StructVec);

                    JointHelper.CreateRigidJoint("-1", "Structure", part[4], "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                }
                else
                {
                    strStructure = "Structure";
                    structure = -1;
                }
                //--------------------------
                //set cabeltray occurrence attributes
                //--------------------------
                //Calculate attributes for the bottom lugs based on the pipe radius
                double ctWidth = 1.1 * dWidth;
                double ctLength = 4 * dWidth;
                double thickness = 0.1; //(4 in)'provide const thickness for CT '0.5 * dHeight

                for (int i = 5; i <= noOfParts; i++)
                {
                    (componentDictionary[part[i]]).SetPropertyValue(ctWidth, "IJUAHgrOccGeometry", "Width");
                    (componentDictionary[part[i]]).SetPropertyValue(ctLength, "IJUAHgrOccLength", "Length");
                    (componentDictionary[part[i]]).SetPropertyValue(thickness, "IJUAHgrThickness", "Thickness");
                }
                 //---------------------------------------
                 // Set offset value
                 //---------------------------------------
                double dPosOffset = 0;
                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    if (routeConnectionValue == 2 || routeConnectionValue == 4 || routeConnectionValue == 6 || routeConnectionValue == 8)
                        dPosOffset = -dHeight - thickness - trayOffset;
                }
                //Create the Joint between the CableTray and the route
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    if (bbxLow == -1)
                        JointHelper.CreatePrismaticJoint("-1", strBBXLow, part[5], "Right", Plane.XY, Plane.XY, Axis.X, Axis.X, -thickness / 2 - trayOffset, -(ctWidth - dWidth) / 2.0);
                    else
                        JointHelper.CreatePrismaticJoint(part[bbxLow], strBBXLow, part[5], "Right", Plane.XY, Plane.XY, Axis.X, Axis.X, -thickness / 2 - trayOffset, -(ctWidth - dWidth) / 2.0);
                else
                    if (bbxLow == -1)
                        JointHelper.CreateRigidJoint("-1", strBBXLow, part[5], "Right", Plane.XY, Plane.XY, Axis.X, Axis.X, -thickness / 2 - trayOffset, -(ctWidth - dWidth) / 2, dPosOffset);
                    else
                        JointHelper.CreateRigidJoint(part[bbxLow], strBBXLow, part[5], "Right", Plane.XY, Plane.XY, Axis.X, Axis.X, -thickness / 2 - trayOffset, -(ctWidth - dWidth) / 2, dPosOffset);


                JointHelper.CreateRigidJoint(part[5], "Left", part[1], "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, ctLength / 2);
                JointHelper.CreateRigidJoint(part[5], "Left", part[2], "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -ctLength / 2);
                JointHelper.CreateRigidJoint(part[5], "Right", part[3], "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, ctLength / 2);
                JointHelper.CreateRigidJoint(part[5], "Right", part[4], "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, -ctLength / 2);

                //add prismatic joint to the HgrBeam itself.
                JointHelper.CreatePrismaticJoint(part[1], "BeginCap", part[1], "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                JointHelper.CreatePrismaticJoint(part[2], "BeginCap", part[2], "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                JointHelper.CreatePrismaticJoint(part[3], "EndCap", part[3], "BeginCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                JointHelper.CreatePrismaticJoint(part[4], "EndCap", part[4], "BeginCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add point on joint to reduce evaluation abortion.
                JointHelper.CreatePointOnPlaneJoint(part[1], "EndCap", "-1", leg1StrPort, Plane.XY);
                JointHelper.CreatePointOnPlaneJoint(part[2], "EndCap", "-1", leg2StrPort, Plane.XY);
                JointHelper.CreatePointOnPlaneJoint(part[3], "BeginCap", "-1", leg3StrPort, Plane.XY);
                JointHelper.CreatePointOnPlaneJoint(part[4], "BeginCap", "-1", leg4StrPort, Plane.XY);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePointOnPlaneJoint(part[5], "Center", structure.ToString(),strStructure, Plane.ZX);
                for (int i = 6; i <= noOfParts; i++)
                {
                    JointHelper.CreateRigidJoint(part[5], "Center", part[i], "Center", Plane.XY, Plane.XY, Axis.X, Axis.X, tierSpacing * (i - 5), 0, 0);
                }

            }



            catch (Exception exception)
            {
                Type myType = this.GetType();
                CmnException exception1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                throw exception1;
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

                    routeConnections.Add(new ConnectionInfo(part[1], 1)); // partindex, routeindex

                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
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

                    structConnections.Add(new ConnectionInfo(part[1], leg1StructIdx)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(part[2], leg2StructIdx)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(part[3], leg3StructIdx)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(part[4], leg4StructIdx)); // partindex, routeindex


                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
                }
            }
        }
        public void AttachConnectionsToGBBR(string lowConnectionIndex, string highConnectionIndex, out double width, out double height)
        {
            try
            {
                BoundingBoxHelper.CreateStandardBoundingBoxes(true);
                BoundingBox boundingBox;

                //==========================
                //1. Load standard bounding box definition
                //==========================
                boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.GlobalSupportedZ);

                width = boundingBox.Width;
                height = boundingBox.Height;                
                double dAngleRY_GZ = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "WORLD", PortAxisType.Z, OrientationAlong.Direct);
                double dAngleSZ_GZ = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, "WORLD", PortAxisType.Z, OrientationAlong.Direct);
                double PI2 = 0.5 * (4 * Math.Atan(1));
                if (Math.Abs(dAngleRY_GZ - PI2) < 0.00001)
                {
                    height = width;
                    width = height;
                    JointHelper.CreateRigidJoint("-1", "GBBR_Z_Low", lowConnectionIndex, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint("-1", "GBBR_Z_High", highConnectionIndex, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else
                {
                    if ((dAngleSZ_GZ > PI2 && dAngleRY_GZ >= PI2) || (dAngleSZ_GZ <= PI2 && dAngleRY_GZ < PI2))
                    {
                        JointHelper.CreateRigidJoint("-1", "GBBR_Z_Low", lowConnectionIndex, "Connection", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0.0, width, 0);
                        JointHelper.CreateRigidJoint("-1", "GBBR_Z_High", highConnectionIndex, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -width, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint("-1", "GBBR_Z_Low", lowConnectionIndex, "Connection", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, height, 0, 0);
                        JointHelper.CreateRigidJoint("-1", "GBBR_Z_High", highConnectionIndex, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, -height, 0, 0);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        public string[] GetStructPortNamesInOrder(string portName, string orderType)
        {
            if (orderType != "LeftRightStructsAlongRoute" && orderType != "RearFrontStructsAlongRoute")
            {
            }
            string[] structPorts = new string[6];
            Position pos1 =null, pos2 =null;
            Vector structVector = new Vector();
            if (SupportHelper.SupportingObjects.Count == 1)
            {
                structPorts[1] = "Structure";
                structPorts[2] = "Structure";
            }
            else
            {
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    pos1 = SupportingHelper.SupportingObjectInfo(1).GeometricCenter;
                    pos2 = SupportingHelper.SupportingObjectInfo(2).GeometricCenter;
                }

                if (pos1 != null && pos2 != null)
                    structVector = new Vector(pos2.X - pos1.X, pos2.Y - pos1.Y, pos2.Z - pos1.Z);

                Matrix4X4 routePortOrientation = new Matrix4X4();
                routePortOrientation = RefPortHelper.PortLCS(portName);
                Vector RouteXaxis = new Vector(routePortOrientation.XAxis.X, routePortOrientation.XAxis.Y, routePortOrientation.XAxis.Z);
                Vector RouteZaxis = new Vector(routePortOrientation.ZAxis.X, routePortOrientation.ZAxis.Z, routePortOrientation.ZAxis.Z);
                double dotProduct;
                if (orderType == "LeftRightStructsAlongRoute")
                {
                    Vector routeYAxis = RouteZaxis.Cross(RouteXaxis);
                    dotProduct = routeYAxis.Dot(structVector);
                }
                else
                    dotProduct = RouteXaxis.Dot(structVector);

                if (dotProduct > 0)
                {
                    structPorts[1] = "Structure";
                    structPorts[2] = "Struct_2";
                }
                else
                {
                    structPorts[1] = "Struct_2";
                    structPorts[2] = "Structure";
                }
            }
            return structPorts;
        }
    }

}