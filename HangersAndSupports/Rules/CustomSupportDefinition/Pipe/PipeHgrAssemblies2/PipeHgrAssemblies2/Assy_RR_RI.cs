//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_RI.cs
//   PipeHgrAssemblies2,Ingr.SP3D.Content.Support.Rules.Assy_RR_RI
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27-11-2015      Vinay   DI-CP-276798	Replace the use of any HS_Utility parts
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
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

    public class Assy_RR_RI : CustomSupportDefinition
    {
        private const string CLAMP = "Clamp";
        private const string ROD1 = "Rod1";
        private const string ROD2 = "Rod2";
        private const string BEAMATT1 = "BeamAtt1";
        private const string BEAMATT2 = "BeamAtt2";
        private const string EYENUT1 = "EyeNut1";
        private const string EYENUT2 = "EyeNut2";
        private const string EYENUT3 = "EyeNut3";
        private const string EYENUT4 = "EyeNut4";
        string[] ShearLug = new string[4];

        private const string STRUCTCONN = "StructConn";
        private const string ROUTECONN = "RouteConn";

        String clamp, rod, beamAtt, eyeNut, LugPart;
        Double clampAngle;
        int withShearLug;

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
                   BusinessObject supportPart = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                   Collection<PartInfo> parts = new Collection<PartInfo>();

                   clamp=(string)((PropertyValueString)support.GetPropertyValue("IJUAHgrRiserSupport", "Clamp")).PropValue;
                   rod = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrRiserSupport", "Rod")).PropValue;
                   eyeNut = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrRiserSupport", "RodFitting")).PropValue;
                   beamAtt = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrRiserSupport", "StructureAtt")).PropValue;
                   withShearLug = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrRiserSupport", "ShearLug")).PropValue;
                   LugPart = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAHSA_LugPart", "LugPart")).PropValue;

                   parts.Add(new PartInfo(CLAMP, clamp));
                   parts.Add(new PartInfo(ROD1, rod));
                   parts.Add(new PartInfo(ROD2, rod));
                   parts.Add(new PartInfo(BEAMATT1, beamAtt));
                   parts.Add(new PartInfo(BEAMATT2, beamAtt));
                   parts.Add(new PartInfo(EYENUT1, eyeNut));
                   parts.Add(new PartInfo(EYENUT2, eyeNut));
                   parts.Add(new PartInfo(EYENUT3, eyeNut));
                   parts.Add(new PartInfo(EYENUT4, eyeNut));
                   parts.Add(new PartInfo(STRUCTCONN, "Log_Conn_Part_1"));
                   parts.Add(new PartInfo(ROUTECONN, "Log_Conn_Part_1"));

                   if (withShearLug == 1)
                   {
                        for(int i = 0;i<= 3;i++)
                        {
                             ShearLug[i] = "shearLug"+ i+1;
                             parts.Add(new PartInfo(ShearLug[i], LugPart));
                        }                   
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

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent[] shearLug = new SupportComponent[4];
                BusinessObject nclampPart = componentDictionary[CLAMP].GetRelationship("madeFrom", "part").TargetObjects[0];

                int noOfStructures;
                // Get the Structure Collection
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    noOfStructures = 1;
                else
                    noOfStructures = SupportHelper.SupportingObjects.Count;

                // Get the Pipe OD without Insulation
                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeOutsideDiameter = routeInfo.OutsideDiameter; 

                // Get the Clamp Angle Attribute 
                clampAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrRiserSupport", "ClampAngle")).PropValue;

                // Get the Route Connection Toggle
                //=== ========== == ==== ===========
                // Check if the 'Structure' is Sloped - First Structure Port is not correct if structure is not sloped.
                double structZAngle;
                Boolean slopedStruct1, slopedStruct2 = false;
                structZAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, OrientationAlong.Global_Z);

                if (HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3), Math.Round(Math.PI / 2, 3)) == true || HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3), Math.Round(Math.PI, 3)) == true || HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3), 0) == true)
                    slopedStruct1 = false;
                else
                    slopedStruct1 = true;

                // Check if 'Struct_2' is sloped. 'Struct_2' port is always correct but we need to adjust the Beam Attachment Joints for sloping Structure
                if (noOfStructures > 1)
                {
                    structZAngle = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Z, OrientationAlong.Global_Z);

                    if (HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3), Math.Round(Math.PI, 3)) == true || HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3), 0) == true)
                        slopedStruct2 = false;
                    else
                        slopedStruct2 = true;
                }

                // ====== ======
                // Create Joints
                // ====== ======
                Matrix4X4 structureHangerPort = new Matrix4X4(), routeHangerPort = new Matrix4X4();

                //Get the Structure Reference Port     
                structureHangerPort = RefPortHelper.PortLCS("Structure");
                Position structurePosition = new Position();
                Vector structureX = new Vector(), structureY = new Vector(), structureZ = new Vector();

                //Get the Structure Port Orientation
                structurePosition = new Position(structureHangerPort.Origin.X, structureHangerPort.Origin.Y, structureHangerPort.Origin.Z);
                structureX = new Vector(structureHangerPort.XAxis.X, structureHangerPort.XAxis.Y, structureHangerPort.XAxis.Z);
                structureZ = new Vector(structureHangerPort.ZAxis.X, structureHangerPort.ZAxis.Y, structureHangerPort.ZAxis.Z);
                structureY = structureZ.Cross(structureX);

                //Get the Route Reference Port
                routeHangerPort = RefPortHelper.PortLCS("Route");
                Position routePosition = new Position();
                Vector routeX = new Vector(), routeY = new Vector(), routeZ = new Vector();

                //Get the Route Port Orientation
                routePosition = new Position(routeHangerPort.Origin.X, routeHangerPort.Origin.Y, routeHangerPort.Origin.Z);
                routeX = new Vector(routeHangerPort.XAxis.X, routeHangerPort.XAxis.Y, routeHangerPort.XAxis.Z);
                routeZ = new Vector(routeHangerPort.ZAxis.X, routeHangerPort.ZAxis.Y, routeHangerPort.ZAxis.Z);
                routeY = routeZ.Cross(routeX);

                //Get a Vector Along The Route's Axis
                Position routeStartPosition = SupportedHelper.SupportedObjectInfo(1).StartLocation;
                Position routeEndPosition = SupportedHelper.SupportedObjectInfo(1).EndLocation;
                Vector routeAxis = routeEndPosition.Subtract(routeStartPosition);

                //Compare the Route Axis to the Route reference ports X and Z axis to determine which is parallel.
                //Connect the Pipe Logical Connection such that its X-Axis is parallel to the pipe
                //PortAxisType parallelAxis;
                Vector projectedStructureAxis = new Vector();
                double angleToStructure;
                if (AngleBetweenVectors(routeAxis, routeX) <= (Math.Atan(1) * 4.0) / 4 || AngleBetweenVectors(routeAxis, routeX) >= 3 * (Math.Atan(1) * 4.0) / 4)
                {
                    //Route X Axis is Parallel to Route
                    //Check if the Route X Axis is pointing towards or away from the Supporting Structure
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.X) > 0)
                        //Route X Points towards structure
                        JointHelper.CreateRigidJoint(ROUTECONN, "Connection", "-1", "Route", Plane.ZX, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                    else
                    {
                        //Route X Points away from structure
                        routeX = new Vector(-routeX.X, -routeX.Y, -routeX.Z);   //(Flip Route X Vector so it points towards structure)
                        JointHelper.CreateRigidJoint(ROUTECONN, "Connection", "-1", "Route", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, 0, 0, 0);
                    }

                    //Project the Structure X Axis into the Route Y-Z Plane
                    projectedStructureAxis = ProjectVectorIntoPlane(structureX, routeX);

                    //Get the Angle From the Route Y Axis to the Projected Structure Axis
                    angleToStructure = routeY.Angle(projectedStructureAxis, routeX);
                }
                else
                {
                    //Route Z Axis is Parallel to Route
                    JointHelper.CreateRigidJoint(ROUTECONN, "Connection", "-1", "Route", Plane.ZX, Plane.ZX, Axis.X, Axis.Z, 0, 0, 0);

                    //Project the Structure X Axis into the Route X-Y Plane
                    projectedStructureAxis = ProjectVectorIntoPlane(structureX, routeZ);

                    //Get the Angle From the Route Y Axis to the Projected Structure Axis
                    angleToStructure = routeY.Angle(projectedStructureAxis, routeZ);
                }


                double clampToggle = 0, clampCorrectionAngle = 0, projection;
                Vector globalZ = new Vector(0, 0, 1);
                Vector globalX = new Vector(1, 0, 0);
                Position locRoute, locStruct;

                Matrix4X4 matrix = RefPortHelper.PortLCS("Route");
                routeX = matrix.XAxis;
                routeZ = matrix.ZAxis;
                locRoute = matrix.Origin;
                routeY = routeZ.Cross(routeX);
                matrix = RefPortHelper.PortLCS("Structure");
                structureX = matrix.XAxis;
                locStruct = matrix.Origin;

                // Project the Vectors into the Horizontal Plane
                routeY.Set(routeY.X, routeY.Y, 0);
                structureX.Set(structureX.X, structureX.Y, 0);

                // Check the location of the "Structure" port to see if the support needs to be flipped by 180 degrees
                if (noOfStructures > 1)
                {
                    clampCorrectionAngle = 0;
                    Vector routeToStructure = new Vector();
                    // Create a Vector from the Route port to the Structure Port
                    routeToStructure.Set(locStruct.X - locRoute.X, locStruct.Y - locRoute.Y, 0);
                    double tempAngle = routeY.Angle(structureX, globalZ) + Math.PI/2;

                    // Rotate RouteY by ClampCorrectionAngle
                    Vector tempV = new Vector();
                    tempV.Z = routeY.Z;
                    tempV.X = routeY.X * Math.Cos(tempAngle) - routeY.Y * Math.Sin(tempAngle);
                    tempV.Y = routeY.Y * Math.Cos(tempAngle) + routeY.X * Math.Sin(tempAngle);

                    // Project the Rotated RouteY (TempV) in the direction of the vector from Route to Structure (RouteToStruct)
                    // if this negative, then the support must be toggled by 180 degreed because the structures where selected in the opposite order
                    projection = tempV.Dot(routeToStructure) / tempV.Length;

                    if (projection >= 0)
                        clampToggle = Math.PI;
                    else
                        clampToggle = 0;
                }
                else
                {
                    clampToggle = 0;
                    clampCorrectionAngle = clampCorrectionAngle + Math.PI / 2;
                }

                // Joint to connect PipeClamp to the Pipe
                JointHelper.CreateAngularRigidJoint(CLAMP, "Route", ROUTECONN, "Connection", new Vector(0, 0, 0), new Vector(angleToStructure + clampAngle + clampToggle + clampCorrectionAngle, 0, 0));

                double offset1 = (double)((PropertyValueDouble)nclampPart.GetPropertyValue("IJUAhsOffset1", "Offset1")).PropValue;
                double offset2 = (double)((PropertyValueDouble)nclampPart.GetPropertyValue("IJUAhsOffset2", "Offset2")).PropValue;

                // Joints Connecting the Eye Nuts and Rods to the Clamp
                JointHelper.CreateRigidJoint(CLAMP, "Wing", EYENUT1, "Eye", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -offset1 - offset2);

                JointHelper.CreateRigidJoint(CLAMP, "Wing", EYENUT2, "Eye", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRigidJoint(EYENUT1, "RodEnd", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRigidJoint(EYENUT2, "RodEnd", ROD2, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                //Prismatic Joints on the Rods to Allow them to be Variable length

                JointHelper.CreatePrismaticJoint(ROD1, "RodEnd2", ROD1, "RodEnd1", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                JointHelper.CreatePrismaticJoint(ROD2, "RodEnd2", ROD2, "RodEnd1", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                // Joints Connecting the Eye Nuts to the Top of the Rods
                // If the Structure is sloped we must allow the Eye to rotate around the rod so that the Beam Attachment can slope with the structure.
                Plane planeA, planeB;
                Axis axisA, axisB;

                if (Configuration == 1)
                {
                    planeA = Plane.XY;
                    planeB = Plane.XY;
                    axisA = Axis.X;
                    axisB = Axis.X;
                }
                else
                {
                    planeA = Plane.XY;
                    planeB = Plane.XY;
                    axisA = Axis.X;
                    axisB = Axis.Y;

                }

                if (slopedStruct1)
                    JointHelper.CreateRevoluteJoint(EYENUT3, "RodEnd", ROD1, "RodEnd2", Axis.Z, Axis.Z);
                else
                    JointHelper.CreateRigidJoint(EYENUT3, "RodEnd", ROD1, "RodEnd2", planeA, planeB, axisA, axisB, 0, 0, 0);

                if (noOfStructures > 1)
                {
                    if (slopedStruct2)
                        JointHelper.CreateRevoluteJoint(EYENUT4, "RodEnd", ROD2, "RodEnd2", Axis.Z, Axis.Z);
                    else
                        JointHelper.CreateRigidJoint(EYENUT4, "RodEnd", ROD2, "RodEnd2", planeA, planeB, axisA, axisB, 0, 0, 0);
                }
                else
                {
                    if (slopedStruct1)
                        JointHelper.CreateRevoluteJoint(EYENUT4, "RodEnd", ROD2, "RodEnd2", Axis.Z, Axis.Z);
                    else
                        JointHelper.CreateRigidJoint(EYENUT4, "RodEnd", ROD2, "RodEnd2", planeA, planeB, axisA, axisB, 0, 0, 0);
                }

                // Joints Connecting the Beam Attachments to the Eye Nuts

                JointHelper.CreateRevoluteJoint(EYENUT3, "Eye", BEAMATT1, "Pin", Axis.X, Axis.X);

                JointHelper.CreateRevoluteJoint(EYENUT4, "Eye", BEAMATT2, "Pin", Axis.X, Axis.X);

                // Joints Connecting the Beam Attachments to the Supporting Structures
                // The 'Structure' Reference Port Z-Axis is not Perpindicular to the Steel Face if the Steel is not sloped. Use a Vertical Joint to
                // Correct this. If the Steel is sloped then we can attach the Beam Attachment directly to the "Structure" Port.
                // "Struct_2" Port appears to be fine, and does not need this connection object
                if (slopedStruct1 == false)
                {
                    JointHelper.CreateSphericalJoint(STRUCTCONN, "Connection", "-1", "Structure");

                    JointHelper.CreateGlobalAxesAlignedJoint(STRUCTCONN, "Connection", Axis.Z, Axis.Z);
                }
                else
                    JointHelper.CreateAngularRigidJoint(STRUCTCONN, "Connection", "-1", "Structure", new Vector(0, 0, 0), new Vector(Math.PI, 0, 0));

                // Now Add the Beam Attachment to the Logical Connection Part

                JointHelper.CreatePlanarJoint(BEAMATT1, "Structure", STRUCTCONN, "Connection", Plane.XY, Plane.NegativeXY, 0);

                if (noOfStructures > 1)
                    JointHelper.CreatePlanarJoint(BEAMATT2, "Structure", "-1", "Struct_2", Plane.XY, Plane.XY, 0);
                else
                    JointHelper.CreatePlanarJoint(BEAMATT2, "Structure", STRUCTCONN, "Connection", Plane.NegativeXY, Plane.XY, 0);

                // Joints For the Shear Lugs
                 if (withShearLug == 1)
                {
                    BusinessObject[] sherlugPart = new BusinessObject[4];
                    for (int index = 0; index <= 3; index++)
                    {
                        shearLug[index] = componentDictionary[ShearLug[index]];
                        sherlugPart[index] = shearLug[index].GetRelationship("madeFrom", "part").TargetObjects[0];
                    }

                    for (int index = 0; index <= 3; index++)
                    {
                        shearLug[index].SetPropertyValue(pipeOutsideDiameter, "IJUAhsRoutePort", "RPDiameter");
                    }
                    Double lugAngle, clampWidth = 0, lugLength = 0, offsetX = 0;

                    clampWidth = (double)((PropertyValueDouble)nclampPart.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue / 2;
                    lugLength = (double)((PropertyValueDouble)sherlugPart[0].GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                    offsetX = clampWidth + lugLength;

                    for (int index = 0; index <= 3; index++)
                    {
                        lugAngle = lugAngle = index * (Math.PI / 2) + (Math.PI / 4);
                        JointHelper.CreateAngularRigidJoint(ShearLug[index], "Route", CLAMP, "Route", new Vector(offsetX, 0, 0), new Vector(Math.PI + lugAngle, 0, 0));
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
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(CLAMP, 1)); // partindex, routeindex

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
                    //structConnections.Add(new ConnectionInfo(BEAMATT1, 1)); // partindex, routeindex

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
        private double AngleBetweenVectors(Vector vector1, Vector vector2)
        {
            return Math.Acos(vector1.Dot(vector2) / ((vector1.Length * vector2.Length)));
        }

        private Vector ProjectVectorIntoPlane(Vector vector, Vector planeNormal)
        {
            Vector Normal = new Vector();

            Normal = new Vector(planeNormal.X, planeNormal.Y, planeNormal.Z);
            Normal.Length = vector.Dot(Normal);

            return vector.Subtract(Normal);
        }
    }
}
