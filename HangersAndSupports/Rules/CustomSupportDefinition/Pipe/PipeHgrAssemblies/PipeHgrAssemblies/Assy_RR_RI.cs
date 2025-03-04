//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_RI.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_RI
//   Author       : Rajeswari
//   Creation Date: 17-April-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 17-April-2013  Rajeswari CR-CP-224484 C#.Net HS_Assembly Project Creation  
//   22-02-2015       PVK   TR-CP-264951  Resolve coverity issues found in November 2014 report
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

        String clamp, rod, beamAtt, eyeNut;
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
                   Collection<PartInfo> parts = new Collection<PartInfo>();

                   clamp=(string)((PropertyValueString)support.GetPropertyValue("IJUAHgrRiserSupport", "Clamp")).PropValue;
                   rod = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrRiserSupport", "Rod")).PropValue;
                   eyeNut = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrRiserSupport", "RodFitting")).PropValue;
                   beamAtt = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrRiserSupport", "StructureAtt")).PropValue;
                   withShearLug = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrRiserSupport", "ShearLug")).PropValue;

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
                             parts.Add(new PartInfo(ShearLug[i], "Utility_USER_FIXED_BOX_1"));
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
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========
                // Shear Lug Dimensions
                Double width = 0, length = 0, thickness = 0;

                if (withShearLug == 1)
                {
                    // Get the Shear Lug Dimensions from the Support
                    width = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrRiserSupport", "ShearLugW")).PropValue;
                    length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrRiserSupport", "ShearLugL")).PropValue;
                    thickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrRiserSupport", "ShearLugT")).PropValue;

                    // Set the Dimensions on the Shear Lug Parts
                    const double CONST_1 = 39.3700787;
                    string bom = "Shear Lug - " + Math.Round(length * CONST_1, 3) + " x " + Math.Round(width * CONST_1, 3) + " x " + Math.Round(thickness * CONST_1, 3) + " in ";
                    BusinessObject[] sherlugPart = new BusinessObject[4];

                    for (int index = 0; index <= 3; index++)
                    {
                        shearLug[index] = componentDictionary[ShearLug[index]];
                        sherlugPart[index] = shearLug[index].GetRelationship("madeFrom", "part").TargetObjects[0];
                    }

                    for (int index = 0; index <= 3; index++)
                    {
                        shearLug[index].SetPropertyValue(length, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                        shearLug[index].SetPropertyValue(width, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                        shearLug[index].SetPropertyValue(thickness, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                        shearLug[index].SetPropertyValue(bom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");
                    }
                }

                //=== ========== == ==== ===========
                // Get the Angle in the Horizontal Plane from the Route Y Axis to the Structure X Axis. This will be used to
                // Rotate the RiserClamp about the Pipe to Align it with the Structure. If it is Single Structure then the
                // clamp should be aligned with the structure. If its Two Structures, then the clamp should be crossing the structures.
                //=== ========== == ==== ===========
                Vector routeX, routeY, routeZ, structureX, globalZ, globalX;
                Position locRoute, locStruct;
                   
                Matrix4X4 matrix = RefPortHelper.PortLCS("Route");
                routeX = matrix.XAxis;
                routeZ = matrix.ZAxis;
                locRoute =matrix.Origin;
                routeY = routeZ.Cross(routeX);
                matrix = RefPortHelper.PortLCS("Structure");
                structureX = matrix.XAxis;
                locStruct = matrix.Origin;

                // Project the Vectors into the Horizontal Plane
                routeY.Set(routeY.X, routeY.Y, 0);
                structureX.Set(structureX.X, structureX.Y, 0);
                globalZ = new Vector(0, 0, 1);
                globalX = new Vector(1, 0, 0);

                // Get the Angle From Route Y to Structure X
                Double clampCorrectionAngle, clampToggle, projection, structZAngle;

                clampCorrectionAngle = routeY.Angle(structureX, globalZ);
                if (noOfStructures > 1)
                    clampCorrectionAngle = clampCorrectionAngle + Math.PI / 2;

                // Check the location of the "Structure" port to see if the support needs to be flipped by 180 degrees
                // Depends on order of selection for two structures and on the ClampCorrectionAngle
                if (noOfStructures > 1)
                {
                    Vector routeToStructure = new Vector();
                    // Create a Vector from the Route port to the Structure Port
                    routeToStructure.Set(locStruct.X - locRoute.X, locStruct.Y - locRoute.Y, 0);

                    // Rotate RouteY by ClampCorrectionAngle
                    Vector tempV = new Vector();
                    tempV.Z = routeY.Z;
                    tempV.X = routeY.X * Math.Cos(clampCorrectionAngle) - routeY.Y * Math.Sin(clampCorrectionAngle);
                    tempV.Y = routeY.Y * Math.Cos(clampCorrectionAngle) + routeY.X * Math.Sin(clampCorrectionAngle);

                    // Project the Rotated RouteY (TempV) in the direction of the vector from Route to Structure (RouteToStruct)
                    // if this negative, then the support must be toggled by 180 degreed because the structures where selected in the opposite order
                    projection = tempV.Dot(routeToStructure) / tempV.Length;

                    if (projection >= 0)
                        clampToggle = 0;
                    else
                        clampToggle = Math.PI;
                    }
                else
                    clampToggle = 0;

                // Check if the 'Structure' is Sloped - First Structure Port is not correct if structure is not sloped.
                Boolean slopedStruct1, slopedStruct2=false;
                structZAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, OrientationAlong.Global_Z);

                if (HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3) , Math.Round(Math.PI / 2, 3))==true || HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3) , Math.Round(Math.PI, 3))==true || HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3) , 0)==true)
                    slopedStruct1 = false;
                else
                    slopedStruct1 = true;

                // Check if 'Struct_2' is sloped. 'Struct_2' port is always correct but we need to adjust the Beam Attachment Joints for sloping Structure
                if (noOfStructures > 1)
                {
                    structZAngle = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Z, OrientationAlong.Global_Z);

                    if (HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3) , Math.Round(Math.PI, 3))==true || HgrCompareDoubleService.cmpdbl(Math.Round(structZAngle, 3) , 0)==true)
                        slopedStruct2 = false;
                    else
                        slopedStruct2 = true;
                }

                // ====== ======
                // Create Joints
                // ====== ======

                // Joints Connecting the Riser Clamp to the Pipe
                // ============
                // This is an example of how the MakeAngularRigidJoint can be used to rotate a port with respect to another
                // port.  Note: The angle must be in radians.
                // ============

                JointHelper.CreateAngularRigidJoint(CLAMP, "Route", "-1", "Route", new Vector(0, 0, 0), new Vector(0, routeZ.Angle(globalZ, routeY), clampAngle + clampCorrectionAngle + clampToggle));

                // Joints Connecting the Eye Nuts and Rods to the Clamp
                JointHelper.CreateRigidJoint(EYENUT1, "Eye", CLAMP, "LeftPin", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRigidJoint(EYENUT2, "Eye", CLAMP, "RightPin", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", EYENUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", EYENUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                // Prismatic Joints on the Rods to Allow them to be Variable length

                JointHelper.CreatePrismaticJoint(ROD1, "BotExThdRH", ROD1, "TopExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                JointHelper.CreatePrismaticJoint(ROD2, "BotExThdRH", ROD2, "TopExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                // Joints Connecting the Eye Nuts to the Top of the Rods
                // If the Structure is sloped we must allow the Eye to rotate around the rod so that the Beam Attachment can slope with the structure.
                Plane planeA, planeB;
                Axis axisA, axisB;

                if (Configuration == 1)
                {
                    planeA=Plane.XY;
                    planeB=Plane.XY;
                    axisA=Axis.X;
                    axisB=Axis.X;
                }
                else
                {
                    planeA=Plane.XY;
                    planeB=Plane.XY;
                    axisA=Axis.X;
                    axisB=Axis.Y;

                }

                if (slopedStruct1)
                    JointHelper.CreateRevoluteJoint(EYENUT3, "InThdRH", ROD1, "TopExThdRH", Axis.Z, Axis.Z);
                else
                    JointHelper.CreateRigidJoint(EYENUT3, "InThdRH", ROD1, "TopExThdRH", planeA, planeB, axisA, axisB, 0, 0, 0);

                if (noOfStructures > 1)
                {
                    if (slopedStruct2)
                        JointHelper.CreateRevoluteJoint(EYENUT4, "InThdRH", ROD2, "TopExThdRH", Axis.Z, Axis.Z);
                    else
                        JointHelper.CreateRigidJoint(EYENUT4, "InThdRH", ROD2, "TopExThdRH", planeA, planeB, axisA, axisB, 0, 0, 0);
                }
                else
                {
                    if (slopedStruct1)
                        JointHelper.CreateRevoluteJoint(EYENUT4, "InThdRH", ROD2, "TopExThdRH", Axis.Z, Axis.Z);
                    else
                        JointHelper.CreateRigidJoint(EYENUT4, "InThdRH", ROD2, "TopExThdRH", planeA, planeB, axisA, axisB, 0, 0, 0);
                    }

                // Joints Connecting the Beam Attachments to the Eye Nuts

                JointHelper.CreateRevoluteJoint(EYENUT3, "Eye", BEAMATT1, "Pin", Axis.X, Axis.Y);

                JointHelper.CreateRevoluteJoint(EYENUT4, "Eye", BEAMATT2, "Pin", Axis.X, Axis.Y);

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

                JointHelper.CreatePlanarJoint(BEAMATT1, "Structure", STRUCTCONN, "Connection", Plane.XY, Plane.XY,0);

                if (noOfStructures > 1)
                    JointHelper.CreatePlanarJoint(BEAMATT2, "Structure", "-1", "Struct_2", Plane.XY, Plane.NegativeXY, 0);
                else
                    JointHelper.CreatePlanarJoint(BEAMATT2, "Structure", STRUCTCONN, "Connection", Plane.XY, Plane.XY, 0);

                // Joints For the Shear Lugs
                if (withShearLug == 1)
                {
                    Double lugAngle, offsetX, offsetY, offsetZ;

                    offsetZ = (double)((PropertyValueDouble)nclampPart.GetPropertyValue("IJUAHgrAnvil_FIG40", "G2")).PropValue/2;
                    for (int index = 0; index <= 3; index++)
                    {
                        lugAngle = index * (Math.PI / 2) + (Math.PI / 4);
                        offsetX = (pipeOutsideDiameter / 2 + width / 2) * Math.Sin(lugAngle);
                        offsetY = (pipeOutsideDiameter / 2 + width / 2) * Math.Cos(lugAngle);
                        JointHelper.CreateAngularRigidJoint(ShearLug[index], "StartOther", CLAMP, "Route", new Vector(offsetX, offsetY, offsetZ), new Vector(0, 0, -lugAngle));
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
                    structConnections.Add(new ConnectionInfo(BEAMATT1, 1)); // partindex, routeindex

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
    }
}
