//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   StrutRiser.cs
//   Strut_Assy,Ingr.SP3D.Content.Support.Rules.StrutRiser
//   Author       :  Vijay
//   Creation Date:  12/07/2013   
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12/07/2013     Vijay    CR-CP-224475 Convert HS_S3DStrut_Assy to C# .Net 
//   04/11/2014     PVK      CR-CP-245790 Modify the exsisting .Net Strut_Assy to new URS Strut supports
//   10/05/2016     PVK      TR-CP-294449	‘Shear Lug Side’ property for strut riser assembly doesn’t work properly
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

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

    public class StrutRiser : CustomSupportDefinition
    {
        //Part Index's
        private const string STRUT_1A = "STRUT_1A";
        private const string STRUT_1B = "STRUT_1B";
        private const string STRUT_2A = "STRUT_2A";
        private const string STRUT_2B = "STRUT_2B";
        private const string PIPE_ATT = "PIPE_ATT";
        private const string STRUCT_ATT_1 = "STRUCT_ATT_1";
        private const string STRUCT_ATT_2 = "STRUCT_ATT_2";


        //Logical Connections
        private const string LOG_CONN = "LOG_CONN";
        private const string STRUCT_CONN = "STRUCT_CONN";
        private const string STRUCT_CONN_2 = "STRUCT_CONN_2";

        //Part Classes
        private string strutA;
        private string strutB;
        private string pipeAtt;
        private string structAtt;
        private string shearLug;
        private int shearLugQuantity;
        private int shearLugSide;
        static int shearLugCount = 1;
        string[] shearLugParts = new string[shearLugCount];

        //Attributes
        private double riserAngle;
        private double strut1Angle;
        private int strut1EndOrientation;
        private double strut2Angle;
        private int strut2EndOrientation;

        private double offset1;
        private double offset2;

        // Collections for the Weld Data and Weld Part Index's
        Collection<StrutAssemblyServices.WeldData> weldCollection = new Collection<StrutAssemblyServices.WeldData>();

        //Booleans
        private Boolean twoStructures;
        private string struct1APartClass, struct1BPartClass, pipeAtt1PartClass;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    if (support.SupportsInterface("IJUAHgrURSCommon"))
                    {
                        string family = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrURSCommon", "Family")).PropValue;
                        string type = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrURSCommon", "Type")).PropValue;
                        StrutAssemblyServices.CheckSupportWithFamilyAndType(this, family, type);
                    }

                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                    //Get the Part Classes / Part Numbers
                    if (part.SupportsInterface("IJUAhsStrutA"))
                        strutA = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsStrutA", "StrutA")).PropValue;
                    else
                        strutA = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsStrutRiser", "StrutA")).PropValue;

                    if (part.SupportsInterface("IJUAhsStrutB"))
                        strutB = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsStrutB", "StrutB")).PropValue;
                    else
                        strutB = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsStrutRiser", "StrutB")).PropValue;

                    pipeAtt = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsStrutPipeAtt", "PipeAttachment")).PropValue;
                    structAtt = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsStrutStructAtt", "StructAttachment")).PropValue;

                    if (part.SupportsInterface("IJUAhsStrutRiserShearLug"))
                        shearLug = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsStrutRiserShearLug", "ShearLug")).PropValue;
                    else
                        shearLug = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsStrutRiserLugs", "ShearLug")).PropValue;
                    

                    //Get the Shear Lug Quantity and Total Count
                    if (part.SupportsInterface("IJUAhsStrutRiserShearLugQty"))
                        shearLugQuantity = (int)((PropertyValueInt)support.GetPropertyValue("IJUAhsStrutRiserShearLugQty", "ShearLugQuantity")).PropValue;                  
                    else
                        shearLugQuantity = (int)((PropertyValueInt)support.GetPropertyValue("IJUAhsStrutRiserLugs", "ShearLugQuantity")).PropValue;
                    if (part.SupportsInterface("IJUAhsStrutRiserShearLugSide"))
                        shearLugSide = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsStrutRiserShearLugSide", "ShearLugSide")).PropValue;
                    else
                        shearLugSide = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsStrutRiserLugs", "ShearLugSide")).PropValue;

                    if (shearLugSide == 1 || shearLugSide == 2)
                        shearLugCount = shearLugQuantity;
                    else
                        shearLugCount = shearLugQuantity * 2;

                    //Joints Connecting the Riser Clamp to the Logical Connection
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "The support can only be placed by point.", "", "StrutRiser.cs", 105);
                        return null;
                    }


                    //Get PartClass
                    GetPartClassValue(strutA, ref struct1APartClass);
                    GetPartClassValue(strutB, ref struct1BPartClass);
                    GetPartClassValue(pipeAtt, ref pipeAtt1PartClass);

                    //Add the Parts
                    parts.Add(new PartInfo(STRUT_1A, strutA, struct1APartClass));
                    parts.Add(new PartInfo(STRUT_1B, strutB, struct1BPartClass));
                    parts.Add(new PartInfo(STRUT_2A, strutA, struct1APartClass));
                    parts.Add(new PartInfo(STRUT_2B, strutB, struct1BPartClass));
                    parts.Add(new PartInfo(PIPE_ATT, pipeAtt, pipeAtt1PartClass));
                    parts.Add(new PartInfo(STRUCT_ATT_1, structAtt));
                    parts.Add(new PartInfo(STRUCT_ATT_2, structAtt));

                    Array.Resize(ref shearLugParts, shearLugCount);
                    if (shearLugCount > 0 && shearLug != "")
                    {
                        for (int nIndex = 0; nIndex < shearLugCount; nIndex++)
                        {
                            shearLugParts[nIndex] = "SHEAR_LUG_" + nIndex + 1;
                            parts.Add(new PartInfo(shearLugParts[nIndex], shearLug));
                        }
                    }

                    //Logical Connections
                    parts.Add(new PartInfo(LOG_CONN, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(STRUCT_CONN, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(STRUCT_CONN_2, "Log_Conn_Part_1"));

                    // Add the Weld Objects from the Weld Sheet
                    //weldCollection = StrutAssemblyServices.AddStrutWeldsFromCatalog(this, parts, "IJUAhsRStrutWelds");

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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject supportPart = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                //Get the Support attributes
                if (supportPart.SupportsInterface("IJUAhsStrutRiserClampAngle"))
                    riserAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutRiserClampAngle", "RiserClampAngle")).PropValue;
                else
                    riserAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutRiserAngle", "RiserClampAngle")).PropValue;
                if (supportPart.SupportsInterface("IJUAhsStrutRiserAngle1"))
                    strut1Angle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutRiserAngle1", "Strut1Angle")).PropValue;
                else
                    strut1Angle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutRiserAngle", "Strut1Angle")).PropValue;
                if (supportPart.SupportsInterface("IJUAhsStrutRiserAngle2"))
                    strut2Angle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutRiserAngle2", "Strut2Angle")).PropValue;
                else
                    strut2Angle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutRiserAngle", "Strut2Angle")).PropValue;
                if (supportPart.SupportsInterface("IJUAhsStrut1RiserEndOrient"))
                    strut1EndOrientation = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsStrut1RiserEndOrient", "Strut1EndOrientation")).PropValue;
                else
                    strut1EndOrientation = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsStrutRiserEnd", "Strut1EndOrientation")).PropValue;
                if (supportPart.SupportsInterface("IJUAhsStrut2RiserEndOrient"))
                    strut2EndOrientation = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsStrut2RiserEndOrient", "Strut2EndOrientation")).PropValue;
                else
                    strut2EndOrientation = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsStrutRiserEnd", "Strut2EndOrientation")).PropValue;

                if (supportPart.SupportsInterface("IJUAhsStrutOffset1"))
                    offset1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutOffset1", "Offset1")).PropValue;
                else
                    offset1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutOffsets", "Offset1")).PropValue;
                if (supportPart.SupportsInterface("IJUAhsStrutOffset2"))
                    offset2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutOffset2", "Offset2")).PropValue;
                else
                    offset2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsStrutOffsets", "Offset2")).PropValue;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = pipeInfo.OutsideDiameter;

                if (SupportHelper.SupportingObjects.Count != 0)
                {
                    if (SupportHelper.SupportingObjects.Count == 2)
                        twoStructures = true;
                    else
                        twoStructures = false;
                }
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
                PortAxisType parallelAxis;
                Vector projectedStructureAxis = new Vector();
                double angleToStructure;
                if (AngleBetweenVectors(routeAxis, routeX) <= (Math.Atan(1) * 4.0) / 4 || AngleBetweenVectors(routeAxis, routeX) >= 3 * (Math.Atan(1) * 4.0) / 4)
                {
                    //Route X Axis is Parallel to Route
                    parallelAxis = PortAxisType.X;

                    //Check if the Route X Axis is pointing towards or away from the Supporting Structure
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.X) > 0)
                        //Route X Points towards structure
                        JointHelper.CreateRigidJoint(LOG_CONN, "Connection", "-1", "Route", Plane.ZX, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                    else
                    {
                        //Route X Points away from structure
                        routeX = new Vector(-routeX.X, -routeX.Y, -routeX.Z);   //(Flip Route X Vector so it points towards structure)
                        JointHelper.CreateRigidJoint(LOG_CONN, "Connection", "-1", "Route", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, 0, 0, 0);
                    }

                    //Project the Structure X Axis into the Route Y-Z Plane
                    projectedStructureAxis = ProjectVectorIntoPlane(structureX, routeX);

                    //Get the Angle From the Route Y Axis to the Projected Structure Axis
                    angleToStructure = routeY.Angle(projectedStructureAxis, routeX);
                }
                else
                {
                    //Route Z Axis is Parallel to Route
                    parallelAxis = PortAxisType.Z;

                    JointHelper.CreateRigidJoint(LOG_CONN, "Connection", "-1", "Route", Plane.ZX, Plane.ZX, Axis.X, Axis.Z, 0, 0, 0);

                    //Project the Structure X Axis into the Route X-Y Plane
                    projectedStructureAxis = ProjectVectorIntoPlane(structureX, routeZ);

                    //Get the Angle From the Route Y Axis to the Projected Structure Axis
                    angleToStructure = routeY.Angle(projectedStructureAxis, routeZ);
                }

                //Determine the Orientation of the First Structure Port
                if (parallelAxis == PortAxisType.X)
                {
                    if (AngleBetweenVectors(routeX, structureZ) <= (Math.Atan(1) * 4.0) / 4 || AngleBetweenVectors(routeX, structureZ) >= 3 * (Math.Atan(1) * 4.0) / 4)
                        //Structure Z Parallel to Pipe
                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                    {
                        //Connect the Structure Logical Connection to the Structure (Structure Y Axis is Normal)
                        if (RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortAxisType.Y) > 0)
                            //Structure Y Axis Points towards the route
                            JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                        else
                            //Structure Y axis Points away from the route
                            JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, 0, 0, 0);
                    }

                }
                else
                {
                    if (AngleBetweenVectors(routeZ, structureZ) <= (Math.Atan(1) * 4.0) / 4 || AngleBetweenVectors(routeZ, structureZ) >= 3 * (Math.Atan(1) * 4.0) / 4)
                        //Structure Z Parallel to Pipe
                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                    {
                        //Connect the Structure Logical Connection to the Structure (Structure Y Axis is Normal)
                        if (RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortAxisType.Y) > 0)
                            //Structure Y Axis Points towards the route
                            JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                        else
                            //Structure Y axis Points away from the route
                            JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, 0, 0, 0);
                    }
                }

                //Determine the Orientation of the Second Structure Port
                if (twoStructures)
                {
                    Vector struct2X = new Vector(), struct2Z = new Vector();

                    //Get the 2nd Structure Reference Port
                    Matrix4X4 structure2HangerPort = new Matrix4X4();

                    //Get the Structure Port Orientation
                    struct2X = new Vector(structure2HangerPort.XAxis.X, structure2HangerPort.XAxis.Y, structure2HangerPort.XAxis.Z);
                    struct2Z = new Vector(structure2HangerPort.ZAxis.X, structure2HangerPort.ZAxis.Y, structure2HangerPort.ZAxis.Z);

                    if (parallelAxis == PortAxisType.X)
                    {
                        if (AngleBetweenVectors(routeX, struct2Z) <= (Math.Atan(1) * 4.0) / 4 || AngleBetweenVectors(routeX, struct2Z) >= 3 * (Math.Atan(1) * 4.0) / 4)
                            //Struct 2 Z Pointing Towards Route
                            JointHelper.CreateRigidJoint("-1", "Struct_2", STRUCT_CONN_2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                        {
                            if (RefPortHelper.DistanceBetweenPorts("Struct_2", "Route", PortAxisType.Y) > 0)
                                //Struct 2 Y Pointing Towards Route
                                JointHelper.CreateRigidJoint(STRUCT_CONN_2, "Connection", "-1", "Struct_2", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                            else
                                //Struct 2 Y Pointing Away From Route
                                JointHelper.CreateRigidJoint(STRUCT_CONN_2, "Connection", "-1", "Struct_2", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, 0, 0, 0);
                        }
                    }
                    else
                    {
                        if (AngleBetweenVectors(routeZ, struct2Z) <= (Math.Atan(1) * 4.0) / 4 || AngleBetweenVectors(routeZ, struct2Z) >= 3 * (Math.Atan(1) * 4.0) / 4)
                            //Struct 2 Z Pointing Towards Route
                            JointHelper.CreateRigidJoint("-1", "Struct_2", STRUCT_CONN_2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                        {
                            if (RefPortHelper.DistanceBetweenPorts("Struct_2", "Route", PortAxisType.Y) > 0)
                                //Struct 2 Y Pointing Towards Route
                                JointHelper.CreateRigidJoint(STRUCT_CONN_2, "Connection", "-1", "Struct_2", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                            else
                                //Struct 2 Y Pointing Away From Route
                                JointHelper.CreateRigidJoint(STRUCT_CONN_2, "Connection", "-1", "Struct_2", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, 0, 0, 0);
                        }
                    }
                }
                else
                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", STRUCT_CONN_2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                if (twoStructures == false)
                    angleToStructure = angleToStructure + Math.Atan(1) * 4.0 / 2;

                //Toggle will Mirror Support (Rotate 180 Degrees)
                if (Configuration == 2)
                    angleToStructure = angleToStructure + Math.Atan(1) * 4.0;

                //Set the C-C Span of the Riser Clamp
                double clampHeight1 = (double)((PropertyValueDouble)componentDictionary[PIPE_ATT].GetPropertyValue("IJOAhsHeight1", "Height1")).PropValue;
                double clampHeight2 = (double)((PropertyValueDouble)componentDictionary[PIPE_ATT].GetPropertyValue("IJOAhsHeight2", "Height2")).PropValue;
                double clampOffset1 = (double)((PropertyValueDouble)componentDictionary[PIPE_ATT].GetPropertyValue("IJOAhsOffset1", "Offset1")).PropValue;
                double clampOffset2 = (double)((PropertyValueDouble)componentDictionary[PIPE_ATT].GetPropertyValue("IJOAhsOffset2", "Offset2")).PropValue;

                double pinOffset1 = clampHeight1 - clampOffset1;
                double pinOffset2 = clampHeight2 - clampOffset2;

                if (offset1 > 0)  //use the keyed in value for the pipe clamp cc
                {
                    componentDictionary[PIPE_ATT].SetPropertyValue(offset1 + pinOffset1, "IJOAhsHeight1", "Height1");
                    componentDictionary[PIPE_ATT].SetPropertyValue(offset1, "IJOAhsOffset1", "Offset1");
                }
                else    //use the catalog values
                {
                    componentDictionary[PIPE_ATT].SetPropertyValue(clampHeight1, "IJOAhsHeight1", "Height1");
                    componentDictionary[PIPE_ATT].SetPropertyValue(clampOffset1, "IJOAhsOffset1", "Offset1");
                }

                if (offset2 > 0)  //use the keyed in value for the pipe clamp cc
                {
                    componentDictionary[PIPE_ATT].SetPropertyValue(offset2 + pinOffset2, "IJOAhsHeight2", "Height2");
                    componentDictionary[PIPE_ATT].SetPropertyValue(offset2, "IJOAhsOffset2", "Offset2");
                }
                else    //use the catalog values
                {
                    componentDictionary[PIPE_ATT].SetPropertyValue(clampHeight2, "IJOAhsHeight2", "Height2");
                    componentDictionary[PIPE_ATT].SetPropertyValue(clampOffset2, "IJOAhsOffset2", "Offset2");
                }

                //Set the Attributes on the Shear Lugs
                double lugThick=0;
                double lugHeight=0;
                double clampWidth=0;
                double lugWidth=0;
                if (shearLugCount > 0 && shearLug != "")
                {
                    BusinessObject part = componentDictionary[PIPE_ATT].GetRelationship("madeFrom", "part").TargetObjects[0];
                    BusinessObject lugPart = componentDictionary[shearLugParts[0]].GetRelationship("madeFrom", "part").TargetObjects[0];
                    clampWidth = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                    lugWidth = (double)((PropertyValueDouble)lugPart.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                    lugThick = (double)((PropertyValueDouble)lugPart.GetPropertyValue("IJUAhsThickness1", "Thickness1")).PropValue;
                    lugHeight = (double)((PropertyValueDouble)lugPart.GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;

                    //Get the Angle Between the Lugs
                    double lugAngle = (2 * Math.Atan(1) * 4.0) / shearLugQuantity;

                    //Get the Offset Angle (If there are 4, 8, etc... lugs)
                    double lugAngleDelta;
                    if (shearLugQuantity % 4 == 0)
                        lugAngleDelta = lugAngle / 2;
                    else
                        lugAngleDelta = 0;
                    switch (shearLugSide)
                    {
                        case 1:
                            //Lugs on Top Side of Clamp
                            for (int nIndex = 0; nIndex < shearLugQuantity; nIndex++)
                            {
                                if (componentDictionary[shearLugParts[nIndex]].SupportsInterface("IJUAhsRoutePort"))
                                {
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(pipeDiameter, "IJUAhsRoutePort", "RPDiameter");
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(nIndex * lugAngle + lugAngleDelta, "IJUAhsRoutePort", "RPRotationAroundPipe");
                                }
                                if (componentDictionary[shearLugParts[nIndex]].SupportsInterface("IJOAhsRoutePort"))
                                {
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(pipeDiameter, "IJOAhsRoutePort", "RPDiameter");
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(nIndex * lugAngle + lugAngleDelta, "IJOAhsRoutePort", "RPRotationAroundPipe");
                                }
                            }
                            break;
                        case 2:
                            //Lugs on Bottom Side of Clamp
                            for (int nIndex = 0; nIndex < shearLugQuantity; nIndex++)
                            {
                                if (componentDictionary[shearLugParts[nIndex]].SupportsInterface("IJUAhsRoutePort"))
                                {
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(pipeDiameter, "IJUAhsRoutePort", "RPDiameter");
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(nIndex * lugAngle + lugAngleDelta, "IJUAhsRoutePort", "RPRotationAroundPipe");
                                }
                                if (componentDictionary[shearLugParts[nIndex]].SupportsInterface("IJOAhsRoutePort"))
                                {
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(pipeDiameter, "IJOAhsRoutePort", "RPDiameter");
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(nIndex * lugAngle + lugAngleDelta, "IJOAhsRoutePort", "RPRotationAroundPipe");
                                }
                            }
                            break;
                        case 3:
                            //Lugs on Both Sides of Clamp
                            for (int nIndex = 0; nIndex < shearLugQuantity; nIndex++)
                            {
                                if (componentDictionary[shearLugParts[nIndex]].SupportsInterface("IJUAhsRoutePort"))
                                {
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(pipeDiameter, "IJUAhsRoutePort", "RPDiameter");
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(nIndex * lugAngle + lugAngleDelta, "IJUAhsRoutePort", "RPRotationAroundPipe");
                                }
                                if (componentDictionary[shearLugParts[nIndex]].SupportsInterface("IJOAhsRoutePort"))
                                {
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(pipeDiameter, "IJOAhsRoutePort", "RPDiameter");
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(nIndex * lugAngle + lugAngleDelta, "IJOAhsRoutePort", "RPRotationAroundPipe");
                                }
                            }
                            for (int nIndex = shearLugQuantity; nIndex < shearLugCount; nIndex++)
                            {
                                if (componentDictionary[shearLugParts[nIndex]].SupportsInterface("IJUAhsRoutePort"))
                                {
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(pipeDiameter, "IJUAhsRoutePort", "RPDiameter");
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue((nIndex - shearLugQuantity) * lugAngle + lugAngleDelta, "IJUAhsRoutePort", "RPRotationAroundPipe");
                                }
                                if (componentDictionary[shearLugParts[nIndex]].SupportsInterface("IJOAhsRoutePort"))
                                {
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue(pipeDiameter, "IJOAhsRoutePort", "RPDiameter");
                                    componentDictionary[shearLugParts[nIndex]].SetPropertyValue((nIndex - shearLugQuantity) * lugAngle + lugAngleDelta, "IJOAhsRoutePort", "RPRotationAroundPipe");
                                }
                            }
                            break;
                    }
                }

                //Joints Connecting the Riser Clamp to the Logical Connection
                JointHelper.CreateAngularRigidJoint(PIPE_ATT, "Route", LOG_CONN, "Connection", new Vector(0, 0, 0), new Vector(angleToStructure + riserAngle, 0, 0));

                //Joints Connecting the Strut B Ends to the Pipe Clamp
                JointHelper.CreateAngularRigidJoint(STRUT_1B, "Port1", PIPE_ATT, "Wing", new Vector(0, 0, 0), new Vector(0, strut1Angle, 0));
                JointHelper.CreateAngularRigidJoint(STRUT_2B, "Port1", PIPE_ATT, "Side", new Vector(0, 0, 0), new Vector(0, strut2Angle, 0));

                //Joints Connecting the Strut B Ends to the Strut A Parts
                if (strut1EndOrientation == 1)
                    JointHelper.CreateRevoluteJoint(STRUT_1A, "Port2", STRUT_1B, "Port2", Axis.Z, Axis.Z);
                else if (strut1EndOrientation == 2)
                    JointHelper.CreateRigidJoint(STRUT_1B, "Port2", STRUT_1A, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    JointHelper.CreateRigidJoint(STRUT_1B, "Port2", STRUT_1A, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                if (strut2EndOrientation == 1)
                    JointHelper.CreateRevoluteJoint(STRUT_2A, "Port2", STRUT_2B, "Port2", Axis.Z, Axis.Z);
                else if (strut2EndOrientation == 2)
                    JointHelper.CreateRigidJoint(STRUT_2B, "Port2", STRUT_2A, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    JointHelper.CreateRigidJoint(STRUT_2B, "Port2", STRUT_2A, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                //Prismatic Joints for the Variable Length Struts
                JointHelper.CreatePrismaticJoint(STRUT_1A, "Port1", STRUT_1A, "Port2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                JointHelper.CreatePrismaticJoint(STRUT_2A, "Port1", STRUT_2A, "Port2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                //Joints Connecting the Structural Attachments to the Strut A Parts
                JointHelper.CreateRevoluteJoint(STRUCT_ATT_1, "Pin", STRUT_1A, "Port1", Axis.Y, Axis.Y);
                JointHelper.CreateRevoluteJoint(STRUCT_ATT_2, "Pin", STRUT_2A, "Port1", Axis.Y, Axis.Y);

                //Joints Connectiong the Structural Attachments to the Structure
                JointHelper.CreatePlanarJoint(STRUCT_ATT_1, "Structure", STRUCT_CONN, "Connection", Plane.XY, Plane.XY, 0);
                JointHelper.CreatePlanarJoint(STRUCT_ATT_2, "Structure", STRUCT_CONN_2, "Connection", Plane.XY, Plane.XY, 0);

                //Joints Connecting the Shear Lugs to the Riser Clamp
                if (shearLugCount > 0 && shearLug != "")
                {
                    for (int nIndex = 0; nIndex < shearLugCount; nIndex++)
                    {
                        if (shearLugSide == 1)
                            JointHelper.CreateRigidJoint(PIPE_ATT, "Route", shearLugParts[nIndex], "Route", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, clampWidth / 2 + lugWidth / 2);
                        else if (shearLugSide == 2)
                            JointHelper.CreateRigidJoint(PIPE_ATT, "Route", shearLugParts[nIndex], "Route", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, -clampWidth / 2 - lugWidth / 2);
                        else
                            if (nIndex < shearLugQuantity)
                                JointHelper.CreateRigidJoint(PIPE_ATT, "Route", shearLugParts[nIndex], "Route", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, clampWidth / 2 + lugWidth / 2);
                            else
                                JointHelper.CreateRigidJoint(PIPE_ATT, "Route", shearLugParts[nIndex], "Route", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, -clampWidth / 2 - lugWidth / 2);
                    }
                }

                //Joints For Weld Objects
                StrutAssemblyServices.WeldData weld = new StrutAssemblyServices.WeldData();
                double width2=0;
                if (weldCollection!=null)
                {
                    for (int weldCount = 0; weldCount < weldCollection.Count; weldCount++)
                    {
                        weld = weldCollection[weldCount];
                        switch(weld.connection)
                        {
                            case "A":
                                width2=(double)((PropertyValueDouble)componentDictionary[STRUCT_ATT_1].GetPropertyValue("IJUAhsWidth2", "Width2")).PropValue;
                                JointHelper.CreateRigidJoint(STRUCT_ATT_1, "Structure", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, width2/2);
                                
                                break;
                            case "B":
                                if (shearLugCount>0 && shearLug!="")
                                JointHelper.CreateRigidJoint( shearLugParts[0], "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0,-lugThick/2, 0);
                                 break;

                            case "C":
                                if (shearLugCount>0 && shearLug!="")
                                JointHelper.CreateRigidJoint( shearLugParts[0], "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -lugThick/2, 0);
                                 break;

                            case "D":
                                if (shearLugCount>0 && shearLug!="")
                                JointHelper.CreateRigidJoint( shearLugParts[0], "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, -lugThick/2, -lugHeight/2, -lugWidth/2);
                                 break;
                        }
                    }
                }


                //Add Dimension Control Points
                Note note;
                ControlPoint controlPoint;
                note = CreateNote("Pipe CL", PIPE_ATT, "Route", new Position(0, 0, 0), "", true, 2, 1, out controlPoint);
                note = CreateNote("Pin 1", PIPE_ATT, "Wing", new Position(0, 0, 0), "", true, 2, 1, out controlPoint);
                note = CreateNote("Pin 2", PIPE_ATT, "Side", new Position(0, 0, 0), "", true, 2, 1, out controlPoint);
                note = CreateNote("Structure 1", STRUCT_ATT_1, "Structure", new Position(0, 0, 0), "", true, 2, 1, out controlPoint);
                note = CreateNote("Structure 2", STRUCT_ATT_2, "Structure", new Position(0, 0, 0), "", true, 2, 1, out controlPoint);
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

                    routeConnections.Add(new ConnectionInfo(PIPE_ATT, 1));       //partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(STRUCT_ATT_1, 1));      //partindex, routeindex
                    if (twoStructures == true)
                        structConnections.Add(new ConnectionInfo(STRUCT_ATT_2, 2));      //partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(STRUCT_ATT_2, 1));      //partindex, routeindex
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
        public override void OrientLocalCoordinateSystem()
        {
            Matrix4X4 routeIJHangerPort = new Matrix4X4(), structureIJHangerPort = new Matrix4X4();
            //Get the Route Reference Port     
            routeIJHangerPort = RefPortHelper.PortLCS("Route");
            //Get the Structure Reference Port
            structureIJHangerPort = RefPortHelper.PortLCS("Structure");
            bool IsRouteVertical = false;
            Vector RouteVector = new Vector();
            if (SupportedHelper.SupportedObjectInfo(1).StartLocation.Z < SupportedHelper.SupportedObjectInfo(1).EndLocation.Z)
                RouteVector = new Vector(SupportedHelper.SupportedObjectInfo(1).EndLocation.X - SupportedHelper.SupportedObjectInfo(1).StartLocation.X, SupportedHelper.SupportedObjectInfo(1).EndLocation.Y - SupportedHelper.SupportedObjectInfo(1).StartLocation.Y, SupportedHelper.SupportedObjectInfo(1).EndLocation.Z - SupportedHelper.SupportedObjectInfo(1).StartLocation.Z);
            else
                RouteVector = new Vector(SupportedHelper.SupportedObjectInfo(1).StartLocation.X - SupportedHelper.SupportedObjectInfo(1).EndLocation.X, SupportedHelper.SupportedObjectInfo(1).StartLocation.Y - SupportedHelper.SupportedObjectInfo(1).EndLocation.Y, SupportedHelper.SupportedObjectInfo(1).StartLocation.Z - SupportedHelper.SupportedObjectInfo(1).EndLocation.Z);

            RouteVector.Length = 1.0;
            if (Math.Acos(new Vector(0, 0, 1).Dot(RouteVector)) > (5 * Math.Atan(1) * 4.0) / 180)
                IsRouteVertical = false;
            else
                IsRouteVertical = true;

            if (IsRouteVertical == true)
            {
                if (routeIJHangerPort.Origin.Z > structureIJHangerPort.Origin.Z)
                    JointHelper.CreateRigidJoint(PIPE_ATT, "Route", "-1", "LocalCS", Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.Y, 0, 0, 0);
                else
                    JointHelper.CreateRigidJoint(PIPE_ATT, "Route", "-1", "LocalCS", Plane.YZ, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);
            }
            else
                JointHelper.CreateRigidJoint(PIPE_ATT, "Route", "-1", "LocalCS", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
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
        /// <summary>
        /// This method will Check either its Part or PartClass and return PSL(PartSelectionRule) if it is PartClass or return empty string for Part.
        /// </summary>
        /// <param name="partOrPartClassName">Name of the PartClass</param>
        /// <param name="partSelectionRule">Return the PartSelectionRule</param>
        /// <returns></returns>
        /// <code>
        /// GetPartClassValue(partClassName, ref partClassValue)
        /// </code>
        public void GetPartClassValue(string partOrPartClassName, ref string partSelectionRule)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            BusinessObject partclass = catalogBaseHelper.GetPartClass(partOrPartClassName);
            if (partclass is PartClass)
            {
                partSelectionRule = partclass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
            }
            else
            {
                partSelectionRule = "";
            }
        }
    }
}
