//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_BBX.cs
//   OglaendAssemblies,Ingr.SP3D.Content.Support.Rules.Cantilever
//   Author       :  Durga Prasad
//   Creation Date:  13-12-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
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
    public class Cantilever : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        #region Global Constants, Fields and properties
        //Part Keys
        private
        const string SECTION = "SECTION";
        const string BASE_ANGLE = "BASE_ANGLE";
        const string STRUCT_CONN = "STRUCT_CONN";

        //Steel Members - Stores Data related to steel member (Part Number, Depth, Width etc...)
        OglaendAssemblyServices.HSSteelMember section;

        //Fields
        string sectionName,baseAngle;
        #endregion

        #region Get Assembly Catalog Parts
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support supportOcc = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //Get Support Attributes
                    sectionName = (string)((PropertyValueString)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGHorSection", "HorSection")).PropValue;
                    baseAngle = (string)((PropertyValueString)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGBaseAngle", "BaseAngle")).PropValue;

                    //Add Parts
                    OglaendAssemblyServices.AddPart(this, SECTION, sectionName, parts);
                    OglaendAssemblyServices.AddPart(this, BASE_ANGLE, baseAngle, parts);
                    OglaendAssemblyServices.AddPart(this, STRUCT_CONN, "Log_Conn_Part_1", parts);

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
        #endregion

        #region Get Assembly Joints
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                SP3D.Support.Middle.Support supportOcc = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                BusinessObject supportPart = supportOcc.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                // Get the Steel Cross Section Data
                section = OglaendAssemblyServices.GetSectionDataFromPartIndex(this, SECTION);

                //Load standard bounding box definition
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                //=========================
                //1. Get bounding box boundary objects dimension information
                //=========================

                boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                //====== ======
                //2. Retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry
                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | boundingBoxHeight
                // |____________________|
                //    boundingBoxWidth

                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                //Get and Set Support/Part Attributes
                double totalOffset=0;
                double offset=0;
                double cantileverLength=0;
                double baseAngleThickness=0;
                double cantileverRotationAngle=0;

                offset = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJOAhsOGOffset", "Offset")).PropValue;
                cantileverLength = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGHSLengthR", "HSLength")).PropValue;
                baseAngleThickness = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[BASE_ANGLE], "IJUAHgrOGThickness", "Thickness")).PropValue;

                PropertyValueCodelist cardinalPoint6Section = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCP", "CP6");
                componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)6, "IJOAhsSteelCP", "CP6");
                componentDictionary[SECTION].SetPropertyValue(cantileverLength - baseAngleThickness, "IJUAHgrOccLength", "Length");
                componentDictionary[BASE_ANGLE].SetPropertyValue((double)0.005,"IJUAhsHolePort1", "HP1PosX");
                componentDictionary[BASE_ANGLE].SetPropertyValue((double)-0.0085, "IJUAhsHolePort1", "HP1PosY");
                componentDictionary[BASE_ANGLE].SetPropertyValue((double)0.0055, "IJUAhsHolePort3", "HP3PosZ");
                componentDictionary[BASE_ANGLE].SetPropertyValue((double)0.005, "IJUAhsHolePort3", "HP3PosX");
                componentDictionary[BASE_ANGLE].SetPropertyValue(-baseAngleThickness, "IJUAhsHolePort3", "HP3PosY");

                //Calculate Ports Orientation Angle
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (supportingType == "Steel")
                {
                    if ((routeStructAngle > (Math.PI / 4) && routeStructAngle < (3*(Math.PI / 4))) || (routeStructAngle > (5*(Math.PI / 4)) && routeStructAngle < (7*(Math.PI / 4))))
                    {
                        if(HgrCompareDoubleService.cmpdbl(offset,0)==true)
                            supportOcc.SetPropertyValue(boundingBoxWidth / 2, "IJOAhsOGOffset", "Offset");
                        else
                            offset = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJOAhsOGOffset", "Offset")).PropValue;

                        //Change Support Configurations
                        switch(Configuration)
                        {
                            case 1:
                                cantileverRotationAngle = 0;
                                totalOffset = offset + (section.depth/2);
                                break;
                            case 2:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = offset + (section.depth/2);
                                break;
                            case 3:
                                cantileverRotationAngle = 0;
                                totalOffset = -offset - (section.depth/2);
                                break;
                            case 4:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = -offset - (section.depth/2);
                                break;
                            case 5:
                                cantileverRotationAngle = Math.PI/2;
                                totalOffset = offset + (section.width/2);
                                break;
                            case 6:
                                cantileverRotationAngle = 3*Math.PI/2;
                                totalOffset = offset + (section.width/2);
                                break;
                            case 7:
                                cantileverRotationAngle = Math.PI/2;
                                totalOffset = -offset - (section.width/2);
                                break;
                            case 8:
                                cantileverRotationAngle = 3*Math.PI/2;
                                totalOffset = -offset - (section.width/2);
                                break;
                        }

                        //Add Joints
                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure",Plane.XY,Plane.XY,Axis.X,Axis.X,0,0,totalOffset);
                        JointHelper.CreateAngularRigidJoint(STRUCT_CONN, "Connection", BASE_ANGLE, "Hole3", new Vector(0,0,0),new Vector(3*Math.PI/2,0,cantileverRotationAngle));
                        JointHelper.CreateRigidJoint(SECTION, "BeginFlex", BASE_ANGLE, "Hole1",Plane.ZX,Plane.NegativeYZ,Axis.X,Axis.NegativeZ,0,0,0);
                    }
                    else
                    {
                        if (HgrCompareDoubleService.cmpdbl(offset, 0) == true)
                            supportOcc.SetPropertyValue((double)0, "IJOAhsOGOffset", "Offset");
                        else
                            offset = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJOAhsOGOffset", "Offset")).PropValue;

                        //Get Distance Between Ports
                        double projectionY = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.Y);
                        double projectionZ = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.Z);

                        //Get Normal Vector along the Surface
                        Matrix4X4 hangerPort = RefPortHelper.PortLCS("Structure");
                        Vector structureZVector = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);

                        //Get the Vector Normal to the Structure Face (May Not be the Z Axis in certain Cases)
                        IPlane iplane = (IPlane)supportOcc.SupportingFaces[0];
                        Vector structureFaceNormal = new Vector(iplane.Normal.X, iplane.Normal.Y, iplane.Normal.Z);

                        //Change Support Configurations
                        switch (Configuration)
                        {
                            case 1:
                                cantileverRotationAngle = 0;
                                totalOffset = offset;
                                break;
                            case 2:
                                cantileverRotationAngle = Math.PI/2;
                                totalOffset = offset;
                                break;
                            case 3:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = offset; 
                                break;
                            case 4:
                                cantileverRotationAngle = 3 * Math.PI / 2;
                                totalOffset = offset; 
                                break;
                            case 5:
                                cantileverRotationAngle = 0;
                                totalOffset = offset;
                                break;
                            case 6:
                                cantileverRotationAngle = Math.PI / 2;
                                totalOffset = offset; 
                                break;
                            case 7:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = offset; 
                                break;
                            case 8:
                                cantileverRotationAngle = 3 * Math.PI / 2;
                                totalOffset = offset; 
                                break;
                        }

                        if (Math.Abs(projectionZ) > Math.Abs(projectionY))
                        {
                            if (OglaendAssemblyServices.AngleBetweenVectors(structureFaceNormal, structureZVector) < Math.PI / 2)
                            {
                                if (projectionY > 0)
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, totalOffset, 0);
                                }
                                else
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -totalOffset, 0);

                                }
                            }
                            else
                            {
                                if (projectionY > 0)
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, totalOffset, 0);
                                }
                                else
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -totalOffset, 0);

                                }
                            }
                        }
                        else
                        {
                            if (OglaendAssemblyServices.AngleBetweenVectors(structureFaceNormal, structureZVector) < Math.PI / 2)
                            {
                                if (projectionY > 0)
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, totalOffset, 0);
                                }
                                else
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, totalOffset, 0);

                                }
                            }
                            else
                            {
                                if (projectionY > 0)
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -totalOffset, 0);
                                }
                                else
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -totalOffset, 0);

                                }
                            }
                        }

                        //Add Remaining joints
                        JointHelper.CreateAngularRigidJoint(STRUCT_CONN, "Connection", BASE_ANGLE, "Hole3", new Vector(0, 0, 0), new Vector(3 * Math.PI / 2, 0, (Math.PI / 2) + cantileverRotationAngle));
                        JointHelper.CreateRigidJoint(SECTION, "BeginFlex", BASE_ANGLE, "Hole1", Plane.ZX, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, 0, 0, 0);
                    }
                }
                else
                {
                    if (HgrCompareDoubleService.cmpdbl(offset, 0) == true)
                        supportOcc.SetPropertyValue(boundingBoxWidth / 2, "IJOAhsOGOffset", "Offset");
                    else
                        offset = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJOAhsOGOffset", "Offset")).PropValue;

                    //Change Support Configurations
                    switch (Configuration)
                    {
                        case 1:
                            cantileverRotationAngle = 0;
                            totalOffset = offset + (section.depth/2);
                            break;
                        case 2:
                            cantileverRotationAngle = Math.PI;
                            totalOffset = offset + (section.depth/2);
                            break;
                        case 3:
                            cantileverRotationAngle = 0;
                            totalOffset = -offset - (section.depth/2);
                            break;
                        case 4:
                            cantileverRotationAngle = Math.PI;
                            totalOffset = -offset - (section.depth/2);
                            break;
                        case 5:
                            cantileverRotationAngle = Math.PI / 2;
                            totalOffset = offset + (section.width/2);
                            break;
                        case 6:
                            cantileverRotationAngle = 3 * Math.PI / 2;
                            totalOffset = offset + (section.width/2);
                            break;
                        case 7:
                            cantileverRotationAngle = Math.PI / 2;
                            totalOffset = -offset - (section.width/2);
                            break;
                        case 8:
                            cantileverRotationAngle = 3 * Math.PI / 2;
                            totalOffset = -offset - (section.width/2);
                            break;
                    }

                    //Add Joints
                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, totalOffset, 0);
                    JointHelper.CreateAngularRigidJoint(STRUCT_CONN, "Connection", BASE_ANGLE, "Hole3", new Vector(0, 0, 0), new Vector(3 * Math.PI / 2, 0,(Math.PI / 2) + cantileverRotationAngle));
                    JointHelper.CreateRigidJoint(SECTION, "BeginFlex", BASE_ANGLE, "Hole1", Plane.ZX, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, 0, 0, 0);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion

        #region ConfigurationCount
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
        #endregion

        #region Get Route Connections
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
                    int numberOfRoutes = SupportHelper.SupportedObjects.Count;
                    for (int routeIndex = 1; routeIndex <= numberOfRoutes; routeIndex++)
                    {
                        routeConnections.Add(new ConnectionInfo(SECTION, routeIndex)); // partindex, routeindex     
                    }

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
        #endregion

        #region Get Struct Connections
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

                    structConnections.Add(new ConnectionInfo(BASE_ANGLE, 1)); // partindex, routeindex

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
        #endregion

        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject supportOcc)
        {
            string bomDesciption = "";
            try
            {
                string supportNumber = (string)((PropertyValueString)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsSupportNumber", "SupportNumber")).PropValue;
                bomDesciption = supportNumber;

                return bomDesciption;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}
