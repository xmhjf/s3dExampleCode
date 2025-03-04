//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_L_SHAPE.cs
//   Halfen_Assy,Ingr.SP3D.Content.Support.Rules.HALFEN_L_SHAPE
//   Author       :  Hema
//   Creation Date:  12.12.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who        change description
//   -----------     ---       ------------------
//    Hema         12.12.2012   CR-CP-224495 C#.Net HS_Halfen_Assy Project Creation
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;

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
    [SymbolVersion("1.0.0.0")]
    public class HALFEN_L_SHAPE : CustomSupportDefinition
    {
        //Constants
        private const string BASE_CHANNEL = "BASE_CHANNEL";
        private const string LOWER_CONNECTOR = "LOWER_CONNECTOR";
        private const string VERT_CHANNEL = "VERT_CHANNEL";
        private const string UPPER_CONNECTOR = "UPPER_CONNECTOR";
        private const string TK_1 = "TK_1";
        private const string TK_2 = "TK_2";
        private const string TK_3 = "TK_3";
        private const string TK_4 = "TK_4";
        private const string Steel = "Steel";
        private const string Slab = "Slab";

        private string channelPartKeys;
        private double arm_L;
        private double shoeHeight;
        private double overLength;
        public int numOfStruct;
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

                    arm_L = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenLShape", "ARM1_L")).PropValue;
                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenShoeH", "SHOE_H")).PropValue;

                    PropertyValueCodelist base_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenLBaseSize", "BASE_SIZE");
                    long base_Size = (long)base_SizeCodelist.PropValue;

                    PropertyValueCodelist connector_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenLConnectorSize", "CONNECTOR_SIZE");
                    long connector_Size = (long)connector_SizeCodelist.PropValue;

                    overLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenOverlength", "OVERLENGTH")).PropValue;

                    String lowerConnectorType;
                    if (connector_Size == 2)
                        lowerConnectorType = "HALFEN_HCS_VT63_22_1";
                    else
                    {
                        if (connector_Size == 3 || connector_Size == 4)
                            lowerConnectorType = "HALFEN_HCS_VT63_23";
                        else
                            lowerConnectorType = "HALFEN_HCS_VT63_21";
                    }
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

                    
                    //Determine whether connecting to Steel or a Slab
                    if (supportingType == "Steel")
                    {
                        parts.Add(new PartInfo(BASE_CHANNEL, "HALFEN_HZL_63_63_1"));
                        parts.Add(new PartInfo(LOWER_CONNECTOR, lowerConnectorType));
                        parts.Add(new PartInfo(VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                        parts.Add(new PartInfo(UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_Size));
                        parts.Add(new PartInfo(TK_1, "HALFEN_HCS_TK_2"));
                        parts.Add(new PartInfo(TK_2, "HALFEN_HCS_TK_2"));
                        parts.Add(new PartInfo(TK_3, "HALFEN_HCS_TK_2"));
                        parts.Add(new PartInfo(TK_4, "HALFEN_HCS_TK_2"));
                    }
                    else if (supportingType == "Slab")
                    {
                        parts.Add(new PartInfo(BASE_CHANNEL, "HALFEN_HZL_63_63_1"));
                        parts.Add(new PartInfo(LOWER_CONNECTOR, lowerConnectorType));
                        parts.Add(new PartInfo(VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                        parts.Add(new PartInfo(UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_Size));
                    }
                    //Return the collection of Catalog Parts
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
                //Load standard bounding box definition
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent upper_Connector = componentDictionary[UPPER_CONNECTOR];
                SupportComponent lower_Connector = componentDictionary[LOWER_CONNECTOR];
                SupportComponent base_Channel = componentDictionary[BASE_CHANNEL];

                PropertyValueCodelist connector_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenLConnectorSize", "CONNECTOR_SIZE");
                long connector_Size = (long)connector_SizeCodelist.PropValue;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                }
                else
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }

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

                double BBXWidth = BBX.Width;
                double BBXHeight = BBX.Height;

                int structConfig;
                double channelWidth = 0.063;
                double inset = 0.005;
                double lower_Connector_W = channelWidth + (0.005 * 2.0);

                int routeConnectionValue = Configuration;

                double byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                if (Math.Abs(byPointAngle2) > Math.PI / 2.0) //The structure is oriented in the standard direction
                {
                    if (Math.Abs(byPointAngle1) < Math.PI / 2.0)
                        structConfig = 1;
                    else
                        structConfig = 2;
                }
                else //The structure is oriented in the opposite direction
                {
                    if (Math.Abs(byPointAngle1) < Math.PI / 2.0)
                        structConfig = 3;
                    else
                        structConfig = 4;
                }
                //Look up component properties
                BusinessObject upper_ConnectorPart = upper_Connector.GetRelationship("madeFrom", "part").TargetObjects[0];

                double upper_Connector_T = (double)((PropertyValueDouble)upper_ConnectorPart.GetPropertyValue("IJUAHgrT", "T")).PropValue;
                double bolt_CC = (double)((PropertyValueDouble)upper_ConnectorPart.GetPropertyValue("IJUAHgrBolt_CC", "Bolt_CC")).PropValue;
                double bolt_CC2 = (double)((PropertyValueDouble)upper_ConnectorPart.GetPropertyValue("IJUAHgrBolt_CC2", "Bolt_CC2")).PropValue;

                if (connector_Size == 4)
                {
                    PropertyValueCodelist configCodelist = (PropertyValueCodelist)lower_Connector.GetPropertyValue("IJOAHgrConfig", "Config");
                    configCodelist.PropValue = 1;
                    lower_Connector.SetPropertyValue(configCodelist.PropValue, "IJOAHgrConfig", "Config");
                }
                //Start Joints here
                //Here is the place by structure stuff
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    //Add joint between vertical channel and upper connector
                    JointHelper.CreateRigidJoint(UPPER_CONNECTOR, "Base", VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, upper_Connector_T, 0, 0);

                    if (routeConnectionValue == 1) //places to the LEFT of route
                    {
                        //Add Joint Between the Top connector and the Overhead beam
                        JointHelper.CreatePrismaticJoint(UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0.0, 0.0);
                        //Add joints between Route and horizontal channel
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, BBXWidth / 2.0 - arm_L / 2.0);
                    }
                    else //places to the RIGHT of route
                    {
                        //Add Joint Between the Top connector and the Overhead beam
                        JointHelper.CreatePrismaticJoint(UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0.0, 0.0);
                        //Add joints between Route and horizontal channel
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + arm_L / 2.0);

                    }
                }
                else //PLACE BY POINT
                {
                    double horizontallength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortDistanceType.Horizontal);

                    //Add joint between vertical channel and upper connector
                    JointHelper.CreateRigidJoint(UPPER_CONNECTOR, "Base", VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                    //Add Joint Between the Top connector and the Overhead beam
                    JointHelper.CreatePlanarJoint(UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);

                    //Add joint between Route and horizontal channel
                    if (structConfig == 1 || structConfig == 3)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, horizontallength - arm_L + inset - lower_Connector_W / 2.0, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, -horizontallength + arm_L - inset + lower_Connector_W / 2.0, 0);
                }
                //Set Horizontal Arm Length
                base_Channel.SetPropertyValue(arm_L - inset, "IJOAHgrOccLength", "Length");

                //Add joint between horizontal channel and connector
                JointHelper.CreateRigidJoint(LOWER_CONNECTOR, "Right", BASE_CHANNEL, "EndMiddle", Plane.XY, Plane.YZ, Axis.X, Axis.Y, 0.0, 0.0, 0);

                //Add joint between vertical channel and lower connector
                JointHelper.CreateRigidJoint(LOWER_CONNECTOR, "Base", VERT_CHANNEL, "EndMiddle", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -overLength, 0, 0);

                //Flexible Member (Vertical)
                JointHelper.CreatePrismaticJoint(VERT_CHANNEL, "EndMiddle", VERT_CHANNEL, "BeginMiddle", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                if (supportingType == "Steel")
                {
                    //Add joints between upper connector and clamps
                    JointHelper.CreateRigidJoint(TK_1, "BottomOfClamp", UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                    JointHelper.CreateRigidJoint(TK_2, "BottomOfClamp", UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                    JointHelper.CreateRigidJoint(TK_3, "BottomOfClamp", UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                    JointHelper.CreateRigidJoint(TK_4, "BottomOfClamp", UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    int numberOfRoutes = SupportHelper.SupportedObjects.Count;

                    //For Clamp
                    int clamp_Begin = 1;
                    int numberOfPart = numberOfRoutes;
                    int clamp_End = numberOfPart;

                    for (int index = clamp_Begin; index <= clamp_End; index++)
                    {
                        channelPartKeys = "BASE_CHANNEL";
                        int connectToRoute = index;
                        routeConnections.Add(new ConnectionInfo(channelPartKeys, connectToRoute));
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
                        structConnections.Add(new ConnectionInfo(UPPER_CONNECTOR, 1));
                        structConnections.Add(new ConnectionInfo(TK_1, 1));
                        structConnections.Add(new ConnectionInfo(TK_2, 1));
                        structConnections.Add(new ConnectionInfo(TK_3, 1));
                        structConnections.Add(new ConnectionInfo(TK_4, 1));
                    }
                    else
                    {
                        structConnections.Add(new ConnectionInfo(UPPER_CONNECTOR, 1));
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
    }
}



