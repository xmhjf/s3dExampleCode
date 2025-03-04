//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_U_SHAPE.cs
//   Halfen_Assy,Ingr.SP3D.Content.Support.Rules.HALFEN_U_SHAPE
//   Author       :  Hema
//   Creation Date:  12.12.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who        change description
//   -----------     ---       ------------------
//    Hema         12.12.2012   CR-CP-224495 C#.Net HS_Halfen_Assy Project Creation
// 22-Jan-2015       PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
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
    [SymbolVersion("1.0.0.0")]
    public class HALFEN_U_SHAPE : CustomSupportDefinition
    {
        //Constants
        //Use these parts if vertical channels are necessary
        private const string BASE_CHANNEL = "BASE_CHANNEL";
        private const string L_LOWER_CONNECTOR = "L_LOWER_CONNECTOR";
        private const string L_VERT_CHANNEL = "L_VERT_CHANNEL";
        private const string L_UPPER_CONNECTOR = "L_UPPER_CONNECTOR";
        private const string R_LOWER_CONNECTOR = "R_LOWER_CONNECTOR";
        private const string R_VERT_CHANNEL = "R_VERT_CHANNEL";
        private const string R_UPPER_CONNECTOR = "R_UPPER_CONNECTOR";
        private const string L_TK_1 = "L_TK_1";
        private const string L_TK_2 = "L_TK_2";
        private const string L_TK_3 = "L_TK_3";
        private const string L_TK_4 = "L_TK_4";
        private const string R_TK_1 = "R_TK_1";
        private const string R_TK_2 = "R_TK_2";
        private const string R_TK_3 = "R_TK_3";
        private const string R_TK_4 = "R_TK_4";

        //Use these parts if vertical channels are unnecessary
        private const string ALT_L_LOWER_CONNECTOR = "ALT_L_LOWER_CONNECTOR";
        private const string ALT_L_VERT_CHANNEL="ALT_L_VERT_CHANNEL";
        private const string ALT_L_UPPER_CONNECTOR = "ALT_L_UPPER_CONNECTOR";
        private const string ALT_R_LOWER_CONNECTOR = "ALT_R_LOWER_CONNECTOR";
        private const string ALT_R_VERT_CHANNEL = "ALT_R_VERT_CHANNEL";
        private const string ALT_R_UPPER_CONNECTOR = "ALT_R_UPPER_CONNECTOR";
        private const string ALT_L_TK_1 = "ALT_L_TK_1";
        private const string ALT_L_TK_2 = "ALT_L_TK_2";
        private const string ALT_L_TK_3 = "ALT_L_TK_3";
        private const string ALT_L_TK_4 = "ALT_L_TK_4";
        private const string ALT_R_TK_1 = "ALT_R_TK_1";
        private const string ALT_R_TK_2 = "ALT_R_TK_2";
        private const string ALT_R_TK_3 = "ALT_R_TK_3";
        private const string ALT_R_TK_4 = "ALT_R_TK_4";

        private const string Steel = "Steel";
        private const string Slab = "Slab";
        private const string SlabSteel = "SlabSteel";
        private const string SteelSlab = "SteelSlab";

        private string channelPartKeys;
        private string connection;
        private double shoeHeight;
        private double length;
        private double width;
        public int base_SizeCodelistValue;
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

                    //Determine whether connecting to Steel or a Slab
                    if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        connection = Steel;
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab)) //Two Slabs
                            connection = Slab;
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Member)) //Slab then Steel
                            connection = SlabSteel;
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab)) //Steel then Slab
                            connection = SteelSlab;
                    }
                    else
                    {
                        if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                            connection = Steel;
                        else if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab)
                            connection = Slab;
                    }

                    width = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenUShape", "WIDTH")).PropValue;
                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenShoeH", "SHOE_H")).PropValue;

                    PropertyValueCodelist base_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenUBaseSize", "BASE_SIZE");
                    long base_SizeCodelistValue = (long)base_SizeCodelist.PropValue;

                    length = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenLength", "LENGTH")).PropValue;

                    PropertyValueCodelist connector_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenUConnectorSize", "CONNECTOR_SIZE");
                    long connector_SizeCodelistValue = (long)connector_SizeCodelist.PropValue;

                    String lowerConnectorType;
                    if (connector_SizeCodelistValue == 2)
                        lowerConnectorType = "HALFEN_HCS_VT63_22_1";
                    else
                        lowerConnectorType = "HALFEN_HCS_VT63_2" + connector_SizeCodelistValue;

                    if (connection == Steel)
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                            parts.Add(new PartInfo(BASE_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(ALT_L_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(ALT_L_VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(ALT_L_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(ALT_R_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(ALT_R_VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(ALT_R_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(ALT_L_TK_1, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_L_TK_2, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_L_TK_3, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_L_TK_4, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_R_TK_1, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_R_TK_2, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_R_TK_3, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_R_TK_4, "HALFEN_HCS_TK_2"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(BASE_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(L_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(L_VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(L_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(R_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(R_VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(R_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(L_TK_1, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(L_TK_2, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(L_TK_3, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(L_TK_4, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(R_TK_1, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(R_TK_2, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(R_TK_3, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(R_TK_4, "HALFEN_HCS_TK_2"));
                        }
                    }
                    if (connection == SlabSteel || connection == SteelSlab)
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                            parts.Add(new PartInfo(BASE_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(ALT_L_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(ALT_L_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(ALT_R_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(ALT_R_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(ALT_L_TK_1, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_L_TK_2, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_L_TK_3, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_L_TK_4, "HALFEN_HCS_TK_2"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(BASE_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(L_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(L_VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(L_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(R_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(R_VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(R_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(L_TK_1, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(L_TK_2, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(L_TK_3, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(L_TK_4, "HALFEN_HCS_TK_2"));
                        }
                    }
                    if (connection == Slab)
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                            parts.Add(new PartInfo(BASE_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(ALT_L_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(ALT_L_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(ALT_R_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(ALT_R_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                        }

                        else
                        {
                            parts.Add(new PartInfo(BASE_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(L_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(L_VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(L_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(R_LOWER_CONNECTOR, lowerConnectorType));
                            parts.Add(new PartInfo(R_VERT_CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(R_UPPER_CONNECTOR, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                        }
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
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                //Load standard bounding box definition
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;

                PropertyValueCodelist base_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenUBaseSize", "BASE_SIZE");
                long base_SizeCodelistValue = (long)base_SizeCodelist.PropValue;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent base_Channel = componentDictionary[BASE_CHANNEL];

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                }
                else
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }

                double BBXWidth = BBX.Width;
                double BBXHeight = BBX.Height;

                Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();
                string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
                string rightStructPort = idxStructPort[0];
                string leftStructPort = idxStructPort[1];

                double inset = 0.005;
                double channel_W = 0.063;
                int structCount = SupportHelper.SupportingObjects.Count;

                double byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);
                int structConfig = 0;

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

                double upper_Connector_T;
                double bolt_CC;
                double bolt_CC2;
                if (base_SizeCodelistValue > 3)
                {
                    SupportComponent alt_L_Upper_Connector = componentDictionary[ALT_L_UPPER_CONNECTOR];
                    BusinessObject alt_L_Upper_Connectorpart = alt_L_Upper_Connector.GetRelationship("madeFrom", "part").TargetObjects[0];
                    upper_Connector_T = (double)((PropertyValueDouble)alt_L_Upper_Connectorpart.GetPropertyValue("IJUAHgrT", "T")).PropValue;
                    bolt_CC = (double)((PropertyValueDouble)alt_L_Upper_Connectorpart.GetPropertyValue("IJUAHgrBolt_CC", "Bolt_CC")).PropValue;
                    bolt_CC2 = (double)((PropertyValueDouble)alt_L_Upper_Connectorpart.GetPropertyValue("IJUAHgrBolt_CC2", "Bolt_CC2")).PropValue;
                }
                else
                {
                    SupportComponent L_Upper_Connector = componentDictionary[L_UPPER_CONNECTOR];
                    BusinessObject L_Upper_Connectorpart = L_Upper_Connector.GetRelationship("madeFrom", "part").TargetObjects[0];
                    upper_Connector_T = (double)((PropertyValueDouble)L_Upper_Connectorpart.GetPropertyValue("IJUAHgrT", "T")).PropValue;
                    bolt_CC = (double)((PropertyValueDouble)L_Upper_Connectorpart.GetPropertyValue("IJUAHgrBolt_CC", "Bolt_CC")).PropValue;
                    bolt_CC2 = (double)((PropertyValueDouble)L_Upper_Connectorpart.GetPropertyValue("IJUAHgrBolt_CC2", "Bolt_CC2")).PropValue;
                }
                double lower_Connector_W = channel_W + (0.005 * 2.0);
                double lower_Connector_H = 0.165;
                if (base_SizeCodelistValue > 3 && length > 470.0 - upper_Connector_T - lower_Connector_H)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "The length specified is too large for base type.", "", "HALFEN_U_SHAPE.cs", 1);

                //Start Joints here
                if (base_SizeCodelistValue < 4)
                {
                    //WITH VERTICAL CHANNEL
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)   //Here is the place by structure stuff
                    {
                        //Add joint between vertical channels and upper connectors
                        JointHelper.CreateRigidJoint(L_UPPER_CONNECTOR, "Base", L_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, upper_Connector_T, 0, 0);
                        JointHelper.CreateRigidJoint(R_UPPER_CONNECTOR, "Base", R_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, upper_Connector_T, 0, 0);
                        //Add Joint Between the Top connectors and the Overhead beam
                        JointHelper.CreatePrismaticJoint(L_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0.0, 0.0);
                        JointHelper.CreatePlanarJoint(R_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                        //Add joints between Route and horizontal channel
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, BBXWidth / 2.0 - width / 2.0);
                    }
                    else    //PLACE BY POINT 
                    {
                        double L_Vertical_Length = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Vertical);
                        double R_Vertical_Length = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Vertical);
                        double vertical_Offset = L_Vertical_Length - R_Vertical_Length;
                        double L_Horizontal_Length = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                        double R_Horizontal_Length = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);

                        if (connection == Steel && structCount == 1)
                        {
                            //Add joint between vertical channels and upper connectors
                            JointHelper.CreateRigidJoint(L_UPPER_CONNECTOR, "Base", L_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            JointHelper.CreateRigidJoint(R_UPPER_CONNECTOR, "Base", R_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(L_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(R_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, L_Horizontal_Length - width + inset * 2.0 - lower_Connector_W / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, -L_Horizontal_Length + width - inset * 2.0 + lower_Connector_W / 2.0, 0);
                        }

                        if (connection == Slab && structCount == 1)
                        {
                            //Add joint between vertical channels and upper connectors
                            JointHelper.CreateRigidJoint(L_UPPER_CONNECTOR, "Base", L_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            JointHelper.CreateRigidJoint(R_UPPER_CONNECTOR, "Base", R_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(L_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(R_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, -BBXWidth / 2.0 - width / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + width / 2.0, 0);
                        }

                        if (connection == Steel && structCount > 1)
                        {
                            width = L_Horizontal_Length + R_Horizontal_Length - channel_W;
                            //Add joint between vertical channels and upper connectors
                            JointHelper.CreateRigidJoint(L_UPPER_CONNECTOR, "Base", L_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            JointHelper.CreateRigidJoint(R_UPPER_CONNECTOR, "Base", R_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(L_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(R_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, vertical_Offset);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, R_Horizontal_Length - width + inset * 2.0 - lower_Connector_W / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, -L_Horizontal_Length + width - inset * 2.0 + lower_Connector_W / 2.0, 0);
                        }

                        if (connection == Slab && structCount > 1)
                        {
                            //Add joint between vertical channels and upper connectors
                            JointHelper.CreateRigidJoint(L_UPPER_CONNECTOR, "Base", L_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            JointHelper.CreateRigidJoint(R_UPPER_CONNECTOR, "Base", R_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(L_UPPER_CONNECTOR, "Base", "-1", rightStructPort, Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(R_UPPER_CONNECTOR, "Base", "-1", leftStructPort, Plane.XY, Plane.XY, vertical_Offset);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, -BBXWidth / 2.0 - width / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + width / 2.0, 0);
                        }

                        if (connection == SlabSteel)
                        {
                            //Add joint between vertical channels and upper connectors
                            JointHelper.CreateRigidJoint(L_UPPER_CONNECTOR, "Base", L_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            JointHelper.CreateRigidJoint(R_UPPER_CONNECTOR, "Base", R_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(L_UPPER_CONNECTOR, "Base", "-1", rightStructPort, Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(R_UPPER_CONNECTOR, "Base", "-1", leftStructPort, Plane.XY, Plane.XY, vertical_Offset);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, -BBXWidth / 2.0 - width / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + width / 2.0, 0);
                        }

                        if (connection == SteelSlab)
                        {
                            //Add joint between vertical channels and upper connectors
                            JointHelper.CreateRigidJoint(L_UPPER_CONNECTOR, "Base", L_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            JointHelper.CreateRigidJoint(R_UPPER_CONNECTOR, "Base", R_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, upper_Connector_T, 0, 0);
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(L_UPPER_CONNECTOR, "Base", "-1", rightStructPort, Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(R_UPPER_CONNECTOR, "Base", "-1", leftStructPort, Plane.XY, Plane.XY, vertical_Offset);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, -BBXWidth / 2.0 - width / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + width / 2.0, 0);
                        }
                    }

                    //Set Horizontal Arm Length
                    base_Channel.SetPropertyValue(width - inset * 2.0, "IJOAHgrOccLength", "Length");

                    //Add joint between horizontal channel and connectors
                    JointHelper.CreateRigidJoint(L_LOWER_CONNECTOR, "Right", BASE_CHANNEL, "EndMiddle", Plane.XY, Plane.YZ, Axis.X, Axis.Y, 0.0, 0.0, 0.0);
                    JointHelper.CreateRigidJoint(R_LOWER_CONNECTOR, "Right", BASE_CHANNEL, "BeginMiddle", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0.0, 0.0, 0.0);

                    // Add joint between vertical channels and lower connectors
                    JointHelper.CreateRigidJoint(L_LOWER_CONNECTOR, "Base", L_VERT_CHANNEL, "EndMiddle", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -0.1, 0.0, 0);
                    JointHelper.CreateRigidJoint(R_LOWER_CONNECTOR, "Top", R_VERT_CHANNEL, "EndMiddle", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.1, 0.0, 0);

                    //Flexible Member (Vertical)
                    JointHelper.CreatePrismaticJoint(L_VERT_CHANNEL, "EndMiddle", L_VERT_CHANNEL, "BeginMiddle", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    JointHelper.CreatePrismaticJoint(R_VERT_CHANNEL, "EndMiddle", R_VERT_CHANNEL, "BeginMiddle", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    if (connection == Steel)
                    {
                        //Add joints between upper connector and clamps
                        JointHelper.CreateRigidJoint(L_TK_1, "BottomOfClamp", L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_2, "BottomOfClamp", L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_3, "BottomOfClamp", L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_4, "BottomOfClamp", L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(R_TK_1, "BottomOfClamp", R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(R_TK_2, "BottomOfClamp", R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(R_TK_3, "BottomOfClamp", R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(R_TK_4, "BottomOfClamp", R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
                    }

                    if (connection == SlabSteel)
                    {
                        //Add joints between upper connector and clamps
                        JointHelper.CreateRigidJoint(L_TK_1, "BottomOfClamp", L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_2, "BottomOfClamp", L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_3, "BottomOfClamp", L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_4, "BottomOfClamp", L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
                    }

                    if (connection == SteelSlab)
                    {
                        //Add joints between upper connector and clamps
                        JointHelper.CreateRigidJoint(L_TK_1, "BottomOfClamp", R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_2, "BottomOfClamp", R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_3, "BottomOfClamp", R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(L_TK_4, "BottomOfClamp", R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
                    }
                }
                else
                {
                    //WITHOUT VERTICAL CHANNEL
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)   //Here is the place by structure stuff
                    {
                        //Add joint between vertical channels and upper connectors
                        JointHelper.CreateRigidJoint(ALT_L_UPPER_CONNECTOR, "Base", ALT_L_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, upper_Connector_T, 0, 0);
                        JointHelper.CreateRigidJoint(ALT_R_UPPER_CONNECTOR, "Base", ALT_R_VERT_CHANNEL, "BeginMiddle", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, upper_Connector_T, 0, 0);
                        //Add Joint Between the Top connectors and the Overhead beam
                        JointHelper.CreatePrismaticJoint(ALT_L_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0.0, 0.0);
                        JointHelper.CreatePlanarJoint(ALT_R_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                        //Add joints between Route and horizontal channel
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, BBXWidth / 2.0 - width / 2.0);
                    }
                    else    //PLACE BY POINT 
                    {
                        double L_Vertical_Length = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Vertical);
                        double R_Vertical_Length = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Vertical);
                        double vertical_Offset = L_Vertical_Length - R_Vertical_Length;
                        double L_Horizontal_Length = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                        double R_Horizontal_Length = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);

                        if (connection == Steel && structCount == 1)
                        {
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(ALT_L_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(ALT_R_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, L_Horizontal_Length - width + inset * 2.0 - lower_Connector_W / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, -L_Horizontal_Length + width - inset * 2.0 + lower_Connector_W / 2.0, 0);
                        }

                        if (connection == Slab && structCount == 1)
                        {
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(ALT_L_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(ALT_R_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, -BBXWidth / 2.0 - width / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + width / 2.0, 0);
                        }

                        if (connection == Steel && structCount > 1)
                        {
                            width = L_Horizontal_Length + R_Horizontal_Length - channel_W;
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(ALT_L_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(ALT_R_UPPER_CONNECTOR, "Base", "-1", "Structure", Plane.XY, Plane.XY, vertical_Offset);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, R_Horizontal_Length - width + inset * 2.0 - lower_Connector_W / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, -L_Horizontal_Length + width - inset * 2.0 + lower_Connector_W / 2.0, 0);
                        }

                        if (connection == Slab && structCount > 1)
                        {
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(ALT_L_UPPER_CONNECTOR, "Base", "-1", rightStructPort, Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(ALT_R_UPPER_CONNECTOR, "Base", "-1", leftStructPort, Plane.XY, Plane.XY, vertical_Offset);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, -BBXWidth / 2.0 - width / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + width / 2.0, 0);
                        }

                        if (connection == SlabSteel)
                        {
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(ALT_L_UPPER_CONNECTOR, "Base", "-1", rightStructPort, Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(ALT_R_UPPER_CONNECTOR, "Base", "-1", leftStructPort, Plane.XY, Plane.XY, vertical_Offset);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, -BBXWidth / 2.0 - width / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + width / 2.0, 0);
                        }

                        if (connection == SteelSlab)
                        {
                            //Add Joint Between the Top connectors and the Overhead beam
                            JointHelper.CreatePlanarJoint(ALT_L_UPPER_CONNECTOR, "Base", "-1", rightStructPort, Plane.XY, Plane.XY, 0);
                            JointHelper.CreatePlanarJoint(ALT_R_UPPER_CONNECTOR, "Base", "-1", leftStructPort, Plane.XY, Plane.XY, vertical_Offset);

                            //Add joint between Route and horizontal channel
                            if (structConfig == 1 || structConfig == 3)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -shoeHeight, -BBXWidth / 2.0 - width / 2.0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", BASE_CHANNEL, "BeginTop", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -shoeHeight, BBXWidth / 2.0 + width / 2.0, 0);
                        }
                    }

                    //Set Horizontal Arm Length
                    base_Channel.SetPropertyValue(width - inset * 2.0, "IJOAHgrOccLength", "Length");

                    //Add Joint Between the Top connectors lower end and the lower connector
                    JointHelper.CreatePrismaticJoint(ALT_L_UPPER_CONNECTOR, "Top", ALT_L_LOWER_CONNECTOR, "Middle", Plane.ZX, Plane.ZX, Axis.Z, Axis.NegativeZ, 0, 0);
                    JointHelper.CreatePrismaticJoint(ALT_R_UPPER_CONNECTOR, "Top", ALT_R_LOWER_CONNECTOR, "Middle", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.Z, 0, 0);

                    // Add joint between horizontal channel and connectors
                    JointHelper.CreateRigidJoint(ALT_L_LOWER_CONNECTOR, "Right", BASE_CHANNEL, "EndMiddle", Plane.XY, Plane.YZ, Axis.X, Axis.Y, 0.0, 0.0, 0);
                    JointHelper.CreateRigidJoint(ALT_R_LOWER_CONNECTOR, "Right", BASE_CHANNEL, "BeginMiddle", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0.0, 0.0, 0);

                    //Flexible Member (Vertical)

                    if (connection == Steel)
                    {
                        //Add joints between upper connector and clamps
                        JointHelper.CreateRigidJoint(ALT_L_TK_1, "BottomOfClamp", ALT_L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_2, "BottomOfClamp", ALT_L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_3, "BottomOfClamp", ALT_L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_4, "BottomOfClamp", ALT_L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_R_TK_1, "BottomOfClamp", ALT_R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_R_TK_2, "BottomOfClamp", ALT_R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_R_TK_3, "BottomOfClamp", ALT_R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_R_TK_4, "BottomOfClamp", ALT_R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
                    }

                    if (connection == SlabSteel)
                    {
                        //Add joints between upper connector and clamps
                        JointHelper.CreateRigidJoint(ALT_L_TK_1, "BottomOfClamp", ALT_L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_2, "BottomOfClamp", ALT_L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_3, "BottomOfClamp", ALT_L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_4, "BottomOfClamp", ALT_L_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
                    }

                    if (connection == SteelSlab)
                    {
                        //Add joints between upper connector and clamps
                        JointHelper.CreateRigidJoint(ALT_L_TK_1, "BottomOfClamp", ALT_R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_2, "BottomOfClamp", ALT_R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2.0, bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_3, "BottomOfClamp", ALT_R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, -bolt_CC / 2.0);
                        JointHelper.CreateRigidJoint(ALT_L_TK_4, "BottomOfClamp", ALT_R_UPPER_CONNECTOR, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2.0, bolt_CC / 2.0);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    int numberOfRoutes = SupportHelper.SupportedObjects.Count;

                    //For Clamp
                    int clamp_Begin = 1;
                    int numOfPart = numberOfRoutes;
                    int clamp_End = numOfPart;

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

                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    PropertyValueCodelist base_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenUBaseSize", "BASE_SIZE");
                    long base_SizeCodelistValue = (long)base_SizeCodelist.PropValue;

                    if (connection == Steel)
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                                structConnections.Add(new ConnectionInfo(ALT_L_UPPER_CONNECTOR, 1));
                                structConnections.Add(new ConnectionInfo(ALT_R_UPPER_CONNECTOR, 2));
                                structConnections.Add(new ConnectionInfo(ALT_L_TK_1, 1));
                                structConnections.Add(new ConnectionInfo(ALT_L_TK_2, 1));
                                structConnections.Add(new ConnectionInfo(ALT_L_TK_3, 1));
                                structConnections.Add(new ConnectionInfo(ALT_L_TK_4, 1));
                                structConnections.Add(new ConnectionInfo(ALT_R_TK_1, 2));
                                structConnections.Add(new ConnectionInfo(ALT_R_TK_2, 2));
                                structConnections.Add(new ConnectionInfo(ALT_R_TK_3, 2));
                                structConnections.Add(new ConnectionInfo(ALT_R_TK_4, 2));
                        }
                        else
                        {
                                structConnections.Add(new ConnectionInfo(L_UPPER_CONNECTOR, 1));
                                structConnections.Add(new ConnectionInfo(R_UPPER_CONNECTOR, 2));
                                structConnections.Add(new ConnectionInfo(L_TK_1, 1));
                                structConnections.Add(new ConnectionInfo(L_TK_2, 1));
                                structConnections.Add(new ConnectionInfo(L_TK_3, 1));
                                structConnections.Add(new ConnectionInfo(L_TK_4, 1));
                                structConnections.Add(new ConnectionInfo(R_TK_1, 2));
                                structConnections.Add(new ConnectionInfo(R_TK_2, 2));
                                structConnections.Add(new ConnectionInfo(R_TK_3, 2));
                                structConnections.Add(new ConnectionInfo(R_TK_4, 2));
                        }
                    }
                    if (connection == SlabSteel)
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                                structConnections.Add(new ConnectionInfo(ALT_L_UPPER_CONNECTOR, 1));
                                structConnections.Add(new ConnectionInfo(ALT_R_UPPER_CONNECTOR, 2));
                                structConnections.Add(new ConnectionInfo(ALT_R_TK_1, 2));
                                structConnections.Add(new ConnectionInfo(ALT_R_TK_2, 2));
                                structConnections.Add(new ConnectionInfo(ALT_R_TK_3, 2));
                                structConnections.Add(new ConnectionInfo(ALT_R_TK_4, 2));
                        }
                        else
                        {
                                structConnections.Add(new ConnectionInfo(L_UPPER_CONNECTOR, 1));
                                structConnections.Add(new ConnectionInfo(R_UPPER_CONNECTOR, 2));
                                structConnections.Add(new ConnectionInfo(R_TK_1, 2));
                                structConnections.Add(new ConnectionInfo(R_TK_2, 2));
                                structConnections.Add(new ConnectionInfo(R_TK_3, 2));
                                structConnections.Add(new ConnectionInfo(R_TK_4, 2));
                        }
                    }

                    if (connection == SteelSlab)
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                                structConnections.Add(new ConnectionInfo(ALT_L_UPPER_CONNECTOR, 1));
                                structConnections.Add(new ConnectionInfo(ALT_R_UPPER_CONNECTOR, 2));
                                structConnections.Add(new ConnectionInfo(ALT_L_TK_1, 1));
                                structConnections.Add(new ConnectionInfo(ALT_L_TK_2, 1));
                                structConnections.Add(new ConnectionInfo(ALT_L_TK_3, 1));
                                structConnections.Add(new ConnectionInfo(ALT_L_TK_4, 1));
                        }
                        else
                        {
                                structConnections.Add(new ConnectionInfo(L_UPPER_CONNECTOR, 1));
                                structConnections.Add(new ConnectionInfo(R_UPPER_CONNECTOR, 2));
                                structConnections.Add(new ConnectionInfo(L_TK_1, 1));
                                structConnections.Add(new ConnectionInfo(L_TK_2, 1));
                                structConnections.Add(new ConnectionInfo(L_TK_3, 1));
                                structConnections.Add(new ConnectionInfo(L_TK_4, 1));
                        }
                    }

                    if (connection == Slab)
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                                structConnections.Add(new ConnectionInfo(ALT_L_UPPER_CONNECTOR, 1));
                                structConnections.Add(new ConnectionInfo(ALT_R_UPPER_CONNECTOR, 2));
                        }
                        else
                        {
                                structConnections.Add(new ConnectionInfo(L_UPPER_CONNECTOR, 1));
                                structConnections.Add(new ConnectionInfo(R_UPPER_CONNECTOR, 2));
                        }
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
        /// <summary>
        /// This method differentiate multiple structure input object based on relative position.
        /// </summary>
        /// <param name="IsOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
        /// </code>
        public String[] GetIndexedStructPortName(Boolean[] IsOffsetApplied)
        {
            String[] structurePort = new String[2];
            int structureCount = SupportHelper.SupportingObjects.Count;
            int i;

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

            if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
            {
                structurePort[0] = "Structure";
                structurePort[1] = "Structure";
            }
            else
            {
                structurePort[0] = "Structure";
                structurePort[1] = "Struct_2";

                if (structureCount > 1)
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {
                        for (i = 0; i <= 1; i++)
                        {
                            double angle = 0;
                            if ((supportingType == "Steel") && IsOffsetApplied[i] == false)
                            {
                                angle = RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
                            }
                            //the port is the right structure port
                            if (Math.Abs(angle) < Math.PI / 2.0)
                            {
                                if (i == 0)
                                {
                                    structurePort[0] = "Struct_2";
                                    structurePort[1] = "Structure";
                                }
                            }
                            //the port is the left structure port
                            else
                            {
                                if (i == 1)
                                {
                                    structurePort[0] = "Struct_2";
                                    structurePort[1] = "Structure";
                                }
                            }
                        }
                    }
                }
                else
                    structurePort[1] = "Structure";
            }
            //switch the OffsetApplied flag
            if (structurePort[0] == "Struct_2")
            {
                Boolean flag = IsOffsetApplied[0];
                IsOffsetApplied[0] = IsOffsetApplied[1];
                IsOffsetApplied[1] = flag;
            }

            return structurePort;
        }
        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();  
        ///</code>
        public Boolean[] GetIsLugEndOffsetApplied()
        {
            try
            {
                Collection<BusinessObject> StructureObjects;
                Boolean[] isOffsetApplied = new Boolean[2];

                //first route object is set as primary route object
                StructureObjects = SupportHelper.SupportingObjects;
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

                isOffsetApplied[0] = true;
                isOffsetApplied[1] = true;

                if (StructureObjects != null)
                {
                    if (StructureObjects.Count >= 2)
                    {
                        for (int index = 0; index <= 1; index++)
                        {
                            //check if offset is to be applied
                            const double RouteStructAngle = Math.PI / 180.0;
                            Boolean varRuleApplied = true;

                            if (supportingType == "Steel")
                            {
                                //if angle is within 1 degree, regard as parallel case
                                //Also check for Sloped structure                                
                                MemberPart memberPart = (MemberPart)SupportHelper.SupportingObjects[index];
                                ICurve memberCurve = memberPart.Axis;

                                Vector supportedVector = new Vector();
                                Vector supportingVector = new Vector();

                                if (SupportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                                {
                                    Position startLocation = new Position(SupportedHelper.SupportedObjectInfo(1).StartLocation);
                                    Position endLocation = new Position(SupportedHelper.SupportedObjectInfo(1).EndLocation);
                                    supportedVector = new Vector(endLocation - startLocation);
                                }
                                if (memberCurve is ILine)
                                {
                                    ILine line = (ILine)memberCurve;
                                    supportingVector = line.Direction;
                                }

                                double angle = GetAngleBetweenVectors(supportingVector, supportedVector);
                                double refAngle1 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Y, OrientationAlong.Global_X) - Math.PI / 2;
                                double refAngle2 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.X, OrientationAlong.Global_X);

                                if (angle < (refAngle1 + 0.001) && angle > (refAngle1 - 0.001))
                                    angle = angle - Math.Abs(refAngle1);
                                else if (angle < (refAngle2 + 0.001) && angle > (refAngle2 - 0.001))
                                    angle = angle - Math.Abs(refAngle2);
                                else
                                    angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                if (Math.Abs(angle) < RouteStructAngle || Math.Abs(angle - Math.PI) < RouteStructAngle)
                                    varRuleApplied = false;
                            }

                            isOffsetApplied[index] = varRuleApplied;
                        }
                    }
                }

                return isOffsetApplied;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied Method of Bline_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="vector1">The vector1(Vector).</param>
        /// <param name="vector2">The vector2(Vector).</param>
        /// <returns>double</returns>        
        /// <code>
        ///       ContentHelper contentHelper = new ContentHelper();
        ///       double value;
        ///       value = contentHelper. GetAngleBetweenVectors(vector1, vector2 );
        ///</code>
        public double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double arcCos = 0.0;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dotProd), 1) == false)
                {
                    arcCos = Math.PI / 2 - Math.Atan(dotProd / Math.Sqrt(1 - dotProd * dotProd));
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, -1) == true)
                {
                    arcCos = Math.PI;
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, 1) == true)
                {
                    arcCos = 0;
                }
                return arcCos;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}



