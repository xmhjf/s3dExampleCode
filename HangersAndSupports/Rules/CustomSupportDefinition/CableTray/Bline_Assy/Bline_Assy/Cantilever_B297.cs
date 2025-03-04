//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Cantilever_B297.cs
//   Bline_Assy,Ingr.SP3D.Content.Support.Rules.Cantilever_B297
//   Author       :  Vijaya
//   Creation Date:  19.Oct.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19.Oct.2012    Vijaya   CR-CP-219114-Initial Creation
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
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
    public class Cantilever_B297 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string Channel = "Channel";
        private const string TopChannel = "TopChannel";
        private const string BotChannel = "BotChannel";
        private const string Plate = "Plate";

        private string channelType;
        private double channelLength;
        private double sectionChLength;
        private string channelMaterial;
        private string channelFinish;
        private double plateWidth;
        private double plateDepth;
        private int channelMaterialCodelistValue;
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

                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType");
                    long sectionSizeCodelistValue = (long)sectionSizeCodelist.PropValue;

                    channelLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyChannelLength", "A")).PropValue;
                    sectionChLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyChannelLength", "SecChLength")).PropValue;
                    PropertyValueString PlateSize = (PropertyValueString)support.GetPropertyValue("IJUAHgrBlineAssyPlateDim", "PlateSize");

                    PropertyValueCodelist channelMaterialCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial");
                    channelMaterialCodelistValue = (int)channelMaterialCodelist.PropValue;

                    PropertyValueCodelist channelFinishCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish");
                    long channelFinishCodelistValue = (long)channelFinishCodelist.PropValue;

                    plateWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrBlineAssyPlateDim", "Width")).PropValue;
                    plateDepth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrBlineAssyPlateDim", "Depth")).PropValue;

                    if (channelFinishCodelistValue == -1)
                    {
                        int defaultvalue = 1;
                        support.SetPropertyValue(defaultvalue, "IJOAHgrAssyB22Channel", "ChannelFinish");
                        channelFinishCodelistValue = (long)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish")).PropValue;
                    }

                    channelMaterial = channelMaterialCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)channelMaterialCodelistValue).DisplayName;
                    channelFinish = channelFinishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)channelFinishCodelistValue).DisplayName;
                    channelType = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)sectionSizeCodelistValue).DisplayName;

                    CommonBline_Assembly cmnBline_Assembly = new CommonBline_Assembly();
                    Dictionary<String, String> parameters = new Dictionary<string, string>();

                    if (channelMaterial.ToLower() == "steel")
                    {
                        parameters.Add("ChannelFinish", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish").ToString());
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial").ToString());
                        parameters.Add("ChannelType", support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType").ToString());
                    }
                    else
                    {
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial").ToString());
                        parameters.Add("ChannelType", support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType").ToString());
                    }

                    string ChannelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssySteelChAUX", "IJUAHgrAssySteelChAUX", "ChannelPartNo", parameters);

                    parts.Add(new PartInfo(Channel, ChannelPartNumber));
                    parts.Add(new PartInfo(TopChannel, ChannelPartNumber));
                    parts.Add(new PartInfo(BotChannel, ChannelPartNumber));
                    parts.Add(new PartInfo(Plate, "Utility_TWO_HOLE_PLATE_" + PlateSize.PropValue));

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
                SupportComponent channel = componentDictionary[Channel];
                SupportComponent Topchannel = componentDictionary[TopChannel];
                SupportComponent Botchannel = componentDictionary[BotChannel];
                SupportComponent BoltedPlate = componentDictionary[Plate];

                BusinessObject TopchannelPart;
                CrossSection CrossSection;

                TopchannelPart = Topchannel.GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection = (CrossSection)TopchannelPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = (double)((PropertyValueDouble)CrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)CrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double cableTrayDepth = cabletrayInfo.Depth;

                PropertyValueCodelist channelHolePatternCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChHolePattern", "ChannelHolePattern");
                int channelHolePatternValue = (int)channelHolePatternCodelist.PropValue;

                string channelInterface = "IJOAHgrBLineChannel" + channelType;

                //For Top Channel            
                Topchannel.SetPropertyValue(channelHolePatternValue, channelInterface, "HolePattern");
                Topchannel.SetPropertyValue(channelLength, "IJUAHgrOccLength", "Length");
                Topchannel.SetPropertyValue(channelMaterialCodelistValue, channelInterface, "MaterialType");

                //For Bottom Channel            
                Botchannel.SetPropertyValue(channelHolePatternValue, channelInterface, "HolePattern");
                Botchannel.SetPropertyValue(channelLength, "IJUAHgrOccLength", "Length");
                Botchannel.SetPropertyValue(channelMaterialCodelistValue, channelInterface, "MaterialType");

                //For Bolted Plate
                BoltedPlate.SetPropertyValue(plateWidth, "IJOAHgrUtility_TWO_HOLE_PLATE", "WIDTH");
                BoltedPlate.SetPropertyValue(plateDepth, "IJOAHgrUtility_TWO_HOLE_PLATE", "DEPTH");

                //For Vertical Channel            
                channel.SetPropertyValue(channelHolePatternValue, channelInterface, "HolePattern");
                channel.SetPropertyValue(channelMaterialCodelistValue, channelInterface, "MaterialType");

                if (sectionChLength < plateWidth)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "The channel length should be more than the Bolted Plate width. Resetting the length to Plate Width", "", "Cantilever_B297.cs", 1);
                    channel.SetPropertyValue(plateWidth, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    channel.SetPropertyValue(sectionChLength, "IJUAHgrOccLength", "Length");
                }

                double angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);
                double angle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Z, OrientationAlong.Global_Z);


                //Create the Joint between the Bounding Box Low and Struct Reference Ports
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType != SupportingObjectType.Member)
                    {
                        //9574
                        JointHelper.CreateRigidJoint(Channel, "Neutral", "-1", "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.X, -steelDepth / 2, -cableTrayDepth, 0);
                    }
                    else
                    {

                        if (System.Math.Abs(angle) < (System.Math.PI) / 2)
                            //10598
                            JointHelper.CreateRigidJoint(Channel, "Neutral", "-1", "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.Y, -steelDepth / 2, -cableTrayDepth, 0);
                        else
                            //2406
                            JointHelper.CreateRigidJoint(Channel, "Neutral", "-1", "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeY, -steelDepth / 2, -cableTrayDepth, 0);
                    }
                }
                else
                {

                    if (System.Math.Abs(angle) < (System.Math.PI) / 2)
                        //10598
                        JointHelper.CreateRigidJoint(Channel, "Neutral", "-1", "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.Y, -steelDepth / 2, -cableTrayDepth, 0);
                    else
                        //2406
                        JointHelper.CreateRigidJoint(Channel, "Neutral", "-1", "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeY, -steelDepth / 2, -cableTrayDepth, 0);
                }

                JointHelper.CreateRigidJoint(Plate, "TopStructure", Channel, "Neutral", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, -steelDepth / 2, 0);
                JointHelper.CreateRigidJoint(TopChannel, "BeginCap", Plate, "BotStructure", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, steelDepth / 2, 0);
                JointHelper.CreateRigidJoint(BotChannel, "BeginCap", Plate, "BotStructure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0, steelDepth / 2);
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
                    routeConnections.Add(new ConnectionInfo(Channel, 1));
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
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections.Add(new ConnectionInfo(Channel, 1));
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

        #region ICustomHgrBOMDescription Members
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string BOMString = "";
            try
            {
                int finishValue;
                int hardwareValue;

                PropertyValueCodelist finishCodelist;
                finishCodelist = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrAssyB297", "Finish");
                finishValue = finishCodelist.PropValue;

                PropertyValueCodelist hardwareCodelist;
                hardwareCodelist = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrAssyHardware", "Hardware");
                hardwareValue = hardwareCodelist.PropValue;

                if (finishValue == 0)
                    finishValue = 1;
                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishValue).DisplayName;

                if (hardwareValue == 0)
                    hardwareValue = 1;
                string hardware = hardwareCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(hardwareValue).DisplayName;

                //Get the CT Width and CT Height
                CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                double cableTrayWidth = cabletrayInfo.Width;
                cableTrayWidth = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, cableTrayWidth, UnitName.DISTANCE_INCH);

                BOMString = "B-Line Assy Cantilever Bracket, Type 297 for " + cableTrayWidth + " in Tray Width, " + "Finish: " + finish + ", " + hardware + " Hardware";
                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - Cantilever_B297" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}



