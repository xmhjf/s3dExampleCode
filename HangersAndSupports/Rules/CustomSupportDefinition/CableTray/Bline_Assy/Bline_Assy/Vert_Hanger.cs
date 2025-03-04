//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Vert_Hanger.cs
//   Bline_Assy,Ingr.SP3D.Content.Support.Rules.Vert_Hanger
//   Author       :  Vijaya
//   Creation Date:  29.Oct.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29.Oct.2012    Vijaya    CR-CP-219114-Initial Creation
//   07.Sep.2015    PR   TR 277225	B-Line Hangers do not place correctly 
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
    public class Vert_Hanger : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string Channel = "Channel";
        private const string LeftHangerRod = "LeftHangerRod";
        private const string RightHangerRod = "RightHangerRod";
        private const string SplicePlate = "SplicePlate";
        private const string Nut1 = "Nut1";
        private const string Nut2 = "Nut2";
        private const string Nut3 = "Nut3";
        private const string Nut4 = "Nut4";
        private const string Nut5 = "Nut5";
        private const string Nut6 = "Nut6";
        string channelType;
        string spliceType;
        string rodLength;
        string channelMaterial;
        int channelMaterialCodelistValue;
       
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
                    CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);                   
                
                    //Cionvert meters to inches
                                
                    double cableTrayWidth = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, cabletrayInfo.Width, UnitName.DISTANCE_INCH);
                    double cableTrayDepth = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, cabletrayInfo.Depth, UnitName.DISTANCE_INCH);   
                       
                   
                    PropertyValueCodelist channelTypeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType");
                    long channelTypeCodelistValue = (long)channelTypeCodelist.PropValue;
                    spliceType = support.GetPropertyValue("IJUAHgrAssyVertHanger", "SplicePlate").ToString();

                    PropertyValueCodelist availRodLengthCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyRodLength", "RodLength");
                    int rodLengthCodelist = (int)availRodLengthCodelist.PropValue;

                    PropertyValueCodelist channelMaterialCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial");
                    channelMaterialCodelistValue = (int)channelMaterialCodelist.PropValue;

                    PropertyValueCodelist channelFinishCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish");
                    long channelFinishCodelistValue = (long)channelFinishCodelist.PropValue;

                    if (channelFinishCodelistValue == -1)
                    {
                        int defaultvalue = 1;
                        support.SetPropertyValue(defaultvalue, "IJOAHgrAssyB22Channel", "ChannelFinish");
                        channelFinishCodelistValue = (long)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish")).PropValue;
                    }

                    rodLength = availRodLengthCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)rodLengthCodelist).DisplayName;
                    channelMaterial = channelMaterialCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)channelMaterialCodelistValue).DisplayName;                    
                    channelType = channelTypeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)channelTypeCodelistValue).DisplayName;                   

                    CommonBline_Assembly cmnBline_Assembly = new CommonBline_Assembly();
                    Dictionary<String,String> parameters =  new Dictionary<string,string>();
                    string channelPartNumber;
                    if (channelMaterial.ToLower() == "steel")
                    {
                        parameters.Add("ChannelFinish", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish").ToString());
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial").ToString());
                        parameters.Add("ChannelType", support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType").ToString());
                        channelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssySteelChAUX", "IJUAHgrAssySteelChAUX", "ChannelPartNo", parameters);
                    }
                    else
                    {
                        
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial").ToString());
                        parameters.Add("ChannelType", support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType").ToString());
                        channelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssyChannelsAUX", "IJUAHgrAssyChannelsAUX", "ChannelPartNo", parameters);                     
                    }
                    
                    string splicePlate=spliceType + Math.Round(cableTrayDepth)+"_"+ Math.Round(cableTrayWidth)+"_"+Math.Round(cableTrayDepth);
                    parts.Add(new PartInfo(Channel, channelPartNumber));
                    parts.Add(new PartInfo(SplicePlate, splicePlate));                   
                    parts.Add(new PartInfo(LeftHangerRod, "ATR_0.5_" + rodLength));
                    parts.Add(new PartInfo(RightHangerRod, "ATR_0.5_" + rodLength));
                    parts.Add(new PartInfo(Nut1, "Anvil_HEX_NUT_3"));
                    parts.Add(new PartInfo(Nut2, "Anvil_HEX_NUT_3"));
                    parts.Add(new PartInfo(Nut3, "Anvil_HEX_NUT_3"));
                    parts.Add(new PartInfo(Nut4, "Anvil_HEX_NUT_3"));
                    parts.Add(new PartInfo(Nut5, "Anvil_HEX_NUT_3"));
                    parts.Add(new PartInfo(Nut6, "Anvil_HEX_NUT_3"));             
                   
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
                SupportComponent splicePlate = componentDictionary[SplicePlate];
                SupportComponent leftHangerRod = componentDictionary[LeftHangerRod];
                SupportComponent rightHangerRod = componentDictionary[RightHangerRod];
                SupportComponent nut1 = componentDictionary[Nut1];
                SupportComponent nut2 = componentDictionary[Nut2];
                SupportComponent nut3 = componentDictionary[Nut3];
                SupportComponent nut4 = componentDictionary[Nut4];
                SupportComponent nut5 = componentDictionary[Nut5];
                SupportComponent nut6 = componentDictionary[Nut6];

                BusinessObject ChannelPart;
                CrossSection CrossSection;

                ChannelPart = channel.GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection = (CrossSection)ChannelPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = (double)((PropertyValueDouble)CrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)CrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;

                BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double cableTrayWidth = BBX.Width;
                double cableTrayDepth = BBX.Height;

                // ========================================
                // Set Values of Part Occurance Attributes
                // ========================================
                // For SplicePlate
                BusinessObject splicePlatepart = splicePlate.GetRelationship("madeFrom", "part").TargetObjects[0];
                double TH = (double)((PropertyValueDouble)splicePlatepart.GetPropertyValue("IJUABlineVertHanger", "TH")).PropValue;
                double HH = (double)((PropertyValueDouble)splicePlatepart.GetPropertyValue("IJUABlineVertHanger", "HH")).PropValue;
                double LS = (double)((PropertyValueDouble)splicePlatepart.GetPropertyValue("IJUABlineVertHanger", "LS")).PropValue;
                double LP = (double)((PropertyValueDouble)splicePlatepart.GetPropertyValue("IJUABlineVertHanger", "LP")).PropValue;
                double WT = (double)((PropertyValueDouble)splicePlatepart.GetPropertyValue("IJUAHgrBlineTrayWidth", "WT")).PropValue;
                double trayWidth = (double)((PropertyValueDouble)splicePlatepart.GetPropertyValue("IJUAHgrBlineTrayWebThick", "TrayWT")).PropValue;

                double portsDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                double rodLength = portsDistance - steelDepth - HH / 2;

                BusinessObject nutpart = nut1.GetRelationship("madeFrom", "part").TargetObjects[0];
                double nut_T = (double)((PropertyValueDouble)nutpart.GetPropertyValue("IJUAHgrAnvil_hex_nut", "T")).PropValue;

                //For both rods
                leftHangerRod.SetPropertyValue(rodLength, "IJUAHgrOccLength", "Length");
                rightHangerRod.SetPropertyValue(rodLength, "IJUAHgrOccLength", "Length");

                //For Channel
                PropertyValueCodelist channelHolePatternCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChHolePattern", "ChannelHolePattern");
                int channelHolePatternValue = (int)channelHolePatternCodelist.PropValue;

                string channelInterface = "IJOAHgrBLineChannel" + channelType;

                channel.SetPropertyValue(cableTrayWidth + 2 * LP, "IJUAHgrOccLength", "Length");
                channel.SetPropertyValue(channelMaterialCodelistValue, channelInterface, "MaterialType");
                channel.SetPropertyValue(channelHolePatternValue, channelInterface, "HolePattern");

                //Create Joints 

                //Add joints between SplicePlate and Route
                JointHelper.CreateRevoluteJoint(SplicePlate, "TrayPort", "-1", "Route", Axis.Y, Axis.Y);
                JointHelper.CreateGlobalAxesAlignedJoint(SplicePlate, "TrayPort", Axis.X, Axis.Z);

                //Add a Vertical Joint to the Rod Z axis - LeftHangerRod
                JointHelper.CreateGlobalAxesAlignedJoint(LeftHangerRod, "TopExThdRH", Axis.Z, Axis.Z);

                //Create the Prsimatic Joint between top and bottom ports of the left hanger rod
                JointHelper.CreatePrismaticJoint(LeftHangerRod, "TopExThdRH", LeftHangerRod, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                //Create the Flexible (Prismatic) Joint between SplicePlate and rod
                JointHelper.CreateRevoluteJoint(SplicePlate, "RodHole1", LeftHangerRod, "BotExThdRH", Axis.Y, Axis.NegativeY);

                //Add a Vertical Joint to the Rod Z axis - RightHangerRod
                JointHelper.CreateGlobalAxesAlignedJoint(RightHangerRod, "TopExThdRH", Axis.Z, Axis.Z);

                //Create the Prsimatic Joint between top and bottom ports of the right hanger rod
                JointHelper.CreatePrismaticJoint(RightHangerRod, "TopExThdRH", RightHangerRod, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                //Create the Flexible (Prismatic) Joint between SplicePlate and rod
                JointHelper.CreateRevoluteJoint(SplicePlate, "RodHole2", RightHangerRod, "BotExThdRH", Axis.Y, Axis.NegativeY);

                //Add a Rigid Joint between the bottom nut and the rod
                JointHelper.CreateRigidJoint(LeftHangerRod, "BotExThdRH", Nut1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.008 - TH / 2, 0, 0);
                JointHelper.CreateRigidJoint(LeftHangerRod, "BotExThdRH", Nut2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, nut_T + TH / 2 - 0.008, 0, 0);
                JointHelper.CreateRigidJoint(RightHangerRod, "BotExThdRH", Nut3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.008 - TH / 2, 0, 0);
                JointHelper.CreateRigidJoint(RightHangerRod, "BotExThdRH", Nut4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, nut_T + TH / 2 - 0.008, 0, 0);

                //Add Jointa between the Top nut and the rod
                JointHelper.CreateRigidJoint(LeftHangerRod, "TopExThdRH", Nut5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(RightHangerRod, "TopExThdRH", Nut6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                //Add Joints between the Rod and the channel
                JointHelper.CreateRigidJoint(Channel, "Neutral", LeftHangerRod, "TopExThdRH", Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeX, steelDepth / 2, -(0.5 * WT + trayWidth + LS), 0);
                JointHelper.CreatePrismaticJoint(Channel, "Neutral", RightHangerRod, "TopExThdRH", Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.Y, steelDepth / 2, 0);

                //Add joints between the Channel and the structure                
                JointHelper.CreatePointOnPlaneJoint(Channel, "BeginCap", "-1", "Structure", Plane.XY);
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
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string BOMString = "";
            try
            {
                string splicePlate;              
                string rodSize;               

                splicePlate = supportOrComponent.GetPropertyValue("IJUAHgrAssyVertHanger", "SplicePlate").ToString();
                rodSize = supportOrComponent.GetPropertyValue("IJUAHgrAssyVertHanger", "RodSize").ToString();             

                //Get the CT Width and CT Height
                CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);

               
                double cableTrayWidth =  cabletrayInfo.Width ;
                double cableTrayHeight = cabletrayInfo.Depth ;
                cableTrayWidth = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, cableTrayWidth, UnitName.DISTANCE_INCH);
                cableTrayHeight = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, cableTrayHeight, UnitName.DISTANCE_INCH);


                BOMString = "B-Line Assy Vertical Hanger for " + cableTrayHeight + " in Tray Height, " + cableTrayWidth + " in Traywidth, with Splice Plate Type: " + splicePlate + " and with Rods of size: " + rodSize;
                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - Vert_Hanger" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}





