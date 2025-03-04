//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Vert_Support.cs
//   Bline_Assy,Ingr.SP3D.Content.Support.Rules.Vert_Support
//   Author       :  Pavan
//   Creation Date:  24.Oct.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   24.Oct.2012    Pavan    CR-CP-219114-Initial Creation
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
    public class Vert_Support : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string Channel = "Channel";
        private const string Clamp_Guide = "Clamp_Guide";        
        private string clampPartNumber;
        private string clampType;
        private double width;       
        private string channelPartNumber;
        private string channelType;
        private string channelMaterial;
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
                    
                    clampType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyClampType", "ClampType")).PropValue;

                    //Get width and height of the Cabletray(s)  
                    CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    width = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, cabletrayInfo.Width, UnitName.DISTANCE_INCH);

                    clampPartNumber = clampType + "_" + width;

                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType");                  

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

                    channelMaterial = channelMaterialCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)channelMaterialCodelistValue).DisplayName;
                    channelType = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)sectionSizeCodelist.PropValue).DisplayName;

                    support.SetPropertyValue(channelMaterialCodelistValue, "IJOAHgrAssyB22Channel", "ChannelFinish");

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

                    channelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssySteelChAUX", "IJUAHgrAssySteelChAUX", "ChannelPartNo", parameters);

                    parts.Add(new PartInfo(Channel, channelPartNumber));
                    parts.Add(new PartInfo(Clamp_Guide, clampPartNumber));

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
                SupportComponent clampguide = componentDictionary[Clamp_Guide];

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
                // For Clamp
                BusinessObject clampguidecatpart = clampguide.GetRelationship("madeFrom", "part").TargetObjects[0];
                double clampHeight = (double)((PropertyValueDouble)clampguidecatpart.GetPropertyValue("IJUAHgrBlineClampDim", "ClampHeight")).PropValue;
                double clampLength = (double)((PropertyValueDouble)clampguidecatpart.GetPropertyValue("IJUAHgrBlineClampDim", "ClampLength")).PropValue;

                PropertyValueCodelist inorOutCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                int inorOutValue = (int)inorOutCodeList.PropValue;

                PropertyValueCodelist clampCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyClampGuide", "ClampGuide");
                int clampValue = (int)clampCodeList.PropValue;

                string hardwareValue = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyHardware", "Hardware")).PropValue;

                clampguide.SetPropertyValue(inorOutValue, "IJOAHgrBlineInsideOutside", "Inside_Outside");
                clampguide.SetPropertyValue(hardwareValue, "IJUAHgrBlineWithHardware", "WithHardware");
                clampguide.SetPropertyValue(clampValue, "IJOAHgrBlineClampGuide", "Clamp_Guide");

                // For Channel
                PropertyValueCodelist channelHolePatternCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChHolePattern", "ChannelHolePattern");
                int channelHolePatternValue = (int)channelHolePatternCodelist.PropValue;

                PropertyValueCodelist channelMaterialCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial");
                int channelMaterialCodelistValue = (int)channelMaterialCodelist.PropValue;

                channel.SetPropertyValue(channelHolePatternValue, "IJOAHgrBLineChannelB22", "HolePattern");
                channel.SetPropertyValue(channelMaterialCodelistValue, "IJOAHgrBLineChannelB22", "MaterialType");
                channel.SetPropertyValue((cableTrayWidth + (4 * clampLength)), "IJUAHgrOccLength", "Length");

                JointHelper.CreateRigidJoint(Channel, "Neutral", Clamp_Guide, "TrayPort", Plane.YZ, Plane.XY, Axis.Z, Axis.Y, steelDepth / 2.0, 0, 0);
                JointHelper.CreateRigidJoint(Clamp_Guide, "TrayPort", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.NegativeY, cableTrayDepth / 2.0, 0, 0);
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
                string clampType = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAHgrAssyClampType", "ClampType")).PropValue;
                string Hardware = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAHgrAssyHardware", "Hardware")).PropValue;

                PropertyValueCodelist insideOutsideCodeList = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                int insideOutsideValue = (int)insideOutsideCodeList.PropValue;

                if (insideOutsideValue == 0)
                    insideOutsideValue = 1;

                string insideOutside = insideOutsideCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(insideOutsideValue).DisplayName;

                PropertyValueCodelist clampGuideCodeList = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrAssyClampGuide", "ClampGuide");
                int clampGuideValue = (int)clampGuideCodeList.PropValue;

                if (clampGuideValue == 0)
                    clampGuideValue = 1;

                string clampGuide = clampGuideCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(clampGuideValue).DisplayName;

                //Get the CT Width and CT Height
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;

                BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double cableTrayWidth = BBX.Width;
                

                cableTrayWidth = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, cableTrayWidth, UnitName.DISTANCE_INCH);

                BOMString = "B-Line Assy Vertical Structural Support for " + cableTrayWidth + " in Tray Width, with CT " + clampGuide + " type: " + clampType + " mounted, " + insideOutside + ", " + Hardware + " Hardware";

                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - Vert_Support" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}





