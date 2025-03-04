//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bracket_B494.cs
//   Bline_Assy,Ingr.SP3D.Content.Support.Rules.Bracket_B494
//   Author       :  Vijaya
//   Creation Date:  25.Oct.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   25.Oct.2012    Vijaya  CR-CP-219114-Initial Creation
//   07.Sep.2015    PR      TR 277225	B-Line Hangers do not place correctly 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   04/07/2016     Siva    DM-CP-296666 	B-Line Trapeze Assembly Support doesn't place on cableway turn feature. 
//   20/07/2016     Siva    TR-CP-298361  Fix the new coveirty issues reported on H&S Content in July Coverity report  
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
    public class Bracket_B494 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string Channel1 = "Channel1";
        private const string Channel2 = "Channel2";
        private const string Brace = "Brace";
        private const string ANGLECON1 = "ANGLECON1";
        private const string ANGLECON2 = "ANGLECON2";
        private const string BoltedPlate = "BoltedPlate";

        string channelType;
        double channelLength;
        double sectionChLength;
        string channelMaterial;

        double plateWidth;
        double plateDepth;
        int channelMaterialCodelistValue;
        int currentFaceNumber;
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

                    channelLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyChannelLength", "A")).PropValue;
                    sectionChLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyChannelLength", "SecChLength")).PropValue;

                    PropertyValueString plateSize = (PropertyValueString)support.GetPropertyValue("IJUAHgrBlineAssyPlateDim", "PlateSize");

                    PropertyValueCodelist channelMaterialCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB42Channel", "ChannelMaterial");
                    channelMaterialCodelistValue = (int)channelMaterialCodelist.PropValue;

                    PropertyValueCodelist channelFinishCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB42Channel", "ChannelFinish");
                    long channelFinishCodelistValue = (long)channelFinishCodelist.PropValue;

                    plateWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrBlineAssyPlateDim", "Width")).PropValue;
                    plateDepth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrBlineAssyPlateDim", "Depth")).PropValue;

                    if (channelFinishCodelistValue == -1)
                    {
                        int defaultvalue = 1;
                        support.SetPropertyValue(defaultvalue, "IJOAHgrAssyB42Channel", "ChannelFinish");
                        channelFinishCodelistValue = (long)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB42Channel", "ChannelFinish")).PropValue;
                    }

                    channelMaterial = channelMaterialCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)channelMaterialCodelistValue).DisplayName;
                    channelType = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)sectionSizeCodelist.PropValue).DisplayName;

                    CommonBline_Assembly cmnBline_Assembly = new CommonBline_Assembly();
                    Dictionary<String, String> parameters = new Dictionary<string, string>();
                    string channelPartNumber;
                    if (channelMaterial.ToLower() == "steel")
                    {
                        parameters.Add("ChannelFinish", support.GetPropertyValue("IJOAHgrAssyB42Channel", "ChannelFinish").ToString());
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyB42Channel", "ChannelMaterial").ToString());
                        parameters.Add("ChannelType", support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType").ToString());
                        channelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssySteelChAUX", "IJUAHgrAssySteelChAUX", "ChannelPartNo", parameters);
                    }
                    else
                    {
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyB42Channel", "ChannelMaterial").ToString());
                        parameters.Add("ChannelType", support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType").ToString());
                        channelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssyChannelsAUX", "IJUAHgrAssyChannelsAUX", "ChannelPartNo", parameters);
                    }
                    parts.Add(new PartInfo(Channel1, channelPartNumber));
                    parts.Add(new PartInfo(Channel2, channelPartNumber));
                    parts.Add(new PartInfo(Brace, "Utility_USER_VARIABLE_BOX_1"));
                    parts.Add(new PartInfo(ANGLECON1, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(ANGLECON2, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(BoltedPlate, "Utility_TWO_HOLE_PLATE_" + plateSize.PropValue));

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
                SupportComponent channel1 = componentDictionary[Channel1];
                SupportComponent channel2 = componentDictionary[Channel2];
                SupportComponent brace = componentDictionary[Brace];
                SupportComponent angleCon1 = componentDictionary[ANGLECON1];
                SupportComponent angleCon2 = componentDictionary[ANGLECON2];
                SupportComponent boltedPlate = componentDictionary[BoltedPlate];
                BusinessObject ChannelPart;
                CrossSection CrossSection;
                //Get SteelWidth and SteelDepth the channel
                ChannelPart = channel1.GetRelationship("madeFrom", "part").TargetObjects[0];
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
                // =========================================
                // For SplicePlate
                BusinessObject boltedPlatePart = boltedPlate.GetRelationship("madeFrom", "part").TargetObjects[0];
                double boltedPlateHeight = (double)((PropertyValueDouble)boltedPlatePart.GetPropertyValue("IJOAHgrUtility_TWO_HOLE_PLATE", "DEPTH")).PropValue;
                double C = (double)((PropertyValueDouble)boltedPlatePart.GetPropertyValue("IJOAHgrUtility_TWO_HOLE_PLATE", "C")).PropValue;
                double boltedPlateThickness = (double)((PropertyValueDouble)boltedPlatePart.GetPropertyValue("IJUAHgrUtility_TWO_HOLE_PLATE", "THICKNESS")).PropValue;

                BusinessObject bracePart = brace.GetRelationship("madeFrom", "part").TargetObjects[0];
                double braceDepth = (double)((PropertyValueDouble)bracePart.GetPropertyValue("IJOAHgrUtility_VARIABLE_BOX", "DEPTH")).PropValue;

                double braceLength = Math.Sqrt(((boltedPlateHeight - C - steelDepth) * (boltedPlateHeight - C - steelDepth)) + ((channelLength - steelDepth) * (channelLength - steelDepth)));

                PropertyValueCodelist channelHolePatternCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChHolePattern", "ChannelHolePattern");
                int channelHolePatternValue = (int)channelHolePatternCodelist.PropValue;

                string channelInterface = "IJOAHgrBLineChannel" + channelType;

                //For Vertical Channel
                channel1.SetPropertyValue(channelHolePatternValue, channelInterface, "HolePattern");
                channel1.SetPropertyValue(channelMaterialCodelistValue, channelInterface, "MaterialType");
                channel1.SetPropertyValue(channelLength, "IJUAHgrOccLength", "Length");

                //For Horizontal Channel
                channel2.SetPropertyValue(channelHolePatternValue, channelInterface, "HolePattern");
                channel2.SetPropertyValue(channelMaterialCodelistValue, channelInterface, "MaterialType");

                if (sectionChLength < plateWidth)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "The channel length should be more than the Bolted Plate width. Resetting the length to Plate Width", "", "Cantilever_B409.cs", 1);
                    channel2.SetPropertyValue(plateWidth, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    channel2.SetPropertyValue(sectionChLength, "IJUAHgrOccLength", "Length");
                }
                // For Brace
                brace.SetPropertyValue(steelWidth, "IJOAHgrUtility_VARIABLE_BOX", "WIDTH");
                brace.SetPropertyValue(braceLength, "IJUAHgrOccLength", "Length");
                brace.SetPropertyValue(0.01, "IJOAHgrUtility_VARIABLE_BOX", "DEPTH");

                //For Bolted Plate
                boltedPlate.SetPropertyValue(plateWidth, "IJOAHgrUtility_TWO_HOLE_PLATE", "WIDTH");
                boltedPlate.SetPropertyValue(plateDepth, "IJOAHgrUtility_TWO_HOLE_PLATE", "DEPTH");


                //Get the current face number          

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType != SupportingObjectType.Member)
                        currentFaceNumber = 513;
                    else
                        currentFaceNumber = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                }
                else
                    currentFaceNumber = 513;

                //Create Joints
                double angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);
                double angle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Z, OrientationAlong.Global_Z);

                Plane CfgIdxPlaneA, CfgIdxPlaneB;
                CfgIdxPlaneA = CfgIdxPlaneB = Plane.XY;

                Axis CfgIdxAxisA, CfgIdxAxisB, Plate1CfgAxisA, Plate1CfgAxisB, Plate2CfgAxisA, Plate2CfgAxisB;
                CfgIdxAxisA = CfgIdxAxisB = Plate1CfgAxisA = Plate1CfgAxisB = Plate2CfgAxisA = Plate2CfgAxisB = Axis.X;

                if ((SupportHelper.SupportingObjects.Count != 0))
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Wall) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab))
                    {
                        if (Configuration == 1)
                        {
                            CfgIdxPlaneA = Plane.YZ;
                            CfgIdxPlaneB = Plane.XY;
                            CfgIdxAxisA = Axis.Y;
                            CfgIdxAxisB = Axis.X;
                        }
                        else
                        {
                            CfgIdxPlaneA = Plane.YZ;
                            CfgIdxPlaneB = Plane.XY;
                            CfgIdxAxisA = Axis.Y;
                            CfgIdxAxisB = Axis.NegativeX;
                        }

                        Plate1CfgAxisA = Axis.Y;
                        Plate1CfgAxisB = Axis.X;
                        Plate2CfgAxisA = Axis.Y;
                        Plate2CfgAxisB = Axis.NegativeX;
                    }
                    else if (System.Math.Abs(angle) < (System.Math.PI) / 2)
                    {
                        if (currentFaceNumber == 513)
                        {
                            if (Configuration == 1)
                            {
                                CfgIdxPlaneA = Plane.YZ;
                                CfgIdxPlaneB = Plane.XY;
                                CfgIdxAxisA = Axis.Y;
                                CfgIdxAxisB = Axis.NegativeY;
                            }
                            else
                            {
                                CfgIdxPlaneA = Plane.YZ;
                                CfgIdxPlaneB = Plane.XY;
                                CfgIdxAxisA = Axis.Y;
                                CfgIdxAxisB = Axis.Y;
                            }
                        }
                        else
                        {
                            CfgIdxPlaneA = Plane.YZ;
                            CfgIdxPlaneB = Plane.XY;
                            CfgIdxAxisA = Axis.Y;
                            CfgIdxAxisB = Axis.Y;
                        }
                        Plate1CfgAxisA = Axis.Y;
                        Plate1CfgAxisB = Axis.Y;
                        Plate2CfgAxisA = Axis.Y;
                        Plate2CfgAxisB = Axis.NegativeY;
                    }
                    else
                    {
                        if (currentFaceNumber == 513)
                        {
                            if (Configuration == 1)
                            {
                                CfgIdxPlaneA = Plane.YZ;
                                CfgIdxPlaneB = Plane.XY;
                                CfgIdxAxisA = Axis.Y;
                                CfgIdxAxisB = Axis.Y;
                            }
                            else
                            {
                                CfgIdxPlaneA = Plane.YZ;
                                CfgIdxPlaneB = Plane.XY;
                                CfgIdxAxisA = Axis.Y;
                                CfgIdxAxisB = Axis.NegativeY;
                            }
                        }
                        else
                        {
                            CfgIdxPlaneA = Plane.YZ;
                            CfgIdxPlaneB = Plane.XY;
                            CfgIdxAxisA = Axis.Y;
                            CfgIdxAxisB = Axis.NegativeY;
                        }
                        Plate1CfgAxisA = Axis.Y;
                        Plate1CfgAxisB = Axis.NegativeY;
                        Plate2CfgAxisA = Axis.Y;
                        Plate2CfgAxisB = Axis.Y;
                    }

                //Create the Joint between the Bounding Box Low and Struct Reference Ports     
                JointHelper.CreateRigidJoint(Channel2, "Neutral", "-1", "Structure", CfgIdxPlaneA, CfgIdxPlaneB, CfgIdxAxisA, CfgIdxAxisB, -steelDepth / 2 - boltedPlateThickness, -cableTrayDepth / 2 - steelDepth - steelWidth / 2, 0);

                //Add joints between bolted plate and the vertical channel
                JointHelper.CreateRigidJoint(BoltedPlate, "TopStructure", Channel2, "Neutral", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, -steelDepth / 2 - boltedPlateThickness, 0);

                // Add joints between horizontal channel and bolted plate
                JointHelper.CreateRigidJoint(Channel1, "BeginCap", BoltedPlate, "BotStructure", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, steelDepth / 2, -steelDepth);

                //Add Connection for the start of the Angled plate 9444
                JointHelper.CreateRigidJoint(ANGLECON1, "Connection", Channel1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, steelDepth, -steelDepth / 2, 0);

                //Add Connection for Angled plate and connection
                JointHelper.CreateRevoluteJoint(ANGLECON1, "Connection", Brace, "StartOther", Plate1CfgAxisA, Plate1CfgAxisB);

                //Flexible Member
                JointHelper.CreatePrismaticJoint(Brace, "StartOther", Brace, "EndOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add joints between Connection and Bolted Plate9572
                JointHelper.CreateRigidJoint(ANGLECON2, "Connection", BoltedPlate, "BotStructure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, 0, boltedPlateHeight, 0);

                //Add Connection for Angled plate and connection
                JointHelper.CreateCylindricalJoint(ANGLECON2, "Connection", Brace, "EndOther", Plate2CfgAxisA, Plate2CfgAxisB, 0);

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
                    routeConnections.Add(new ConnectionInfo(Channel1, 1));
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
                    structConnections.Add(new ConnectionInfo(Channel1, 1));
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
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string BOMString = "";
            try
            {
                int finishValue;
                int hardwareValue;

                PropertyValueCodelist finishCodelist;
                finishCodelist = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrAssyB494", "Finish");
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
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;

                BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double cableTrayWidth = BBX.Width;
                double cableTrayDepth = BBX.Height;

                cableTrayWidth = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, cableTrayWidth, UnitName.DISTANCE_INCH);

                BOMString = "B-Line Assy Cantilever Bracket, Type 494 for " + cableTrayWidth + " in Tray Width, " + "Finish: " + finish + ", " + hardware + " Hardware";
                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - Cantilever_B494_BOM" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}