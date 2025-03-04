//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   BeamSupCTWChannel.cs
//   Bline_Assy,Ingr.SP3D.Content.Support.Rules.BeamSupCTWChannel
//   Author       :  Pavan
//   Creation Date:  17.Oct.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17.Oct.2012    Pavan    CR-CP-219114-Initial Creation
//   07.Sep.2015    PR       TR 277225	B-Line Hangers do not place correctly
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
    public class BeamSupCTWChannel : CustomSupportDefinition, ICustomHgrBOMDescription
    {
            private string ClampPartKeys;
            private string ChannelPartKeys;
            private string partNumber;
            private string clampType;
            private string refPortName;

            private double[] width = new double[5];
            private double[] height = new double[5];
            double[] clampLength = new double[5];
            double width_Inches;

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

                            int numOfRoutes = SupportHelper.SupportedObjects.Count;                         

                            //configuration definition of the assembly
                            //For Clamp
                            int clamp_Begin = 1;
                            int numOfPart = numOfRoutes;
                            int clamp_End = numOfPart;

                            clampType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyClampType", "ClampType")).PropValue;

                            //Get width and height of the Cabletray(s)  
                            for (int index = 1; index <= numOfRoutes; index++)
                            {
                                Array.Resize(ref width, numOfRoutes);
                                Array.Resize(ref height, numOfRoutes);

                                CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(index);
                                width[index - 1] = cabletrayInfo.Width;
                                height[index - 1] = cabletrayInfo.Depth;
                            }
                            for (int index = clamp_Begin; index <= clamp_End; index++)
                            {
                                width_Inches = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, width[index - 1], UnitName.DISTANCE_INCH);
                                ClampPartKeys = "Clamp" + index.ToString();
                                partNumber = clampType + "_" + width_Inches;
                                parts.Add(new PartInfo(ClampPartKeys, partNumber));
                            }
                            //For Channel
                            int channel_Begin = clamp_End + 1;
                            int channel_End = clamp_End + numOfPart;

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

                            for (int index = channel_Begin; index <= channel_End; index++)
                            {
                                ChannelPartKeys = "Channel" + index.ToString();
                                parts.Add(new PartInfo(ChannelPartKeys, channelPartNumber));
                            }
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
            public override void ConfigureSupport(Collection<SupportComponent> SupCompColl)
            {
               try
               {
                   Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                int numOfRoutes = SupportHelper.SupportedObjects.Count;
                
                //For Clamp
                int clamp_Begin = 1;
                int numOfPart = numOfRoutes;
                int clamp_End = numOfPart;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;                

                //Get width and height of the Cabletray(s)  
                for (int index = 1; index <= numOfRoutes; index++)
                {
                    Array.Resize(ref width, numOfRoutes);
                    Array.Resize(ref height, numOfRoutes);

                    CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(index);
                    width[index - 1] = cabletrayInfo.Width;
                    height[index - 1] = cabletrayInfo.Depth;
                }

                for (int index = clamp_Begin; index <= clamp_End; index++)
                {
                    if (index == clamp_Begin)
                    {
                        refPortName = "Route";
                    }
                    else
                    {
                        int refPortnum = index - clamp_Begin + 1;
                        refPortName = "Route_" + refPortnum;
                    }

                    ClampPartKeys = "Clamp" + index.ToString();

                    //Get Clamp Length of all the clamps
                    Array.Resize(ref clampLength, numOfRoutes);
                    SupportComponent clampPartOcc = componentDictionary[ClampPartKeys];
                    BusinessObject clampPart = clampPartOcc.GetRelationship("madeFrom", "part").TargetObjects[0];
                    clampLength[index - 1] = (double)((PropertyValueDouble)clampPart.GetPropertyValue("IJUAHgrBlineClampDim", "ClampLength")).PropValue;

                    PropertyValueCodelist insideOutsideCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                    int insideOutsideValue = (int)insideOutsideCodeList.PropValue;

                    PropertyValueCodelist clampGuideCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyClampGuide", "ClampGuide");
                    int clampGuideValue = (int)clampGuideCodeList.PropValue;

                    string hardwareValue = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyHardware", "Hardware")).PropValue;

                    clampPartOcc.SetPropertyValue(insideOutsideValue, "IJOAHgrBlineInsideOutside", "Inside_Outside");
                    clampPartOcc.SetPropertyValue(clampGuideValue, "IJOAHgrBlineClampGuide", "Clamp_Guide");
                    clampPartOcc.SetPropertyValue(hardwareValue, "IJUAHgrBlineWithHardware", "WithHardware");

                    //2340
                    JointHelper.CreateRigidJoint(ClampPartKeys, "TrayPort", "-1", refPortName, Plane.XY, Plane.XY, Axis.Y, Axis.Y, (1.05 * height[index - 1]/2), 0, 0);
                }

                //For channel
                int channel_Begin = clamp_End + 1;
                int channel_End = clamp_End + numOfPart;

                for (int index = channel_Begin; index <= channel_End; index++)
                {
                    ChannelPartKeys = "Channel" + index.ToString();

                    if (index == (channel_End + 1))
                    {
                        refPortName = "Route";
                    }
                    else
                    {
                        int refportnum = index - channel_Begin + 1;
                        refPortName = "Route_" + refportnum;
                    }

                    SupportComponent channelPartOcc = componentDictionary[ChannelPartKeys];

                    BusinessObject channelPart;
                    CrossSection CrossSection;

                    channelPart = channelPartOcc.GetRelationship("madeFrom", "part").TargetObjects[0];
                    CrossSection = (CrossSection)channelPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                    double[] steelDepth = new double[numOfRoutes];
                    steelDepth[index - numOfRoutes - 1] = (double)((PropertyValueDouble)CrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                    PropertyValueCodelist ChannelHolePatternCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChHolePattern", "ChannelHolePattern");
                    int channelHolePatternValue = (int)ChannelHolePatternCodeList.PropValue;

                    channelPartOcc.SetPropertyValue(channelHolePatternValue, "IJOAHgrBLineChannelB22", "HolePattern");
                    channelPartOcc.SetPropertyValue(channelMaterialCodelistValue, "IJOAHgrBLineChannelB22", "MaterialType");

                    double channelLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyChannelLength", "ChannelLength")).PropValue;

                    if (channelLength < (width[index - numOfRoutes - 1] + (4 * clampLength[index - numOfRoutes - 1])))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "The Channel length should be more than the sum of Tray width and four times Clamp length.", "", "BeamSupCTWChannel.cs", 1);
                        channelPartOcc. SetPropertyValue( width[index - numOfRoutes - 1]+ (4 * clampLength[index - numOfRoutes - 1]), "IJUAHgrOccLength", "Length");                       
                    }
                    else
                    {
                       channelPartOcc.SetPropertyValue(channelLength, "IJUAHgrOccLength", "Length");
                    }

                    int clampPart = index - numOfPart;
                    ClampPartKeys = "Clamp" + clampPart.ToString();
                    //2340
                    JointHelper.CreateRigidJoint(ChannelPartKeys, "Neutral", ClampPartKeys, "TrayPort", Plane.YZ, Plane.XY, Axis.Z, Axis.Y, steelDepth[index - numOfRoutes - 1] / 2.0, 0, 0);
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
                        int numOfRoutes = SupportHelper.SupportedObjects.Count;

                        //For Clamp
                        int clamp_Begin = 1;
                        int numOfPart = numOfRoutes;
                        int clamp_End = numOfPart;

                        for (int index = clamp_Begin; index <= clamp_End; index++)
                        {
                            ClampPartKeys = "Clamp" + index.ToString();
                            int connecttoroute = index - clamp_Begin + 1;
                            routeConnections.Add(new ConnectionInfo(ClampPartKeys, connecttoroute));
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
                        Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                        structConnections.Add(new ConnectionInfo("Clamp1", 1));
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
                    string clamptype = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAHgrAssyClampType", "ClampType")).PropValue;
                    string hardware = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAHgrAssyHardware", "Hardware")).PropValue;

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

                    BOMString = "B-Line Assy Beam Supported Cable Tray (With Channel Strut) with CT " + clampGuide + " type: " + clamptype + " mounted, " + insideOutside + ", " + hardware + " Hardware";

                    return BOMString;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in BOM of Assembly - BeamSupCTWChannel" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
            #endregion
    }
}







