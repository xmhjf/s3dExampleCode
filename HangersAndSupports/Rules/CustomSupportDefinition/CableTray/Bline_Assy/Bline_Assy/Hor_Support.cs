
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Hor_Support.cs
//   Bline_Assy,Ingr.SP3D.Content.Support.Rules.Hor_Support
//   Author       :  Vijaya
//   Creation Date:  1.Nov.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   1.Nov.2012     Vijaya   CR-CP-219114-Initial Creation
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
    public class Hor_Support : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private string[] ClampPartKeys;
        private string[] ChannelPartKeys;
        private string[] ConnetionPartKeys;
        private string partkey;
        private string partNumber;
        int clamp_Begin;
        int clamp_End;
        int channel_Begin;
        int channel_End;
        int connection_Begin;
        int connection_End;
        int numOfPart;
        int numOfRoutes;
        int channelIndex;
        int connectionIndex;
        int clampIndex;
        double[] width;
        double[] height;
        int currentFaceNumber;
        double width_Inches;
        Collection<PartInfo> parts;
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
                    parts = new Collection<PartInfo>();
                    numOfRoutes = SupportHelper.SupportedObjects.Count;

                    //For  Clamp
                    clamp_Begin = 1;
                    numOfPart = numOfRoutes;
                    clamp_End = numOfPart;

                    //For Channel
                    channel_Begin = clamp_End + 1;
                    channel_End = clamp_End + numOfRoutes;

                    //For Connection Object
                    connection_Begin = channel_End + 1;
                    connection_End = channel_End + numOfRoutes;

                    //Get the occurrence attributes values
                    string clampType = support.GetPropertyValue("IJUAHgrAssyClampType", "ClampType").ToString();

                    PropertyValueCodelist ChannelMaterialCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial");
                    string ChannelMaterial = ChannelMaterialCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)ChannelMaterialCodelist.PropValue).DisplayName;

                    PropertyValueCodelist ChannelFinishValueCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish");
                    long channelFinishCodelistValue = (long)ChannelFinishValueCodelist.PropValue;

                    if (channelFinishCodelistValue == -1)
                    {
                        int defaultvalue = 1;
                        support.SetPropertyValue(defaultvalue, "IJOAHgrAssyB22Channel", "ChannelFinish");
                        channelFinishCodelistValue = (long)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish")).PropValue;
                    }

                    CommonBline_Assembly cmnBline_Assembly = new CommonBline_Assembly();
                    Dictionary<String, String> parameters = new Dictionary<string, string>();
                    string ChannelPartNumber;

                    //Finish exists only when the channel is steel
                    if (ChannelMaterial.ToLower() == "steel")
                    {
                        parameters.Add("ChannelFinish", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelFinish").ToString());
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial").ToString());
                        parameters.Add("ChannelType", support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType").ToString());
                        ChannelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssySteelChAUX", "IJUAHgrAssySteelChAUX", "ChannelPartNo", parameters);
                    }
                    else
                    {
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial").ToString());
                        parameters.Add("ChannelType", support.GetPropertyValue("IJUAHgrAssyChannelType", "ChannelType").ToString());
                        ChannelPartNumber = cmnBline_Assembly.GetPartFromTable("BLineAssyChannelsAUX", "IJUAHgrAssyChannelsAUX", "ChannelPartNo", parameters);
                    }

                    //Get width and height of the Cabletray(s) 
                    width = new double[numOfRoutes];
                    height = new double[numOfRoutes];
                    ClampPartKeys = new string[numOfRoutes];
                    ChannelPartKeys = new string[numOfRoutes];
                    ConnetionPartKeys = new string[numOfRoutes];
             

                    for (int index = 0; index < numOfRoutes; index++)
                    {
                        CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(index + 1);
                        width[index] = cabletrayInfo.Width;
                        height[index] = cabletrayInfo.Depth;
                    }
                    for (int clampIndex = clamp_Begin; clampIndex <= clamp_End; clampIndex++)
                    {
                       
                        width_Inches = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, width[clampIndex - 1], UnitName.DISTANCE_INCH);
                        ClampPartKeys[clampIndex - 1] = "Clamp" + clampIndex.ToString();
                        partNumber = clampType + "_" + Math.Round(width_Inches);
                        parts.Add(new PartInfo(ClampPartKeys[clampIndex - 1], partNumber));
                    }
                    for (int channelIndex = channel_Begin; channelIndex <= channel_End; channelIndex++)
                    {
                        int index = channelIndex - clamp_End - 1;
                        ChannelPartKeys[index] = "Channel" + (channelIndex - clamp_End).ToString();
                        parts.Add(new PartInfo(ChannelPartKeys[index], ChannelPartNumber));
                    }

                    for (int connectionIndex = connection_Begin; connectionIndex <= connection_End; connectionIndex++)
                    {
                        int index = connectionIndex - channel_End - 1;
                        ConnetionPartKeys[index] = "Connection" + (connectionIndex - channel_End).ToString();
                        parts.Add(new PartInfo(ConnetionPartKeys[index], "Log_Conn_Part_1"));
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
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //Get Port Locations
                string hangerPortName;
                string refPortName;
                Plane ConfRoutePlaneA, ConfRoutePlaneB;
                Axis ConfRouteAxisA, ConfRouteAxisB, ConfConnAxisA, ConfConnAxisB;

                int confConnIndex;
                int confRouteIndex;

                double routePlaneOffset;
                double routeAxisOffset;
                double routeOriginOffset;

                double connPlaneOffset;
                double connAxisOffset;

                double connRouteAxisOffset;
                double connRouteOriginOffset;

                double[] portsDistance = new double[numOfRoutes];
                Matrix4X4[] portOrientation = new Matrix4X4[numOfRoutes];

                //BusinessObject HangerPort;
                for (int index = 1; index <= numOfRoutes; index++)
                {
                    if (index == 1)
                    {
                        hangerPortName = "Route";
                        portOrientation[index - 1] = RefPortHelper.PortLCS("Route");
                    }
                    else
                    {
                        hangerPortName = "Route_" + index;
                        portOrientation[index - 1] = RefPortHelper.PortLCS(hangerPortName);
                    }

                 }

                //Get the Current Location in the Route Connection Cycle

                for (int index = 1; index <= numOfRoutes; index++)
                {
                    if (index == 1)
                        refPortName = "Route";
                    else
                        refPortName = "Route_" + index;
                    //Get the distance between the Route and Structure ports
                    portsDistance[index - 1] = RefPortHelper.DistanceBetweenPorts(refPortName, "Structure", PortDistanceType.Horizontal);
                }

                // Set Values of Part Occurance Attributes

                double[] steelWidth = new double[channel_End];
                double[] steelDepth = new double[channel_End];

                for (int index = channel_Begin; index <= channel_End; index++)
                {
                    SupportComponent channel = componentDictionary[ChannelPartKeys[index - clamp_End - 1]];
                    BusinessObject channelPart = channel.GetRelationship("madeFrom", "part").TargetObjects[0];
                    CrossSection crossSection = (CrossSection)channelPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                    //Get dSteelWidth and dSteelDepth of all the channels
                    steelWidth[index - 1] = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                    steelDepth[index - 1] = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                    // For Channel
                    PropertyValueCodelist channelHolePatternCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrAssyChHolePattern", "ChannelHolePattern");
                    int channelHolePattern = (int)channelHolePatternCodelist.PropValue;

                    PropertyValueCodelist channelMaterialCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyB22Channel", "ChannelMaterial");
                    int channelMaterialValue = (int)channelMaterialCodelist.PropValue;

                    channel.SetPropertyValue(channelHolePattern, "IJOAHgrBLineChannelB22", "HolePattern");
                    channel.SetPropertyValue(channelMaterialValue, "IJOAHgrBLineChannelB22", "MaterialType");
                }

                for (int index = clamp_Begin; index <= clamp_End; index++)
                {
                    SupportComponent clamp = componentDictionary[ClampPartKeys[index - 1]];
                    BusinessObject clampPart = clamp.GetRelationship("madeFrom", "part").TargetObjects[0];

                    //for Clamp
                    PropertyValueCodelist inorOutCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                    int inorOut = (int)inorOutCodeList.PropValue;

                    PropertyValueCodelist clampGuideCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyClampGuide", "ClampGuide");
                    int clampGuideValue = (int)clampGuideCodeList.PropValue;

                    string hardware = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyHardware", "Hardware")).PropValue;

                    clamp.SetPropertyValue(inorOut, "IJOAHgrBlineInsideOutside", "Inside_Outside");
                    clamp.SetPropertyValue(hardware, "IJUAHgrBlineWithHardware", "WithHardware");
                    clamp.SetPropertyValue(clampGuideValue, "IJOAHgrBlineClampGuide", "Clamp_Guide");
                }

                //Get the current face number
                if (support.SupportingObjects.Count == 0)
                {
                    currentFaceNumber = 513;
                    goto CREATE_JOINTS;
                }
                else if (support.SupportingObjects.Equals(null))
                {
                    currentFaceNumber = 513;
                    goto CREATE_JOINTS;
                }
                if ((SupportHelper.SupportingObjects.Count != 0))
                    currentFaceNumber = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
				else
					currentFaceNumber = 513;	

            CREATE_JOINTS: 
                //Create Joints
                for (int index = 1; index <= numOfRoutes; index++)
                {
                    if (index == 1)
                    {
                        channelIndex = channel_Begin;
                        connectionIndex = connection_Begin;
                        refPortName = "Route";
                        clampIndex = clamp_Begin;
                    }
                    else
                    {
                        channelIndex = channelIndex + 1;
                        connectionIndex = connectionIndex + 1;
                        refPortName = "Route_" + index;
                        clampIndex = clampIndex + 1;
                    }
                    double angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                    connPlaneOffset = (portsDistance[0] + width[0]) - (portsDistance[index - 1] + width[index - 1]);
                    connAxisOffset = portOrientation[0].Origin.Z - portOrientation[index - 1].Origin.Z;
                    if (System.Math.Abs(angle) < (System.Math.PI) / 2)
                    {
                        confConnIndex = 74;
                        confRouteIndex = 9445;

                        connPlaneOffset = -connPlaneOffset;
                        connAxisOffset = -connAxisOffset;

                        routePlaneOffset = steelDepth[channelIndex - 1];
                        routeAxisOffset = width[index - 1];
                        routeOriginOffset = steelWidth[channelIndex - 1] / 2;

                        connRouteAxisOffset = height[index - 1] / 2 + steelDepth[channelIndex - 1];
                        if (currentFaceNumber == 514)
                            connRouteOriginOffset = 0;
                        else
                            connRouteOriginOffset = steelWidth[channelIndex - 1];
                    }
                    else
                    {
                        confConnIndex = 10;
                        confRouteIndex = 9381;

                        connPlaneOffset = -connPlaneOffset;
                        connAxisOffset = -connAxisOffset;

                        routePlaneOffset = -steelDepth[channelIndex - 1];
                        routeAxisOffset = width[index - 1];
                        routeOriginOffset = -steelWidth[channelIndex - 1] / 2;

                        connRouteAxisOffset = -height[index - 1] / 2 - steelDepth[channelIndex - 1];
                        if (currentFaceNumber == 514)
                            connRouteOriginOffset = -steelWidth[channelIndex - 1];
                        else
                            connRouteOriginOffset = 0;
                    }

                    if (confConnIndex == 74 && confRouteIndex == 9445)
                    {
                        ConfConnAxisA = Axis.Y;
                        ConfConnAxisB = Axis.X;
                        ConfRoutePlaneA = Plane.ZX;
                        ConfRoutePlaneB = Plane.XY;
                        ConfRouteAxisA = Axis.X;
                        ConfRouteAxisB = Axis.X;
                    }
                    else
                    {
                        //confConnIndex == 10 && confRouteIndex == 9381
                        ConfConnAxisA = Axis.Y;
                        ConfConnAxisB = Axis.NegativeX;
                        ConfRoutePlaneA = Plane.ZX;
                        ConfRoutePlaneB = Plane.NegativeXY;
                        ConfRouteAxisA = Axis.X;
                        ConfRouteAxisB = Axis.X;
                    }

                    string connectionBegin = ConnetionPartKeys[0];

                    //Add joint between Channel and the Beam
                    JointHelper.CreatePlanarJoint(ChannelPartKeys[index - 1], "BeginCap", "-1", "Structure", Plane.XY, Plane.XY, 0);

                    //Add joint (Flexible) between Channel ports
                    JointHelper.CreatePrismaticJoint(ChannelPartKeys[index - 1], "BeginCap", ChannelPartKeys[index - 1], "EndCap", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);

                    //Add joint between Channel and the Connection Object
                    JointHelper.CreateRevoluteJoint(ChannelPartKeys[index - 1], "EndCap", ConnetionPartKeys[index - 1], "Connection", ConfConnAxisA, ConfConnAxisB);

                    if (channelIndex == channel_Begin)
                    {
                        //Add joint between Connection Object and the Route
                        JointHelper.CreateRigidJoint(ConnetionPartKeys[index - 1], "Connection", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, width[index - 1], connRouteAxisOffset, connRouteOriginOffset);
                    }
                    else
                    {
                        //Add joint between first Cconnection Object and the Connection Objects on other Routes
                        JointHelper.CreateRigidJoint(ConnetionPartKeys[index - 1], "Connection", connectionBegin, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, connPlaneOffset, connAxisOffset, 0);
                    }

                    //Add joint between Connection Object and the Clamp
                    JointHelper.CreateRigidJoint(ConnetionPartKeys[index - 1], "Connection", ClampPartKeys[index - 1], "TrayPort", ConfRoutePlaneA, ConfRoutePlaneB, ConfRouteAxisA, ConfRouteAxisB, routePlaneOffset, routeAxisOffset, routeOriginOffset);
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

                    for (int clampIndex = clamp_Begin; clampIndex <= clamp_End; clampIndex++)
                    {
                        partkey = ClampPartKeys[clampIndex - 1];
                        int connecttoroute = clampIndex - clamp_Begin + 1;
                        routeConnections.Add(new ConnectionInfo(partkey, connecttoroute));
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
                    structConnections.Add(new ConnectionInfo(ClampPartKeys[0], 1));
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                string hardware;
                string clampType;
                long insideOuside;
                long clampGuide;
                int cableTrayCount = SupportHelper.SupportedObjects.Count;

                clampType = SupportOrComponent.GetPropertyValue("IJUAHgrAssyClampType", "ClampType").ToString();

                PropertyValueCodelist insideOutsideCodelist;
                insideOutsideCodelist = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                insideOuside = insideOutsideCodelist.PropValue;

                hardware = SupportOrComponent.GetPropertyValue("IJUAHgrAssyHardware", "Hardware").ToString();

                PropertyValueCodelist clampGuideCodelist;
                clampGuideCodelist = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrAssyClampGuide", "ClampGuide");
                clampGuide = clampGuideCodelist.PropValue;

                if (insideOuside == 0)
                    insideOuside = 1;
                string insideOusideValue = insideOutsideCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(insideOutsideCodelist.PropValue).DisplayName;

                if (clampGuide == 0)
                    clampGuide = 1;
                string clampGuideValue = clampGuideCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(clampGuideCodelist.PropValue).DisplayName;

                if (cableTrayCount > 1)
                    BOMString = "B-Line Assy Horizontal Structural Support with CT " + clampGuideValue + " type: " + clampType + " mounted " + insideOusideValue + ", " + hardware + " Hardware, with Multi levels of Cable Tray";
                else
                    BOMString = "B-Line Assy Horizontal Structural Support with CT " + clampGuideValue + " type: " + clampType + " mounted " + insideOusideValue + ", " + hardware + " Hardware, with One level of Cable Tray";

                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - Hor_Suuport BOM" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }

}


