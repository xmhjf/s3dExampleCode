//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   BeamSupCTWoChannel.cs
//   Bline_Assy,Ingr.SP3D.Content.Support.Rules.BeamSupCTWoChannel
//   Author       :  Pavan
//   Creation Date:  5.Nov.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   5.Nov.2012     Pavan   CR-CP-219114-Initial Creation
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
    public class BeamSupCTWoChannel : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private string PartKeys;
        private string partNumber;
        private string clampType;
        private string refPortName;
        private double[] width = new double[5];
        private double[] height = new double[5];
        double width_Inches;
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

                    //For Left Clamp
                    int leftClamp_Begin = 1;
                    int numOfPart = numOfRoutes;
                    int leftClamp_End = numOfPart;

                    for (int index = leftClamp_Begin; index <= leftClamp_End; index++)
                    {
                        width_Inches = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, width[index - 1], UnitName.DISTANCE_INCH);
                        PartKeys = "LeftClamp" + index.ToString();
                        partNumber = clampType + "_" + width_Inches;
                        parts.Add(new PartInfo(PartKeys, partNumber));
                    }

                    //For Right Clamp
                    int rightClamp_Begin = leftClamp_End + 1;
                    int rightClamp_End = leftClamp_End + numOfPart;

                    for (int index = rightClamp_Begin; index <= rightClamp_End; index++)
                    {
                        width_Inches = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, width[index - numOfPart - 1], UnitName.DISTANCE_INCH);
                        PartKeys = "RightClamp" + index.ToString();
                        partNumber = clampType + "_" + width_Inches;
                        parts.Add(new PartInfo(PartKeys, partNumber));
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

                //For Left Clamp
                int leftClamp_Begin = 1;
                int numOfPart = numOfRoutes;
                int leftClamp_End = numOfPart;
                double flangeThickness = 0;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                    {
                        flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                    }
                }

                for (int index = leftClamp_Begin; index <= leftClamp_End; index++)
                {
                    if (index == leftClamp_Begin)
                    {
                        refPortName = "Route";
                    }
                    else
                    {
                        int refPortnum = index - leftClamp_Begin + 1;
                        refPortName = "Route_" + refPortnum;
                    }
                    PartKeys = "LeftClamp" + index.ToString();

                    SupportComponent LeftClampOcc = componentDictionary[PartKeys];
                    BusinessObject LeftClampPart = LeftClampOcc.GetRelationship("madeFrom", "part").TargetObjects[0];
                    //Get Values
                    double leftClampThickness = (double)((PropertyValueDouble)LeftClampPart.GetPropertyValue("IJUAHgrBlineClamp1249X", "TC")).PropValue;
                    double leftLB = (double)((PropertyValueDouble)LeftClampPart.GetPropertyValue("IJUAHgrBlineClamp1249X", "LB")).PropValue;
                    double leftClampHeight = flangeThickness + 2 * leftClampThickness;

                    PropertyValueCodelist insideOutsideCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                    int insideOutsideValue = (int)insideOutsideCodeList.PropValue;

                    int flipClamp = 1;

                    //Set Values
                    LeftClampOcc.SetPropertyValue(insideOutsideValue, "IJOAHgrBlineInsideOutside", "Inside_Outside");
                    LeftClampOcc.SetPropertyValue(flipClamp, "IJOAHgrBlineFlipClamp", "Flip_Clamp");
                    LeftClampOcc.SetPropertyValue(leftClampHeight, "IJUAHgrBlineClamp1249X", "HC");

                    if ((SupportHelper.SupportingObjects.Count != 0))
                    {
                        if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType != SupportingObjectType.Member)
                            //2404
                            JointHelper.CreatePrismaticJoint(PartKeys, "TrayPort", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY, -flangeThickness / 2.0, 0);
                        else
                            //1380
                            JointHelper.CreatePrismaticJoint(PartKeys, "TrayPort", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -flangeThickness / 2.0, 0);
                    }
                    else
                        //1380
                        JointHelper.CreatePrismaticJoint(PartKeys, "TrayPort", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -flangeThickness / 2.0, 0);

                    JointHelper.CreateCylindricalJoint(PartKeys, "TrayPort", "-1", refPortName, Axis.Z, Axis.NegativeZ, 0);
                }

                //For Right Clamp
                int rightClamp_Begin = leftClamp_End + 1;
                int rightClamp_End = leftClamp_End + numOfPart;

                for (int index = rightClamp_Begin; index <= rightClamp_End; index++)
                {
                    PartKeys = "RightClamp" + index.ToString();

                    if (index == (leftClamp_End + 1))
                    {
                        refPortName = "Route";
                    }
                    else
                    {
                        int refPortnum = index - rightClamp_Begin + 1;
                        refPortName = "Route_" + refPortnum;
                    }

                    SupportComponent RightClampOcc = componentDictionary[PartKeys];
                    BusinessObject RightClampPart = RightClampOcc.GetRelationship("madeFrom", "part").TargetObjects[0];
                    //Get Values
                    double rightClampThickness = (double)((PropertyValueDouble)RightClampPart.GetPropertyValue("IJUAHgrBlineClamp1249X", "TC")).PropValue;
                    double rightLB = (double)((PropertyValueDouble)RightClampPart.GetPropertyValue("IJUAHgrBlineClamp1249X", "LB")).PropValue;
                    double rightClampHeight = flangeThickness + 2 * rightClampThickness;

                    PropertyValueCodelist insideOutsideCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                    int insideOutsideValue = (int)insideOutsideCodeList.PropValue;

                    int flipClamp = 2;

                    //Set Values
                    RightClampOcc.SetPropertyValue(insideOutsideValue, "IJOAHgrBlineInsideOutside", "Inside_Outside");
                    RightClampOcc.SetPropertyValue(flipClamp, "IJOAHgrBlineFlipClamp", "Flip_Clamp");
                    RightClampOcc.SetPropertyValue(rightClampHeight, "IJUAHgrBlineClamp1249X", "HC");

                    if ((SupportHelper.SupportingObjects.Count != 0))
                    {
                        if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType != SupportingObjectType.Member)
                            //2404
                            JointHelper.CreatePrismaticJoint(PartKeys, "TrayPort", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY, -flangeThickness / 2.0, 0);
                        else
                            //1380
                            JointHelper.CreatePrismaticJoint(PartKeys, "TrayPort", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -flangeThickness / 2.0, 0);
                    }
                    else
                        //1380
                        JointHelper.CreatePrismaticJoint(PartKeys, "TrayPort", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, -flangeThickness / 2.0, 0);

                    JointHelper.CreateCylindricalJoint(PartKeys, "TrayPort", "-1", refPortName, Axis.Z, Axis.NegativeZ, 0);
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

                    //For Left Clamp
                    int leftClamp_Begin = 1;
                    int numOfPart = numOfRoutes;
                    int leftClamp_End = numOfPart;

                    for (int index = leftClamp_Begin; index <= leftClamp_End; index++)
                    {
                        PartKeys = "LeftClamp" + index.ToString();
                        int connectToRoute = index - leftClamp_Begin + 1;
                        routeConnections.Add(new ConnectionInfo(PartKeys, connectToRoute));
                    }

                    //For Right Clamp
                    int rightClamp_Begin = leftClamp_End + 1;
                    int rightClamp_End = leftClamp_End + numOfPart;

                    for (int index = rightClamp_Begin; index <= rightClamp_End; index++)
                    {
                        PartKeys = "RightClamp" + index.ToString();
                        int connectToRoute = index - rightClamp_Begin + 1;
                        routeConnections.Add(new ConnectionInfo(PartKeys, connectToRoute));
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
                    structConnections.Add(new ConnectionInfo("LeftClamp1", 1));
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
                string hardware = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAHgrAssyHardware", "Hardware")).PropValue;

                PropertyValueCodelist insideOutsideCodeList = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                int insideOutsideValue = (int)insideOutsideCodeList.PropValue;

                if (insideOutsideValue == 0)
                    insideOutsideValue = 1;

                string insideOutside = insideOutsideCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(insideOutsideValue).DisplayName;

                PropertyValueCodelist clampGuideCodeList = (PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJUAHgrAssyClampGuide", "ClampGuide");
                int clampGuideValue = (int)clampGuideCodeList.PropValue;

                if (clampGuideValue == 0)
                    clampGuideValue = 1;

                string clampGuide = clampGuideCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(clampGuideValue).DisplayName;

                BOMString = "B-Line Assy Beam Supported Cable Tray (Without Channel Strut) with CT " + clampGuide + " type: " + clampType + " mounted, " + insideOutside + ", " + hardware + " Hardware";

                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - BeamSupCTWoChannel" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}





