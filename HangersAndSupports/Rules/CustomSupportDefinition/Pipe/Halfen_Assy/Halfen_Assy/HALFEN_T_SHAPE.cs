//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_T_SHAPE.cs
//   Halfen_Assy,Ingr.SP3D.Content.Support.Rules.HALFEN_T_SHAPE
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
    public class HALFEN_T_SHAPE : CustomSupportDefinition
    {
        //Constants
        //Use this with normal base
        private const string T_TOP = "T_TOP";
        private const string CHANNEL = "CHANNEL";
        private const string BASE = "BASE";
        private const string TK_1 = "TK_1";
        private const string TK_2 = "TK_2";
        private const string TK_3 = "TK_3";
        private const string TK_4 = "TK_4";

        //Use this without a channel
        private const string ALT_BASE = "ALT_BASE";
        private const string ALT_TK_1 = "ALT_TK_1";
        private const string ALT_TK_2 = "ALT_TK_2";
        private const string ALT_TK_3 = "ALT_TK_3";
        private const string ALT_TK_4 = "ALT_TK_4";

        private string topKeys;
        private double shoeHeight;
        private double length;
        private double T_TopLength;
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

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenShoeH", "SHOE_H")).PropValue;

                    PropertyValueCodelist base_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenTBaseSize", "BASE_SIZE");
                    long base_SizeCodelistValue = (long)base_SizeCodelist.PropValue;

                    length = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenLength", "LENGTH")).PropValue;
                    T_TopLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHalfenTShape", "T_TOPLENGTH")).PropValue;

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
                        if (base_SizeCodelistValue > 3)
                        {
                            parts.Add(new PartInfo(T_TOP, "HALFEN_HCS_VT63_31_1"));
                            parts.Add(new PartInfo(ALT_BASE, "HALFEN_HCS_VT63_15"));
                            parts.Add(new PartInfo(ALT_TK_1, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_TK_2, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_TK_3, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(ALT_TK_4, "HALFEN_HCS_TK_2"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(T_TOP, "HALFEN_HCS_VT63_31_1"));
                            parts.Add(new PartInfo(CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(BASE, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
                            parts.Add(new PartInfo(TK_1, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(TK_2, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(TK_3, "HALFEN_HCS_TK_2"));
                            parts.Add(new PartInfo(TK_4, "HALFEN_HCS_TK_2"));
                        }
                    }
                    else if (supportingType == "Slab")
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                            parts.Add(new PartInfo(T_TOP, "HALFEN_HCS_VT63_31_1"));
                            parts.Add(new PartInfo(ALT_BASE, "HALFEN_HCS_VT63_15"));
                        }
                        else
                        {
                            parts.Add(new PartInfo(T_TOP, "HALFEN_HCS_VT63_31_1"));
                            parts.Add(new PartInfo(CHANNEL, "HALFEN_HZL_63_63_1"));
                            parts.Add(new PartInfo(BASE, "HALFEN_HCS_VT63_1" + base_SizeCodelistValue));
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
                //Load standard bounding box definition
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;
               
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent T_Top = componentDictionary[T_TOP];

                PropertyValueCodelist base_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenTBaseSize", "BASE_SIZE");
                long base_SizeCodelistValue = (long)base_SizeCodelist.PropValue;

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

                int routeConnectionValue = Configuration;
                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter + 2.0 * routeInfo.InsulationThickness;

                if (base_SizeCodelistValue > 3 && length - pipeDiameter / 2.0 - shoeHeight > 545.0)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "The length specified is too large for HCS-VT63-15/0.", "", "HALFEN_T_SHAPE.cs", 1);
               
                double byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                double base_T;
                double bolt_CC;
                double bolt_CC2;
                if (base_SizeCodelistValue > 3)
                {
                    SupportComponent alt_base = componentDictionary[ALT_BASE];
                    BusinessObject ALT_BASEpart = alt_base.GetRelationship("madeFrom", "part").TargetObjects[0];
                    base_T = (double)((PropertyValueDouble)ALT_BASEpart.GetPropertyValue("IJUAHgrT", "T")).PropValue;
                    bolt_CC = (double)((PropertyValueDouble)ALT_BASEpart.GetPropertyValue("IJUAHgrBolt_CC", "Bolt_CC")).PropValue;
                    bolt_CC2 = (double)((PropertyValueDouble)ALT_BASEpart.GetPropertyValue("IJUAHgrBolt_CC2", "Bolt_CC2")).PropValue;
                }
                else
                {
                    SupportComponent Base = componentDictionary[BASE];
                    BusinessObject Basepart = Base.GetRelationship("madeFrom", "part").TargetObjects[0];
                    base_T = (double)((PropertyValueDouble)Basepart.GetPropertyValue("IJUAHgrT", "T")).PropValue;
                    bolt_CC = (double)((PropertyValueDouble)Basepart.GetPropertyValue("IJUAHgrBolt_CC", "Bolt_CC")).PropValue;
                    bolt_CC2 = (double)((PropertyValueDouble)Basepart.GetPropertyValue("IJUAHgrBolt_CC2", "Bolt_CC2")).PropValue;
                }
                //Set T Top Horizontal Length
                T_Top.SetPropertyValue(T_TopLength, "IJOAHgrT_TopLength", "T_TopLength");
                //Start Joints here
                if (base_SizeCodelistValue < 4)
                {
                    //WITH VERTICAL CHANNEL
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)   //Here is the place by structure stuff
                    {
                        //Add joint between Route and T Top
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_High", T_TOP, "Top", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, shoeHeight, -BBXWidth / 2);
                        //Add joint between channel and top
                        JointHelper.CreateRigidJoint(T_TOP, "Middle", CHANNEL, "BeginMiddle", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                        //Add joint between vertical channel and base
                        JointHelper.CreateRigidJoint(BASE, "Base", CHANNEL, "EndMiddle", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, base_T, 0, 0);
                        //Add Joint Between the Base and the Supporting beam
                        JointHelper.CreatePrismaticJoint(BASE, "Base", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0);
                    }
                    else    //PLACE BY POINT 
                    {
                        //Add joint between Route and T Top
                        JointHelper.CreateRigidJoint("-1", "BBR_High", T_TOP, "Top", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, shoeHeight, -BBXWidth / 2, 0);
                        //Add joint between channel and top
                        JointHelper.CreateRigidJoint(T_TOP, "Middle", CHANNEL, "BeginMiddle", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                        // Add joint between vertical channel and base
                        if (routeConnectionValue == 1)
                        {
                            JointHelper.CreateRigidJoint(BASE, "Base", CHANNEL, "EndMiddle", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, base_T, 0, 0);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint(BASE, "Base", CHANNEL, "EndMiddle", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, base_T, 0, 0);
                        }
                        //Add Joint Between the Base and the Supporting beam
                        JointHelper.CreatePlanarJoint(BASE, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                    }
                    // Flexible Member (Vertical)
                    JointHelper.CreatePrismaticJoint(CHANNEL, "EndMiddle", CHANNEL, "BeginMiddle", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    if (supportingType == "Steel")
                    {
                        //Add joints between upper connector and clamps
                        JointHelper.CreateRigidJoint(TK_1, "BottomOfClamp", BASE, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2, -bolt_CC / 2);
                        JointHelper.CreateRigidJoint(TK_2, "BottomOfClamp", BASE, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2, bolt_CC / 2);
                        JointHelper.CreateRigidJoint(TK_3, "BottomOfClamp", BASE, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2, -bolt_CC / 2);
                        JointHelper.CreateRigidJoint(TK_4, "BottomOfClamp", BASE, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2, bolt_CC / 2);
                    }
                }
                else
                {
                    //WITHOUT VERTICAL CHANNEL
                    //Here is the place by structure stuff
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        //Add joint between Route and T Top
                        JointHelper.CreateRigidJoint("-1", "BBSR_High", T_TOP, "Top", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, shoeHeight, -BBXWidth / 2, 0);
                    }
                    else //PLACE BY POINT
                    {
                        //Add joint between Route and T Top
                        JointHelper.CreateRigidJoint("-1", "BBR_High", T_TOP, "Top", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, shoeHeight, -BBXWidth / 2, 0);
                    }
                    //Add joint between channel and top
                    JointHelper.CreatePrismaticJoint(T_TOP, "Middle", ALT_BASE, "Top", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.Z, 0, 0);
                    //Add Joint Between the Base and the Supporting beam
                    JointHelper.CreatePlanarJoint(ALT_BASE, "Base", "-1", "Structure", Plane.XY, Plane.XY, 0);
                    if (supportingType == "Steel")
                    {
                        //Add joints between upper connector and clamps
                        JointHelper.CreateRigidJoint(ALT_TK_1, "BottomOfClamp", ALT_BASE, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2, -bolt_CC / 2);
                        JointHelper.CreateRigidJoint(ALT_TK_2, "BottomOfClamp", ALT_BASE, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, bolt_CC2 / 2, bolt_CC / 2);
                        JointHelper.CreateRigidJoint(ALT_TK_3, "BottomOfClamp", ALT_BASE, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2, -bolt_CC / 2);
                        JointHelper.CreateRigidJoint(ALT_TK_4, "BottomOfClamp", ALT_BASE, "Base", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0.0, -bolt_CC2 / 2, bolt_CC / 2);
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
                    int numOfRoutes = SupportHelper.SupportedObjects.Count;

                    //For Clamp
                    int clamp_Begin = 1;
                    int numOfPart = numOfRoutes;
                    int clamp_End = numOfPart;

                    for (int index = clamp_Begin; index <= clamp_End; index++)
                    {
                        topKeys = "T_TOP";
                        int connecttoroute = index;
                        routeConnections.Add(new ConnectionInfo(topKeys, connecttoroute));
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
                    PropertyValueCodelist base_SizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHalfenTBaseSize", "BASE_SIZE");
                    long base_SizeCodelistValue = (long)base_SizeCodelist.PropValue;

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
                        if (base_SizeCodelistValue > 3)
                        {
                                structConnections.Add(new ConnectionInfo(ALT_BASE, 1));
                                structConnections.Add(new ConnectionInfo(ALT_TK_1, 1));
                                structConnections.Add(new ConnectionInfo(ALT_TK_2, 1));
                                structConnections.Add(new ConnectionInfo(ALT_TK_3, 1));
                                structConnections.Add(new ConnectionInfo(ALT_TK_4, 1));
                        }
                        else
                        {
                                structConnections.Add(new ConnectionInfo(BASE, 1));
                                structConnections.Add(new ConnectionInfo(TK_1, 1));
                                structConnections.Add(new ConnectionInfo(TK_2, 1));
                                structConnections.Add(new ConnectionInfo(TK_3, 1));
                                structConnections.Add(new ConnectionInfo(TK_4, 1));
                        }
                    }
                    else
                    {
                        if (base_SizeCodelistValue > 3)
                        {
                                structConnections.Add(new ConnectionInfo(ALT_BASE, 1));
                        }
                        else
                        {
                                structConnections.Add(new ConnectionInfo(BASE, 1));
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
    }
}



