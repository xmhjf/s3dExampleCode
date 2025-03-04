//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_LE_DL.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_LE_DL
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
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
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class Assy_LE_DL : CustomSupportDefinition
    {
        //Constants
        private const string HGR_PIPE = "HGR_PIPE"; 
        private const string CONNECTION = "CONNECTION"; 
        private const string CONNECTION1 = "CONNECTION1"; 

        string pipeSize,Orientation;
        double extension;        
        
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

                    pipeSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrPipeSize", "PipeSize")).PropValue;

                    extension = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssy_LE_DL", "EXTENSION")).PropValue;

                    Orientation = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssy_LE_DL", "Orientation")).PropValue;

                    parts.Add(new PartInfo(HGR_PIPE, pipeSize));
                    parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(CONNECTION1, "Log_Conn_Part_1"));

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
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
 
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                NominalDiameter pipeDiameter = new NominalDiameter();
                pipeDiameter.Size = pipeInfo.NominalDiameter.Size;
                pipeDiameter.Size = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeDiameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);

                BusinessObject hgrPipePart = componentDictionary[HGR_PIPE].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)hgrPipePart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double hgrPipeRadius = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue / 2;
                Matrix4X4 routePortOrientation = new Matrix4X4();
                Matrix4X4 structPortOrientation = new Matrix4X4();
                BusinessObject turnRefPort = RefPortHelper.ReferencePort("TurnRef");
                double byPointRouteOffset1 = 0, byPointRouteOffset2 = 0, byPointRouteOffset3 = 0, routeStructAngle1 = 0, routeStructAngle2 = 0, routeStructAngle3 = 0, routeStructAngle4 = 0;

                Plane routePlaneA = new Plane();
                Plane routePlaneB = new Plane();
                Axis routeAxisA = new Axis();
                Axis routeAxisB = new Axis();
                string elbowPositon =string.Empty;

                const double CONST_1 = 90.001; const double CONST_2 = 60.999;
                const int CONST_3 = 150; const int CONST_4 = 210; const int CONST_5 = 181; const int CONST_6 = 179; const int CONST_7 = 91; const int CONST_8 = 89; const int CONST_9 = 182; const int CONST_10 = 178;

                if (turnRefPort != null)
                {
                    double valueAdj = 0;
                    double CALC1 = Math.Sqrt(Math.Abs((((pipeInfo.BendRadius + pipeDiameter.Size / 2) * (pipeInfo.BendRadius + pipeDiameter.Size / 2)) - ((pipeInfo.BendRadius + hgrPipeRadius) * (pipeInfo.BendRadius + hgrPipeRadius)))));
                    if (Math.Round((pipeInfo.BendAngle * 180 / Math.PI), 3) < CONST_1 && Math.Round((pipeInfo.BendAngle * 180 / Math.PI), 3) > CONST_2)
                        valueAdj = pipeInfo.BendRadius - CALC1;
                    else
                    {
                        CALC1 = hgrPipeRadius / Math.Tan(pipeInfo.BendAngle / 2);
                        valueAdj = pipeInfo.BendRadius - CALC1 - pipeInfo.FaceToCenter;
                    }
                    routePortOrientation = RefPortHelper.PortLCS("TurnRef");
                    structPortOrientation = RefPortHelper.PortLCS("Structure");
                    if (routePortOrientation.Origin.Z > structPortOrientation.Origin.Z)
                        elbowPositon = "Above";
                    else
                        elbowPositon = "Below";
                    //Get Port Orientations
                    routeStructAngle1 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, "TurnRef", PortAxisType.X, OrientationAlong.Direct);
                    routeStructAngle2 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, "TurnRef", PortAxisType.Y, OrientationAlong.Direct);
                    routeStructAngle3 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, "TurnRef", PortAxisType.X, OrientationAlong.Direct);
                    routeStructAngle4 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, "TurnRef", PortAxisType.Z, OrientationAlong.Direct);
                    if (Math.Round((routeStructAngle4 * 180 / Math.PI), 3) >= CONST_3 && Math.Round((routeStructAngle4 * 180 / Math.PI), 3) <= CONST_4)
                        if (Orientation != "Horizontal")
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Vertical Dummy Leg cannot be placed on a horizontal turn", "", "Assy_LE_DL.cs", 1);
                            return;
                        }


                    if (Orientation == "Horizontal")
                    {
                        byPointRouteOffset1 = hgrPipeRadius;
                        byPointRouteOffset2 = -valueAdj;
                        byPointRouteOffset3 = -hgrPipeRadius;

                        routePlaneA = Plane.ZX;
                        routePlaneB = Plane.ZX;
                        routeAxisA = Axis.X;
                        routeAxisB = Axis.Z;
                    }
                    else
                    {
                        if (elbowPositon == "Below")
                        {
                            byPointRouteOffset1 = -valueAdj;
                            byPointRouteOffset2 = hgrPipeRadius;
                            byPointRouteOffset3 = hgrPipeRadius;
 
                            routePlaneA = Plane.ZX;
                            routePlaneB = Plane.NegativeZX;
                            routeAxisA = Axis.X;
                            routeAxisB = Axis.NegativeX;
                        }
                        else
                        {
                            byPointRouteOffset1 = valueAdj;
                            byPointRouteOffset2 = hgrPipeRadius;
                            byPointRouteOffset3 = -hgrPipeRadius;
 
                            routePlaneA = Plane.ZX;
                            routePlaneB = Plane.ZX;
                            routeAxisA = Axis.X;
                            routeAxisB = Axis.NegativeX;
                        }
                    }

                }
                //Create Joints
                if (turnRefPort != null)
                {
                    if (Orientation == "Horizontal")
                        JointHelper.CreateRigidJoint("-1", "TurnRef", HGR_PIPE, "BeginCap", routePlaneA, routePlaneB, routeAxisA, routeAxisB, byPointRouteOffset3, byPointRouteOffset1, byPointRouteOffset2);
                    else
                    {
                        JointHelper.CreateSphericalJoint(CONNECTION1, "Connection", "-1", "TurnRef");
                        JointHelper.CreateGlobalAxesAlignedJoint(CONNECTION1, "Connection", Axis.Z, Axis.Z);

                        JointHelper.CreateRigidJoint(CONNECTION1, "Connection", HGR_PIPE, "BeginCap", routePlaneA, routePlaneB, routeAxisA, routeAxisB, byPointRouteOffset3, byPointRouteOffset1, byPointRouteOffset2);

                    }

                    //Create a Cylindrical Joint between the two ports of the HgrPipe
                    JointHelper.CreatePrismaticJoint(HGR_PIPE, "BeginCap", HGR_PIPE, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);


                    //Connection to HgrPipe
                    JointHelper.CreateSphericalJoint(CONNECTION, "Connection", HGR_PIPE, "EndCap");


                    //Connection to Structure
                    if (Orientation == "Horizontal")
                    {
                        if ((routeStructAngle1 * 180 / Math.PI) < CONST_5 && (routeStructAngle1 * 180 / Math.PI) > CONST_6)
                            //This Joint will connect to the Ends of the Flanges
                            JointHelper.CreateTranslationalJoint(CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, extension / 2);             
                        else
                            //This Joint will connect to the Middle of the WEB
                            if ((routeStructAngle2 * 180 / Math.PI) < CONST_5 && (routeStructAngle2 * 180 / Math.PI) > CONST_6)
                                JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, -extension / 2);                   
                            else
                                JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, extension / 2);

                    }
                    else
                    {
                        //Elbow Perpendicular to the Stucture
                        if ((routeStructAngle3 * 180 / Math.PI) < CONST_7 && (routeStructAngle3 * 180 / Math.PI) > CONST_8)
                            if ((routeStructAngle1 * 180 / Math.PI) < CONST_7 && (routeStructAngle1 * 180 / Math.PI) > CONST_8)
                                //This Joint will connect to the Ends of the Flanges
                                JointHelper.CreateTranslationalJoint(CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, extension);
                            else
                                //This Joint will connect to the Middle of the WEB
                                if ((routeStructAngle1 * 180 / Math.PI) < CONST_9 && (routeStructAngle1 * 180 / Math.PI) > CONST_10)
                                    JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, extension);
                                else
                                    JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, -extension);         
                        else //Elbow Parallel to the Stucture
                            if ((routeStructAngle2 * 180 / Math.PI) < CONST_7 && (routeStructAngle2 * 180 / Math.PI) > CONST_8)
                                //This Joint will connect to the Ends of the Flanges
                                JointHelper.CreateTranslationalJoint(CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, extension);                       
                            else
                                if (elbowPositon == "Below")
                                    //This Joint will connect to the Middle of the WEB
                                    if ((routeStructAngle2 * 180 / Math.PI) < CONST_5 && (routeStructAngle2 * 180 / Math.PI) > CONST_6)
                                        JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, -extension);                            
                                    else
                                        JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, extension);                            
                                else
                                    if ((routeStructAngle2 * 180 / Math.PI) < CONST_5 && (routeStructAngle2 * 180 / Math.PI) > CONST_6)
                                        JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, extension);                            
                                    else
                                        JointHelper.CreateTranslationalJoint(CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, -extension);
                    }
                }
                else
                {
                    //Not placed on a turn - Therefore use the Route Port (Cap, Tee, Pipe, etc...)
                    double planeAngle1, planeAngle2;

                    CommonAssembly commonAssembly = new CommonAssembly();

                    planeAngle1 = commonAssembly.GetRouteStructConfigAngle(this, "Structure", "Route", PortAxisType.Z);
                    planeAngle2 = commonAssembly.GetRouteStructConfigAngle(this, "Structure", "Route", PortAxisType.Y);

                    JointHelper.CreatePrismaticJoint(HGR_PIPE, "BeginCap", HGR_PIPE, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    if (Orientation == "HORIZONTAL")
                        if (planeAngle1 < Math.PI / 4 || planeAngle1 > 3 * Math.PI / 4)
                            JointHelper.CreateRigidJoint(HGR_PIPE, "BeginCap", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, hgrPipeRadius, hgrPipeRadius);                
                        else
                            if (planeAngle2 < Math.PI / 4)
                                JointHelper.CreateRigidJoint(HGR_PIPE, "BeginCap", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, hgrPipeRadius, 0, hgrPipeRadius);                   
                            else
                                JointHelper.CreateRigidJoint(HGR_PIPE, "BeginCap", "-1", "Route", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, hgrPipeRadius, 0, hgrPipeRadius);     
                    else
                        JointHelper.CreateRigidJoint(HGR_PIPE, "BeginCap", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, hgrPipeRadius, hgrPipeRadius);


                    //Connection to HgrPipe
                    JointHelper.CreateSphericalJoint(CONNECTION, "Connection", HGR_PIPE, "EndCap");


                    if (planeAngle1 < Math.PI / 4 || planeAngle1 > 3 * Math.PI / 4)
                        JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, -extension);           
                    else
                        if (planeAngle2 < Math.PI / 4)
                            JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, extension);           
                        else
                            JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.ZX, Plane.ZX, -extension);     
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
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(HGR_PIPE, 1)); // partindex, routeindex

                        //Return the collection of Route connection information
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
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections.Add(new ConnectionInfo(CONNECTION, 1)); // partindex, routeindex

                    //Return the collection of Structure connection information.
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

