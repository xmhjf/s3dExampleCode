//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_YK5.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_YK5
//   Author       :  BS
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     BS      CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net 
//   22-Jan-2015     PVK     TR-CP-264951  Resolve coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
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

    public class SFS5380_YK5 : CustomSupportDefinition
    {
        private const string PIPECLAMPYK5 = "PipeClamp_YK5";
        private const string RODYK5 = "Rod_YK5";
        private const string EYENUTYK5 = "EyeNut_YK5";
        private const string CONNOBJ1YK5 = "ConnObj1_YK5";
        private const string CONNOBJ2YK5 = "ConnObj2_YK5";
        public int Index { get; set; }

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(2);
                    double pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                    PropertyValue clampsize;
                    Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                    parameter.Add("ClampSizeIn", pipeDiameter.ToString());
                    GenericHelper.GetDataByRule("FINLSubSrv_SFS5371_Clamp", "IJUAFINLSrv_5371_Clamp", "ClampPart", parameter, out clampsize);

                    int loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                    parts.Add(new PartInfo(PIPECLAMPYK5 + Index, "FINLCmp_SFS5371_" + clampsize));
                    parts.Add(new PartInfo(RODYK5 + Index, "FINLCmp_SFS5381_" + loadClass));
                    parts.Add(new PartInfo(EYENUTYK5 + Index, "FINLCmp_SFS5389_" + loadClass));
                    parts.Add(new PartInfo(CONNOBJ1YK5 + Index, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(CONNOBJ2YK5 + Index, "Log_Conn_Part_1"));

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
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                string sAssmDesc = support.SupportDefinition.PartDescription;

                if (sAssmDesc.Contains("AK2") || sAssmDesc.Contains("AK4") || sAssmDesc.Contains("AR2") || sAssmDesc.Contains("AR4"))
                {
                    JointHelper.CreateRigidJoint("-1", "Route_2", CONNOBJ2YK5 + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, 0, 0);
                    if (SupportHelper.SupportedObjects.Count == 2)
                        JointHelper.CreatePrismaticJoint(CONNOBJ2YK5 + Index, "Connection", PIPECLAMPYK5 + Index, "Route", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0);
                    else
                        JointHelper.CreatePrismaticJoint(CONNOBJ2YK5 + Index, "Connection", PIPECLAMPYK5 + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                }
                else
                {
                    double RouteAngle = Math3d.Deg(RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Route_2", PortAxisType.X, OrientationAlong.Direct));
                    if (HgrCompareDoubleService.cmpdbl(Math.Round(RouteAngle, 0), 0) == true || HgrCompareDoubleService.cmpdbl(Math.Round(RouteAngle, 0), 180) == true)
                        JointHelper.CreateRigidJoint("-1", "Route_2", PIPECLAMPYK5 + Index, "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, 0, 0);
                    else
                    {
                        JointHelper.CreateRigidJoint("-1", "Route_2", CONNOBJ2YK5 + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreatePrismaticJoint(CONNOBJ2YK5 + Index, "Connection", PIPECLAMPYK5 + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0);
                    }
                }

                JointHelper.CreateRigidJoint(PIPECLAMPYK5 + Index, "Pin", CONNOBJ1YK5 + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRevoluteJoint(CONNOBJ1YK5 + Index, "Connection", EYENUTYK5 + Index, "Eye", Axis.X, Axis.X);

                JointHelper.CreateRigidJoint(EYENUTYK5 + Index, "InThdRH", RODYK5 + Index, "BotExThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreatePrismaticJoint(RODYK5 + Index, "BotExThdRH", RODYK5 + Index, "TopExThdRH", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);

                JointHelper.CreateGlobalAxesAlignedJoint(RODYK5 + Index, "BotExThdRH", Axis.Z, Axis.Z);

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(PIPECLAMPYK5 + Index, 1));

                    //Return the collection of Route connection information.
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
                    structConnections.Add(new ConnectionInfo(RODYK5 + Index, 1));

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


