//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_YK2.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_YK2
//   Author       : Rajeswari
//   Creation Date: 27-June-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 01-July-2013  Rajeswari CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net
// 22-Jan-2015   PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //--------------------------------------------------------------------------------

    public class SFS5380_YK2 : CustomSupportDefinition
    {
        private const string ATTACHMENTB = "ATTACHMENTB_YK2";
        private const string CLEVIS = "CLEVIS_YK2";
        private const string ROD = "ROD_YK2";
        private const string LOGOBJ = "LOGOBJ_YK2";
        public int Index { get; set; }

        int loadClass;
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

                    // GET variable values from XLS files
                    loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                    parts.Add(new PartInfo(ATTACHMENTB + Index, "FINLCmp_SFS5387_" + loadClass));
                    parts.Add(new PartInfo(CLEVIS + Index, "FINLCmp_SFS5390_" + loadClass));
                    parts.Add(new PartInfo(ROD + Index, "FINLCmp_SFS5381_" + loadClass));
                    parts.Add(new PartInfo(LOGOBJ + Index, "Log_Conn_Part_1"));

                    // Return the collection of Catalog Parts
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                string assmDescription = support.SupportDefinition.PartDescription;

                // Rotate Top sub-assy 0/90 degree

                double hexNutTakeOut = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_SFS5380_YK1", "IJUAFINLSrv_5380YK1", "HexNutTakeOut", "IJUAFINLSrv_5380YK1", "LoadClass", loadClass.ToString());
                double hexNutThickness = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_SFS5380_YK1", "IJUAFINLSrv_5380YK1", "HexNutThk", "IJUAFINLSrv_5380YK1", "LoadClass", loadClass.ToString());

                if (assmDescription.IndexOf("AK2") > 0 || assmDescription.IndexOf("AK4") > 0 || assmDescription.IndexOf("AR2") > 0 || assmDescription.IndexOf("AR4") > 0)
                {
                    double routeXStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);
                    if (assmDescription.IndexOf("AK2") > 0 || assmDescription.IndexOf("AR2") > 0)
                    {
                        JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreateTranslationalJoint(ATTACHMENTB + Index, "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0);
                    }
                    else if (assmDescription.IndexOf("AK4") > 0 || assmDescription.IndexOf("AR4") > 0)
                    {
                        if (Math.Abs((routeXStructAngle) * 180 / Math.PI - 90) < 2)
                        {
                            JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0, 0);

                            JointHelper.CreatePrismaticJoint(ATTACHMENTB + Index, "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                            JointHelper.CreatePrismaticJoint(ATTACHMENTB + Index, "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0);
                        }
                    }
                    JointHelper.CreateRevoluteJoint(CLEVIS + Index, "Pin", ATTACHMENTB + Index, "Hole", Axis.X, Axis.X);

                    JointHelper.CreateRigidJoint(CLEVIS + Index, "InThdRH", ROD + Index, "TopExThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0, 0);

                    JointHelper.CreatePrismaticJoint(ROD + Index, "BotExThdRH", ROD + Index, "TopExThdRH", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);

                    JointHelper.CreateGlobalAxesAlignedJoint(CLEVIS + Index, "InThdRH", Axis.NegativeZ, Axis.NegativeZ);
                }
                else
                {
                    JointHelper.CreateRigidJoint("-1", "Structure", ATTACHMENTB + Index, "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0, 0);

                    JointHelper.CreateRevoluteJoint(CLEVIS + Index, "Pin", ATTACHMENTB + Index, "Hole", Axis.X, Axis.X);

                    JointHelper.CreateRigidJoint(CLEVIS + Index, "InThdRH", ROD + Index, "TopExThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0, 0);

                    JointHelper.CreatePrismaticJoint(ROD + Index, "BotExThdRH", ROD + Index, "TopExThdRH", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);

                    JointHelper.CreateGlobalAxesAlignedJoint(CLEVIS + Index, "InThdRH", Axis.NegativeZ, Axis.NegativeZ);
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
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 1;
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
                    routeConnections.Add(new ConnectionInfo(ROD + Index, 1)); // partindex, routeindex

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
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections.Add(new ConnectionInfo(ATTACHMENTB + Index, 1)); // partindex, routeindex

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
