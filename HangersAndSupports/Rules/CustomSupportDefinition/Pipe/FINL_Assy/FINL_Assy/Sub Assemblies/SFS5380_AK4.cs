//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AK4.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK4
//   Author       : Rajeswari
//   Creation Date: 27-June-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 27-June-2013  Rajeswari CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net 
// 20-05-2014    PVK       TR-CP-237153  Problems in placement of .Net HS_FINL_Assy, HS_FINL_SupAssy.
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
    public class SFS5380_AK4 : CustomSupportDefinition
    {
        private const string CROSSBEAM = "CROSSBEAM_AK4";
        private const string HEXNUT1 = "HEXNUT1_AK4";
        private const string HEXNUT22 = "HEXNUT22_AK4";
        private const string HEXNUT3 = "HEXNUT3_AK4";
        private const string HEXNUT4 = "HEXNUT4_AK4";
        private const string LOGOBJ = "LOGOBJ_AK4";
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

                    loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;
                    if (loadClass == 1 || loadClass == 6)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Cannot place loadclass 1 for AK4. Becasue SFS5388 does not have loadclass 1.", "", "SFS5380_AK4.cs", 58);
                        return null;
                    }

                    parts.Add(new PartInfo(CROSSBEAM + Index, "FINLCmp_SFS5388_" + loadClass));
                    parts.Add(new PartInfo(HEXNUT1 + Index, "FINLCmp_SFS4032_" + loadClass));
                    parts.Add(new PartInfo(HEXNUT22 + Index, "FINLCmp_SFS4032_" + loadClass));
                    parts.Add(new PartInfo(HEXNUT3 + Index, "FINLCmp_SFS4032_" + loadClass));
                    parts.Add(new PartInfo(HEXNUT4 + Index, "FINLCmp_SFS4032_" + loadClass));
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
                double userLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAFINL_BeamLength", "BeamLength")).PropValue;
                double shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINL_ShoeH", "ShoeH")).PropValue;
                                
                double takeOut = 0, hexNutThickness = 0;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = pipeInfo.OutsideDiameter;

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

                if (RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct) < pipeInfo.NominalDiameter.Size * 25.4 / 1000 + 0.1)
                {
                    if (supportingType == "Steel")
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AK4.cs", 105);
                        return;
                    }
                }

                componentDictionary[CROSSBEAM + Index].SetPropertyValue(userLength, "IJOAFINL_UserLength", "UserLength");

                if (assmDescription.IndexOf("YK3") > 0)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.Y, -(pipeDiameter / 2 + shoeHeight), 0, 0);

                        JointHelper.CreatePrismaticJoint(LOGOBJ + Index, "Connection", CROSSBEAM + Index, "Route", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0);

                        JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole1", HEXNUT1 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - hexNutThickness, 0, 0);

                        JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole1", HEXNUT22 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - 2 * hexNutThickness, 0, 0);

                        JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole2", HEXNUT3 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - hexNutThickness, 0, 0);

                        JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole2", HEXNUT4 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - 2 * hexNutThickness, 0, 0);
                    }
                    else
                    {
                        if (supportingType == "Slab")
                        {
                            JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, pipeDiameter / 2 + shoeHeight, 0, 0);

                            JointHelper.CreatePointOnAxisJoint(LOGOBJ + Index, "Connection", "-1", "Route", Axis.NegativeZ);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.Y, -(pipeDiameter / 2 + shoeHeight), 0, 0);

                            JointHelper.CreatePrismaticJoint(LOGOBJ + Index, "Connection", CROSSBEAM + Index, "Route", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0);
                        }
                        JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole1", HEXNUT1 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - hexNutThickness, 0, 0);

                        JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole1", HEXNUT22 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - 2 * hexNutThickness, 0, 0);

                        JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole2", HEXNUT3 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - hexNutThickness, 0, 0);

                        JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole2", HEXNUT4 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - 2 * hexNutThickness, 0, 0);
                    }
                }
                else
                {
                    JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.Y, -(pipeDiameter / 2 + shoeHeight), 0, 0);

                    JointHelper.CreatePrismaticJoint(LOGOBJ + Index, "Connection", CROSSBEAM + Index, "Route", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0);

                    JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole1", HEXNUT1 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - hexNutThickness, 0, 0);

                    JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole1", HEXNUT22 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - 2 * hexNutThickness, 0, 0);

                    JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole2", HEXNUT3 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - hexNutThickness, 0, 0);

                    JointHelper.CreateRigidJoint(CROSSBEAM + Index, "Hole2", HEXNUT4 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, takeOut - 2 * hexNutThickness, 0, 0);
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
                    routeConnections.Add(new ConnectionInfo(CROSSBEAM + Index, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(CROSSBEAM + Index, 1)); // partindex, routeindex

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
