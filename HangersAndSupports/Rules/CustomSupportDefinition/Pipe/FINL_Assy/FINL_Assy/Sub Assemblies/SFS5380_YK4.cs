//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_YK4.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_YK4
//   Author       : Rajeswari
//   Creation Date: 27-June-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 01-July-2013  Rajeswari CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;

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
    public class SFS5380_YK4 : CustomSupportDefinition
    {
        private const string ROD = "ROD_YK4";
        private const string WASHERPLATE = "WASHERPLATE_YK4";
        private const string HEXNUT1 = "HEXNUT1_YK4";
        private const string HEXNUT2 = "HEXNUT2_YK4";
        private const string LOGOBJ = "LOGOBJ_YK4";
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

                    parts.Add(new PartInfo(ROD + Index, "FINLCmp_SFS5381_" + loadClass));
                    parts.Add(new PartInfo(WASHERPLATE + Index, "FINLCmp_SFS4683_" + loadClass));
                    parts.Add(new PartInfo(HEXNUT1 + Index, "FINLCmp_SFS4032_" + loadClass));
                    parts.Add(new PartInfo(HEXNUT2 + Index, "FINLCmp_SFS4032_" + loadClass));
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

                double takeOut = 0, boltHeight = 0, plateThickness = 0;
                double lengthV = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                takeOut = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5381", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                boltHeight = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS4032", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                plateThickness = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS4683", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());

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

                if (assmDescription.IndexOf("AK2") > 0 || assmDescription.IndexOf("AK4") > 0 || assmDescription.IndexOf("AR2") > 0 || assmDescription.IndexOf("AR4") > 0)
                {
                  
                    int currentFaceNumber;

                    if (SupportHelper.SupportingObjects.Count != 0)
                    {
                        try
                        {
                            currentFaceNumber = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                        }
                        catch
                        {
                            currentFaceNumber = 0;
                        }
                    }
                    else
                        currentFaceNumber = -1;

                    double routeXStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);
                    if (supportingType == "Slab")
                    {
                        if (Math.Abs((routeXStructAngle) * 180 / Math.PI - 90) < 2)
                        {
                            if (currentFaceNumber == 514)
                                JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, 0, 0, 0);
                        }
                        else
                            JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        if (assmDescription.IndexOf("AK2") > 0 || assmDescription.IndexOf("AR2") > 0)
                            JointHelper.CreatePlanarJoint(WASHERPLATE + Index, "TopStructure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePrismaticJoint(WASHERPLATE + Index, "TopStructure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        if (assmDescription.IndexOf("AK2") > 0 || assmDescription.IndexOf("AR2") > 0)
                            JointHelper.CreatePlanarJoint(WASHERPLATE + Index, "TopStructure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePrismaticJoint(WASHERPLATE + Index, "TopStructure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                        if (assmDescription.IndexOf("AK2") > 0 || assmDescription.IndexOf("AR2") > 0)
                            JointHelper.CreatePlanarJoint(WASHERPLATE + Index, "TopStructure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePrismaticJoint(WASHERPLATE + Index, "TopStructure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    }
                    JointHelper.CreateRigidJoint(WASHERPLATE + Index, "Hole", ROD + Index, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -plateThickness / 2 - takeOut, 0, 0);

                    JointHelper.CreateRigidJoint(ROD + Index, "TopExThdRH", HEXNUT1 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -boltHeight + takeOut, 0, 0);

                    JointHelper.CreateRigidJoint(ROD + Index, "TopExThdRH", HEXNUT2 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -boltHeight + takeOut - boltHeight, 0, 0);

                    JointHelper.CreatePrismaticJoint(ROD + Index, "BotExThdRH", ROD + Index, "TopExThdRH", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);
                }
                else
                {
                    componentDictionary[ROD + Index].SetPropertyValue(lengthV + plateThickness / 2 + takeOut, "IJUAHgrOccLength", "Length");

                    JointHelper.CreateRigidJoint("-1", "Structure", WASHERPLATE + Index, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    JointHelper.CreateRigidJoint(WASHERPLATE + Index, "Hole", ROD + Index, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -plateThickness / 2 - takeOut, 0, 0);

                    JointHelper.CreateRigidJoint(ROD + Index, "TopExThdRH", HEXNUT1 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -boltHeight + takeOut, 0, 0);

                    JointHelper.CreateRigidJoint(ROD + Index, "TopExThdRH", HEXNUT2 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -boltHeight + takeOut - boltHeight, 0, 0);

                    JointHelper.CreatePrismaticJoint(ROD + Index, "BotExThdRH", ROD + Index, "TopExThdRH", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);
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
                    structConnections.Add(new ConnectionInfo(WASHERPLATE + Index, 1)); // partindex, routeindex

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
