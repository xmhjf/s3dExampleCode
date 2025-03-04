//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5858.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5858
//   Author       :  Manikanth
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who            change description
//   -----------     ---            ------------------
//   20-Jun-2013     Manikanth   CR-CP-224492- Convert HS_FINL_Assy to C# .Net 
//  11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5858 : CustomSupportDefinition, ICustomHgrBOMDescription
    {

        public bool Override { get; set; }
        double rotation, nomPipeDiaMetric;
        string material;
        string VERPLATE = "VerPlate", HORPLATE = "HorPlate", CONNOBJECT = "ConnObj", PIPECLAMP1 = "PipeClamp1", PIPECLAMP2 = "PipeClamp2", ROUTECONNOBJECT = "RouteConnObj";
        public int Clamps { get; set; }
        public int Index { get; set; }

        public override Collection<PartInfo> Parts
        {

            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    if (!Override)
                        if (support.SupportsInterface("IJUAFINLClamps"))
                            Clamps = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLClamps", "Clamps")).PropValue;
                        else
                            Clamps = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLClamps1", "Clamps")).PropValue;
                    if (support.SupportsInterface("IJUAFINLRot"))
                        rotation = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLRot", "Rot")).PropValue;
                    else
                        rotation = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLRot1", "Rot")).PropValue;
                    material = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLMaterial", "Material")).PropValue;

                    //To get Pipe Nom Dia
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    NominalDiameter currentDiameter = new NominalDiameter();
                    currentDiameter = pipeInfo.NominalDiameter;

                    if (currentDiameter.Units != "mm")
                        nomPipeDiaMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, currentDiameter.Size, UnitName.NPD_INCH, UnitName.NPD_MILLIMETER) / 1000;
                    else
                        nomPipeDiaMetric = currentDiameter.Size / 1000;

                    NominalDiameter minNominalDiameter = new NominalDiameter();
                    minNominalDiameter.Size = 10;
                    minNominalDiameter.Units = "mm";
                    NominalDiameter maxNominalDiameter = new NominalDiameter();
                    maxNominalDiameter.Size = 150;
                    maxNominalDiameter.Units = "mm";
                    NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 90 }, "mm");

                    //check valid pipe size
                    if (IsPipeSizeValid(currentDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5858.cs", 86);
                        return parts;
                    }

                    if (Clamps == 1)
                    {
                        parts.Add(new PartInfo(VERPLATE + " _ " + Index, "Utility_USER_FIXED_BOX_1"));
                        parts.Add(new PartInfo(HORPLATE + " _ " + Index, "Utility_USER_FIXED_BOX_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT + " _ " + Index, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(CONNOBJECT + " _ " + Index, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(PIPECLAMP1 + " _ " + Index, "FINLCmp_SFS5370"));
                        parts.Add(new PartInfo(PIPECLAMP2 + " _ " + Index, "FINLCmp_SFS5370"));
                    }
                    else
                    {
                        parts.Add(new PartInfo(VERPLATE + " _ " + Index, "Utility_USER_FIXED_BOX_1"));
                        parts.Add(new PartInfo(HORPLATE + " _ " + Index, "Utility_USER_FIXED_BOX_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT + " _ " + Index, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(CONNOBJECT + " _ " + Index, "Log_Conn_Part_1"));
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                double shoeH = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5858_Dim", "IJUAFINLSrv_SFS5858_Dim", "Shoe_H", "IJUAFINLSrv_SFS5858_Dim", "Pipe_Nom_Dia_m", Convert.ToString(nomPipeDiaMetric * 1000));
                double B = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5858_Dim", "IJUAFINLSrv_SFS5858_Dim", "B", "IJUAFINLSrv_SFS5858_Dim", "Pipe_Nom_Dia_m", Convert.ToString(nomPipeDiaMetric * 1000));
                double T = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5858_Dim", "IJUAFINLSrv_SFS5858_Dim", "T", "IJUAFINLSrv_SFS5858_Dim", "Pipe_Nom_Dia_m", Convert.ToString(nomPipeDiaMetric * 1000));
                double L = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5858_Dim", "IJUAFINLSrv_SFS5858_Dim", "L", "IJUAFINLSrv_SFS5858_Dim", "Pipe_Nom_Dia_m", Convert.ToString(nomPipeDiaMetric * 1000));

                double clampThickness = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_PipeND_mm", "PipeND", nomPipeDiaMetric - 0.003, nomPipeDiaMetric + 0.003);
                double clampWidth = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAFINL_Width", "Width", "IJUAFINL_PipeND_mm", "PipeND", nomPipeDiaMetric - 0.003, nomPipeDiaMetric + 0.003);
                double clampInnerDia = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAHgrPipe_Dia", "PIPE_DIA", "IJUAFINL_PipeND_mm", "PipeND", nomPipeDiaMetric - 0.003, nomPipeDiaMetric + 0.003);

                string vertPlateBOM = "Plate SFS 2022 " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, shoeH - clampThickness - T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[VERPLATE + " _ " + Index]).SetPropertyValue(shoeH - clampThickness - T, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                (componentDictionary[VERPLATE + " _ " + Index]).SetPropertyValue(T, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                (componentDictionary[VERPLATE + " _ " + Index]).SetPropertyValue(L, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                (componentDictionary[VERPLATE + " _ " + Index]).SetPropertyValue(vertPlateBOM, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                string horPlateBOM = "Plate SFS 2022 " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, B, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[HORPLATE + " _ " + Index]).SetPropertyValue(T, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                (componentDictionary[HORPLATE + " _ " + Index]).SetPropertyValue(B, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                (componentDictionary[HORPLATE + " _ " + Index]).SetPropertyValue(L, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                (componentDictionary[HORPLATE + " _ " + Index]).SetPropertyValue(horPlateBOM, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                if (!Override)
                    JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT + " _ " + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);


                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", VERPLATE + " _ " + Index, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, clampInnerDia / 2 + clampThickness, 0, 0);

                JointHelper.CreateRigidJoint(VERPLATE + " _ " + Index, "EndOther", HORPLATE + " _ " + Index, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                if (Clamps == 1)
                {
                    double hyp, X, Y, varAngle;

                    varAngle = rotation * (180 / Math.PI);

                    if (Math.Abs(varAngle) > 40 + 0.0001 && Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(Math.Abs(varAngle)), 45) == false)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "WARNING: " + "Clamp rotation cannot exceed 40 degrees..", "", "SFS5858.cs", 162);

                    if (Configuration == 2)             //if toggled, mirror the support
                        varAngle = -varAngle;

                    hyp = 0.1;    // This is just an arbitrary number.
                    X = Math.Sin(rotation - Math.PI / 2) * hyp;
                    Y = Math.Cos(rotation - Math.PI / 2) * hyp;


                    //Add a joint between Connection 1 and the clamp pin so the clamp can spin
                    JointHelper.CreateRevoluteJoint(CONNOBJECT + " _ " + Index, "Connection", PIPECLAMP1 + " _ " + Index, "Pin", Axis.X, Axis.Y);
                    //Add a joint between Route and the Clamp center so the clamp can spin
                    JointHelper.CreateCylindricalJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", PIPECLAMP1 + " _ " + Index, "Route", Axis.X, Axis.Y, 0);
                    //Add a flexible joint to the clamp
                    JointHelper.CreateCylindricalJoint(PIPECLAMP1 + " _ " + Index, "Route", PIPECLAMP1 + " _ " + Index, "Pin", Axis.Z, Axis.Z, 0);

                    //set material attribute
                    (componentDictionary[PIPECLAMP1 + " _ " + Index]).SetPropertyValue(material, "IJOAFINL_Material", "Material");
                    (componentDictionary[PIPECLAMP2 + " _ " + Index]).SetPropertyValue(material, "IJOAFINL_Material", "Material");

                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", CONNOBJECT + " _ " + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, Y, X, -L / 2 + clampWidth / 2);
                    JointHelper.CreateRigidJoint(PIPECLAMP1 + " _ " + Index, "Route", PIPECLAMP2 + " _ " + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, L - clampWidth, 0);

                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of FINL_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                return 2;
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

                    routeConnections.Add(new ConnectionInfo(ROUTECONNOBJECT + " _ " + Index, 1)); // partindex, routeindex


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

                    structConnections.Add(new ConnectionInfo(HORPLATE + " _ " + Index, 1)); // partindex, routeindex

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

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            double angle = rotation * (180 / Math.PI);
            string BOMdesc;

            PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
            NominalDiameter currentDiameter = new NominalDiameter();
            currentDiameter = pipeInfo.NominalDiameter;
            if (Symbols.HgrCompareDoubleService.cmpdbl(rotation, 0) == true)
                BOMdesc = "Pipe slide SFS 5858 DN " + currentDiameter;
            else
                BOMdesc = "Pipe slide SFS 5858 DN " + currentDiameter + " - " + angle;

            return BOMdesc;
        }
    }
}