//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5860.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5860
//   Author       :  Manikanth
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
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
    public class SFS5860 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        public bool Override { get; set; }
        double rotation, dNomPipeDiaMetric;
        string material;
        string ENDPLATE1 = "EndPlate1", ENDPLATE2 = "EndPlate2", CSECTION = "CSection", CONNOBJECT = "ConnObj", PIPECLAMP1 = "PipeClamp1", PIPECLAMP2 = "PipeClamp2", ROUTECONNOBJECT = "RouteConnObj";
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

                    if (Override == false)
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
                        dNomPipeDiaMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, currentDiameter.Size, UnitName.NPD_INCH, UnitName.NPD_MILLIMETER) / 1000;
                    else
                        dNomPipeDiaMetric = currentDiameter.Size / 1000;

                    NominalDiameter minNominalDiameter = new NominalDiameter();
                    minNominalDiameter.Size = 600;
                    minNominalDiameter.Units = "mm";
                    NominalDiameter maxNominalDiameter = new NominalDiameter();
                    maxNominalDiameter.Size = 1200;
                    maxNominalDiameter.Units = "mm";
                    NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 1100 }, "mm");

                    //check valid pipe size
                    if (IsPipeSizeValid(currentDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5860.cs", 83);
                        return parts;
                    }

                    if (Clamps == 1)
                    {
                        parts.Add(new PartInfo(ENDPLATE1 + " _ " + Index, "Util_End_Plate_Metric"));
                        parts.Add(new PartInfo(ENDPLATE2 + " _ " + Index, "Util_End_Plate_Metric"));
                        parts.Add(new PartInfo(CSECTION + " _ " + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT + " _ " + Index, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(CONNOBJECT + " _ " + Index, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(PIPECLAMP1 + " _ " + Index, "FINLCmp_SFS5856"));
                        parts.Add(new PartInfo(PIPECLAMP2 + " _ " + Index, "FINLCmp_SFS5856"));
                    }
                    else
                    {
                        parts.Add(new PartInfo(ENDPLATE1 + " _ " + Index, "Util_End_Plate_Metric"));
                        parts.Add(new PartInfo(ENDPLATE2 + " _ " + Index, "Util_End_Plate_Metric"));
                        parts.Add(new PartInfo(CSECTION + " _ " + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT + " _ " + Index, "Log_Conn_Part_1"));
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

                double shoeH = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5860_Dim", "IJUAFINLSrv_SFS5860_Dim", "Shoe_H", "IJUAFINLSrv_SFS5860_Dim", "Pipe_Nom_Dia_m", Convert.ToString(dNomPipeDiaMetric * 1000));
                double K = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5860_Dim", "IJUAFINLSrv_SFS5860_Dim", "K", "IJUAFINLSrv_SFS5860_Dim", "Pipe_Nom_Dia_m", Convert.ToString(dNomPipeDiaMetric * 1000));
                double T = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5860_Dim", "IJUAFINLSrv_SFS5860_Dim", "T", "IJUAFINLSrv_SFS5860_Dim", "Pipe_Nom_Dia_m", Convert.ToString(dNomPipeDiaMetric * 1000));
                double A = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5860_Dim", "IJUAFINLSrv_SFS5860_Dim", "A", "IJUAFINLSrv_SFS5860_Dim", "Pipe_Nom_Dia_m", Convert.ToString(dNomPipeDiaMetric * 1000));
                double L = 500.0 / 1000;

                double clampThickness = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5856", "IJUAFINL_T", "T", "IJUAFINL_PipeND_mm", "PipeND", dNomPipeDiaMetric - 0.003, dNomPipeDiaMetric + 0.003);
                double clampWidth = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5856", "IJUAFINL_B", "B", "IJUAFINL_PipeND_mm", "PipeND", dNomPipeDiaMetric - 0.003, dNomPipeDiaMetric + 0.003);
                double clampInnerDia = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5856", "IJUAHgrPipe_Dia", "PIPE_DIA", "IJUAFINL_PipeND_mm", "PipeND", dNomPipeDiaMetric - 0.003, dNomPipeDiaMetric + 0.003);
                double alpha, wM, aM, hM, radiusM, offsetM;

                hM = shoeH - A - clampThickness;
                radiusM = clampInnerDia / 2 + clampThickness;
                wM = K - 2 * T;

                alpha = Math.Acos(2.0 / 3);
                if (wM <= radiusM * 2)
                    alpha = Math.Acos(((wM / 2.0) / radiusM)) * 180 / Math.PI;

                aM = radiusM - (radiusM * Math.Sin(alpha * 1.74532925199433E-02));

                if (radiusM < wM / 2)
                    aM = radiusM;

                offsetM = Math.Sqrt(Math.Pow(clampInnerDia / 2 + clampThickness, 2) - Math.Pow(K / 2 - T, 2));

                string endPlateBom = "End Plate - Pipe-Base L " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, hM + aM - (radiusM - offsetM), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + ", Overall L " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, Math.Round((hM + aM), 3), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + ", " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, K - 2 * T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Radius " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, clampInnerDia / 2 + clampThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[ENDPLATE1 + " _ " + Index]).SetPropertyValue(clampInnerDia / 2 + clampThickness, "IJOAHgrUtilMetricRadius", "Radius");
                (componentDictionary[ENDPLATE1 + " _ " + Index]).SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                (componentDictionary[ENDPLATE1 + " _ " + Index]).SetPropertyValue(K - 2 * T, "IJOAHgrUtilMetricW", "W");
                (componentDictionary[ENDPLATE1 + " _ " + Index]).SetPropertyValue(shoeH - A - clampThickness, "IJOAHgrUtilMetricH", "H");
                (componentDictionary[ENDPLATE1 + " _ " + Index]).SetPropertyValue(endPlateBom, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                (componentDictionary[ENDPLATE2 + " _ " + Index]).SetPropertyValue(clampInnerDia / 2 + clampThickness, "IJOAHgrUtilMetricRadius", "Radius");
                (componentDictionary[ENDPLATE2 + " _ " + Index]).SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                (componentDictionary[ENDPLATE2 + " _ " + Index]).SetPropertyValue(K - 2 * T, "IJOAHgrUtilMetricW", "W");
                (componentDictionary[ENDPLATE2 + " _ " + Index]).SetPropertyValue(shoeH - A - clampThickness, "IJOAHgrUtilMetricH", "H");
                (componentDictionary[ENDPLATE2 + " _ " + Index]).SetPropertyValue(endPlateBom, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                if (!Override)  //this will be overriden by super assembly
                    JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT + " _ " + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                //Add the joints between route and vertical plate
                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", ENDPLATE1 + " _ " + Index, "Route", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, 0, 0, L / 2 - clampWidth / 2 + T / 2);
                //Add the joints between vertical plate and horizontal plate
                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", ENDPLATE2 + " _ " + Index, "Route", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, 0, 0, -L / 2 + clampWidth / 2 + T / 2);

                double offset = Math.Sqrt(Math.Pow(clampInnerDia / 2 + clampThickness, 2) - Math.Pow(K / 2, 2));
                //Setting attribute of vertical plate  L, WIDTH, DEPTH, BOM_DESC
                double dF = clampInnerDia / 2 + shoeH - offset;

                string CSectionBOM = "Plate SFS 2022 " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, 2 * dF + K - 2 * T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "X " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[CSECTION + " _ " + Index]).SetPropertyValue(L, "IJOAHgrUtility_L", "L");
                (componentDictionary[CSECTION + " _ " + Index]).SetPropertyValue(K, "IJOAHgrUtility_Depth", "Depth");
                (componentDictionary[CSECTION + " _ " + Index]).SetPropertyValue(dF, "IJOAHgrUtility_Width", "Width");
                (componentDictionary[CSECTION + " _ " + Index]).SetPropertyValue(T, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                (componentDictionary[CSECTION + " _ " + Index]).SetPropertyValue(T, "IJOAHgrUtility_WebTh", "WebTh");
                (componentDictionary[CSECTION + " _ " + Index]).SetPropertyValue(CSectionBOM, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", CSECTION + " _ " + Index, "EndCap", Plane.YZ, Plane.YZ, Axis.Z, Axis.NegativeZ, L / 2, K / 2, dF + offset);

                if (Clamps == 1)
                {
                    double hyp, X, Y, varAngle;

                    varAngle = rotation * (180 / Math.PI);

                    if (Math.Abs(varAngle) > 40 + 0.0001 && Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(Math.Abs(varAngle)), 45) == false)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "WARNING: " + "Clamp rotation cannot exceed 40 degrees..", "", "SFS5860.cs", 191);

                    if (Configuration == 2)             //if toggled, mirror the support
                        varAngle = -varAngle;

                    hyp = 0.1;    // This is just an arbitrary number.
                    X = Math.Sin(rotation - Math.PI / 2) * hyp;
                    Y = Math.Cos(rotation - Math.PI / 2) * hyp;


                    //Add a joint between Connection 1 and the clamp pin so the clamp can spin
                    JointHelper.CreateRevoluteJoint(CONNOBJECT + " _ " + Index, "Connection", PIPECLAMP1 + " _ " + Index, "Pin1", Axis.X, Axis.Y);
                    //Add a joint between Route and the Clamp center so the clamp can spin
                    JointHelper.CreateCylindricalJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", PIPECLAMP1 + " _ " + Index, "Route", Axis.X, Axis.Y, 0);
                    //Add a flexible joint to the clamp
                    JointHelper.CreateCylindricalJoint(PIPECLAMP1 + " _ " + Index, "Route", PIPECLAMP1 + " _ " + Index, "Pin1", Axis.Z, Axis.Z, 0);

                    if (dNomPipeDiaMetric > 49.00 / 1000.00)
                    {
                        //set material attribute
                        (componentDictionary[PIPECLAMP1 + " _ " + Index]).SetPropertyValue(material, "IJOAFINL_Material", "Material");
                        (componentDictionary[PIPECLAMP2 + " _ " + Index]).SetPropertyValue(material, "IJOAFINL_Material", "Material");

                        //Add the joints to the clamps
                        JointHelper.CreateRigidJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", CONNOBJECT + " _ " + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, Y, X, -L / 2 + clampWidth / 2);
                        JointHelper.CreateRigidJoint(PIPECLAMP1 + " _ " + Index, "Route", PIPECLAMP2 + " _ " + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, L - clampWidth, 0);
                    }
                    else
                    {
                        (componentDictionary[PIPECLAMP1 + " _ " + Index]).SetPropertyValue(material, "IJOAFINL_Material", "Material");
                        JointHelper.CreateRigidJoint(ROUTECONNOBJECT + " _ " + Index, "Connection", CONNOBJECT + " _ " + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, Y, X, 0);
                    }
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

                    structConnections.Add(new ConnectionInfo(CSECTION + " _ " + Index, 1)); // partindex, routeindex


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
            try
            {
                double angle = rotation * (180 / Math.PI);
                string BOMdesc;
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                NominalDiameter currentDiameter = new NominalDiameter();
                currentDiameter = pipeInfo.NominalDiameter;
                if (Symbols.HgrCompareDoubleService.cmpdbl(rotation, 0) == true)
                    BOMdesc = "Pipe slide SFS 5860 DN " + currentDiameter;
                else
                    BOMdesc = "Pipe slide SFS 5860 DN " + currentDiameter + " -" + angle;

                return BOMdesc;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}