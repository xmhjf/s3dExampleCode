//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5396_C.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5396_C
//   Author       :  Vijaya
//   Creation Date:  26.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   26-Jun-2013     Vijaya   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;

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
    public class SFS5396_C : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string PIPE5396C = "PIPECLAMP5396_C";
        private const string PLATE4HOLE5396C = "PLATE4HOLE5396_C";
        private const string PLATE5396C = "PLATE5396_C";
        private int footTypeC;
        static double pipeDiameterMetric, structDistance;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist footTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLFootTypeC", "FootTypeC");
                    footTypeC = footTypeCodeList.PropValue;

                    // To get Pipe Nominal Diameter
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    NominalDiameter currentDiameter = new NominalDiameter(), minNominalDiameter = new NominalDiameter(), maxNominalDiameter = new NominalDiameter();
                    currentDiameter = pipeInfo.NominalDiameter;


                    //to change NomPipeDia to metric unit
                    if (pipeInfo.NominalDiameter.Units != "mm")
                        pipeDiameterMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, currentDiameter.Size, UnitName.NPD_INCH, UnitName.NPD_MILLIMETER);
                    else
                        pipeDiameterMetric = currentDiameter.Size / 1000;
                    if (pipeInfo.NominalDiameter.Units != "mm")
                        currentDiameter.Size = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, currentDiameter.Size, UnitName.NPD_INCH, UnitName.NPD_MILLIMETER);

                    minNominalDiameter.Size = 50;
                    minNominalDiameter.Units = "mm";
                    maxNominalDiameter.Size = 500;
                    maxNominalDiameter.Units = "mm";

                    bool pipeSizeValid = true;
                    if (currentDiameter.Size < minNominalDiameter.Size || currentDiameter.Size > maxNominalDiameter.Size)
                        pipeSizeValid = false;

                    if (!pipeSizeValid)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5396_C.cs", 79);
                        return parts;
                    }

                    parts.Add(new PartInfo(PIPE5396C, "Utility_USER_FIXED_CYL_1"));

                    if (footTypeC == 1)// with clamps
                        parts.Add(new PartInfo(PLATE4HOLE5396C, "Util_FourHolePl_Metric_1"));

                    else if (footTypeC == 2)
                        parts.Add(new PartInfo(PLATE5396C, "Util_Plate_Metric_1"));

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

                if (footTypeC == 1)    // on slab
                {
                    if (supportingType == "Steel")
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Support is not to be placed on steel.", "", "SFS5396_C.cs", 127);
                }
                else if (footTypeC != 2)
                {    // on steel
                    if (supportingType == "Slab")
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Support is not to be placed on Slab.", "", "SFS5396_C.cs", 132);
                }

                double P = 0.0, L = 0.0, T = 0.0, C = 0.0, D = 0.0, PipeDia2 = 0.0, pipeLength, inset = 0.0;

                PropertyValue PipeSizeA, PipeSizeB;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                GenericHelper.GetDataByRule("FINLSrv_SFS5396_C_Dim", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Size_A", parameter, out PipeSizeA);
                GenericHelper.GetDataByRule("FINLSrv_SFS5396_C_Dim", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Size_B", parameter, out PipeSizeB);
                string pipeSizeA = PipeSizeA.ToString(), pipeSizeB = PipeSizeB.ToString(), plateBOM = string.Empty, pipeSize = string.Empty;



                P = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_C_Dim", "IJUAFINLSrv_SFS5396_C_Dim", "P", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                L = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_C_Dim", "IJUAFINLSrv_SFS5396_C_Dim", "L", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                T = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_C_Dim", "IJUAFINLSrv_SFS5396_C_Dim", "T", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                C = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_C_Dim", "IJUAFINLSrv_SFS5396_C_Dim", "C", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                D = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_C_Dim", "IJUAFINLSrv_SFS5396_C_Dim", "D", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                PipeDia2 = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_C_Dim", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Dia2", "IJUAFINLSrv_SFS5396_C_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());

                structDistance = RefPortHelper.DistanceBetweenPorts("TurnRef", "Structure", PortDistanceType.Vertical);

                if (footTypeC == 1)
                {
                    pipeLength = structDistance + P - T;
                    inset = (L - C) / 2;
                    plateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " +
                        MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" +
                    MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + ", 4 eq spaced " +
                    MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, D, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " dia. holes " +
                    MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, inset, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " from edges";

                    //four hole plate
                    componentDictionary[PLATE4HOLE5396C].SetPropertyValue(L, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[PLATE4HOLE5396C].SetPropertyValue(L, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[PLATE4HOLE5396C].SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                    componentDictionary[PLATE4HOLE5396C].SetPropertyValue(D, "IJOAHgrUtilMetricHoleSize", "HoleSize");
                    componentDictionary[PLATE4HOLE5396C].SetPropertyValue(inset, "IJOAHgrUtilMetricC", "C");
                    componentDictionary[PLATE4HOLE5396C].SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                    JointHelper.CreateRigidJoint(PIPE5396C, "StartOther", PLATE4HOLE5396C, "BotStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0, 0.0);

                }
                else if (footTypeC == 2)
                {
                    pipeLength = structDistance + P - T;
                    plateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " +
                      MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" +
                      MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    //four hole plate
                    componentDictionary[PLATE5396C].SetPropertyValue(L, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[PLATE5396C].SetPropertyValue(L, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[PLATE5396C].SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                    componentDictionary[PLATE5396C].SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                    JointHelper.CreateRigidJoint(PIPE5396C, "StartOther", PLATE5396C, "BotStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0, 0.0);

                }
                else
                    pipeLength = structDistance + P;

                string Category = SupportedHelper.SupportedObjectInfo(1).MaterialCategory;

                if (Category.ToUpper() == "CARBON STEELS") //  //carbon steel
                {
                    support.SetPropertyValue("Carbon Steel SFS 2005", "IJUAFINLMaterial", "Material");
                    pipeSize = "Carbon Steel Pipe " + pipeSizeA;

                }
                else if (Category.ToUpper() == "STAINLESS STEELS")    //stainles steel
                {
                    support.SetPropertyValue("Stainless Steel SFS 4161", "IJUAFINLMaterial", "Material");
                    pipeSize = "Stainless Steel Pipe " + pipeSizeB;
                }
                else
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Support is not to be placed on Slab.", "", "SFS5396_C.cs", 208);

                string pipeBom = pipeSize + ", Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pipeLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[PIPE5396C].SetPropertyValue(pipeLength, "IJOAHgrUtility_USER_FIXED_CYL", "L");
                componentDictionary[PIPE5396C].SetPropertyValue(PipeDia2 / 2, "IJOAHgrUtility_USER_FIXED_CYL", "RADIUS");
                componentDictionary[PIPE5396C].SetPropertyValue(pipeBom, "IJOAHgrUtility_USER_FIXED_CYL", "BOM_DESC");

                JointHelper.CreateRigidJoint("-1", "TurnRef", PIPE5396C, "EndOther", Plane.XY, Plane.XY, Axis.X, Axis.X, P, 0.0, 0.0);

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of FINL_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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

                    routeConnections.Add(new ConnectionInfo(PIPE5396C, 1));

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

                    structConnections.Add(new ConnectionInfo(PIPE5396C, 1));

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
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDescription;
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                PropertyValueCodelist footTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLFootTypeC", "FootTypeC");
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                string footTypeC = footTypeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(footTypeCodeList.PropValue).DisplayName;

                bomDescription = "Foot SFS 5396 " + footTypeC + " DN " + pipeInfo.NominalDiameter.Size + " x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, Math.Round(structDistance, 3), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                return bomDescription;

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5396_C" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

