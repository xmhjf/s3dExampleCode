//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5396_B.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5396_B
//   Author       :  Vijaya
//   Creation Date:  26.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   26-Jun-2013     Vijaya   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
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
    public class SFS5396_B : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string PIPECLAMP5396B = "PIPECLAMP5396_B";
        private const string TSSECTION5396B = "TSSECTION5396_B";
        private const string PLATE4HOLE5396B = "PLATE4HOLE5396_B";
        private const string PLATE5396B = "PLATE5396_B";
        private int footTypeB;
        static double pipeDiameterMetric = 0.0, structAngle, structDistance;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist footTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLFootTypeB", "FootTypeB");
                    footTypeB = footTypeCodeList.PropValue;

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

                    if (pipeSizeValid == false)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5396_B.cs", 80);
                        return parts;
                    }

                    parts.Add(new PartInfo(PIPECLAMP5396B, "FINLCmp_SFS5370"));
                    parts.Add(new PartInfo(TSSECTION5396B, "Utility_USER_FIXED_BOX_1"));

                    if (footTypeB == 1)// with clamps                    
                        parts.Add(new PartInfo(PLATE4HOLE5396B, "Util_FourHolePl_Metric_1"));
                    else if (footTypeB == 2)
                        parts.Add(new PartInfo(PLATE5396B, "Util_Plate_Metric_1"));

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
                double P = 0.0, L = 0.0, B = 0.0, T = 0.0, C = 0.0, D = 0.0, tsWidth, tsFlangeThickness, tsWebThickness, tsDepth, tsLength, inset = 0.0;

                PropertyValue TSSize;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                GenericHelper.GetDataByRule("FINLSrv_SFS5396_B_Dim", "IJUAFINLSrv_SFS5396_B_Dim", "TS_Size", parameter, out TSSize);
                string tsSize = TSSize.ToString(), plateBOM = string.Empty;


                P = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_B_Dim", "IJUAFINLSrv_SFS5396_B_Dim", "P", "IJUAFINLSrv_SFS5396_B_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                L = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_B_Dim", "IJUAFINLSrv_SFS5396_B_Dim", "L", "IJUAFINLSrv_SFS5396_B_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                B = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_B_Dim", "IJUAFINLSrv_SFS5396_B_Dim", "B", "IJUAFINLSrv_SFS5396_B_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                T = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_B_Dim", "IJUAFINLSrv_SFS5396_B_Dim", "T", "IJUAFINLSrv_SFS5396_B_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                C = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_B_Dim", "IJUAFINLSrv_SFS5396_B_Dim", "C", "IJUAFINLSrv_SFS5396_B_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());
                D = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5396_B_Dim", "IJUAFINLSrv_SFS5396_B_Dim", "D", "IJUAFINLSrv_SFS5396_B_Dim", "Pipe_Nom_Dia_m", pipeDiameterMetric.ToString().Trim());

                //Get Crosssection Values                
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", tsSize, out  tsWidth, out   tsFlangeThickness, out  tsWebThickness, out  tsDepth);


                //Check the Structure is beside
                structAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, OrientationAlong.Global_Z);


                if (structAngle < Math.PI / 4)
                    structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                else if (structAngle > 3 * Math.PI / 4)
                    structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                else
                    structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);


                JointHelper.CreateRigidJoint("-1", "Route", PIPECLAMP5396B, "Route", Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeX, 0, 0, 0);
                if (footTypeB == 1)
                {
                    tsLength = structDistance - P - T;
                    inset = (L - C) / 2;
                    plateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " +
                        MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" +
                    MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, B, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + ", 4 eq spaced " +
                    MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, D, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " dia. holes " +
                    MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, inset, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " from edges";

                    //four hole plate
                    componentDictionary[PLATE4HOLE5396B].SetPropertyValue(L, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[PLATE4HOLE5396B].SetPropertyValue(B, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[PLATE4HOLE5396B].SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                    componentDictionary[PLATE4HOLE5396B].SetPropertyValue(D, "IJOAHgrUtilMetricHoleSize", "HoleSize");
                    componentDictionary[PLATE4HOLE5396B].SetPropertyValue(inset, "IJOAHgrUtilMetricC", "C");
                    componentDictionary[PLATE4HOLE5396B].SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                    JointHelper.CreateRigidJoint(TSSECTION5396B, "EndOther", PLATE4HOLE5396B, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0, 0.0);

                }
                else if (footTypeB == 2)
                {
                    tsLength = structDistance - P - T;
                    plateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " +
                      MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" +
                      MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, B, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    //four hole plate
                    componentDictionary[PLATE5396B].SetPropertyValue(L, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[PLATE5396B].SetPropertyValue(B, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[PLATE5396B].SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                    componentDictionary[PLATE5396B].SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                    JointHelper.CreateRigidJoint(TSSECTION5396B, "EndOther", PLATE5396B, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0, 0.0);

                }
                else
                    tsLength = structDistance - P;

                string tSectionBOM = tsSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, Math.Round(tsLength, 3), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[TSSECTION5396B].SetPropertyValue(tsLength, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                componentDictionary[TSSECTION5396B].SetPropertyValue(tsWidth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                componentDictionary[TSSECTION5396B].SetPropertyValue(tsDepth, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                componentDictionary[TSSECTION5396B].SetPropertyValue(tSectionBOM, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                JointHelper.CreateRigidJoint("-1", "Route", TSSECTION5396B, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, P, 0.0, 0.0);

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

                    routeConnections.Add(new ConnectionInfo(PIPECLAMP5396B, 1));

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

                    structConnections.Add(new ConnectionInfo(PIPECLAMP5396B, 1));

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

            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                PropertyValueCodelist footTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLFootTypeB", "FootTypeB");
                string footType = footTypeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(footTypeCodeList.PropValue).DisplayName;
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                return "Foot SFS 5396 " + footType + " DN " + pipeInfo.NominalDiameter.Size + " x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, Math.Round(structDistance, 3), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5396_B" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

