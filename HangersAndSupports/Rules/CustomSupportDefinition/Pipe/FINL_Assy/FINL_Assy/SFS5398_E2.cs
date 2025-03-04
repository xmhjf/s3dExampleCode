//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5398_E2.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5398_E2
//   Author       :  Manikanth
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Manikanth   CR-CP-224492- Convert HS_FINL_Assy to C# .Net 
//  11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    public class SFS5398_E2 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        string size4;
        double span, shoeH, overHang;
        private const string HORSTEEL = "HorSteel";
        private const string VERSTEEL1 = "VerSteel1";
        private const string VERSTEEL2 = "VerSteel2";
        private const int STEELDENSITYKGPERM = 7900;
        PropertyValueCodelist sizeList;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    sizeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize4", "Size4");
                    size4 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;

                    span = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLSpan", "Span")).PropValue;
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;

                    parts.Add(new PartInfo(VERSTEEL1, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(VERSTEEL2, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(HORSTEEL, "Util_Fixed_Box_Metric_1"));
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

                //To get Pipe Nom Dia

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                NominalDiameter currentDiameter = new NominalDiameter();
                currentDiameter = pipeInfo.NominalDiameter;
                double pipeDia = pipeInfo.OutsideDiameter;
                NominalDiameter minNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                NominalDiameter maxNominalDiameter = new NominalDiameter();
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";
                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 10, 1200, 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                //check valid pipe size
                if (IsPipeSizeValid(currentDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5398_E2.cs", 98);
                    return;
                }
                PropertyValue tsSize1, tsSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size4", size4);

                GenericHelper.GetDataByRule("FINLSrv_SFS5398_E2_Dim", "IJUAFINLSrv_SFS5398_E2_Dim", "TS_Size1", parameter, out tsSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5398_E2_Dim", "IJUAFINLSrv_SFS5398_E2_Dim", "TS_Size2", parameter, out tsSize2);
                double maxSpan = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_E2_Dim", "IJUAFINLSrv_SFS5398_E2_Dim", "Max_Span", "IJUAFINLSrv_SFS5398_E2_Dim", "Size4", size4);
                double maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_E2_Dim", "IJUAFINLSrv_SFS5398_E2_Dim", "Max_Len", "IJUAFINLSrv_SFS5398_E2_Dim", "Size4", size4);
                double ma = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_E2_Dim", "IJUAFINLSrv_SFS5398_E2_Dim", "MA", "IJUAFINLSrv_SFS5398_E2_Dim", "Size4", size4);
                double mp = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_E2_Dim", "IJUAFINLSrv_SFS5398_E2_Dim", "MP", "IJUAFINLSrv_SFS5398_E2_Dim", "Size4", size4);
                double fMax = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_E2_Dim", "IJUAFINLSrv_SFS5398_E2_Dim", "Fmax", "IJUAFINLSrv_SFS5398_E2_Dim", "Size4", size4);

                support.SetPropertyValue(ma, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(mp, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(fMax, "IJUAFINLFmax", "Fmax");

                double widthHorTS, flangeThickness, depthhorTS, webThickness;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize1), out widthHorTS, out flangeThickness, out webThickness, out depthhorTS);
                double widthVertTS, depthVertTS, TL;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize2), out widthVertTS, out TL, out webThickness, out depthVertTS);
                double horLength = span;

                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5398_E2.cs", 125);
                    return;
                }

                string horSteelPort = "StartOther";
                if (Configuration == 2)
                {
                    horSteelPort = "EndOther";
                    overHang = -overHang;
                }

                double theta = Math.Abs(Math.PI / 2 - RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "", PortAxisType.Y, OrientationAlong.Global_Z));
                double angledOffset = Math.Tan(theta) * widthHorTS;
                double vertDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                double verticalLength = vertDistance - pipeDia / 2 - shoeH - angledOffset / 2;
                //Checking for max H
                if (verticalLength > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + maxLength * 1000 + "mm exceeded.", "", "SFS5398_E2.cs", 143);
                //If the span length entered is greater than allowed
                if (span > maxSpan)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Maximum allowable length of " + maxSpan * 1000 + "mm exceeded.", "", "SFS5398_E2.cs", 147);
                    return;
                }
                if (overHang > span)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must not exceed the value of Span..", "", "SFS5398_E2.cs", 152);
                    return;
                }

                double vertsSteelWeight = (widthVertTS * depthVertTS * verticalLength - ((widthVertTS - 2 * 5.0 / 1000) * (depthVertTS - 2 * 5.0 / 1000) * verticalLength)) * STEELDENSITYKGPERM;
                double horSteelweight;

                if (size4 == "3")
                    horSteelweight = (widthHorTS * depthhorTS * horLength - ((widthHorTS - 2 * 5.0 / 1000) * (depthhorTS - 2 * 5.0 / 1000) * horLength)) * STEELDENSITYKGPERM;
                else
                    horSteelweight = (widthHorTS * depthhorTS * horLength - ((widthHorTS - 2 * 4.0 / 1000) * (depthhorTS - 2 * 4.0 / 1000) * horLength)) * STEELDENSITYKGPERM;

                string vertBOM = tsSize2 + "Length " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[VERSTEEL1]).SetPropertyValue(widthVertTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[VERSTEEL1]).SetPropertyValue(depthVertTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[VERSTEEL1]).SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[VERSTEEL1]).SetPropertyValue(vertBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[VERSTEEL1]).SetPropertyValue(vertsSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                (componentDictionary[VERSTEEL2]).SetPropertyValue(widthVertTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[VERSTEEL2]).SetPropertyValue(depthVertTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[VERSTEEL2]).SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[VERSTEEL2]).SetPropertyValue(vertBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[VERSTEEL2]).SetPropertyValue(vertsSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string horBOM = tsSize1 + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[HORSTEEL]).SetPropertyValue(widthHorTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL]).SetPropertyValue(depthhorTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL]).SetPropertyValue(horLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL]).SetPropertyValue(horBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL]).SetPropertyValue(horSteelweight, "IJOAHgrUtilMetricInWt", "InputWeight");

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
                    JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, horSteelPort, Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, vertDistance - (pipeDia / 2 + depthhorTS / 2 + shoeH + angledOffset / 2), 0, -overHang);
                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                    {
                       JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, horSteelPort, Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, vertDistance - (pipeDia / 2 + depthhorTS / 2 + shoeH + angledOffset / 2), overHang, 0);
                    }
                    else
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Support can only be placed on Slab By-Point. " + pipeDia * 1000 + shoeH * 1000 + " mm.", "", "SFS5398_E2.cs", 203);
                }
                JointHelper.CreateRigidJoint(HORSTEEL, "StartOther", VERSTEEL1, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, -depthVertTS / 2, 0, -depthhorTS / 2);
                JointHelper.CreateRigidJoint(HORSTEEL, "EndOther", VERSTEEL2, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, depthVertTS / 2, 0, -depthhorTS / 2);
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

                    routeConnections.Add(new ConnectionInfo(HORSTEEL, 1)); // partindex, routeindex


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

                    structConnections.Add(new ConnectionInfo(VERSTEEL1, 1)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(VERSTEEL2, 1)); // partindex, routeindex


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

            sizeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAFINLSize4", "Size4");
            size4 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;
            string BOMdesc = "Gate support E2 -" + size4 + " SFS 5398";

            return BOMdesc;
        }
    }
}