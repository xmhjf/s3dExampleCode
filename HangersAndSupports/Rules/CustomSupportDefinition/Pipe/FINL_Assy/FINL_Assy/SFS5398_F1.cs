//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5398_F1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5398_F1
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
using Ingr.SP3D.ReferenceData.Middle.Services;
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
    public class SFS5398_F1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        string size3, material;
        double span, shoeH, overHang;
        private const string VERSTEEL = "VerSteel";
        private const string HORSTEEL = "HorSteel";
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

                    sizeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");
                    size3 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;

                    span = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLSpan", "Span")).PropValue;
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;
                    material = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLMaterial", "Material")).PropValue;

                    parts.Add(new PartInfo(VERSTEEL, "Util_Fixed_Box_Metric_1"));
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
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5398_F1.cs", 98);
                    return;
                }


                CatalogStructHelper catalogHelper = new CatalogStructHelper();
                PropertyValue tsSize;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size3", size3);

                GenericHelper.GetDataByRule("FINLSrv_SFS5398_F1_Dim", "IJUAFINLSrv_SFS5398_F1_Dim", "TS_Size", parameter, out tsSize);
                double maxSpan = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_F1_Dim", "IJUAFINLSrv_SFS5398_F1_Dim", "Max_Span", "IJUAFINLSrv_SFS5398_F1_Dim", "Size3", size3);
                double maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_F1_Dim", "IJUAFINLSrv_SFS5398_F1_Dim", "Max_Len", "IJUAFINLSrv_SFS5398_F1_Dim", "Size3", size3);
                double ma = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_F1_Dim", "IJUAFINLSrv_SFS5398_F1_Dim", "MA", "IJUAFINLSrv_SFS5398_F1_Dim", "Size3", size3);
                double mp = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_F1_Dim", "IJUAFINLSrv_SFS5398_F1_Dim", "MP", "IJUAFINLSrv_SFS5398_F1_Dim", "Size3", size3);
                double fMax = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_F1_Dim", "IJUAFINLSrv_SFS5398_F1_Dim", "Fmax", "IJUAFINLSrv_SFS5398_F1_Dim", "Size3", size3);

                support.SetPropertyValue(ma, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(mp, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(fMax, "IJUAFINLFmax", "Fmax");

                double widthHorTS, flangeThickness, depthhorTS, webThickness;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize), out widthHorTS, out flangeThickness, out webThickness, out depthhorTS);
                double widthVertTS, depthVertTS, TL;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize), out widthVertTS, out TL, out webThickness, out depthVertTS);

                double horLength = span;

                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5398_F1.cs", 128);
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

                double verticalLength = vertDistance + pipeDia / 2 + shoeH - angledOffset / 2;
                //Checking for max H
                if (verticalLength > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + maxLength * 1000 + "mm exceeded.", "", "SFS5398_F1.cs", 146);
                //If the span length entered is greater than allowed
                if (horLength > maxSpan)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Maximum allowable length of " + maxSpan * 1000 + "mm exceeded.", "", "SFS5398_F1.cs", 150);
                    return;
                }
                if (overHang + pipeDia / 2 + widthVertTS / 2 > horLength / 2)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang cannot exceed " + (horLength / 2 - pipeDia / 2 - widthVertTS / 2) * 1000 + " mm..", "", "SFS5398_F1.cs", 155);
                    return;
                }
                double horSteelweight = (widthHorTS * depthhorTS * horLength - ((widthHorTS - 2 * 4.0 / 1000) * (depthhorTS - 2 * 4.0 / 1000) * horLength)) * STEELDENSITYKGPERM;

                double vertsSteelWeight = (widthVertTS * depthVertTS * verticalLength - ((widthVertTS - 2 * 4.0 / 1000) * (depthVertTS - 2 * 4.0 / 1000) * verticalLength)) * STEELDENSITYKGPERM;

                string vertBOM = tsSize + " Length " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[VERSTEEL]).SetPropertyValue(widthVertTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[VERSTEEL]).SetPropertyValue(depthVertTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[VERSTEEL]).SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertsSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string horBOM = tsSize + " Length " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

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
                    JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, horSteelPort, Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, vertDistance + pipeDia / 2 + shoeH + depthhorTS / 2 + angledOffset / 2, 0, -overHang);
                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                    {
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, horSteelPort, Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, vertDistance + pipeDia / 2 + shoeH + depthhorTS / 2 + angledOffset / 2, overHang, 0);
                    }
                    else
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Support can only be placed on Slab By-Point. ", "", "SFS5398_F1.cs", 196);
                }

                JointHelper.CreateRigidJoint(HORSTEEL, "StartOther", VERSTEEL, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, horLength / 2, 0, depthVertTS / 2);
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

                    structConnections.Add(new ConnectionInfo(VERSTEEL, 1)); // partindex, routeindex


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

            sizeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAFINLSize3", "Size3");
            size3 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;
            string BOMdesc = "Gate support F1 -" + size3 + " SFS 5398";

            return BOMdesc;
        }
    }
}