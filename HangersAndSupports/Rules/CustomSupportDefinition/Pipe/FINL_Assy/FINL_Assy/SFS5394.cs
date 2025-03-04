//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5394.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5394
//   Author       :  Vijay
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Vijay   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
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
    public class SFS5394 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        int application, topFeature;
        double normalPipeDiameterMetric, angle;
        string loadClass;

        private const string CLIP = "Clip_5394";
        private const string SLEEVE = "Sleeve_5394";
        private const string BAR = "Bar_5394";
        private const string NUT1 = "Nut1_5394";
        private const string NUT2 = "Nut2_5394";

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    application = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLApplication", "Application")).PropValue;
                    topFeature = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLTopFeature", "TopFeature")).PropValue;

                    //To get Pipe Nom Dia
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                    if (pipeInfo.NominalDiameter.Units != "mm")
                        normalPipeDiameterMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, UnitName.NPD_INCH, UnitName.NPD_MILLIMETER) / 1000;
                    else
                        normalPipeDiameterMetric = pipeInfo.NominalDiameter.Size / 1000;

                    NominalDiameter minNominalDiameter = new NominalDiameter();
                    minNominalDiameter.Size = 6;
                    minNominalDiameter.Units = "mm";
                    NominalDiameter maxNominalDiameter = new NominalDiameter();
                    maxNominalDiameter.Size = 150;
                    maxNominalDiameter.Units = "mm";
                    NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90 }, "mm");

                    //check valid pipe size
                    if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5394.cs", 77);
                        return parts;
                    }

                    angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Z, OrientationAlong.Global_Z);

                    //Rule Check
                    if ((angle * 180.0 / Math.PI) < 90 && application == 2)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Cantilever should be placed horizontally.", "", "SFS5394.cs", 86);
                        return parts;
                    }
                    else if ((angle * 180.0 / Math.PI) >= 90 && application == 1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe hanger should be placed vertically.", "", "SFS5394.cs", 91);
                        return parts;
                    }

                    if (normalPipeDiameterMetric > (100.0 / 1000.0))
                        loadClass = "2";
                    else
                        loadClass = "1";

                    if (topFeature == 1)     //anchored -1 nut
                    {
                        if (application == 1)        //pipe hanger
                        {
                            parts.Add(new PartInfo(CLIP, "FINLCmp_SFS5394_CLIP"));
                            parts.Add(new PartInfo(SLEEVE, "FINLCmp_SFS5394_SLEEVE"));
                            parts.Add(new PartInfo(BAR, "FINLCmp_SFS5394_TRDBAR_" + loadClass));       //depending on load class
                            parts.Add(new PartInfo(NUT1, "FINLCmp_SFS4032_" + loadClass));             //depending on load class
                        }
                        else
                        {
                            parts.Add(new PartInfo(CLIP, "FINLCmp_SFS5394_CLIP"));
                            parts.Add(new PartInfo(BAR, "FINLCmp_SFS5394_TRDBAR_" + loadClass));       //depending on load class
                            parts.Add(new PartInfo(NUT1, "FINLCmp_SFS4032_" + loadClass));             //depending on load class
                        }
                    }
                    else if (topFeature == 2)      //nuts- 2 nuts                                          
                    {
                        if (application == 1)      //pipe hanger
                        {
                            parts.Add(new PartInfo(CLIP, "FINLCmp_SFS5394_CLIP"));
                            parts.Add(new PartInfo(SLEEVE, "FINLCmp_SFS5394_SLEEVE"));
                            parts.Add(new PartInfo(BAR, "FINLCmp_SFS5394_TRDBAR_" + loadClass));       //depending on load class
                            parts.Add(new PartInfo(NUT1, "FINLCmp_SFS4032_" + loadClass));             //depending on load class
                            parts.Add(new PartInfo(NUT2, "FINLCmp_SFS4032_" + loadClass));             //depending on load class
                        }
                        else
                        {
                            parts.Add(new PartInfo(CLIP, "FINLCmp_SFS5394_CLIP"));
                            parts.Add(new PartInfo(BAR, "FINLCmp_SFS5394_TRDBAR_" + loadClass));       //depending on load class
                            parts.Add(new PartInfo(NUT1, "FINLCmp_SFS4032_" + loadClass));             //depending on load class
                            parts.Add(new PartInfo(NUT2, "FINLCmp_SFS4032_" + loadClass));             //depending on load class
                        }
                    }
                    else        //welded- 0 nuts
                    {
                        if (application == 1)      //pipe hanger
                        {
                            parts.Add(new PartInfo(CLIP, "FINLCmp_SFS5394_CLIP"));
                            parts.Add(new PartInfo(SLEEVE, "FINLCmp_SFS5394_SLEEVE"));
                            parts.Add(new PartInfo(BAR, "FINLCmp_SFS5394_TRDBAR_" + loadClass));       //depending on load class
                        }
                        else
                        {
                            parts.Add(new PartInfo(CLIP, "FINLCmp_SFS5394_CLIP"));
                            parts.Add(new PartInfo(BAR, "FINLCmp_SFS5394_TRDBAR_" + loadClass));       //depending on load class
                        }
                    }

                    return parts;       //Get the collection of parts
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                double rodEndToRoute = 0;

                //Getting Dimension Information
                double E = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5394_CLIP", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double D = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5394_CLIP", "IJUAFINL_D", "D", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double T = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5394_CLIP", "IJUAFINL_T", "T", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double length = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5394_SLEEVE", "IJUAFINL_L", "L", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double length2 = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5394_SLEEVE", "IJUAFINL_L2", "L2", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double C = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5394_SLEEVE", "IJUAFINL_C", "C", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double boltT = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS4032", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass);

                rodEndToRoute = E + (length - length2 - C);

                double lengthV = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                double lengthH = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                double flange = 0, web = 0;

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
                    flange = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                    web = SupportingHelper.SupportingObjectInfo(1).WebThickness;
                    if (topFeature == 1)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "For anchored connection, the support should be placed on concrete", "", "SFS5394.cs", 200);
                }
                else        //Slab
                {
                    if (!(topFeature == 1))
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "For nuts or welded connection, the support should be placed on steel", "", "SFS5394.cs", 205);
                }

                if (topFeature == 1)     //Anchored -1 Nut
                {
                    if (application == 1)    //Pipe hanger
                    {
                        componentDictionary[BAR].SetPropertyValue(lengthV - rodEndToRoute + 2 * boltT + E * (1 - Math.Cos(angle)), "IJUAHgrOccLength", "Length");
                        length = lengthV + 2 * boltT + E * (1 - Math.Cos(angle));
                        support.SetPropertyValue(length, "IJUAFINL_SFS5394_Dim", "L");

                        //Clip, Sleeve, Bar, Nut1
                        JointHelper.CreateRigidJoint("-1", "Route", CLIP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRevoluteJoint(CLIP, "Pin", SLEEVE, "Hole", Axis.X, Axis.X);
                        JointHelper.CreateRigidJoint(SLEEVE, "Pin", BAR, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BAR, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -3 * boltT, 0, 0);
                        JointHelper.CreateGlobalAxesAlignedJoint(BAR, "TopExThdRH", Axis.Z, Axis.Z);
                    }
                    else
                    {
                        //Clip Bar, Nut1
                        componentDictionary[BAR].SetPropertyValue(lengthH - D / 2 - T + 2 * boltT, "IJUAHgrOccLength", "Length");
                        length = lengthH + 2 * boltT;
                        support.SetPropertyValue(length, "IJUAFINL_SFS5394_Dim", "L");
                        JointHelper.CreateRigidJoint("-1", "Route", CLIP, "Route", Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.NegativeZ, 0, 0, 0);
                        JointHelper.CreateRigidJoint(CLIP, "Route", BAR, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -D / 2 - T, 0, 0);
                        JointHelper.CreateRigidJoint(BAR, "TopExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 2 * boltT, 0, 0);
                    }
                }
                else if (topFeature == 2)        //nuts- 2 nuts
                {
                    if (application == 1)         //Pipe hanger
                    {
                        //Clip Sleeve Bar Nut1 Nut2
                        componentDictionary[BAR].SetPropertyValue(lengthV - rodEndToRoute + 2 * boltT + flange + E * (1 - Math.Cos(angle)), "IJUAHgrOccLength", "Length");
                        length = lengthV + 2 * boltT + flange + E * (1 - Math.Cos(angle));
                        support.SetPropertyValue(length, "IJUAFINL_SFS5394_Dim", "L");
                        JointHelper.CreateRigidJoint("-1", "Route", CLIP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRevoluteJoint(CLIP, "Pin", SLEEVE, "Hole", Axis.X, Axis.X);
                        JointHelper.CreateRigidJoint(SLEEVE, "Pin", BAR, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(BAR, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -3 * boltT - flange, 0, 0);
                        JointHelper.CreateRigidJoint(BAR, "BotExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -2 * boltT, 0, 0);
                        JointHelper.CreateGlobalAxesAlignedJoint(BAR, "TopExThdRH", Axis.Z, Axis.Z);
                    }
                    else
                    {
                        //Clip Bar Nut1 Nut2
                        componentDictionary[BAR].SetPropertyValue((lengthH - D / 2 - T) + 2 * boltT + web, "IJUAHgrOccLength", "Length");
                        length = lengthH + 2 * boltT + web;
                        support.SetPropertyValue(length, "IJUAFINL_SFS5394_Dim", "L");
                        JointHelper.CreateRigidJoint("-1", "Route", CLIP, "Route", Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.NegativeZ, 0, 0, 0);
                        JointHelper.CreateRigidJoint(CLIP, "Route", BAR, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -D / 2 - T, 0, 0);
                        JointHelper.CreateRigidJoint(BAR, "TopExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 2 * boltT + web, 0, 0);
                        JointHelper.CreateRigidJoint(BAR, "TopExThdRH", NUT2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, boltT, 0, 0);
                    }
                }
                else        //welded- 0 nuts
                {
                    if (application == 1)        //Pipe hanger
                    {
                        //Clip Sleeve Bar
                        componentDictionary[BAR].SetPropertyValue(lengthV - rodEndToRoute + E * (1 - Math.Cos(angle)), "IJUAHgrOccLength", "Length");
                        length = lengthV + E * (1 - Math.Cos(angle));
                        support.SetPropertyValue(length, "IJUAFINL_SFS5394_Dim", "L");
                        JointHelper.CreateRigidJoint("-1", "Route", CLIP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRevoluteJoint(CLIP, "Pin", SLEEVE, "Hole", Axis.X, Axis.X);
                        JointHelper.CreateRigidJoint(SLEEVE, "Pin", BAR, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateGlobalAxesAlignedJoint(BAR, "TopExThdRH", Axis.Z, Axis.Z);
                    }
                    else
                    {
                        //Clip, Bar
                        componentDictionary[BAR].SetPropertyValue(lengthH - D / 2 - T, "IJUAHgrOccLength", "Length");
                        length = lengthH;
                        support.SetPropertyValue(length, "IJUAFINL_SFS5394_Dim", "L");
                        JointHelper.CreateRigidJoint("-1", "Route", CLIP, "Route", Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.NegativeZ, 0, 0, 0);
                        JointHelper.CreateRigidJoint(CLIP, "Route", BAR, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -D / 2 - T, 0, 0);
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Finl_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    routeConnections.Add(new ConnectionInfo(CLIP, 1));      //partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(BAR, 1));     //partindex, structindex

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

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string BOMString = "";
            try
            {
                //To get Pipe Nom Dia
                double pipeDiameter = ((PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1)).NominalDiameter.Size;
                int applicationBOM = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAFINLApplication", "Application")).PropValue;
                int topFeatureBOMValue = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAFINLTopFeature", "TopFeature")).PropValue;
                double lBOM = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAFINL_SFS5394_Dim", "L")).PropValue;
                string topFeatureBOM, bomPrefix;

                if (applicationBOM == 1)     //Pipe Hanger
                {
                    if (topFeatureBOMValue == 1)
                        topFeatureBOM = "A";
                    else if (topFeatureBOMValue == 2)
                        topFeatureBOM = "B";
                    else
                        topFeatureBOM = "C";
                    bomPrefix = "Pipe hanger";
                }
                else        //Cantilever
                {
                    if (topFeatureBOMValue == 1)
                        topFeatureBOM = "D";
                    else if (topFeatureBOMValue == 2)
                        topFeatureBOM = "E";
                    else
                        topFeatureBOM = "F";
                    bomPrefix = "Cantilever";
                }
                BOMString = bomPrefix + " SFS 5394 DN " + pipeDiameter + " - " + topFeatureBOM + " - " + Math.Round(lBOM * 1000, 0);

                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        #endregion
    }
}
