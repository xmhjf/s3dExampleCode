//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Util_ClampPlastic_Metric.cs
//    Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_ClampPlastic_Metric
//   Author       :  Rajeswari
//   Creation Date:  14/11/2012
//   Description  :  CR-CP-222287-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   14/11/2012   Rajeswari CR-CP-222287-Initial Creation
//   24/05/2013    Rajeswari Resolved TDL Errors
//   31/10/2013     Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Util_ClampPlastic_Metric : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility_Metric,Ingr.SP3D.Content.Support.Symbols.Util_ClampPlastic_Metric"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L1", "L1", 0.999999)]
        public InputDouble m_L1;
        [InputDouble(3, "L2", "L2", 0.999999)]
        public InputDouble m_L2;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(5, "H1", "H1", 0.999999)]
        public InputDouble m_H1;
        [InputDouble(6, "H3", "H3", 0.999999)]
        public InputDouble m_H3;
        [InputDouble(7, "Radius", "Radius", 0.999999)]
        public InputDouble m_Radius;
        [InputString(8, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
        [InputDouble(9, "InputWeight", "InputWeight", 0.999999)]
        public InputDouble m_InputWeight;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("CYLINDER", "CYLINDER")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                Double L1 = m_L1.Value;
                Double L2 = m_L2.Value;
                Double H = m_H.Value;
                Double H1 = m_H1.Value;
                Double H3 = m_H3.Value;
                Double radius = m_Radius.Value;
                Double inputWeight = m_InputWeight.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                if (L1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidL1, "L1 cannot be zero or negative"));
                    return;
                }
                if (L2 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidL2, "L2 cannot be zero or negative"));
                    return;
                }
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidH, "H cannot be zero or negative"));
                    return;
                }
                if (H1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidH1, "H1 cannot be zero or negative"));
                    return;
                }
                if (H3 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidH3, "H3 cannot be zero or negative"));
                    return;
                }
                if (radius <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidRadius, "Radius cannot be zero or negative"));
                    return;
                }

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -H), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                curveCollection.Add(new Line3d(new Position(-H3 / 2.0, -radius - H1 / 2.0, 0), new Position(-H3 / 2.0, -radius - H1 / 2.0, -(H - radius) / 2.0)));
                curveCollection.Add(new Line3d(new Position(-H3 / 2.0, -radius - H1 / 2.0, -(H - radius) / 2.0), new Position(-H3 / 2.0, -L2, -(H - radius) / 2.0 - 0.005)));
                curveCollection.Add(new Line3d(new Position(-H3 / 2.0, -L2, -(H - radius) / 2.0 - 0.005), new Position(-H3 / 2.0, -L2, -H)));
                curveCollection.Add(new Line3d(new Position(-H3 / 2.0, -L2, -H), new Position(-H3 / 2.0, L2, -H)));
                curveCollection.Add(new Line3d(new Position(-H3 / 2.0, L2, -H), new Position(-H3 / 2.0, L2, -(H - radius) / 2.0 - 0.005)));
                curveCollection.Add(new Line3d(new Position(-H3 / 2.0, L2, -(H - radius) / 2.0 - 0.005), new Position(-H3 / 2.0, radius + H1 / 2.0, -(H - radius) / 2.0)));
                curveCollection.Add(new Line3d(new Position(-H3 / 2.0, radius + H1 / 2.0, -(H - radius) / 2.0), new Position(-H3 / 2.0, radius + H1 / 2.0, 0)));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Matrix4X4 matrix = new Matrix4X4();
                Arc3d arc = symbolGeometryHelper.CreateArc(null, radius + H1 / 2.0, Math.PI);
                matrix.Rotate(0, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -H3 / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), H3, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-H3 / 2.0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d cylinder = symbolGeometryHelper.CreateCylinder(null, radius, H3 + 0.00002);
                m_Symbolic.Outputs["CYLINDER"] = cylinder;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Util_ClampPlastic_Metric.cs."));
                return;
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Double radius = (Double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricRadius", "Radius")).PropValue;
                string inputBomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue;

                if (inputBomDescription == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if (inputBomDescription == null)
                        bomDescription = "Plastic Pipe Clamp for " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, radius * 2.0, UnitName.DISTANCE_MILLIMETER) + " dia pipe";
                    else
                        bomDescription = inputBomDescription.Trim();
                }
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrBOMDescription, "Error in BOMDescription of Util_ClampPlastic_Metric.cs."));
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (Double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricInWt", "InputWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrWeightCG, "Error in WeightCG of Util_ClampPlastic_Metric.cs."));
            }
        }
        #endregion
    }
}
