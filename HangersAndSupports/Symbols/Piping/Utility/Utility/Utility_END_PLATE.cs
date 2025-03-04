//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_END_PLATE.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_END_PLATE
//   Author       :  Rajeswari
//   Creation Date:  01/11/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01/11/2012   Rajeswari  CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   30/Dec/2014    PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Linq;
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
    public class Utility_END_PLATE : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_END_PLATE"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "THICKNESS", "THICKNESS", 1)]
        public InputDouble m_THICKNESS;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(5, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_PIPE_DIA;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("BODY", "BODY")]
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

                Double thickness = m_THICKNESS.Value;
                Double W = m_W.Value;
                Double H = m_H.Value;
                Double pipeDiameter = m_PIPE_DIA.Value;

                Double T, A, alpha;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                PropertyValueCodelist thicknessCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrUtility_END_PLATE", "THICKNESS");
                CodelistItem codeList = thicknessCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)thickness);
                if (codeList != null)
                {
                    if (codeList.Value < 1 || codeList.Value > 12)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidPlateThickness, "Thickness should be between 1 to 12"));
                        return;
                    }
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWidthGTZero, "Width should be greater than zero"));
                    return;
                }
                if (codeList != null)

                    T = Double.Parse(codeList.ShortDisplayName) * 25.4 / 1000;
                else
                    T = 0.0127;

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                alpha = Math.Acos(2.0 / 3.0);

                if (W <= pipeDiameter)
                {
                    alpha = Math.Acos((W / 2) / (pipeDiameter / 2)) * 180 / Math.PI;
                }
                A = pipeDiameter / 2 - (pipeDiameter / 2 * Math.Sin(alpha * Math.PI / 180));
                if (pipeDiameter / 2 < W / 2)
                {
                    A = pipeDiameter / 2;
                    alpha = 0;
                }
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();
                Projection3d body;
                if (W > pipeDiameter)
                {
                    curveCollection.Add(new Line3d(new Position(0, W / 2, -pipeDiameter / 2 - H), new Position(0, -W / 2, -pipeDiameter / 2 - H)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2, -pipeDiameter / 2 - H), new Position(0, -W / 2, -pipeDiameter / 2 + A)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2, -pipeDiameter / 2 + A), new Position(0, -pipeDiameter / 2, -pipeDiameter / 2 + A)));

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));

                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI + alpha), new Vector(0, 0, 1));

                    Arc3d topArc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, Math.PI - 2 * alpha);
                    topArc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                    topArc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                    topArc.Transform(matrix);
                    curveCollection.Add(topArc);
                    curveCollection.Add(new Line3d(new Position(0, pipeDiameter / 2, -pipeDiameter / 2 + A), new Position(0, W / 2, -pipeDiameter / 2 + A)));
                    curveCollection.Add(new Line3d(new Position(0, W / 2, -pipeDiameter / 2 + A), new Position(0, W / 2, -pipeDiameter / 2 - H)));

                    body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
                else
                {
                    curveCollection = new Collection<ICurve>();
                    curveCollection.Add(new Line3d(new Position(0, W / 2, -pipeDiameter / 2 - H), new Position(0, -W / 2, -pipeDiameter / 2 - H)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2, -pipeDiameter / 2 - H), new Position(0, -W / 2, -pipeDiameter / 2 + A)));
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI + alpha * Math.PI / 180), new Vector(0, 0, 1));

                    Arc3d topArc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, Math.PI - 2 * alpha * Math.PI / 180);
                    topArc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                    topArc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                    topArc.Transform(matrix);
                    curveCollection.Add(topArc);
                    curveCollection.Add(new Line3d(new Position(0, W / 2, -pipeDiameter / 2 + A), new Position(0, W / 2, -pipeDiameter / 2 - H)));

                    body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_END_PLATE"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                string thickness, bomH, bomW, PipeDiameter;
                string[] bomHUnits, bomUnits;
                double alpha, A, T, H, W, pipeDiameterValue;

                PropertyValueCodelist thicknessList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE", "THICKNESS");
                thickness = thicknessList.PropertyInfo.CodeListInfo.GetCodelistItem(thicknessList.PropValue).DisplayName;

                H = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE", "H")).PropValue;
                bomH = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, H, UnitName.DISTANCE_INCH);
                bomHUnits = bomH.Split(' ');
                W = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE", "W")).PropValue;
                bomW = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, W, UnitName.DISTANCE_INCH);

                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                pipeDiameterValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                PipeDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pipeDiameterValue, UnitName.DISTANCE_INCH);
                bomUnits = PipeDiameter.Split(' ');

                if (thickness != "")
                {
                    T = double.Parse(thickness.Trim()) * 25.4 / 1000;
                }
                else
                {
                    T = 0.009525;
                }
                alpha = Math.Acos(2.0 / 3.0);
                if (W <= pipeDiameterValue)
                {
                    alpha = Math.Acos((W / 2) / pipeDiameterValue / 2) * (180 / Math.PI);
                }
                A = pipeDiameterValue / 2 - (pipeDiameterValue / 2 * Math.Sin(alpha * (Math.PI / 180)));
                if (pipeDiameterValue / 2 < W / 2)
                {
                    A = pipeDiameterValue / 2;
                    alpha = 0;
                }
                A = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, A, UnitName.DISTANCE_INCH).Split(' ').First());
                T = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_INCH).Split(' ').First());
                bomDescription = "End Plate - Pipe-Base L = " + bomH + ", Overall L = " + (Double.Parse(bomHUnits[0]) + A) + "in, " + bomW + " X " + T + "in, Radius " + Double.Parse(bomUnits[0]) / 2 + " " + bomUnits[1];
                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_END_PLATE"));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Double weight, cogX, cogY, cogZ;
                const int getSteelDensityKGPerM = 7900;
                double alpha = 0, arcLength, segmentArea, A;
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE", "H")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE", "W")).PropValue;
                PropertyValueCodelist thicknessList = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE", "THICKNESS");
                double T = double.Parse(thicknessList.PropertyInfo.CodeListInfo.GetCodelistItem(thicknessList.PropValue).DisplayName) * 25.4 / 1000;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double pipeDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;

                if (W <= pipeDiameter)
                {
                    alpha = Math.Acos((W / 2) / (pipeDiameter / 2));
                }
                A = pipeDiameter / 2 - (pipeDiameter / 2 * Math.Sin(alpha));
                if (pipeDiameter / 2 < W / 2)
                {
                    A = pipeDiameter / 2;
                    alpha = 0;
                }
                arcLength = pipeDiameter / 2 * ((180) - 2 * (alpha * 180 / Math.PI)) / 180 * Math.PI;
                segmentArea = 0.5 * (pipeDiameter / 2 * arcLength - (W * (pipeDiameter / 2 - A)));

                weight = ((H + A) * W - segmentArea) * T * getSteelDensityKGPerM;

                cogX = 0;
                cogY = 0;
                cogZ = 0;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_END_PLATE"));
                }
            }
        }
    }
        #endregion
}

