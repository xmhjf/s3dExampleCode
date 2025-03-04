//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_HALF_END_PLATE.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_HALF_END_PLATE
//   Author       :  Hema
//   Creation Date:  07/11/2012
//   Description:

//   Change History: 
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   07/11/2012      Hema    CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		 Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using System.Linq;

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
    [VariableOutputs]
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Utility_HALF_END_PLATE : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_HALF_END_PLATE"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "THICKNESS", "THICKNESS", 1)]
        public InputDouble m_oTHICKNESS;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(5, "OVER", "OVER", 0.999999)]
        public InputDouble m_dOVER;
        [InputDouble(6, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;

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

                Double thickness = m_oTHICKNESS.Value;
                Double W = m_dW.Value;
                Double H = m_dH.Value;
                Double over = m_dOVER.Value;
                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double T, alpha, beta;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                PropertyValueCodelist thicknessCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrUtility_HALF_END_PLATE", "THICKNESS");
                CodelistItem codeList = thicknessCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)thickness);

                if (codeList.Value < 1 || codeList.Value > 12)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidPlateThickness, "Thickness should be between 1 to 12"));
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                if (codeList != null)
                    T = Double.Parse(codeList.ShortDisplayName) * 25.4 / 1000;
                else
                    T = 0.0127;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (W <= pipeDiameter / 2)
                {
                    alpha = Math.Acos((W) / (pipeDiameter / 2)) * 180 / Math.PI;
                    beta = Math.Acos((over) / (pipeDiameter / 2)) * 180 / Math.PI;
                }
                else
                {
                    alpha = 0;
                    beta = Math.Acos((over) / (pipeDiameter / 2)) * 180 / Math.PI;
                }

                Double A = pipeDiameter / 2 - (pipeDiameter / 2 * Math.Sin(alpha * Math.PI / 180));
                Double B = pipeDiameter / 2 - (pipeDiameter / 2 * Math.Sin(beta * Math.PI / 180));

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Projection3d body;
                Matrix4X4 matrix = new Matrix4X4();
                Vector lineVector;
                if (W > pipeDiameter / 2)
                {
                    curveCollection.Add(new Line3d(new Position(0, over, -(H + pipeDiameter / 2)), new Position(0, W, -(H + pipeDiameter / 2))));
                    curveCollection.Add(new Line3d(new Position(0, W, -(H + pipeDiameter / 2)), new Position(0, W, -(pipeDiameter / 2 - A))));
                    curveCollection.Add(new Line3d(new Position(0, W, -(pipeDiameter / 2 - A)), new Position(0, pipeDiameter / 2, -(pipeDiameter / 2 - A))));

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix.SetIdentity();
                    matrix.Rotate(-beta * Math.PI / 180, new Vector(0, 0, 1));
                    Arc3d arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, -alpha + (beta * Math.PI / 180));
                    arc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(-3 * Math.PI / 2, new Vector(1, 0, 0));
                    arc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                    arc.Transform(matrix);
                    curveCollection.Add(arc);

                    curveCollection.Add(new Line3d(new Position(0, over, -(pipeDiameter / 2 - B)), new Position(0, over, -(H + pipeDiameter / 2))));

                    lineVector = new Vector(T, 0, 0);
                    body = new Projection3d(new ComplexString3d(curveCollection), lineVector, lineVector.Length, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
                else
                {
                    curveCollection.Add(new Line3d(new Position(0, over, -(H + pipeDiameter / 2)), new Position(0, W, -(H + pipeDiameter / 2))));
                    curveCollection.Add(new Line3d(new Position(0, W, -(H + pipeDiameter / 2)), new Position(0, W, -(pipeDiameter / 2 - A))));

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(-beta * Math.PI / 180, new Vector(0, 0, 1));
                    Arc3d arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, (-alpha + beta) * Math.PI / 180);
                    arc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(-(3 * Math.PI / 2), new Vector(1, 0, 0));
                    arc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                    arc.Transform(matrix);
                    curveCollection.Add(arc);

                    curveCollection.Add(new Line3d(new Position(0, over, -(pipeDiameter / 2 - B)), new Position(0, over, -(H + pipeDiameter / 2))));

                    lineVector = new Vector(1, 0, 0);
                    body = new Projection3d(new ComplexString3d(curveCollection), lineVector, T, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_HALF_END_PLATE"));
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
                double alpha, beta, T;
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                PropertyValueCodelist thicknesslist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_HALF_END_PLATE", "THICKNESS");
                long thickness = thicknesslist.PropValue;
                string bomthickness = thicknesslist.PropertyInfo.CodeListInfo.GetCodelistItem(thicknesslist.PropValue).DisplayName;
                double H = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_HALF_END_PLATE", "H")).PropValue;
                double W = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_HALF_END_PLATE", "W")).PropValue;
                double over = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_HALF_END_PLATE", "OVER")).PropValue;
                double pipeDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                string bomH = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, H, UnitName.DISTANCE_INCH);
                string bomW = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, W, UnitName.DISTANCE_INCH);
                string bomover = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, over, UnitName.DISTANCE_INCH);
                string bompipeDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, pipeDiameter, UnitName.DISTANCE_INCH);

                string[] strHUnits = bomH.Split(' ');
                string[] strUnits = bompipeDiameter.Split(' ');

                if (bomthickness != null)
                    T = double.Parse(bomthickness.Trim()) * 25.4 / 1000;
                else
                    T = 0.009525;
                if (W <= pipeDiameter / 2)
                {
                    alpha = Math.Acos((W) / (pipeDiameter / 2)) * (180 / Math.PI);
                    beta = Math.Acos((over) / (pipeDiameter / 2)) * (180 / Math.PI);
                }
                else
                {
                    alpha = 0;
                    beta = Math.Acos((over) / (pipeDiameter / 2)) * (180 / Math.PI);
                }
                double A = pipeDiameter / 2 - (pipeDiameter / 2 * Math.Sin(alpha*Math.PI/180));
                double B = pipeDiameter / 2 - (pipeDiameter / 2 * Math.Sin(beta*Math.PI / 180));
                A = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, A, UnitName.DISTANCE_INCH).Split(' ').First());
                T = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T, UnitName.DISTANCE_INCH).Split(' ').First());
                W = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, W, UnitName.DISTANCE_INCH).Split(' ').First());
                over = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, over, UnitName.DISTANCE_INCH).Split(' ').First());
                bomDescription = "Half End Plate - Pipe-Base L = " + bomH + ", Overall L = " + (Double.Parse(strHUnits[0]) + A) + "in, " + (W - over) + "in X " + (T) + "in, Radius " + double.Parse(strUnits[0]) / 2 + " " + strUnits[1];

                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_HALF_END_PLATE"));
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

                Double weight, cogX, cogY, cogZ, alpha = 0, arcLength, segmentArea, A;
                const int getSteelDensityKGPerM = 7900;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_HALF_END_PLATE", "H")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_HALF_END_PLATE", "W")).PropValue;
                PropertyValueCodelist thicknesslist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrUtility_HALF_END_PLATE", "THICKNESS");
                double T = double.Parse(thicknesslist.PropertyInfo.CodeListInfo.GetCodelistItem(thicknesslist.PropValue).DisplayName) * 25.4 / 1000;
                double PipeDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;

                if (W <= PipeDiameter)
                {
                    alpha = Math.Acos((W / 2) / (PipeDiameter / 2));
                }
                A = PipeDiameter / 2 - (PipeDiameter / 2 * Math.Sin(alpha));
                if (PipeDiameter / 2 < W / 2)
                {
                    A = PipeDiameter / 2;
                    alpha = 0;
                }
                arcLength = PipeDiameter / 2 * ((180) - 2 * (alpha * 180 / Math.PI)) / 180 * Math.PI;
                segmentArea = 0.5 * (PipeDiameter / 2 * arcLength - (W * (PipeDiameter / 2 - A)));
                weight = ((H + A) * W - segmentArea) * T * getSteelDensityKGPerM * 0.5;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_HALF_END_PLATE"));
                }
            }
        }

        #endregion
    }
}
