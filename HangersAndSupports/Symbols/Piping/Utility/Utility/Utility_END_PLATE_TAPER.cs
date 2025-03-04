//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_END_PLATE_TAPER.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_END_PLATE_TAPER
//   Author       :  Hema
//   Creation Date:  30.10.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.10.2012      Hema    CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		 Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   30/Dec/2014    PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;
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
    [VariableOutputs]
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Utility_END_PLATE_TAPER : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_END_PLATE_TAPER"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputDouble(3, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(4, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(5, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(6, "ANGLE", "ANGLE", 0.999999)]
        public InputDouble m_dANGLE;

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

                Double R = m_dR.Value;
                Double T = m_dT.Value;
                Double W = m_dW.Value;
                Double H = m_dH.Value;
                Double angle = m_dANGLE.Value;

                if (HgrCompareDoubleService.cmpdbl(T, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidT, "Thickness cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(R, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidRadius, "Radius cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(angle * 180 / Math.PI) % 180, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidIncludedAngle, "Included Angle cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Double W2 = Math.Sin((angle / 2)) * R * 2;
                Double CALC1 = Math.Cos((angle / 2)) * R;

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                curveCollection.Add(new Line3d(new Position(0, W2 / 2, -CALC1), new Position(0, W / 2, -(H + R))));
                curveCollection.Add(new Line3d(new Position(0, W / 2, -(H + R)), new Position(0, -W / 2, -(H + R))));
                curveCollection.Add(new Line3d(new Position(0, -W / 2, -(H + R)), new Position(0, -W2 / 2, -CALC1)));

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI / 2 - angle / 2), new Vector(0, 0, 1));

                Arc3d topArc = symbolGeometryHelper.CreateArc(null, R, 2 * angle / 2);
                topArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                topArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, 0, 0));
                topArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                topArc.Transform(matrix);
                curveCollection.Add(topArc);

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                m_Symbolic.Outputs["BODY"] = body;

            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_END_PLATE_TAPER"));
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
                double RValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "R")).PropValue;
                double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "W")).PropValue;
                double HValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "H")).PropValue;
                double TValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "T")).PropValue;
                double angleValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "ANGLE")).PropValue;

                string R = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, RValue, UnitName.DISTANCE_INCH);
                string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_INCH);
                string H = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, HValue, UnitName.DISTANCE_INCH);
                string T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_INCH);
                string angle = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, angleValue, UnitName.ANGLE_DEGREE);

                bomDescription = "Tapered End Plate, R=" + R + ", W=" + W + ", H=" + H + ", T=" + T + ", Angle=" + angle + " degrees ";

                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_END_PLATE_TAPER"));
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
                Double weight, cogX, cogY, cogZ, segmentArea;
                const int getSteelDensityKGPerM = 7900;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "W")).PropValue;
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "H")).PropValue;
                double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "T")).PropValue;
                double R = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "R")).PropValue;
                double angle = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_END_PLATE_TAPER", "ANGLE")).PropValue;

                double W2 = (Math.Sin(angle / 2.0)) * R * 2.0;
                double arcLength = R * (angle * 180 / Math.PI) / 180 * Math.PI;
                double A = R - (R * Math.Cos(angle / 2.0));
                segmentArea = 0.5 * (R * arcLength - W2 * (R - A));

                weight = ((((H + A) * W2 - segmentArea) * T) - ((W2 / 2.0 - W / 2.0) * ((H + A) * T))) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_END_PLATE_TAPER"));
                }
            }
        }
    }
        #endregion
}

