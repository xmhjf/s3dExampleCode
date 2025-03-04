//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GEN_HOOD.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_HOOD
//   Author       :Sasidhar  
//   Creation Date:6-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   6-11-2012    Sasidhar   CR-CP-222288  Converted HS_Utility VB Project to C# .Net   
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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
    
    public class Utility_GEN_HOOD : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_HOOD"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "OUT_RAD", "OUT_RAD", 0.999999)]
        public InputDouble m_OUT_RAD;
        [InputDouble(3, "IN_RAD", "IN_RAD", 0.999999)]
        public InputDouble m_IN_RAD;
        [InputDouble(4, "STRAP_L", "STRAP_L", 0.999999)]
        public InputDouble m_STRAP_L;
        [InputDouble(5, "HEIGHT", "HEIGHT", 0.999999)]
        public InputDouble m_HEIGHT;
        [InputDouble(6, "LEG", "LEG", 0.999999)]
        public InputDouble m_LEG;
        [InputString(7, "BOM_DESC1", "BOM_DESC1", "No Value")]
        public InputString m_BOM_DESC1;

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

                Double outRad = m_OUT_RAD.Value;
                Double inRad = m_IN_RAD.Value;
                Double strapl = m_STRAP_L.Value;
                Double height = m_HEIGHT.Value;
                Double leg = m_LEG.Value;

                if (strapl == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidStrapL, "Strap Length cannot be zero"));
                    return;
                }
                if (outRad <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidOutRad, "Outside Arc Radius should be greater than zero"));
                    return;
                }
                if (inRad <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidInRad, "Inside Arc Radius should be greater than zero"));
                    return;
                }
                if (leg <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidDimensionL, "Dimension L should be greater than zero"));
                    return;
                }
                if (height == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidDimensionH, "Dimension H cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Double calc1 = Math.Sqrt((height - 0.0127) * (height - 0.0127) + (inRad + leg) * (inRad + leg));
                Double calc2 = Math.Sqrt(Math.Abs((Math.Pow(calc1, 2) - Math.Pow(outRad, 2))));
                Double angle1 = Math.Atan(outRad / calc2) * (180 / Math.PI);
                Double angle2 = Math.Atan((height - 0.0127) / (inRad + leg)) * (180 / Math.PI);
                Double angle3 = angle1 + angle2;
                Double calc3 = Math.Sin(angle3 * (Math.PI / 180)) * calc2;
                Double calc4 = Math.Cos(angle3 * (Math.PI / 180)) * calc2;
                Double angle4 = Math.Acos((inRad + leg - calc4) / (outRad)) * 180 / Math.PI;

                Collection<ICurve> curveCollection = new Collection<ICurve>();

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI - (angle4 * Math.PI / 180)), new Vector(0, 0, 1));

                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, outRad, 2 * (angle4 * Math.PI / 180) - Math.PI);
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -strapl / 2, 0));
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);

                curveCollection.Add(new Line3d(new Position(-strapl / 2, inRad + leg - calc4, -(height - 0.0127 - calc3)), new Position(-strapl / 2, inRad + leg, -(height - 0.0127))));
                curveCollection.Add(new Line3d(new Position(-strapl / 2, inRad + leg, -(height - 0.0127)), new Position(-strapl / 2, inRad + leg, -(height))));
                curveCollection.Add(new Line3d(new Position(-strapl / 2, inRad + leg, -(height)), new Position(-strapl / 2, inRad, -(height))));
                curveCollection.Add(new Line3d(new Position(-strapl / 2, inRad, -(height)), new Position(-strapl / 2, inRad, 0)));

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));

                Arc3d innerArc = symbolGeometryHelper.CreateArc(null, inRad, -Math.PI);
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -strapl / 2, 0));
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                innerArc.Transform(matrix);
                curveCollection.Add(innerArc);

                curveCollection.Add(new Line3d(new Position(-strapl / 2, -inRad, 0), new Position(-strapl / 2, -inRad, -(height))));
                curveCollection.Add(new Line3d(new Position(-strapl / 2, -inRad, -(height)), new Position(-strapl / 2, -inRad - leg, -(height))));
                curveCollection.Add(new Line3d(new Position(-strapl / 2, -inRad - leg, -(height)), new Position(-strapl / 2, -inRad - leg, -(height - 0.0127))));
                curveCollection.Add(new Line3d(new Position(-strapl / 2, -inRad - leg, -(height - 0.0127)), new Position(-strapl / 2, -inRad - leg + calc4, -(height - 0.0127 - calc3))));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), strapl, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GEN_HOOD"));
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
                String bomDescriptionValue = ((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_HOOD", "BOM_DESC1")).PropValue;
                Double inRadiusValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_HOOD", "IN_RAD")).PropValue;
                Double strapLengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_HOOD", "STRAP_L")).PropValue;
                if (bomDescriptionValue == "")
                {
                    bomDescription = "" + inRadiusValue + " Pipe Hood " + strapLengthValue;
                }
                else
                {
                    bomDescription = bomDescriptionValue;
                }

                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_GEN_HOOD"));
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
                Double outRadius = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_HOOD", "OUT_RAD")).PropValue;
                Double inRadius = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_HOOD", "IN_RAD")).PropValue;
                Double strapLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_HOOD", "STRAP_L")).PropValue;
                Double leg = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_HOOD", "LEG")).PropValue;
                Double height = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_HOOD", "HEIGHT")).PropValue;


                Double calc1 = Math.Sqrt((height - 0.0127) * (height - 0.0127) + (inRadius + leg) * (inRadius + leg));
                Double calc2 = Math.Sqrt(Math.Abs((Math.Pow(calc1, 2) - Math.Pow(outRadius, 2))));
                Double angle1 = Math.Atan(outRadius / calc2) * (180 / Math.PI);
                Double angle2 = Math.Atan((height - 0.0127) / (inRadius + leg)) * (180 / Math.PI);
                Double angle3 = angle1 + angle2;
                Double calc3 = Math.Sin(angle3 * (Math.PI / 180)) * calc2;
                Double calc4 = Math.Cos(angle3 * (Math.PI / 180)) * calc2;

                weight = (((outRadius * outRadius * Math.PI * strapLength) - (inRadius * inRadius * Math.PI * strapLength)) / 2 + (leg * 0.01 * strapLength * 2) + (calc3 * calc4 * strapLength) + ((leg - calc4) * (height - 0.01) * strapLength) * 2) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GEN_HOOD"));
                }
            }
        }

        #endregion
    }
}

