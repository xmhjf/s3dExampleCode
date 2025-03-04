//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GEN_U_STRAP.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_U_STRAP
//   Author       :  Hema
//   Creation Date:  02.11.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   02.11.2012     Hema     CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
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
    public class Utility_GEN_U_STRAP : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_U_STRAP"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(3, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(4, "R", "R", 0.999999)]
        public InputDouble m_R;
        [InputDouble(5, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(6, "OL", "OL", 0.999999)]
        public InputDouble m_OL;
        [InputString(7, "BOM_DESC1", "BOM_DESC1", "No Value")]
        public InputString m_BOM_DESC1;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT", "LEFT")]
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

                Double D = m_D.Value;
                Double H = m_H.Value;
                Double R = m_R.Value;
                Double W = m_W.Value;
                Double overlapStrap = m_OL.Value;

                if (H == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidHeight, "Height cannot be zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrRodDiameterGTZero, "Rod Diameter should be greater than zero"));
                    return;
                }
                                
                if (W == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidW, "Width of Strap cannot be zero"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                curveCollection.Add(new Line3d(new Position(-W / 2, -R / 2 + D / 2, 0), new Position(-W / 2, -R / 2 + D / 2, -overlapStrap)));
                curveCollection.Add(new Line3d(new Position(-W / 2, R / 2 - D / 2, 0), new Position(-W / 2, R / 2 - D / 2, -overlapStrap)));

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Matrix4X4 matrix = new Matrix4X4();

                matrix.SetIdentity();
                matrix.Rotate(0, new Vector(0, 0, 1));

                Arc3d strap = symbolGeometryHelper.CreateArc(null, R / 2 - D / 2, Math.PI);
                strap.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                strap.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -W / 2, 0));
                strap.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                strap.Transform(matrix);
                curveCollection.Add(strap);

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), W, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -R / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d cylinderright = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2, H);
                m_Symbolic.Outputs["RIGHT"] = cylinderright;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, R / 2, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d cylinderleft = (Projection3d)symbolGeometryHelper.CreateCylinder(null, D / 2, H);
                m_Symbolic.Outputs["LEFT"] = cylinderleft;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GEN_U_STRAP"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                Double H = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "H")).PropValue;
                Double D = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "D")).PropValue;
                Double R = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "R")).PropValue;
                PropertyValueCodelist nutsCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "NUTS");
                Double nuts = nutsCodelist.PropValue;
                Double thread = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "THREAD")).PropValue;
                string bomDescription = ((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "BOM_DESC1")).PropValue;

                if (bomDescription == "")
                {
                    bomString = "Custom U-Bolt, H=" + H + " D=" + D + " R=" + R + " Nuts=" + nuts + " Thread=" + thread;
                }
                else
                {
                    bomString = bomDescription;
                }

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_GEN_U_STRAP"));
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
                Double D = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "D")).PropValue;
                Double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "H")).PropValue;
                Double R = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "R")).PropValue;
                Double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "W")).PropValue;
                Double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "T")).PropValue;
                Double OL = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_U_STRAP", "OL")).PropValue;
                weight = (((Math.PI) * D / 2 * D / 2 * H) * 2 + ((Math.PI) * (R / 2 - D / 2) * (R / 2 - D / 2) * W) / 2.0 - ((Math.PI) * (R / 2 - D / 2 - T) * (R / 2 - D / 2 - T) * W) / 2 + ((OL * W * T) * 2)) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GEN_U_STRAP"));
                }
            }
        }

        #endregion
    }
}


