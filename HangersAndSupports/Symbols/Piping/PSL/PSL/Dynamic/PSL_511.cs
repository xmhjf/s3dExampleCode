//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_511.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_511
//   Author       :  Hema
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Hema    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
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

    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_511 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_511"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "END_ROT", "END_ROT", 0.999999)]
        public InputDouble END_ROT;
        [InputDouble(3, "Length", "Length", 0.999999)]
        public InputDouble Length;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(7, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(8, "L_MIN", "L_MIN", 0.999999)]
        public InputDouble L_MIN;
        [InputDouble(9, "L_MAX", "L_MAX", 0.999999)]
        public InputDouble L_MAX;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("L_END_BOX", "L_END_BOX")]
        [SymbolOutput("L_END_CYL", "L_END_CYL")]
        [SymbolOutput("L_MID", "L_MID")]
        [SymbolOutput("MID", "MID")]
        [SymbolOutput("R_MID", "R_MID")]
        [SymbolOutput("R_END_BOX", "R_END_BOX")]
        [SymbolOutput("R_END_CYL", "R_END_CYL")]
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
                //-----------------------------------------------------------------------
                // Construction of Physical Aspect 
                //----------------------------------------------------------------------

                Part part = (Part)PartInput.Value;
                Double endRotation = END_ROT.Value;
                Double length = Length.Value;
                Double a = A.Value;
                Double b = B.Value;
                Double c = C.Value;
                Double d = D.Value;
                Double lMin = L_MIN.Value;
                Double lMax = L_MAX.Value;

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                if (length < lMin || length > lMax)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Length " + (length * 1000).ToString() + "is out of Min and Max range: MIN:" + (lMin * 1000).ToString() + "Max:" + (lMax * 1000).ToString());

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, a / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, length + a / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (b <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B value should be greater than zero"));
                    return;
                }
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }
                symbolGeometryHelper.ActivePosition = new Position(-b / 2.0, -c / 2.0, a / 2.0);
                Projection3d leftEndBox = symbolGeometryHelper.CreateBox(null, b, c, b, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                leftEndBox.Transform(matrix);
                m_Symbolic.Outputs["L_END_BOX"] = leftEndBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d leftEndCylinder = symbolGeometryHelper.CreateCylinder(null, b / 2.0, c);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -c / 2.0, a / 2.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                leftEndCylinder.Transform(matrix);
                m_Symbolic.Outputs["L_END_CYL"] = leftEndCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, a / 2.0 + b * 2.0).Subtract(new Position(0, 0, b + a / 2.0));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, b + a / 2.0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d leftMid = symbolGeometryHelper.CreateCylinder(null, a / 2.0, normal.Length);
                m_Symbolic.Outputs["L_MID"] = leftMid;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(0, 0, length + a / 2.0 - b * 2.0).Subtract(new Position(0, 0, b * 2.0 + a / 2.0));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, b * 2.0 + a / 2.0);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d mid = symbolGeometryHelper.CreateCylinder(null, d / 2.0, normal1.Length);
                m_Symbolic.Outputs["MID"] = mid;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal2 = new Position(0, 0, length - b + a / 2.0).Subtract(new Position(0, 0, length + a / 2.0 - b * 2.0));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, length + a / 2.0 - b * 2.0);
                symbolGeometryHelper.SetOrientation(normal2, normal2.GetOrthogonalVector());
                Projection3d rightMid = symbolGeometryHelper.CreateCylinder(null, a / 2.0, normal2.Length);
                m_Symbolic.Outputs["R_MID"] = rightMid;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-b / 2.0, -c / 2.0, length + a / 2.0 - b);
                Projection3d rightEndBox = symbolGeometryHelper.CreateBox(null, b, c, b, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(endRotation, new Vector(0, 0, 1));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rightEndBox.Transform(matrix);
                m_Symbolic.Outputs["R_END_BOX"] = rightEndBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rightEndCylinder = symbolGeometryHelper.CreateCylinder(null, b / 2.0, c);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -c / 2.0, length + a / 2.0));
                matrix.Rotate(endRotation, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                rightEndCylinder.Transform(matrix);
                m_Symbolic.Outputs["R_END_CYL"] = rightEndCylinder;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_511."));
                    return;
                }
            }
        }

        #endregion

        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;
                String size = (String)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_SIZE", "SIZE")).PropValue;
                double weight = 0, maxLoad = 0;
                if (length <= 1)
                {
                    weight = PSLSymbolServices.GetDataByCondition("PSL_511", "IJUAHgrPSL_511", "WEIGHT_0", "IJUAHgrPSL_SIZE", "SIZE", size);
                    maxLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_511", "MAX_LOAD_1")).PropValue;
                    supportComponentBO.SetPropertyValue(maxLoad, "IJOAHgrPSL_511", "MAX_LOAD");
                }
                else if (length <= 2)
                {
                    weight = PSLSymbolServices.GetDataByCondition("PSL_511", "IJUAHgrPSL_511", "WEIGHT_1", "IJUAHgrPSL_SIZE", "SIZE", size);
                    maxLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_511", "MAX_LOAD_2")).PropValue;
                    supportComponentBO.SetPropertyValue(maxLoad, "IJOAHgrPSL_511", "MAX_LOAD");
                }
                else if (length <= 3)
                {
                    weight = PSLSymbolServices.GetDataByCondition("PSL_511", "IJUAHgrPSL_511", "WEIGHT_2", "IJUAHgrPSL_SIZE", "SIZE", size);
                    maxLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_511", "MAX_LOAD_3")).PropValue;
                    supportComponentBO.SetPropertyValue(maxLoad, "IJOAHgrPSL_511", "MAX_LOAD");
                }
                else if (length <= 4)
                {
                    weight = PSLSymbolServices.GetDataByCondition("PSL_511", "IJUAHgrPSL_511", "WEIGHT_3", "IJUAHgrPSL_SIZE", "SIZE", size);
                    maxLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_511", "MAX_LOAD_4")).PropValue;
                    supportComponentBO.SetPropertyValue(maxLoad, "IJOAHgrPSL_511", "MAX_LOAD");
                }
                else if (length <= 5)
                {
                    weight = PSLSymbolServices.GetDataByCondition("PSL_511", "IJUAHgrPSL_511", "WEIGHT_4", "IJUAHgrPSL_SIZE", "SIZE", size);
                    maxLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_511", "MAX_LOAD_5")).PropValue;
                    supportComponentBO.SetPropertyValue(maxLoad, "IJOAHgrPSL_511", "MAX_LOAD");
                }
                else if (length <= 6)
                {
                    weight = PSLSymbolServices.GetDataByCondition("PSL_511", "IJUAHgrPSL_511", "WEIGHT_5", "IJUAHgrPSL_SIZE", "SIZE", size);
                    maxLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_511", "MAX_LOAD_6")).PropValue;
                    supportComponentBO.SetPropertyValue(maxLoad, "IJOAHgrPSL_511", "MAX_LOAD");
                }
                else
                {
                    weight = PSLSymbolServices.GetDataByCondition("PSL_511", "IJUAHgrPSL_511", "WEIGHT_6", "IJUAHgrPSL_SIZE", "SIZE", size);
                    maxLoad = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_511", "MAX_LOAD_6")).PropValue;
                    supportComponentBO.SetPropertyValue(maxLoad, "IJOAHgrPSL_511", "MAX_LOAD");
                }
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, 0, 0, 0);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_511."));
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_SIZE", "SIZE")).PropValue;

                double lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);

                bomDescription = "PSL " + partNumber + "  Rigit Strut, L=" + length;

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_511"));
                return "";
            }
        }
        #endregion
    }
}
