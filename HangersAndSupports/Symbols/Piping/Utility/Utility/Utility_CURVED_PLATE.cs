//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_CURVED_PLATE.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_CURVED_PLATE
//   Author       :Sasidhar  
//   Creation Date:5.11.2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   5.11.2012      Sasidhar   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	  DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   02-Dec-2014    Chethan   DM-CP-265599  Unable to change plate thickness or radius of curved plate
//   30/Dec/2014    PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    
    public class Utility_CURVED_PLATE : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_CURVED_PLATE"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_L;
        [InputDouble(3, "RADIUS", "RADIUS", 0.999999)]
        public InputDouble m_RADIUS;
        [InputDouble(4, "THICKNESS", "THICKNESS", 0.999999)]
        public InputDouble m_THICKNESS;
        [InputDouble(5, "ANGLE", "ANGLE", 0.999999)]
        public InputDouble m_ANGLE;
        [InputString(6, "BOM_DESC", "BOM_DESC", "No Value")]
        public InputString m_BOM_DESC;

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
                Double L = m_L.Value;
                Double radius = m_RADIUS.Value;
                Double thickness = m_THICKNESS.Value;
                Double angle = m_ANGLE.Value;
                Double calc1 = Math.Sin(angle / 2) * radius;                
                Double calc2 = Math.Cos(angle / 2) * radius;
                Double clac3 = Math.Sin(angle / 2) * (radius + thickness);                
                Double calc4 = Math.Cos(angle / 2) * (radius + thickness);

                if (HgrCompareDoubleService.cmpdbl(L , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidLength, "Length cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(angle * 180 / Math.PI), 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidangle, "Angle cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(radius, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidInsideRadius, "Inside Radius cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                if (((angle) * 180.0 / Math.PI) < 180)
                {
                    calc2 = -calc2;
                    calc4 = -calc4;
                }
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI / 2 - angle / 2), new Vector(0, 0, 1));

                Arc3d arc = symbolGeometryHelper.CreateArc(null, radius + thickness, angle);
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -L / 2, 0));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                Line3d line = new Line3d(new Position(-L / 2, clac3, calc4), new Position(-L / 2, calc1, calc2));
                curveCollection.Add(line);

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI / 2 - angle / 2), new Vector(0, 0, 1));

                arc = new Arc3d(symbolGeometryHelper.CreateArc(null, radius, angle));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -L / 2, 0));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                line = new Line3d(new Position(-L / 2, -calc1, calc2), new Position(-L / 2, -clac3, calc4));
                curveCollection.Add(line);

                ComplexString3d side1Collection = new ComplexString3d(curveCollection);
                Vector projectionVector = new Vector(1, 0, 0);
                Projection3d projectionBody = new Projection3d(side1Collection, projectionVector, L, true);

                m_Symbolic.Outputs["BODY"] = projectionBody;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_CURVED_PLATE"));
                    return;
                }
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
                Double radius = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_CURVED_PLATE", "RADIUS")).PropValue;
                Double thickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_CURVED_PLATE", "THICKNESS")).PropValue;
                Double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_CURVED_PLATE", "L")).PropValue;
                Double angle = ((double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_CURVED_PLATE", "ANGLE")).PropValue * 180 / Math.PI);
                weight = (((Math.PI) * (radius + thickness) * (radius + thickness) * L) - ((Math.PI) * (radius) * (radius) * L)) * ((angle) / 360) * getSteelDensityKGPerM;

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_CURVED_PLATE"));
                }
            }
        }
    }
        #endregion
}
