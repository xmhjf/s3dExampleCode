﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_RectPad.cs
//    Generic,Ingr.SP3D.Content.Support.Symbols.Gen_RectPad
//   Author       :  Hema
//   Creation Date:  20-11-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-11-2012     Hema    CR-CP-222274  Converted HS_Generic VB Project to C# .Net 
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   6/11/2013     Vijaya   CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
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

    [CacheOption(CacheOptionType.NonCached)] 
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Gen_RectPad : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Generic,Ingr.SP3D.Content.Support.Symbols.Gen_RectPad"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(4, "Radius", "Radius", 0.999999)]
        public InputDouble m_dRadius;
        [InputDouble(5, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputString(6, "BOM_DESC1", "BOM_DESC1","No Value")]
        public InputString m_oBOM_DESC1;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Rectangle", "Rectangle")]
        [SymbolOutput("HgrPort1", "HgrPort1")]
        [SymbolOutput("HgrPort2", "HgrPort2")]
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
                
                Double L = m_dL.Value;
                Double W = m_dW.Value;
                Double R = m_dRadius.Value;
                Double TH = m_dThickness.Value;
               
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "HgrPort_1", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["HgrPort1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "HgrPort_2", new Position(0, 0, TH), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["HgrPort2"] = port2;

                if (TH == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidTH, "Thickness cannot be zero"));
                    return;
                }
                if (L < 2 * R || W < 2 * R)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidBAndHGTR, "Width and Height should be greater than twice the Radius"));
                    return;
                }
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, R, Math.PI / 2);
                Matrix4X4 matrix = new Matrix4X4();
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Translate(new Vector(-L / 2 + R, -W / 2 + R, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                curveCollection.Add(new Line3d( new Position(-L / 2 + R, -W / 2, 0), new Position(L / 2 - R, -W / 2, 0)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, R, -(Math.PI / 2 + Math.PI));
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI+Math.PI/2, new Vector(0, 0, 1));
                matrix.Translate(new Vector(L / 2 - R, -W / 2 + R, 0));
                arc1.Transform(matrix);
                curveCollection.Add(arc1);

                curveCollection.Add(new Line3d( new Position(L / 2, -W / 2 + R, 0), new Position(L / 2, W / 2 - R, 0)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc2 = symbolGeometryHelper.CreateArc(null, R, Math.PI / 2);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(L / 2 - R, W / 2 - R, 0));
                arc2.Transform(matrix);
                curveCollection.Add(arc2);

                curveCollection.Add(new Line3d( new Position(L / 2 - R, W / 2, 0), new Position(-L / 2 + R, W / 2, 0)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc3 = symbolGeometryHelper.CreateArc(null, R, Math.PI / 2);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Translate(new Vector(-L / 2 + R, W / 2 - R, 0));
                arc3.Transform(matrix);
                curveCollection.Add(arc3);

                curveCollection.Add(new Line3d( new Position(-L / 2, W / 2 - R, 0), new Position(-L / 2, -W / 2 + R, 0)));

                Vector lineVector = new Vector(0, 0, TH);
                Projection3d rectangle = new Projection3d( new ComplexString3d(curveCollection), lineVector, lineVector.Length, true);
                m_Symbolic.Outputs["Rectangle"] = rectangle;
            }
            catch    //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrConstructOutputs, "Error in ConstructOutputs of Gen_RectPad"));
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
                String material = "";
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                PropertyValueCodelist materialCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrGenPadMat", "Material");
                CodelistItem codeList = materialCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(materialCodeList.PropValue);

                if (codeList != null)
                {
                    material = codeList.ShortDisplayName;
                }
                Double thicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                String thickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, thicknessValue, UnitName.DISTANCE_MILLIMETER);

                Double lengthValue = (Double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrGenRectPad", "L")).PropValue;
                String length_mm = (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, (double)(lengthValue), UnitName.DISTANCE_MILLIMETER)).Trim();
                string[] length = length_mm.Split(' ');

                Double widthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrGenRectPad", "W")).PropValue;
                String width_mm = (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_MILLIMETER)).Trim();
                string[] width = width_mm.Split(' ');

                String bomDescription = (String)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrGenPartBOMDesc", "BOM_DESC1")).PropValue;

                if ((bomDescription).ToUpper() == "None")
                    bomString = "";
                else
                    bomString = part.PartDescription + ", L" + length[0] + "x" + width[0] + ", T=" + thickness + ", " + material;
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrBOMDescription, "Error in BOMDescription of Gen_RectPad"));
                }
                return "";
            }
        }
        #endregion
    }
}
