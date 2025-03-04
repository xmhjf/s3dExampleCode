//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_CurvedPlate.cs
//    Generic,Ingr.SP3D.Content.Support.Symbols.Gen_CurvedPlate
//   Author       :  Hema
//   Creation Date:  19-11-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19-11-2012     Hema    CR-CP-222274  Converted HS_Generic VB Project to C# .Net 
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
    public class Gen_CurvedPlate : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Generic,Ingr.SP3D.Content.Support.Symbols.Gen_CurvedPlate"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(3, "Radius", "Radius", 0.999999)]
        public InputDouble m_dRadius;
        [InputDouble(4, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(5, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;     
        [InputString(6, "BOM_DESC1", "BOM_DESC1","No Value")]
        public InputString m_oBOM_DESC1;
        [InputDouble(7, "EarL", "EarL", 0.999999)]
        public InputDouble m_dEarL;

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

                Double angle = m_dAngle.Value;
                Double thickness = m_dT.Value;
                Double radius = m_dRadius.Value;
                Double L = m_dW.Value;               
                Double earL = m_dEarL.Value;
                Collection<ICurve> curveCollection;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidW, "W cannot be zero"));
                    return;
                }
                double calc1 = Math.Sin(angle / 2) * radius;
                double calc2 = Math.Sqrt((radius * radius) - (calc1 * calc1));
                double calc3 = Math.Sin(angle / 2) * (radius + thickness);
                double calc4 = Math.Sqrt(((radius + thickness) * (radius + thickness)) - (calc3 * calc3));
                Matrix4X4 matrix = new Matrix4X4();
                
                if( angle <= Math.PI)
                {                    
                        calc2 = -calc2;
                        calc4 = -calc4;
                }
                if (earL == 0)
                {
                    curveCollection = new Collection<ICurve>();

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d outerArc = symbolGeometryHelper.CreateArc(null, radius + thickness, angle);
                    matrix.Rotate(Math.PI / 2 - angle / 2.0, new Vector(0, 0, 1));
                    matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, -L / 2.0, 0));
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1),new Position(0,0,0));
                    outerArc.Transform(matrix);
                    curveCollection.Add(outerArc);

                    curveCollection.Add(new Line3d( new Position(-L / 2.0, calc3, calc4), new Position(-L / 2.0, calc1, calc2)));

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d innerArc = symbolGeometryHelper.CreateArc(null, radius, angle);
                    matrix = new Matrix4X4();
                    matrix.Rotate(Math.PI / 2 - angle / 2.0, new Vector(0, 0, 1));
                    matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, -L / 2.0, 0));
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1),new Position(0,0,0));
                    innerArc.Transform(matrix);
                    curveCollection.Add(innerArc);

                    curveCollection.Add(new Line3d( new Position(-L / 2.0, -calc1, calc2), new Position(-L / 2.0, -calc3, calc4)));

                    Vector lineVector = new Vector(L, 0, 0);
                    Projection3d body = new Projection3d( new ComplexString3d(curveCollection), lineVector, lineVector.Length, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
                else
                {
                    Double berta = (Math.PI/2 - angle / 2);
                    Double innerDown = Math.Sin(berta) * (radius);
                    Double innerOver = Math.Cos(berta) * (radius);
                    Double down = Math.Sin(berta) * (radius + earL + thickness);
                    Double over = Math.Cos(berta) * (radius + earL + thickness);
                    Double earDown = Math.Cos(berta) * (thickness);
                    Double earOver = Math.Sin(berta) * (thickness);
                    Double outerArcRadius = radius + thickness;
                    Double mystery = outerArcRadius - Math.Sqrt(outerArcRadius * outerArcRadius - thickness * thickness);
                    Double lastUp = Math.Sin(berta) * (earL + mystery);
                    Double lastOver = Math.Cos(berta) * (earL + mystery);
                    Double leg1 = over - earOver - lastOver;
                    Double leg2 = down + earDown - lastUp;
                    Double gramma = (180/180*(Math.PI)) - 2 * (Math.Atan(leg2 / leg1));

                    curveCollection = new Collection<ICurve>();

                    curveCollection.Add(new Line3d( new Position(-L / 2.0, leg1, -leg2), new Position(-L / 2.0, over - earOver, -(down + earDown))));
                    curveCollection.Add(new Line3d( new Position(-L / 2.0, over - earOver, -(down + earDown)), new Position(-L / 2.0, over, -down)));
                    curveCollection.Add(new Line3d( new Position(-L / 2.0, over, -down), new Position(-L / 2.0, innerOver, -innerDown)));

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d outerArc = symbolGeometryHelper.CreateArc(null, radius + thickness, gramma);
                    matrix = new Matrix4X4();
                    matrix.Rotate(Math.PI / 2 - gramma / 2, new Vector(0, 0, 1));
                    matrix.Rotate(Math.PI / 2 - gramma / 2, new Vector(0, 0, 1));
                    matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, -L / 2.0, 0));
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1),new Position(0,0,0));
                    outerArc.Transform(matrix);
                    curveCollection.Add(outerArc);

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d innerArc = symbolGeometryHelper.CreateArc(null, radius, angle);
                    matrix = new Matrix4X4();
                    matrix.Rotate(Math.PI / 2 - angle / 2.0, new Vector(0, 0, 1));
                    matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, -L / 2.0, 0));
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1),new Position(0,0,0));
                    innerArc.Transform(matrix);
                    curveCollection.Add(innerArc);

                    curveCollection.Add(new Line3d( new Position(-L / 2.0, -innerOver, -innerDown), new Position(-L / 2.0, -over, -down)));
                    curveCollection.Add(new Line3d( new Position(-L / 2.0, -over, -down), new Position(-L / 2.0, -(over - earOver), -(down + earDown))));
                    curveCollection.Add(new Line3d( new Position(-L / 2.0, -(over - earOver), -(down + earDown)), new Position(-L / 2.0, -leg1, -leg2)));

                    Vector lineVector = new Vector(L, 0, 0);
                    Projection3d body = new Projection3d( new ComplexString3d(curveCollection), lineVector, lineVector.Length, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
            }
            catch    //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrConstructOutputs, "Error in ConstructOutputs of Gen_CurvedPlate"));
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
                double earLValue;
                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenCurvedPlate", "W")).PropValue;
                double radiusValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenCurvedPlate", "Radius")).PropValue;
                double thicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenCurvedPlate", "T")).PropValue;
                double angleValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenCurvedPlate", "Angle")).PropValue;
                try
                {
                    earLValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenCurvedPlate", "EarL")).PropValue;
                }
                catch
                {
                    earLValue = 0;
                }

                string bomDescription = (String)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrGenPartBOMDesc", "BOM_DESC1")).PropValue;

                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_MILLIMETER);
                string radius = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, radiusValue*2.0, UnitName.DISTANCE_MILLIMETER);
                string thickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, thicknessValue, UnitName.DISTANCE_MILLIMETER);
                string angle = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, angleValue, UnitName.ANGLE_DEGREE);
                string earL = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, earLValue, UnitName.DISTANCE_MILLIMETER);

                double calcWValue = angleValue * radiusValue;
                string calcW = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, calcWValue, UnitName.DISTANCE_MILLIMETER);
                string earBomValue;
                if (earLValue == 0)
                    earBomValue = "";
                else
                {
                    earBomValue = " w/ 2 x " + earL + " ears";
                }
                
                if ((bomDescription) == "None")
                    bomString = " ";
                else if (bomDescription == null)
                {
                    bomString = (radius + " dia x " + thickness + " Curved Plate, " + L + " X " + calcW + earBomValue); 
                }
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrBOMDescription, "Error in BOMDescription of Gen_CurvedPlate"));
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
                Double weight, cogX = 0, cogY = 0, cogZ = 0, earL;
                const int getSteelDensityKGPerM = 7900;
                try
                {
                    earL = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenCurvedPlate", "EarL")).PropValue;
                }
                catch
                {
                    earL = 0;
                }
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenCurvedPlate", "W")).PropValue;
                double radius = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenCurvedPlate", "Radius")).PropValue;
                double thickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenCurvedPlate", "T")).PropValue;
                double angle = ((double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenCurvedPlate", "Angle")).PropValue) * 180 / (Math.PI);

                double earW = thickness * earL * L * getSteelDensityKGPerM;
                weight = earW + ((((Math.PI) * (radius + thickness) * (radius + thickness) * L) - ((Math.PI) * (radius) * (radius) * L)) * (angle / 360.0) * getSteelDensityKGPerM);
               
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrWeightCG, "Error in WeightCG of Gen_CurvedPlate"));
                }
            }
        }
        #endregion
    }
}
