//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bolt.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.Bolt
//   Author       :  PVK
//   Creation Date:  23.SEPT.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   31.May.2012     Ramya     Initial Creation
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [VariableOutputs]
    public class Bolt : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HS_OglaendParts,Ingr.SP3D.Content.Support.Symbols.Bolt"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)] public InputCatalogPart m_oPartInput;
        [InputDouble(2, "RodDiameter", "RodDiameter", 0, true)] public InputDouble m_dRodDiameter;        
        [InputDouble(3, "Offset1", "Offset1", 0, true)]public InputDouble m_dOffset1;
        [InputDouble(4, "Thickness1", "Thickness1", 0, true)] public InputDouble m_dThickness1;
        [InputDouble(5, "Diameter1", "Diameter1", 0, true)] public InputDouble m_dDiameter1;
        [InputDouble(6, "Length1", "Length1", 0, true)] public InputDouble m_dLength1;
        [InputDouble(7, "Shape1Width1", "Shape1Width1", 0, true)]public InputDouble m_dShape1Width1;
        [InputDouble(8, "Shape1Length", "Shape1Length", 0, true)] public InputDouble m_Shape1Length;


        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("RodEnd1", "RodEnd1")]
        [SymbolOutput("RodEnd2", "RodEnd2")]
        [SymbolOutput("Rod", "Rod")]
        [SymbolOutput("TopHead ", "TopHead")]
        [SymbolOutput("HexNut", "HexNut")]

        public AspectDefinition m_oPhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part oPart = null;
                SymbolGeometryHelper oSymbolGeomHlpr = new SymbolGeometryHelper();
                SP3DConnection oConnection = default(SP3DConnection);
                oConnection = OccurrenceConnection;

                oPart = m_oPartInput.Value as Part;

                double dRodDiameter = m_dRodDiameter.Value;                
                double dOffset1 = m_dOffset1.Value;
                double dThickness1 = m_dThickness1.Value;
                double dDiameter1 = m_dDiameter1.Value;
                double dLength1 = m_dLength1.Value;
                double dShape1Width1 = m_dShape1Width1.Value;
                double Shape1Length = m_Shape1Length.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d cylinder1 = (Projection3d)symbolGeometryHelper.CreateCylinder(OccurrenceConnection, dRodDiameter / 2, dLength1);
                m_oPhysicalAspect.Outputs["Rod"] = cylinder1;

                if (dThickness1 > 0)
                {

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, dLength1);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d cylinder2 = (Projection3d)symbolGeometryHelper.CreateCylinder(OccurrenceConnection, dDiameter1 / 2, dThickness1);
                    m_oPhysicalAspect.Outputs["TopHead"] = cylinder2;
                }
                //HexNut//

                if (Shape1Length > 0)
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Line3d oNut11, oNut12, oNut13, oNut14, oNut15, oNut16;
                    Collection<ICurve> oNutColl = new Collection<ICurve>();

                    double dwidth1 = dShape1Width1 * Math.Sin(30 * Math.PI / 180);
                    double dwidth2 = dShape1Width1 * Math.Sin(60 * Math.PI / 180);


                    oNut11 = new Line3d(new Position(-dwidth1, -dwidth2, dOffset1), new Position(dwidth1, -dwidth2, dOffset1));
                    oNut12 = new Line3d(new Position(dwidth1, -dwidth2, dOffset1), new Position(dShape1Width1, 0, dOffset1));
                    oNut13 = new Line3d(new Position(dShape1Width1, 0, dOffset1), new Position(dwidth1, dwidth2, dOffset1));
                    oNut14 = new Line3d(new Position(dwidth1, dwidth2, dOffset1), new Position(-dwidth1, dwidth2, dOffset1));
                    oNut15 = new Line3d(new Position(-dwidth1, dwidth2, dOffset1), new Position(-dShape1Width1, 0, dOffset1));
                    oNut16 = new Line3d(new Position(-dShape1Width1, 0, dOffset1), new Position(-dwidth1, -dwidth2, dOffset1));

                    //dPlate1Thickness
                    oNutColl.Add(oNut11);
                    oNutColl.Add(oNut12);
                    oNutColl.Add(oNut13);
                    oNutColl.Add(oNut14);
                    oNutColl.Add(oNut15);
                    oNutColl.Add(oNut16);


                    ComplexString3d oNutComplxString = new ComplexString3d(oNutColl);
                    Vector NutVector = new Vector(0, 0, 1);
                    Projection3d oNutProjection = new Projection3d(oConnection, oNutComplxString, NutVector, Shape1Length, true);

                    m_oPhysicalAspect.Outputs["HexNut"] = oNutProjection;
                }

                Port port1 = new Port(OccurrenceConnection, oPart, "RodEnd1", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["RodEnd1"] = port1;

                Port port2 = new Port(OccurrenceConnection, oPart, "RodEnd2", new Position(0, 0, dLength1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oPhysicalAspect.Outputs["RodEnd2"] = port2;

            }
            catch (Exception oExc) //General Unhandled exception 
            {
                throw oExc;
            }
        }
        #endregion

        #region ICustomHgrWeightCG Members
        void ICustomHgrWeightCG.WeightCG(SupportComponent supportComponent, ref double weight, ref double cogX, ref double cogY, ref double cogZ)
        {

            ////System WCG Attributes

            Part catalogPart = (Part)supportComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

            try
            {
                weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
            }
            catch
            {
                weight = 0;
            }
            //Center of Gravity
            try
            {
                cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
            }
            catch
            {
                cogX = 0;
            }
            try
            {
                cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
            }
            catch
            {
                cogY = 0;
            }
            try
            {
                cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
            }
            catch
            {
                cogZ = 0;
            }

        }
        #endregion
    }
}

