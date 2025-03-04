
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//
//Copyright 1992 - 2009 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  DesignedHandrail.cs
//
//Abstract
//	This is .NET DesignedHandrail symbol. This class subclasses from HandRailSymbolDefinition.
//
//
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
using Ingr.SP3D.Common.Middle;


//===========================================================================================
//Namespace of this class is Ingr.SP3D.Content.Structure
//It is recommended that customers specify namespace of their symbols to be
//<CompanyName>.SP3D.Content.<Specialization>.
//It is also recommended that if customers want to change this symbol to suit their
//requirements, they should change namespace/symbol name so the identity of the modified
//symbol will be different from the one delivered by Intergraph.
//===========================================================================================
namespace Ingr.SP3D.Content.Structure
{
    //DesignedHandrail class is used as a no-graphics symbol set for the ‘DesignedHandrailPart’ in the catalog.
    //This will be used to avoid re-evaluation and creation of graphic outputs on computation of the Handrail.
    public class DesignedHandrail : HandRailSymbolDefinition
    {
        //=======================================================================================================
        //DefinitionName/ProgID of this symbol is "HandRailSymbols,Ingr.SP3D.Content.Structure.DesignedHandrail"
        //=======================================================================================================

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart catalogPart;
        #endregion

        #region "Definitions of Aspects"

        //SimplePhysical Aspect
        //DesignedHandrail does not have any outputs, but it needs at least one aspect to be a valid symbol definition.
        [Aspect("SimplePhysical", "SimplePhysical Aspect of Handrail", AspectID.SimplePhysical)]
        public AspectDefinition simplePhysicalAspect;

        #endregion
    }
}

