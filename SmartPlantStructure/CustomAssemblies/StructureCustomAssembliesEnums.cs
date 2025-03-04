//-------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  StructureCustomAssembliesEnums.cs
//
//Abstract
//	StructureCustomAssembliesEnums is Structure custom assemblies enums.
//-------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure.Services
{
    /// <summary>
    /// Enumerated values for StructACSizingRule codelist table. This is a lookup definition of the "official" CodeList definition, which can be found in
    /// <![CDATA[<InstallDir>]]>\CatalogData\BulkLoad\DataFiles\AllCodeLists.xls
    /// </summary>
    public enum AssemblyConnectionSizing
    {
        /// <summary>
        /// ByRule
        /// </summary>
        ByRule = 1,
        /// <summary>
        /// UserDefined
        /// </summary>
        UserDefined = 2
    }
}