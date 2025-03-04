//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   FINL_AssyResourceIDs.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.FINL_AssyResourceIDs
//   Author       :  BS
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     BS   CR224491,224492,224485- Convert FINL_Assy to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;


namespace Ingr.SP3D.Content.Support.Rules
{
	/// <summary>
	/// Summary description for CmnResourceIDs
	/// </summary>
    public static class FINL_AssyResourceIDs
	{
		/// <summary>
		/// Given ProgId is not in the correct format. .NET class ProgID should be in the following format: AssemblyName,Namespace.ClassName
		/// </summary>
		public const int CmnIncorrectProgIDFormat = 1;

		/// <summary>
		/// Given ProgID not found in the System or Custom Symbol Configuration XML file in symbol share. Please make sure the assembly for the rule or symbol exists in symbol share and run the UpdateSymbolConfiguratoin command from Project Management.
		/// </summary>
        public const int CmnProgIDNotFoundInSymbolMap = 2;

		/// <summary>
		/// CustomSymbolConfig.xml file is not found at symbol share.
		/// </summary>
        public const int CmnCustomSymbolConfigFileNotFound = 3;

		/// <summary>
		/// SystemSymbolConfig.xml file is not found at symbol share.
		/// </summary>
        public const int CmnSystemSymbolConfigFileNotFound = 4;

		/// <summary>
		/// An assembly mapping to the given ProgID was found in the symbol configuration file, but the assembly cannot be found in symbol share. This can happen if the assembly was deleted later after updating symbol configuration or if the mapping file was copied from another symbol share.
		/// </summary>
        public const int CmnAssemblyNotFoundInSymbolShare = 4;

		/// <summary>
		/// An assembly and class namespace mapping to the given ProgID was found in the symbol configuration file, but the assembly does not contain the class with given namespace and name. Check the source for assembly.
		/// </summary>
        public const int CmnClassNotFoundInAssembly = 6;
    }
}


