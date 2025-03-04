//*******************************************************************
//  Copyright (C) 1999,2003 Intergraph Corporation.  All rights reserved.
//
//  Project: NameRules
//
//  Abstract: The file contains a sample implementation of a naming rule
//       If the object is a part occurrence, then then the name rule is
//           "part name" + "Unique counter"
//       else if the object is not a part occurrence
//           "Class User Name" + "Unique counter"

//
//  Author: Bathina Balakrishna Reddy
//
// 03/21/00      Rob Lemley                  TR CP10966: NameRules.CommonNameRule does not support
//                                           different GSCAD object types
// 01 May 2000   Bathina Balakrishna Reddy   Removed client tier references and moved to middle tier.
// 17 Oct 2000   jpf                         Changed to use part name instead of type string in name
//
//       Modified to use the IJProdModelItem.TypeString (with blanks
//       removed) as the object//s "basename" for naming.
//
//       Fixed to remove extra blank in name introduced by the
//       VB Str() function.  Use Format() function and create
//       zero-padded fixed-width counter field in name.
//
// 1 June 2001   JY                     Modified to make use of IJNamedItem instead of
//                                           IJProdModelItem
//       SS Oct/07/2003
//           TR CP49909: fixed the usage of LocationID
// Oct/18/2003   JPF         Removed using the system parent name as prefix for named object
//
//  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols

//  17.MAR.2015     Ashok  CR-CP-269668  Rewriting the NameRules using 3DAPI
//  12 Jun 2015     Swetha  TR-CP-273048 	Get PartNumber fails for BuiltUp CrossSection
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Common.Rules
{
	public class CommonNameRule : NameRuleBase
	{
		/// <summary>
		/// Computes the name for the given entity.
		/// </summary>
		/// <param name="entity">The business object whose name is being computed.</param>
		/// <param name="parents">The parentsNaming parents of the business object,
		/// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
		/// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
		public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
		{
			if (entity == null)
			{
				throw new ArgumentNullException("The entity to be named is null");
			}

			if (parents == null)
			{
				throw new ArgumentNullException("The naming parents of the business object to be named are null");
			}

			if (activeEntity == null)
			{
				throw new ArgumentNullException("The name rule active entity associated to the business object to be named is null");
			}

			try
			{
				string partName = string.Empty;
				// Get part from part occurrence
				BusinessObject part = base.GetPart(entity);
				if (null != part)
				{
					// Get part number from part
					try
					{
						partName = base.GetPartNumber(part);
					}
					catch (CmnInvalidObjectTypeException ex)
					{
						// If given part does not implement IPart then use generic property of IStructConnection to get the part number

						if (part.SupportsInterface("IJCrossSection"))
						{
							PropertyValue propValue = part.GetPropertyValue("IStructCrossSection", "SectionName");
							PropertyValueString propertyValueString = (PropertyValueString)propValue;
							partName = propertyValueString.PropValue;
						}
						else
						{
							throw ex;
						}
					}
				}
				if (string.IsNullOrEmpty(partName))
				{
					if (entity.SupportsInterface("IJNamedItem"))
					{
						// Get "IIDForTypeString" property value from IJNamedItem interface
						partName = base.GetTypeString(entity);
					}
				}
				// Remove spaces from name
				partName = partName.Replace(" ", "").Trim();

				// Get naming parents string from active entity
				string nameBasis = base.GetNamingParentsString(activeEntity);
				// Check if new naming parent string and existing naming parent string are same
				// If they are same, we do not generate a new sequence number and there is no change to the name of the object.
				if(!partName.Equals(nameBasis))
				{
					// Set "NamingParentsString" property value on IJNameRuleAE interface
					base.SetNamingParentsString(activeEntity, partName);
					long count;
					string locationID = null;
					int range = 1;
					base.GetCountAndLocationID(partName, range, out count, out locationID);
					string name;
					if (!string.IsNullOrEmpty(locationID))
					{
						name = partName + "-" + locationID + "-" + count.ToString("D4"); ;
					}
					else
					{
						name = partName + "-" + count.ToString("D4"); ;
					}
					// Set name on entity
					base.SetName(entity, name);
				}
			}
			catch (Exception)
			{
				throw new Exception("Unexpected error:  CmnNameRules.CommonNameRule.ComputeName");
			}
		}

		/// <summary>
		/// Gets the naming parents from naming rule.
		/// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection.
		/// The parents added here are used in computing the name of the object in ComputeName() of the same interface.
		/// Both these methods are called from naming rule semantic.
		/// </summary>
		/// <param name="entity">BusinessObject for which naming parents are required.</param>
		/// <returns>Collection of BusinessObjects.</returns>
		public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
		{
			Collection<BusinessObject> parents = new Collection<BusinessObject>();
			return parents;
		}
	}
}