//'*******************************************************************
//'  Copyright (C) 2003, Intergraph Corporation.  All rights reserved.
//'
//'  Project: NameRules
//'  File:    SystemNameRule.cls
//'
//'  Abstract: The file contains a sample implementation of a naming rule
//'       If the object is a reference object. This simply returns the
//'    name of the given object. This is implemented separately instead of
//' handling in the property page so that any future enhancements could be addressed easily.
//'
//'  Author: Yalla Kiran Kumar
//'
//'  23-Aug-2004  Created
//'
//   25.MAR.2015  Vijay  CR-CP-269668  Rewriting the NameRules using 3DAPI
//'******************************************************************

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Common.Rules
{
	public class ReferenceNameRule : NameRuleBase
	{
		/// <summary>
		///  Computes the name for the given entity.
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
				string childName = string.Empty;
				INamedItem childNamedItem = entity as INamedItem;
				if (childNamedItem != null)
				{
					childName = childNamedItem.Name;
				}
				else
				{
					childName = base.GetName(entity);
				}
				// Set name on entity
				base.SetName(entity, childName);
			}
			catch (Exception)
			{
				throw new Exception("Unexpected error:  CmnNameRules.ReferenceNameRule.ComputeName");
			}
		}

		/// <summary>
		/// Returns empty collection of naming rules
		/// </summary>
		/// <param name="entity">>BusinessObject for which naming parents are required.</param>
		/// <returns>ReadOnlyCollection of BusinessObjects</returns>
		public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
		{
			return null;
		}
	}
}