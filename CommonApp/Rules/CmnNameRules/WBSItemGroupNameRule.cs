//'*******************************************************************
//'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
//'
//'  Project: NameRules
//'  File:    WBSItemGroupNameRule.cls
//'
//'  Abstract: The file contains a sample implementation of a naming rule which
//'           could be used in conjunction with the Create WBS Items from Piping Parts
//'           command.
//'
//'  Author: Mike Furno
//'
//'  23-May-2006  Created
//'  17.MAR.2015  Aditya  CR-CP-269668  Rewriting the WBSNameRules using 3DAPI

//'******************************************************************

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Common.Rules
{
	/// <summary>
	/// WBSItemGroupNameRule computes name for WBSItem
	/// NameRuleBase Class contains common implementation across all name rules and provides properties and methods for use in implementation of custom name rules.
	/// </summary>
	public class WBSItemGroupNameRule : NameRuleBase
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
				string newName = string.Empty;
				string label = string.Empty;
				string seqID = string.Empty;

				if (!entity.SupportsInterface("IJWBSItemGroup"))
				{
					newName = "WBS Item";
				}
				else
				{
					PropertyValue grpName = entity.GetPropertyValue("IJWBSItemGroup", "GroupName");
					PropertyValueString grpValue = (PropertyValueString)grpName;
					label = grpValue.PropValue;
					PropertyValue sequenceID = entity.GetPropertyValue("IJSequence", "Id");
					PropertyValueString seqIDValue = (PropertyValueString)sequenceID;
					seqID = seqIDValue.PropValue;
					newName = label + "-" + seqID;
				}

				string namingParentsString = base.GetNamingParentsString(activeEntity);

				//Name only needs to be recomputed and if the naming parent string is different from the existing one.
				if (!newName.Equals(namingParentsString))
				{
					//Set the naming parents string to obtained new name.
					base.SetNamingParentsString(activeEntity, newName);
					//Set name on entity
					base.SetName(entity, newName);
				}
			}
			catch (Exception)
			{
				throw new Exception("Unexpected error:  CmnNameRules.WBSItemGroupNameRule.ComputeName");
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