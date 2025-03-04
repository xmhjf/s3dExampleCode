
Author : Chaitanya

Date   : 24th April 2009.

______________________________________________________________________________________
Instructions for using Grids .NET(C# and VB.NET) Naming Rules on End-User Machine
--------------------------------------------------------------------------------------

The .NET implementation of Grids Naming Rules is provided in C# and VB.NET

    C#.NET project is located at "\GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetCS"   
    VB.NET project is located at "\GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetVB"   

____________________________________________________________________
Step 1. Re-reference S3DAPI assemblies suitable for End-User Machine
--------------------------------------------------------------------

--> The .NET NameRule project "\GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetCS\GridsNamingRulesNetCS\GridsNamingRulesNetCS.csproj"
    has references to 
	(a) "\Container\Bin\Assemblies\Release\CommonMiddle.dll"  and
	(b) "\Container\Bin\Assemblies\Release\GridsMiddle.dll"

    For End User, these references are to be removed and replaced with 
	(a) "Core\Container\Bin\Assemblies\Release\CommonMiddle.dll" and
	(b) "Core\Container\Bin\Assemblies\Release\GridsMiddle.dll"

--> Similar Re-referencing is required for "\GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetVB\GridsNamingRulesNetVB\GridsNamingRulesNetVB.vbproj"

____________________________
Step 2. Building the project 
----------------------------

-->  For using C#.NET Naming Rule, build '\GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetCS\GridsNamingRulesNetCS.sln' to 

    generate '\GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetCS\GridsNamingRulesNetCSCAB\Debug\GridsNamingRulesNetCSCAB.CAB" file.


-->  Similarly for VB.NET, build '\GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetVB\GridsNamingRulesNetVB.sln' to 

    generate '\GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetVB\GridsNamingRulesNetVBCAB\Debug\GridsNamingRulesNetVBCAB.CAB'.

_______________________________
Step 3. Deployment instructions 
-------------------------------

-->  Copy the "GridsNamingRulesNetCSCAB.CAB" and "GridsNamingRulesNetVBCAB.CAB" files to any folder under Symbol Share. 

	Ex:- < symbolshare >\NameRules\


_________________________________
Step 4. Configuring the NameRules
---------------------------------

-->  Add the Solver ProgID for the NameRule to 'NamingRules' sheet in "\CatalogData\BulkLoad\DataFiles\GenericNamingRules.xls" in the following format

	"[AssemblyName],[Namespace].[ClassName]|Dir\CABFile.cab". 

Note: Here "Dir" is the relative path of the CAB files in the symbol share.

   
Following are some examples explaining how the Solver ProgID for VB.NET and C#.NET Naming Rules are to be given.


                   Head         TypeName            Name                        SolverProgID
                   ----         --------            ----                        ------------ 

(Actual NameRule)           CSPGElevationPlane     Position		      GSNamingRules.PositionNameRule

( C#.NET NameRule)  A	    CSPGElevationPlane     PositionForCSNET	      GridsNamingRulesNetCS,GridsNamingRulesNetCS.PositionNameRule|NameRules\GridsNamingRulesNetCSCab.CAB

( VB.NET NameRule)  A	    CSPGElevationPlane     PositionForVBNET	      GridsNamingRulesNetVB,GridsNamingRulesNetVB.PositionNameRule|NameRules\GridsNamingRulesNetVBCAB.CAB
	

Note: The above SolverProgIDs strings (.NET NameRules) are conveying that the CAB files are copied to "< symbolshare >\NameRules\" folder.

-->  Put 'A' in the first column ("Head")  of 'NamingRules' sheet in \CatalogData\BulkLoad\DataFiles\GenericNamingRules.xls for the added rows.

-->  For Bulkloading the above xls sheet, invoke '\CatalogData\BulkLoad\Bin\Bulkload.exe' and provide the following details.

	a) Add '\CatalogData\BulkLoad\DataFiles\GenericNamingRules.xls' to "Excel files" list.
		
	b) Select Bulkload Mode as " Add,Modify, or Delete records in existing catalog ". ( AMD mode)

	c) Provide the catalog information i.e, server name, Catalog database and Catalog_SCHEMA database names.
		
	d) Provide 'log file' path and 'Symbol share' path in respective textboxes.
 
	e) Click 'Load' to BulkLoad the GenericNamingRules.xls sheet.


_____________________________
Step 5. Testing the NameRules
-----------------------------

-->  After successfully bulkloading GenericNamingRules.xls sheet in AMD mode , test the new Naming Rules.

-->  Create a "CSPGElevationPlane" type (i.e, Elevation Plane / Z-Plane) and open its Property page. 

-->  On the property page select "PositionForCSNET" or "PositionForVBNET" from "Name Rule" ComboBOx to apply the .NET naming rules.


