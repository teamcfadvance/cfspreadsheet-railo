<!--- TODO: only supporting HSSF (non-xslx format) for now --->
<cffunction name="SpreadsheetRead" returntype="any" output="false">
	<cfargument name="src" type="string" required="true" hint="Path to an existing workbook file on disk" />
	<cfargument name="sheetName" type="string" required="false" hint="Sheet name to activate" />
	<cfargument name="sheet" type="numeric" required="false" hint="Sheet number to activate" />

	<cfif structKeyExists(arguments, "sheetName") and structKeyExists(arguments, "sheet")>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
					message="Invalid Argument Combination" 
					detail="Either specify a 'SheetName' OR 'Sheet', but not both.">
	</cfif>	
	
	<cfreturn CreateObject("component", "org.cfpoi.spreadsheet.Spreadsheet").init( argumentCollection=arguments ) />
</cffunction>
