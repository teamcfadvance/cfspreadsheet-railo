<cffunction name="SpreadsheetNew" returntype="org.cfpoi.spreadsheet.Spreadsheet" output="false">
	<cfargument name="sheetName" type="string" required="false" />
	<cfargument name="xmlFormat" type="boolean" required="false" default="false" />
	
	<!--- TODO: only supporting HSSF (non-xslx format) for now --->
	<cfreturn CreateObject("component", "org.cfpoi.spreadsheet.Spreadsheet").init() />
</cffunction>
