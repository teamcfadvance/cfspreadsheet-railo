<cffunction name="SpreadsheetAddRows" returntype="void" output="false">
	<cfargument name="spreadsheet" type="org.cfpoi.spreadsheet.Spreadsheet" required="true" />
	<cfargument name="data" type="query" required="true" />
	<cfargument name="row" type="numeric" required="false" />
	
	<cfset var args = StructNew() />
	
	<cfset args.data = arguments.data />
	
	<cfif StructKeyExists(arguments, "row")>
		<cfset args.row = arguments.row />
	</cfif>
	
	<cfset arguments.spreadsheet.addRows(argumentcollection = args) />
</cffunction>