<cffunction name="SpreadsheetAddRows" returntype="void" output="false">
	<cfargument name="spreadsheet" type="org.cfpoi.spreadsheet.Spreadsheet" required="true" />
	<cfargument name="data" type="query" required="true" />
	<cfargument name="row" type="numeric" required="false" />
	<cfargument name="column" type="numeric" required="false" />
	<cfargument name="insert" type="boolean" default="true" />
	
	<cfset var args = StructNew() />

	<cfset args.data   = arguments.data />
	<cfset args.insert = arguments.insert />
	
	<cfif StructKeyExists(arguments, "row")>
		<cfset args.row = arguments.row />
	</cfif>
	<cfif StructKeyExists(arguments, "column")>
		<cfset args.column = arguments.column />
	</cfif>
	
	<cfset arguments.spreadsheet.addRows(argumentcollection = args) />
</cffunction>