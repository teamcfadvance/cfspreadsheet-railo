<cffunction name="SpreadsheetAddRow" returntype="void" output="false">
	<cfargument name="spreadsheet" type="org.cfpoi.spreadsheet.Spreadsheet" required="true" />
	<cfargument name="data" type="string" required="true" hint="Delimited list of values" />
	<cfargument name="delimiter" type="string" required="false" default="," />
	<cfargument name="row" type="numeric" required="false" />
	<cfargument name="column" type="numeric" required="false" />
	
	<cfset var args = StructNew() />
	
	<cfset args.data = arguments.data />
	<cfset args.delimiter = arguments.delimiter />
	
	<cfif StructKeyExists(arguments, "row")>
		<cfset args.startRow = arguments.row />
	</cfif>
	
	<cfif StructKeyExists(arguments, "column")>
		<cfset args.startColumn = arguments.column />
	</cfif>
	
	<cfset arguments.spreadsheet.addRow(argumentcollection = args) />
</cffunction>