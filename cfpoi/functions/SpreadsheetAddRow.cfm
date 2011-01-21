<!--- Compatibility note: ACF does not yet support a "delimiter" argument --->
<cffunction name="SpreadsheetAddRow" returntype="void" output="false">
	<cfargument name="spreadsheet" type="org.cfpoi.spreadsheet.Spreadsheet" required="true" />
	<cfargument name="data" type="string" required="true" hint="Delimited list of values" />
	<cfargument name="row" type="numeric" required="false" />
	<cfargument name="column" type="numeric" required="false" />
	<cfargument name="insert" type="boolean" default="true" />
	<cfargument name="delimiter" type="string" default="," />
	
	<cfset var args = StructNew() />
	
	<cfset args.data 		= arguments.data />
	<cfset args.insert 		= arguments.insert />
	<cfset args.delimiter 	= arguments.delimiter />
	
	<cfif StructKeyExists(arguments, "row")>
		<cfset args.startRow = arguments.row />
	</cfif>
	
	<cfif StructKeyExists(arguments, "column")>
		<cfset args.startColumn = arguments.column />
	</cfif>
	
	<cfset arguments.spreadsheet.addRow(argumentcollection = args) />
</cffunction>