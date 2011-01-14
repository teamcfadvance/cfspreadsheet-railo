<cffunction name="SpreadsheetAddColumn" returntype="void" output="false">
	<cfargument name="spreadsheet" type="org.cfpoi.spreadsheet.Spreadsheet" required="true" />
	<cfargument name="data" type="string" required="true" hint="Delimited list of values" />
	<cfargument name="startRow" type="numeric" required="false" />
	<cfargument name="startColumn" type="numeric" required="false" />
	<cfargument name="insert" type="boolean" required="false" />
	<cfargument name="delimiter" type="string" required="false" default="," />
	
	<cfset var args = StructNew() />
	
	<cfset args.data = arguments.data />
	<cfset args.delimiter = arguments.delimiter />
	
	<cfif StructKeyExists(arguments, "startRow")>
		<cfset args.startRow = arguments.startRow />
	</cfif>
	
	<cfif StructKeyExists(arguments, "startColumn")>
		<cfset args.column = arguments.startColumn />
	</cfif>
	
	<cfif StructKeyExists(arguments, "insert")>
		<cfset args.insert = arguments.insert />
	</cfif>
	
	<cfset arguments.spreadsheet.addColumn( argumentCollection=args ) />
</cffunction>