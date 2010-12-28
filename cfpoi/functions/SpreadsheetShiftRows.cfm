<cffunction name="SpreadsheetShiftRows" returntype="void" output="false">
	<cfargument name="spreadsheet" type="org.cfpoi.spreadsheet.Spreadsheet" required="true" />
	<cfargument name="start" type="numeric" required="true" />
	<cfargument name="end" type="numeric" required="false" />
	<cfargument name="rows" type="numeric" required="false" />
	
	<cfset var args = StructNew() />
	
	<cfset args.startRow = arguments.start />
	
	<cfif StructKeyExists(arguments, "end")>
		<cfset args.endRow = arguments.end />
	</cfif>
	
	<cfif StructKeyExists(arguments, "rows")>
		<cfset args.offset = arguments.rows />
	</cfif>
	
	<cfset arguments.spreadsheet.shiftRows(argumentcollection = args) />
</cffunction>
