<cfcomponent>

    <cffunction name="install" returntype="struct" output="no" hint="called from Railo to install application">
    	<cfargument name="error" type="struct">
        <cfargument name="path" type="string">
        <cfargument name="config" type="struct">
        
		<cfset var result = {status = true, message = ""} />
		<cfset var serverPath = expandPath('{railo-web-directory}') />
		
		<cftry>
			
			<!--- Export the CFPOI component --->
	        <cfzip
	            action = "unzip"
	            destination = "#serverPath#/components/org"
	            file = "#path#cfpoi.zip"
	            overwrite = "yes"
	            recurse = "yes"
	            storePath = "yes"/>

			<!--- Export the functions --->
	        <cfzip
	            action = "unzip"
	            destination = "#serverPath#/library/function"
	            file = "#path#functions.zip"
	            overwrite = "yes"
	            recurse = "yes"
	            storePath = "false"/>

			<!--- Export the tag --->
	        <cfzip
	            action = "unzip"
	            destination = "#serverPath#/library/tag"
	            file = "#path#tags.zip"
	            overwrite = "yes"
	            recurse = "yes"
	            storePath = "false"/>
			
			
			<!--- Export the jars --->
	        <cfzip
	            action = "unzip"
	            destination = "#serverPath#/lib"
	            file = "#path#poi-3.7.zip"
	            overwrite = "yes"
	            recurse = "yes"
	            storePath = "false"/>
		        
				<cfsavecontent variable="temp">
					<cfoutput>
						<p>Tags correctly installed. You will need to Restart Railo for the functions to work.</p>
					</cfoutput>				
				</cfsavecontent>
				
				<cfset result.message = temp />
			
			<cfcatch type="any">            
				<cfset result.status = false />
				<cfset result.message = cfcatch.message />
				<cflog file="railo_extension_install" text="Error: #cfcatch.message#">
			</cfcatch>			
        
	   </cftry>
	   
	   <cfreturn result />
	   
    </cffunction>
	
	<cffunction name="uninstall" returntype="struct" output="no" hint="called by Railo to uninstall the application">
        <cfargument name="path" type="any"/>
        <cfargument name="config" type="any"/>
        
		<cfset var processResult = {
			status = true,
			message = ""} />
		<cfset var ssDir = "" />
		<cfset var serverPath = expandPath('{railo-web-directory}') />
		
		<cftry>

			<cfdirectory action="delete" directory="#serverPath#/components/org/cfpoi" recurse="true" />

			<cffile action="delete" file="#serverPath#/lib/poi-3.7-20101029.jar" />
			<cffile action="delete" file="#serverPath#/lib/poi-ooxml-3.7-20101029.jar" />
			<cffile action="delete" file="#serverPath#/lib/poi-ooxml-schemas-3.7-20101029.jar" />
			<cffile action="delete" file="#serverPath#/lib/dom4j-1.6.1.jar" />
			<cffile action="delete" file="#serverPath#/lib/geronimo-stax-api_1.0_spec-1.0.jar" />
			<cffile action="delete" file="#serverPath#/lib/xmlbeans-2.3.0.jar" />
			
			<cfdirectory action="list" directory="#serverPath#/library/function" filter="Spreadsheet*" name="ssDir">
			
			<cfloop query="ssDir">
				<cffile action="delete" file="#ssDir.directory#/#ssDir.name#">
			</cfloop>			
			
			<cfcatch type="any">
				<cflog file="rail_extension_poi" text="#cfcatch.message#">
				<cfset processResult.status = false />
				<cfset processResult.message = cfcatch.message />				
			</cfcatch>
				
		</cftry>
		
		<cfreturn processResult />
	</cffunction>
 </cfcomponent>