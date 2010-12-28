<cfcomponent extends="InstallFolder">

    <cffunction name="validate" returntype="void" output="no" hint="called to validate the entered">
    	<cfargument name="error" type="struct">
        <cfargument name="path" type="string">
        <cfargument name="config" type="struct">
        <cfargument name="step" type="numeric">
     </cffunction>
    
    <cffunction name="install" returntype="string" output="yes" hint="called from Railo to install application">
    	<cfargument name="error" type="struct">
        <cfargument name="path" type="string">
        <cfargument name="config" type="struct">
		<cfset var sReturn = "">
		<cfset var temp = "" >

		<cfset var stReturn = super.install(argumentCollection:arguments)>

		<cfif stReturn.status>
			<cfset sReturn = stReturn.message />
		<cfelse>
			<cfsavecontent variable="sReturn">
				<cfset uninstall(argumentCollection=arguments)>
				<p style="color:red">Tags has not been installed.</p>
				<cfoutput>#stReturn.message#</cfoutput>
			</cfsavecontent>
		</cfif>
        <cfreturn sReturn>
    </cffunction>

   <cffunction name="uninstall" returntype="string" output="no" hint="called by Railo to uninstall the application">
        <cfargument name="path" type="string">
        <cfargument name="config" type="struct">
		<cfset var sReturn = "">

		<cfset var stReturn = super.uninstall(argumentCollection:arguments)>

		<cfif stReturn.status>
			<cfif len(trim(stReturn.message))>
				<cfset sReturn = stReturn.message>
			<cfelse>
				<cfsavecontent variable="sReturn">
					<p>Tags has been successfully removed!</p>
				</cfsavecontent>
			</cfif>
		<cfelse>
			<cfsavecontent variable="sReturn">
				<cfset uninstall(argumentCollection=arguments)>
				<p style="color:red">Some error occurred during the uninstalling proceed:</p>
				<cfoutput>#stReturn.message#</cfoutput>
			</cfsavecontent>
		</cfif>
        <cfreturn sReturn>
    </cffunction>

    <cffunction name="update" returntype="string" output="no" hint="called from Railo to update a existing application">
		<cfreturn install(argumentCollection=arguments)>
    </cffunction>

</cfcomponent>