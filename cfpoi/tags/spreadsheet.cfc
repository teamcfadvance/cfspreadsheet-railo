<cfcomponent name="cfspreadsheet">
    <!--- Meta data --->
    <cfset this.metadata.hint="Handles spreadsheets">
    <cfset this.metadata.attributetype="fixed">
    <cfset this.metadata.attributes = {
        action : { required:false, type:"any", hint="<strong>read</strong>:Reads the contents of an XLS format file. <strong>update</strong>: Adds a new sheet to an existing XLS file. You cannot use the uppdate action to change an existing sheet in a file. For more information, see Usage. <strong>write</strong>: Writes a new XLS format file or overwrites an existing file." }, 
        filename : { required:false, type:"any", hint="The pathname of the file that is written." }, 
        excludeHeaderRow : { required:false, type:"any", hint="If set to true, excludes the headerrow from being included in the query results.The attribute helps when you read Excel as a query. When you specify the headerrow attribute, the column names are retrieved from the header row. But they are also included in the first row of the query. To not include the header row, set true as the attribute value." }, 
        name : { required:false, type:"any", hint="read action: The variable in which to store the spreadsheet file data. Specify name or query. write and update actions: A variable containing CSV-format data or an ColdFusion spreadsheet object containing the data to write. Specify the name or query." }, 
        query : { required:false, type:"any", hint="read action: The query in which to store the converted spreadsheet file. Specify format, name, or query. write and update actions: A query variable containing the data to write. Specify name or query." }, 
        src : { required:false, type:"any", hint="The pathname of the file to read." }, 
        columns : { required:false, type:"any", hint="Column number or range of columns. Specify a single number, a hypen-separated column range, a comma-separated list, or any combination of these; for example: 1,3-6,9." }, 
        columnnames : { required:false, type:"any", hint="Comma-separated column names." }, 
        format : { required:false, type:"any", hint="Format of the data represented by the name variable. All: csv On read, converts an XLS file to a CSV variable. On update or write, Saves a CSV variable as an XLS file. Read only: html Converts an XLS file to an HTML variable. The cfspreadsheet tag always writes spreadsheet data as an XLS file. To write HTML variables or CSV variables as HTML or CSV files, use the cffile tag." }, 
        headerrow : { required:false, type:"any", hint="Row number that contains column names." }, 
        overwrite : { required:false, type:"any", hint="A Boolean value specifying whether to overwrite an existing file." }, 
        password : { required:false, type:"any", hint="Set a password for modifying the sheet. Note: Setting a password of the empty string does no unset password protection entirely; you are still prompted for a password if you try to modify the sheet." }, 
        rows : { required:false, type:"any", hint="The range of rows to read. Specify a single number, a hypen-separated row range, a comma-separated list, or any combination of these; for example: 1,3-6,9." }, 
        sheet : { required:false, type:"any", hint="Number of the sheet. For the read action, you can specify sheet or sheetname." }, 
        sheetname : { required:false, type:"any", hint="Name of the sheet For the read action, you can specify sheet or sheetname. For write and update actions, the specified sheet is renamed according to the value you specify for sheetname." }
    }>
     
    <cffunction name="init" output="no" returntype="void" hint="invoked after tag is constructed">
        <cfargument name="hasEndTag" type="boolean" required="yes">
        <cfargument name="parent" type="component" required="no" hint="the parent cfc custom tag, if there is one">
        <cfset variables.hasEndTag = arguments.hasEndTag />
        <cfset variables.parent = arguments.parent />     
    </cffunction>
 
    <cffunction name="onStartTag" output="yes" returntype="boolean">
        <cfargument name="attributes" type="struct" />
        <cfargument name="caller" type="struct" />
        <cfset variables.attributes = arguments.attributes />
        <cfset var key = "" />
         
        <!--- name or query are required for all operations --->
        <cfif attributeExists('name') and attributeExists('query')>
            <cfthrow type="application"  message="Both 'name' and 'query' Attributes May Not Be Provided"  detail="Only one of either 'name' or 'query' may be provided" />
        </cfif>
         
        <cfif not attributeExists('name') and not attributeExists('query')>
            <cfthrow type="application" message="A 'name' or 'query' Attribute Is Required"  detail="Either 'name' or 'query' must be provided" />
        </cfif>       
         
        <cfswitch expression="#getAttribute('action')#">
 
            <!---        READ        --->
 
            <cfcase value="read">
                <cfset var spreadsheet = CreateObject("component", "org.cfpoi.spreadsheet.Spreadsheet").init() />
                 
                <cfif not attributeExists('src')>
                    <cfthrow type="application" message="Attribute 'src' is Required" detail="The 'src' attribute is required for the read action." />
                </cfif>
         
                <cfif attributeExists('columns') and attributeExists('columnnames')>
                    <cfthrow type="application" message="Both 'columns' and 'columnnames' Attributes May Not Be Provided"  detail="Only one of either 'columns' or 'columnnames' may be provided" />
                </cfif>
                 
                <cfif attributeExists("sheet") and attributeExists("sheetname")>
                    <cfthrow type="application"  message="Both 'sheet' and 'sheetname' Attributes May Not Be Provided"  detail="Only one of either 'sheet' or 'sheetname' may be provided" />
                </cfif>
                 
                <cfif attributeExists("query") and attributeExists("format")>
                    <cfthrow type="application"  message="Both 'query' and 'format' Attributes May Not Be Provided"  detail="Only one of either 'query' or 'format' may be provided" />
                </cfif>
                 
                <cfif attributeExists("name")>
                    <cfset caller[attributes.name] = spreadsheet.read(argumentcollection = attributes) />
                <cfelseif attributeExists("query")>
                    <cfset caller[attributes.query] = spreadsheet.read(argumentcollection = attributes) />
                </cfif>
            </cfcase>
         
            <!---        WRITE,UPDATE        --->
         
            <cfcase value="write,update" delimiters=",">
                <cfif not attributeExists("filename")>
                    <cfthrow type="application" message="Filename Attribute is Required"  detail="The 'filename' attribute must be provided for write and update actions" />
                </cfif>
                 
                <cfset attributes.filepath = attributes.filename />
                 
                <cfif attributes.action eq "update">
                    <cfset attributes.overwrite = true />
                    <cfset attributes.isUpdate = true />
                </cfif>
                 
                <cfif attributeExists("name")>
                    <cfif not attributeExists("format")>
                        <cfset caller[attributes.name].write(argumentcollection = attributes) />
                    <cfelse>
                        <cfset args = StructCopy(attributes) />
                        <cfset args.name = caller[attributes.name] />
                        <cfset spreadsheet = CreateObject("component", "org.cfpoi.spreadsheet.Spreadsheet").init() />
                        <cfset spreadsheet.write(argumentcollection = args) />
                    </cfif>
                <cfelseif attributeExists("query")>
                    <cfset args = StructCopy(attributes) />
                    <!--- <cfset args.query = caller[attributes.query] /> --->
                    <cfset args.query = attributes.query />                   
                    <cfset spreadsheet = CreateObject("component", "org.cfpoi.spreadsheet.Spreadsheet").init() />
                    <cfset spreadsheet.write(argumentcollection = args) />
                </cfif>
            </cfcase>
             
            <!---        CATCH INVALID OR MISSING ATTRIBUTES     --->
             
            <cfdefaultcase>
                <cfthrow type="application" message="Invalid or Missing Action Attribute"  detail="You must provide an action of 'read', 'update', or 'write' for this tag." />
            </cfdefaultcase>
        </cfswitch>       
         
        <cfreturn true>
    </cffunction>
 
    <cffunction name="onEndTag" output="yes" returntype="boolean">
        <cfargument name="attributes" type="struct">
        <cfargument name="caller" type="struct">
         
        <cfreturn false>
    </cffunction>
 
 
    <!---   attributes   --->
    <cffunction name="getAttribute" output="false" access="private" returntype="any">
        <cfargument name="key" required="true" type="String" />
        <cfreturn variables.attributes[key] />
    </cffunction>
 
    <cffunction name="attributeExists" output="false" access="private" returntype="boolean">
        <cfargument name="key" required="true" type="String" />
        <cfreturn structKeyExists(variables.attributes, key) />
    </cffunction>
 
</cfcomponent>