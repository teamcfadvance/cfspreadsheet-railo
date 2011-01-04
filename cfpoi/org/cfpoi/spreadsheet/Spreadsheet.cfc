<cfcomponent 
	displayname="Spreadsheet" 
	output="false" 
	hint="CFC wrapper for the Apache POI project's HSSF (xls) and XSSF (xlsx) classes">
	
	<cffunction name="loadPoi" access="private" output="false" returntype="any">
		<cfargument name="javaclass" type="string" required="true" hint="I am the java class to be loaded" />
		<cfargument name="javainit" type="string" required="false" hint="I am the java initilising parameters" />
		<cfscript>
			//create the loader
			local.paths = arrayNew(1);
			// This points to the jar we want to load. Could also load a directory of .class files
			local.paths[1] = expandPath('{railo-web-directory}'&'/lib/poi-3.7-20101029.jar');
			local.paths[2] = expandPath('{railo-web-directory}'&'/lib/poi-ooxml-3.7-20101029.jar');
			local.paths[3] = expandPath('{railo-web-directory}'&'/lib/poi-ooxml-schemas-3.7-20101029.jar');		
			local.paths[4] = expandPath('{railo-web-directory}'&'/lib/dom4j-1.6.1.jar');		
			local.paths[5] = expandPath('{railo-web-directory}'&'/lib/geronimo-stax-api_1.0_spec-1.0.jar');		
			local.paths[6] = expandPath('{railo-web-directory}'&'/lib/xmlbeans-2.3.0.jar');		
		
			if( NOT structKeyExists( server, "_poiLoader")){
				server._poiLoader = createObject("component", "javaloader.JavaLoader").init(loadPaths = local.paths, trustedSource=true);
			}
			//at this stage we only have access to the class, but we don't have an instance
			var classInstance = server._poiLoader.create( arguments.javaclass);
			/*
			Create the instance, just like createObject("java", "HelloWorld").init();
			*/
			if(structKeyExists(arguments, "javainit")){
				var jclass = classInstance.init( arguments.javainit );						
			} else{
				var jclass = classInstance.init();		
			}
			
		</cfscript>		
		<cfreturn jclass />
	</cffunction>	
	
	
	<!--- CONSTRUCTOR --->
	<cffunction name="init" access="public" output="false" returntype="Spreadsheet">
		<!--- if init is called, assume it's because they want a new workbook with a blank sheet ---->
		<cfset var workbook = CreateObject("java", "org.apache.poi.hssf.usermodel.HSSFWorkbook").init() />
		
		<cfset workbook.createSheet(JavaCast("string", "Sheet1")) />
		<cfset setWorkbook(workbook) />
		<cfset setActiveSheet("Sheet1") />
		
		<cfreturn this />
	</cffunction>
	
	<!--- BASIC READ/WRITE/UPDATE FUNCTIONS --->
	<!--- TODO: need to handle arguments of columns, columnnames, and rows --->
	<cffunction name="read" access="public" output="false" returntype="any" 
			hint="Reads a spreadsheet from disk and returns a Spreadsheet CFC, query, CSV, or HTML">
		<cfargument name="src" type="string" required="true" hint="The full file path to the spreadsheet" />
		<cfargument name="columns" type="string" required="false" />
		<cfargument name="columnnames" type="string" required="false" />
		<cfargument name="format" type="string" required="false" />
		<cfargument name="headerrow" type="numeric" required="false" />
		<cfargument name="query" type="string" required="false" />
		<cfargument name="rows" type="string" required="false" />
		<cfargument name="sheet" type="numeric" required="false" />
		<cfargument name="sheetname" type="string" required="false" />
	
		<cfset var args = StructNew() />
		<cfset var rowIterator = 0 />
		<cfset var row = 0 />
		<cfset var cellIterator = 0 />
		<cfset var csv = "" />
		<cfset var html = "" />
		<cfset var theQuery = "" />
		<cfset var queryColumnName = "" />
		<cfset var queryColumnNames = "" />
		<cfset var i = 0 />
		<cfset var lineSeparator = CreateObject("java", "java.lang.System").getProperty("line.separator") />
		<cfset var returnVal = this />
		
		<cfset args.src = arguments.src />
		
		<cfif StructKeyExists(arguments, "sheet")>
			<cfset args.sheet = arguments.sheet />
		</cfif>
		
		<cfif StructKeyExists(arguments, "sheetname")>
			<cfset args.sheetname = arguments.sheetname />
		</cfif>
		
		<cfif StructKeyExists(arguments, "query")>
			<cfset arguments.format = "query" />
		</cfif>
		
		<cfset setWorkbook(readFromFile(argumentcollection = args)) />
		
		<cfif StructKeyExists(arguments, "format")>
			<cfset rowIterator = getActiveSheet().rowIterator() />
			
			<cfswitch expression="#arguments.format#">
				<cfcase value="csv">
					<cfloop condition="#rowIterator.hasNext()#">
						<cfset row = rowIterator.next() />
						
						<cfset cellIterator = row.cellIterator() />
						
						<cfloop condition="#cellIterator.hasNext()#">
							<cfset csv = csv & getCellValue(row.getRowNum() + 1, cellIterator.next().getColumnIndex() + 1) & "," />
						</cfloop>
						
						<cfset csv = Left(csv, Len(csv) - 1) & lineSeparator />
					</cfloop>
					
					<cfset returnVal = csv />
				</cfcase>
				
				<cfcase value="html">
					<cfloop condition="#rowIterator.hasNext()#">
						<cfset row = rowIterator.next() />
						
						<cfset html = html & "<tr>" />
						
						<cfset cellIterator = row.cellIterator() />
						
						<cfloop condition="#cellIterator.hasNext()#">
							<cfset html = html & Chr(9) & "<td>" & getCellValue(row.getRowNum() + 1, cellIterator.next().getColumnIndex() + 1) & "</td>" />
						</cfloop>
						
						<cfset html = html & "</tr>" & lineSeparator />
					</cfloop>
					
					<cfset returnVal = html />
				</cfcase>

				<cfcase value="query">
					<!--- If a header row is specified, use that for the query column names.
							Otherwise, use COL_1, COL_2, etc. for column names. --->
					<cfif StructKeyExists(arguments, "headerrow")>
						<cfset row = getActiveSheet().getRow(arguments.headerrow - 1) />
					<cfelse>
						<cfset row = getActiveSheet().getRow(0) />
					</cfif>

					<!--- If the sheet is empty the row == null and we have value to iterate over--->
					
					<cfif NOT isNull( getActiveSheet().getRow(0) )>
					
						<cfset cellIterator = row.cellIterator() />

						<cfset i = 1 />
					
						<cfloop condition="#cellIterator.hasNext()#">
							<cfset queryColumnName = getCellValue(row.getRowNum() + 1, cellIterator.next().getColumnIndex() + 1) />
							
							<cfif not StructKeyExists(arguments, "headerrow")>
								<cfset queryColumnName = "COL_" & i />
							</cfif>
							
							<cfset queryColumnNames = queryColumnNames & queryColumnName & "," />
							
							<cfset i = i + 1 />
						</cfloop>
					
						<cfset queryColumnNames = Left(queryColumnNames, Len(queryColumnNames) - 1) />
						
						<cfset query = QueryNew(queryColumnNames) />
						
						<cfset i = 1 />
					
						<cfloop condition="#rowIterator.hasNext()#">
							<cfset QueryAddRow(query, 1) />
							
							<cfset row = rowIterator.next() />
							<cfset cellIterator = row.cellIterator() />
							
							<cfloop condition="#cellIterator.hasNext()#">
								<cfset QuerySetCell(query, ListGetAt(queryColumnNames, i), getCellValue(row.getRowNum() + 1, cellIterator.next().getColumnIndex() + 1)) />
								
								<cfset i = i + 1 />
							</cfloop>
							
							<cfset i = 1 />
						</cfloop>
					<cfelse>
						<cfset query = queryNew("")/>
					</cfif>
					
					<cfset returnVal = query />
				</cfcase>
				
				<cfdefaultcase>
					<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
								message="Invalid Format" 
								detail="Only formats of csv, html, and query are supported" />
				</cfdefaultcase>
			</cfswitch>
		</cfif>
		
		<cfreturn returnVal />
	</cffunction>
	
	<cffunction name="write" access="public" output="false" returntype="void" 
			hint="Writes a spreadsheet to disk">
		<cfargument name="filepath" type="string" required="true" />
		<cfargument name="format" type="string" required="false" />
		<cfargument name="name" type="string" required="false" />
		<cfargument name="overwrite" type="boolean" required="false" default="false" />
		<cfargument name="password" type="string" required="false" />
		<cfargument name="query" type="query" required="false" />
		<cfargument name="sheet" type="numeric" required="false" />
		<cfargument name="sheetname" type="string" required="false" />
		<cfargument name="isUpdate" type="boolean" required="false" default="false" />
		
		<cfset var sheetToWrite = 0 />
		<cfset var row = 0 />
		<cfset var cell = 0 />
		<cfset var queryColumnList = 0 />
		<cfset var i = 0 />
		<cfset var j = 0 />
		<cfset var csvRows = 0 />
		
		<cfif StructKeyExists(arguments, "query") and StructKeyExists(arguments, "format")>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Argument Combination" 
						detail="Both 'query' and 'format' may not be provided." />
		</cfif>
		
		<!--- <cfdump var="#arguments#"/>
		<cfabort> --->
		
		<cfif StructKeyExists(arguments, "sheetname")>

			<cfif arguments.isUpdate and getWorkbook().getSheet(JavaCast("string", arguments.sheetname)) neq "">
				<cfset getWorkbook().removeSheetAt(JavaCast("int", getWorkbook().getSheetIndex(JavaCast("string", arguments.sheetname)))) />
			</cfif>
			
			<cfif getWorkbook().getNumberOfSheets() eq 1 
					and getWorkbook().getSheetAt(JavaCast("int", 0)).getPhysicalNumberOfRows() eq 0>
				<cfset getWorkbook().removeSheetAt(JavaCast("int", 0)) />
			</cfif>

			<cfset sheetToWrite = getWorkbook().createSheet(JavaCast("string", arguments.sheetname)) />
		<cfelseif StructKeyExists(arguments, "sheet")>
			<cfif arguments.isUpdate and arguments.sheet lte getWorkbook().getNumberOfSheets() 
					and getWorkbook().getSheetAt(JavaCast("int", arguments.sheet - 1)) neq "">
				<cfset getWorkbook().removeSheetAt(JavaCast("int", arguments.sheet - 1)) />
			</cfif>

			<cfset sheetToWrite = getWorkbook().createSheet(JavaCast("string", "Sheet" & arguments.sheet)) />
		<cfelse>
		
			<cfif getWorkbook().getNumberOfSheets() eq 0 
				or getWorkbook().getSheetAt(JavaCast("int", getWorkbook().getNumberOfSheets() - 1)).getPhysicalNumberOfRows() eq 0>
				
				<!--- getSheetAt() does not bring back a simple value	
					<cfif getWorkbook().getSheetAt(JavaCast("int", getWorkbook().getNumberOfSheets() - 1)) neq "">
						<cfset getWorkbook().removeSheetAt(JavaCast("int", getWorkbook().getNumberOfSheets() - 1)) />
					</cfif>
				--->
				
				<cfif getWorkbook().getSheetAt(JavaCast("int", getWorkbook().getNumberOfSheets() - 1)).getPhysicalNumberOfRows() EQ 0>
					<cfset getWorkbook().removeSheetAt(JavaCast("int", getWorkbook().getNumberOfSheets() - 1)) />
				</cfif>				
				
				<cfset sheetToWrite = getWorkbook().createSheet(JavaCast("string", "Sheet" & getWorkbook().getNumberOfSheets() + 1)) />
			<cfelse>
				<cfset sheetToWrite = getWorkbook().createSheet(JavaCast("string", "Sheet" & getWorkbook().getNumberOfSheets() + 1)) />
			</cfif>
		</cfif>
		
		<!--- If name is supplied and format isn't, we're just writing the workbook to disk. 
				Otherwise, handle query or CSV accordingly. --->
		<cfif StructKeyExists(arguments, "query")>
			<!--- loop over the query and populate a sheet object --->
			<cfset queryColumnList = arguments.query.ColumnList />
			
			<cfloop query="arguments.query">
				<cfset row = sheetToWrite.createRow(JavaCast("int", arguments.query.CurrentRow - 1)) />
				
				<!--- TODO: should we determine data types and set the cells accordingly 
							or just leave everything as a string? --->
				<cfloop index="i" from="1" to="#ListLen(queryColumnList)#">
					<cfset cell = row.createCell(JavaCast("int", i - 1)) />
					<cfset cell.setCellValue(JavaCast("string", arguments.query[ListGetAt(queryColumnList, i)][arguments.query.CurrentRow])) />
				</cfloop>
			</cfloop>
		<cfelseif StructKeyExists(arguments, "format")>
			<cfif UCase(arguments.format) eq "CSV">
				<!--- for csv format the assumption is it's csv (duh) and one sheet row per line (double duh) --->
				<cfset csvRows = arguments.name.split("\r\n|\n") />
				
				<cfloop index="i" from="1" to="#ArrayLen(csvRows)#">
					<cfset row = sheetToWrite.createRow(JavaCast("int", i - 1)) />
					
					<cfloop index="j" from="1" to="#ListLen(csvRows[i])#">
						<cfset cell = row.createCell(JavaCast("int", j - 1)) />
						<cfset cell.setCellValue(JavaCast("string", ListGetAt(csvRows[i], j))) />
					</cfloop>
				</cfloop>
			<cfelse>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Unsupported Write Format" 
							detail="The format #arguments.format# is not supported for write operations." />
			</cfif>
		</cfif>
		
		<cfif StructKeyExists(arguments, "password") and arguments.password neq "">
			<!--- writeProtectWorkbook takes both a user name and a password, but 
					since CF 9 tag only takes a password, just making up a user name --->
			<!--- TODO: workbook.isWriteProtected() returns true but the workbook opens 
						without prompting for a password --->
			<cfset getWorkbook().writeProtectWorkbook(JavaCast("string", arguments.password), JavaCast("string", "user")) />
		</cfif>
		
		<cfset writeToFile(arguments.filepath, getWorkbook(), arguments.overwrite) />
	</cffunction>
	
	<!--- TODO: CF 9 doesn't allow for overwriting a sheet with the same name on an update, which seems 
				strange to me. Makes sense on a write, but not on an update IMO. --->
	<cffunction name="update" access="public" output="false" returntype="void" 
			hint="Updates a workbook with a new sheet or overwrites an existing sheet with the same name">
		<cfargument name="filepath" type="string" required="true" />
		<cfargument name="format" type="string" required="false" />
		<cfargument name="name" type="string" required="false" />
		<cfargument name="password" type="string" required="false" />
		<cfargument name="query" type="query" required="false" />
		<cfargument name="sheet" type="string" required="false" />
		<cfargument name="sheetname" type="string" required="false" />

		<cfset arguments.workbookToUpdate = read(arguments.filepath).getWorkbook() />
		<cfset arguments.overwrite = true />
		<cfset write(argumentcollection = arguments) />
	</cffunction>
	
	<!--- SPREADSHEET MANIPULATION FUNCTIONS --->
	<!--- sheet functions --->
	<cffunction name="addFreezePane" access="public" output="false" returntype="void" 
			hint="Adds a split ('freeze pane') to the sheet">
		<cfargument name="splitColumn" type="numeric" required="true" 
				hint="Horizontal position of split" />
		<cfargument name="splitRow" type="numeric" required="true" 
				hint="Vertical position of split" />
		<cfargument name="leftmostColumn" type="numeric" required="false" 
				hint="Left column visible in right pane" />
		<cfargument name="topRow" type="numeric" required="false" 
				hint="Top row visible in bottom pane" />
		
		<cfif StructKeyExists(arguments, "leftmostColumn") 
				and not StructKeyExists(arguments, "topRow")>
			<cfset arguments.topRow = arguments.splitRow />
		</cfif>
		
		<cfif StructKeyExists(arguments, "topRow") 
				and not StructKeyExists(arguments, "leftmostColumn")>
			<cfset arguments.leftmostColumn = arguments.splitColumn />
		</cfif>
		
		<!--- createFreezePane() operates on the logical row/column numbers as opposed to physical, 
				so no need for n-1 stuff here --->
		<cfif not StructKeyExists(arguments, "leftmostColumn")>
			<cfset getActiveSheet().createFreezePane(JavaCast("int", arguments.splitColumn), 
													JavaCast("int", arguments.splitRow)) />
		<cfelse>
			<!--- POI lets you specify an active pane if you use createSplitPane() here --->
			<cfset getActiveSheet().createFreezePane(JavaCast("int", arguments.splitColumn), 
													JavaCast("int", arguments.splitRow), 
													JavaCast("int", arguments.leftmostColumn), 
													JavaCast("int", arguments.topRow)) />
		</cfif>
	</cffunction>
	
	<!--- the CF 9 docs seem to be wrong on what the last argument means ... or 
			they're combining split pane and freeze pane --->
	<cffunction name="createSplitPane" access="public" output="false" returntype="void" 
			hint="Adds a split pane to a sheet, which differs from a freeze pane in that it has x and y positioning">
		<cfargument name="xSplitPos" type="numeric" required="true" />
		<cfargument name="ySplitPos" type="numeric" required="true" />
		<cfargument name="leftmostColumn" type="numeric" required="true" />
		<cfargument name="topRow" type="numeric" required="true" />
		<cfargument name="activePane" type="string" required="false" default="UPPER_LEFT" 
				hint="Valid values are LOWER_LEFT, LOWER_RIGHT, UPPER_LEFT, and UPPER_RIGHT" />
		
		<cfset arguments.activePane = Evaluate("getActiveSheet().PANE_#arguments.activePane#") />
		
		<cfset getActiveSheet().createSplitPane(JavaCast("int", arguments.xSplitPos), 
											JavaCast("int", arguments.ySplitPos), 
											JavaCast("int", arguments.leftmostColumn), 
											JavaCast("int", arguments.topRow), 
											JavaCast("int", arguments.activePane)) />
	</cffunction>
	
	<!--- TODO: Should we allow for passing in of a boolean indicating whether or not an image resize 
				should happen (only works on jpg and png)? Currently does not resize. If resize is 
				performed, it does mess up passing in x/y coordinates for image positioning. --->
	<cffunction name="addImage" access="public" output="false" returntype="void" 
			hint="Adds an image to the workbook. Valid argument combinations are filepath + anchor, or imageData + imageType + anchor">
		<cfargument name="filepath" type="string" required="false" />
		<cfargument name="imageData" type="any" required="false" />
		<cfargument name="imageType" type="string" required="false" />
		<cfargument name="anchor" type="string" required="true" />
		
		<cfset var toolkit = CreateObject("java", "java.awt.Toolkit") />
		<!--- For some reason calling creationHelper.createClientAnchor() bombs with a 'could not instantiate object' 
				error, so we'll create the anchor manually later. Just leaving this in here in case it's worth another 
				look. --->
		<!--- <cfset var creationHelper = CreateObject("java", "org.apache.poi.hssf.usermodel.HSSFCreationHelper") /> --->
		<cfset var ioUtils = CreateObject("java", "org.apache.poi.util.IOUtils") />
		<cfset var inputStream = 0 />
		<cfset var bytes = 0 />
		<cfset var picture = 0 />
		<cfset var imgType = "" />
		<cfset var imgTypeIndex = 0 />
		<cfset var imageIndex = 0 />
		<cfset var theAnchor = 0 />
		<!--- TODO: need to look into createDrawingPatriarch() vs. getDrawingPatriarch() 
					since create will kill any existing images. getDrawingPatriarch() throws 
					a null pointer exception when an attempt is made to add a second 
					image to the spreadsheet --->
		<cfset var drawingPatriarch = getActiveSheet().createDrawingPatriarch() />
		
		<!--- we'll need the image type int in all cases --->
		<cfif StructKeyExists(arguments, "filepath")>
			<!--- TODO: better way to determine image type for physical files? using file extension for now --->
			<cfset imgType = UCase(ListLast(arguments.filePath, ".")) />
		<cfelseif StructKeyExists(arguments, "imageType")>
			<cfset imgType = UCase(arguments.imageType) />
		<cfelse>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Could Not Determine Image Type" 
						detail="An image type could not be determined from the filepath or imagetype provided" />
		</cfif>
		
		<cfswitch expression="#imgType#">
			<cfcase value="DIB">
				<cfset imgTypeIndex = getWorkbook().PICTURE_TYPE_DIB />
			</cfcase>
			
			<cfcase value="EMF">
				<cfset imgTypeIndex = getWorkbook().PICTURE_TYPE_EMF />
			</cfcase>
			
			<cfcase value="JPG,JPEG" delimiters=",">
				<cfset imgTypeIndex = getWorkbook().PICTURE_TYPE_JPEG />
			</cfcase>

			<cfcase value="PICT">
				<cfset imgTypeIndex = getWorkbook().PICTURE_TYPE_PICT />
			</cfcase>
			
			<cfcase value="PNG">
				<cfset imgTypeIndex = getWorkbook().PICTURE_TYPE_PNG />
			</cfcase>
			
			<cfcase value="WMF">
				<cfset imgTypeIndex = getWorkbook().PICTURE_TYPE_WMF />
			</cfcase>
			
			<cfdefaultcase>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Invalid Image Type" 
							detail="Valid image types are DIB, EMF, JPG, JPEG, PICT, PNG, and WMF" />
			</cfdefaultcase>
		</cfswitch>
		
		<cfif StructKeyExists(arguments, "filepath") and StructKeyExists(arguments, "anchor")>
			<cfset inputStream = CreateObject("java", "java.io.FileInputStream").init(JavaCast("string", arguments.filepath)) />
			<cfset bytes = ioUtils.toByteArray(inputStream) />
			<cfset inputStream.close() />
		<cfelse>
			<cfset bytes = arguments.imageData />
		</cfif>

		<cfset imageIndex = getWorkbook().addPicture(bytes, JavaCast("int", imgTypeIndex)) />

		<cfset theAnchor = CreateObject("java", "org.apache.poi.hssf.usermodel.HSSFClientAnchor").init() />

		<cfif ListLen(arguments.anchor) eq 4>
			<!--- list is in format startRow, startCol, endRow, endCol --->
			<cfset theAnchor.setRow1(JavaCast("int", ListFirst(arguments.anchor) - 1)) />
			<cfset theAnchor.setCol1(JavaCast("int", ListGetAt(arguments.anchor, 2) - 1)) />
			<cfset theAnchor.setRow2(JavaCast("int", ListGetAt(arguments.anchor, 3) - 1)) />
			<cfset theAnchor.setCol2(JavaCast("int", ListLast(arguments.anchor) - 1)) />
		<cfelseif ListLen(arguments.anchor) eq 8>
			<!--- list is in format dx1, dy1, dx2, dy2, col1, row1, col2, row2 --->
			<cfset theAnchor.setDx1(JavaCast("int", ListFirst(arguments.anchor))) />
			<cfset theAnchor.setDy1(JavaCast("int", ListGetAt(arguments.anchor, 2))) />
			<cfset theAnchor.setDx2(JavaCast("int", ListGetAt(arguments.anchor, 3))) />
			<cfset theAnchor.setDy2(JavaCast("int", ListGetAt(arguments.anchor, 4))) />
			<cfset theAnchor.setRow1(JavaCast("int", ListGetAt(arguments.anchor, 5) - 1)) />
			<cfset theAnchor.setCol1(JavaCast("int", ListGetAt(arguments.anchor, 6) - 1)) />
			<cfset theAnchor.setRow2(JavaCast("int", ListGetAt(arguments.anchor, 7) - 1)) />
			<cfset theAnchor.setCol2(JavaCast("int", ListLast(arguments.anchor) - 1)) />
		<cfelse>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Anchor Argument" 
						detail="The anchor argument must be a comma-delimited list of integers with either 4 or 8 elements" />
		</cfif>
		
		<cfset picture = drawingPatriarch.createPicture(theAnchor, imageIndex) />
		
		<!--- disabling this for now--maybe let people pass in a boolean indicating 
				whether or not they want the image resized? --->
		<!--- if this is a png or jpg, resize the picture to its original size 
				(this doesn't work for formats other than jpg and png) --->
		<!--- <cfif imgTypeIndex eq getWorkbook().PICTURE_TYPE_JPEG 
				or imgTypeIndex eq getWorkbook().PICTURE_TYPE_PNG>
			<cfset picture.resize() />
		</cfif> --->
	</cffunction>
	
	<cffunction name="getInfo" access="public" output="false" returntype="struct" 
			hint="Returns a struct containing the standard properties for the workbook">
		<!--- 
			workbook properties returned in the struct are:
			* AUTHOR
			* CATEGORY
			* COMMENTS
			* CREATIONDATE
			* LASTEDITED
			* LASTAUTHOR
			* LASTSAVED
			* KEYWORDS
			* MANAGER
			* COMPANY
			* SUBJECT
			* TITLE
			* SHEETS
			* SHEETNAMES
			* SPREADSHEETTYPE
		--->
		<cfset var info = StructNew() />
		<cfset var docSummaryInfo = getWorkbook().getDocumentSummaryInformation() />
		<cfset var summaryInfo = getWorkbook().getSummaryInformation() />
		<cfset var i = 0 />
		
		<cfset info.author = summaryInfo.getAuthor() />
		<cfset info.category = docSummaryInfo.getCategory() />
		<cfset info.comments = summaryInfo.getComments() />
		<cfset info.creationdate = summaryInfo.getCreateDateTime() />
		
		<cfset info.lastedited = summaryInfo.getEditTime() />
		<cfif info.lastedited eq 0>
			<cfset info.lastedited = "" />
		<cfelse>
			<cfset info.lastedited = CreateObject("java", "java.util.Date").init(JavaCast("long", summaryInfo.getEditTime())) />
		</cfif>
		
		<cfset info.lastauthor = summaryInfo.getLastAuthor() />
		<cfset info.lastsaved = summaryInfo.getLastSaveDateTime() />
		<cfset info.keywords = summaryInfo.getKeywords() />
		<cfset info.manager = docSummaryInfo.getManager() />
		<cfset info.company = docSummaryInfo.getCompany() />
		<cfset info.subject = summaryInfo.getSubject() />
		<cfset info.title = summaryInfo.getTitle() />
		<cfset info.sheets = getWorkbook().getNumberOfSheets() />
		<cfset info.sheetnames = "" />
		
		<cfif IsNumeric(info.sheets) and info.sheets gt 0>
			<cfloop index="i" from="1" to="#info.sheets#">
				<cfset info.sheetnames = ListAppend(info.sheetnames, getWorkbook().getSheetName(JavaCast("int", i - 1))) />
			</cfloop>
		</cfif>
		
		<cfif getWorkbook().getClass().getName() eq "org.apache.poi.hssf.usermodel.HSSFWorkbook">
			<cfset info.spreadsheettype = "Excel" />
		<cfelseif getWorkbook().getClass().getName() eq "org.apache.poi.xssf.usermodel.XSSFWorkbook">
			<cfset info.spreadsheettype = "Excel (2007)" />
		<cfelse>
			<cfset info.spreadsheettype = "" />
		</cfif>
		
		<cfreturn info />
	</cffunction>
	
	<cffunction name="addInfo" access="public" output="false" returntype="void" 
			hint="Set standard properties on the workbook">
		<cfargument name="props" type="struct" required="true" 
				hint="Valid struct keys are author, category, lastauthor, comments, keywords, manager, company, subject, title" />
		
		<cfset var documentSummaryInfo = 0 />
		<cfset var summaryInfo = 0 />
		<cfset var filename = 0 />
		
		<!--- if the spreadsheet has never been written to disk, getDocumentSummaryInformation() 
				and getSummaryInformation() throw null pointer errors, so we need to do this in a 
				try/catch and write a temp file out to disk (unfortunately) --->
		
		<cftry>
			<cfset documentSummaryInfo = getWorkbook().getDocumentSummaryInformation() />
			<cfset summaryInfo = getWorkbook().getSummaryInformation() />
			<cfcatch type="any">
				<cfset filename = CreateUUID() & ".xls" />
				<cfset writeToFile(ExpandPath(filename), getWorkbook()) />
				<!--- <cfset read(ExpandPath(filename)) /> --->
				
				<cfset documentSummaryInfo = getWorkbook().getDocumentSummaryInformation() />
				<cfset summaryInfo = getWorkbook().getSummaryInformation() />
			</cfcatch>
		</cftry>
		
		
		<cfloop collection="#props#" item="prop">
			<cfswitch expression="#prop#">
				<cfcase value="author">
					<cfset summaryInfo.setAuthor(JavaCast("string", arguments.props.author)) />
				</cfcase>
				
				<cfcase value="category">
					<cfset documentSummaryInfo.setCategory(JavaCast("string", arguments.props.category)) />
				</cfcase>
				
				<cfcase value="lastauthor">
					<cfset summaryInfo.setLastAuthor(JavaCast("string", arguments.props.lastauthor)) />
				</cfcase>
				
				<cfcase value="comments">
					<cfset summaryInfo.setComments(JavaCast("string", arguments.props.comments)) />	
				</cfcase>
				
				<cfcase value="keywords">
					<cfset summaryInfo.setKeywords(JavaCast("string", arguments.props.keywords)) />
				</cfcase>
				
				<cfcase value="manager">
					<cfset documentSummaryInfo.setManager(JavaCast("string", arguments.props.manager)) />
				</cfcase>
				
				<cfcase value="company">
					<cfset documentSummaryInfo.setCompany(JavaCast("string", arguments.props.company)) />
				</cfcase>
				
				<cfcase value="subject">
					<cfset summaryInfo.setSubject(JavaCast("string", arguments.props.subject)) />
				</cfcase>
				
				<cfcase value="title">
					<cfset summaryInfo.setTitle(JavaCast("string", arguments.props.title)) />
				</cfcase>
			</cfswitch>
		</cfloop>
	</cffunction>
	
	<cffunction name="readBinary" access="public" output="false" returntype="binary" 
			hint="Writes the workbook to disk and returns a binary representation of the file">
		<!--- The workbook class has a getBytes() method that returns the sheets (only!) as 
				a byte array, but CF 9 returns a byte array of the entire file. From 
				what I can gather, since the Workbook class isn't serializable we can't 
				accomplish all of this in memory using a ByteArrayOutputStream and 
				ObjectOutputStream. So we have to write the file to disk first, then 
				do a CFFILE readbinary on it. I'm not sure if this is what CF 9 is doing 
				under the hood but the end binary result matches. --->
		<cfset var bytes = 0 />
		<cfset var filename = CreateUUID() & ".tmp" />

		<cfset writeToFile(ExpandPath(filename), getWorkbook()) />
		<!--- <cffile action="readbinary" file="#ExpandPath(filename)#" variable="bytes" /> --->
		<!--- <cffile action="delete" file="#ExpandPath(filename)#" /> --->
		
		<cfreturn bytes />
	</cffunction>
	
	<cffunction name="setFooter" access="public" output="false" returntype="void" 
			hint="Sets the footer values on the sheet">
		<cfargument name="centerFooter" type="string" required="true" />
		<cfargument name="leftFooter" type="string" required="true" />
		<cfargument name="rightFooter" type="string" required="true" />
		
		<cfif arguments.centerFooter neq "">
			<cfset getActiveSheet().getFooter().setCenter(JavaCast("string", arguments.centerFooter)) />
		</cfif>
		
		<cfif arguments.leftFooter neq "">
			<cfset getActiveSheet().getFooter().setLeft(JavaCast("string", arguments.leftFooter)) />
		</cfif>
		
		<cfif arguments.rightFooter neq "">
			<cfset getActiveSheet().getFooter().setRight(JavaCast("string", arguments.rightFooter)) />
		</cfif>
	</cffunction>
	
	<cffunction name="setHeader" access="public" output="false" returntype="void" 
			hint="Sets the header values on the sheet">
		<cfargument name="centerHeader" type="string" required="true" />
		<cfargument name="leftHeader" type="string" required="true" />
		<cfargument name="rightHeader" type="string" required="true" />
		
		<cfif arguments.centerHeader neq "">
			<cfset getActiveSheet().getHeader().setCenter(JavaCast("string", arguments.centerHeader)) />
		</cfif>
		
		<cfif arguments.leftHeader neq "">
			<cfset getActiveSheet().getHeader().setLeft(JavaCast("string", arguments.leftHeader)) />
		</cfif>
		
		<cfif arguments.rightHeader neq "">
			<cfset getActiveSheet().getHeader().setRight(JavaCast("string", arguments.rightHeader)) />
		</cfif>
	</cffunction>
	
	<!--- TODO: implement an addPageNumbers() function to allow for addition of page numbers 
				in header or footer (tons more stuff like this that could easily be added) --->
	
	<!--- row functions --->
	<!--- TODO: SpreadsheetAddRow in CF 9 is not consitent with SpreadsheetAddColumn because there 
				is no option to pass a boolean indicating whether or not the existing row should 
				be overwritten. Instead, it will only insert the new row and shift existing rows 
				down (i.e. increment their row number) by one, which would necessitate a deletion 
				of the existing row if an overwrite is desired. Leaving behavior as is for now but 
				might be worth changing/enhancing later. --->
	<cffunction name="addRow" access="public" output="false" returntype="void" 
			hint="Adds a new row and inserts the data provided in the new row.">
		<cfargument name="data" type="string" required="true" hint="Delimited list of data" />
		<cfargument name="delimiter" type="string" required="true" />
		<cfargument name="startRow" type="numeric" required="false" />
		<cfargument name="startColumn" type="numeric" required="false" />
		
		<cfset var row = 0 />
		<cfset var cell = 0 />
		<cfset var cellNum = 0 />
		<cfset var cellValue = 0 />
		
		<cfif StructKeyExists(arguments, "startRow") and arguments.startRow lte 0>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Row Value" 
						detail="The value for row must be greater than or equal to 1." />
		</cfif>
		
		<cfif StructKeyExists(arguments, "startColumn") and arguments.startColumn lte 0>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Column Value" 
						detail="The value for column must be greater than or equal to 1." />
		</cfif>
		
		<cfif not StructKeyExists(arguments, "startRow")>
			<cfset row = createRow() />
		<cfelse>
			<cfset row = createRow(arguments.startRow - 1) />
		</cfif>
		
		<cfif StructKeyExists(arguments, "startColumn")>
			<cfset cellNum = arguments.startColumn - 1 />
		</cfif>
		
		<!--- TODO: treating all data as strings; need to support data types? --->
		<cfloop list="#arguments.data#" index="cellValue" delimiters="#arguments.delimiter#">
			<cfset cell = createCell(row, cellNum) />
			<cfset cell.setCellValue(JavaCast("string", cellValue)) />
			<cfset cellNum = cellNum + 1 />
		</cfloop>
	</cffunction>
	
	<cffunction name="addRows" access="public" output="false" returntype="void" 
			hint="Adds rows to a sheet from a query object">
		<cfargument name="data" type="query" required="true" />
		<cfargument name="row" type="numeric" required="false" />
		
		<cfset var column = 0 />
		<cfset var theRow = 0 />
		<cfset var rowNum = 0 />
		<cfset var cell = 0 />
		<cfset var cellNum = 0 />
		
		<cfif StructKeyExists(arguments, "row")>
			<cfif arguments.row lte 0>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Invalid Row Value" 
							detail="The value for row must be greater than or equal to 1." />
			<cfelse>
				<cfset rowNum = arguments.row - 1 />
			</cfif>
		<cfelse>
			<cfset rowNum = getActiveSheet().getPhysicalNumberOfRows() />
		</cfif>
		
		<cfloop query="arguments.data">
			<!--- can't just call addRow() here since that function expects a comma-delimited 
					list of data (probably not the greatest limitation ...) and the query 
					data may have commas in it, so this is a bit redundant with the addRow() 
					function --->
			<cfset theRow = createRow(rowNum) />
			
			<!--- odd that you can't specify a start column when adding multiple rows, 
					but that's the way it is in cf 9 so leaving this out for now --->
			<!--- <cfif StructKeyExists(arguments, "startColumn")>
				<cfset cellNum = arguments.startColumn - 1 />
			</cfif> --->


			<!--- TODO: treating all data as strings; need to support data types? --->
			<cfset cellNum = 0 />
			<cfloop list="#arguments.data.ColumnList#" index="column">
				<cfset cell = createCell(theRow, cellNum) />
				<cfset cell.setCellValue(JavaCast("string", arguments.data[column][arguments.data.CurrentRow])) />
				<cfset cellNum = cellNum + 1 />
			</cfloop>

			<cfset rowNum = rowNum + 1 />
		</cfloop>
	</cffunction>
	
	<cffunction name="deleteRow" access="public" output="false" returntype="void" 
			hint="Deletes the data from a row. Does not physically delete the row.">
		<cfargument name="rowNum" type="numeric" required="true" />
		
		<cfif arguments.rowNum lte 0>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Row Value" 
						detail="The value for row must be greater than or equal to 1." />
		</cfif>
		
		<cfset getActiveSheet().removeRow(getActiveSheet().getRow(JavaCast("int", arguments.rowNum - 1))) />
	</cffunction>
	
	<cffunction name="deleteRows" access="public" output="false" returntype="void" 
			hint="Deletes a range of rows">
		<cfargument name="range" type="string" required="true" />
		
		<cfset var rangeValue = 0 />
		<cfset var rangeTest = "^[0-9]{1,}(-[0-9]{1,})?$">
		<cfset var i = 0 />
		
		<!--- Range is a comma-delimited list of ranges, and each value can be either 
				a single number or a range of numbers with a hyphen. --->
		<cfloop list="#arguments.range#" index="rangeValue">
			<cfif REFind(rangeTest, rangeValue) eq 0>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Invalid Range Value" 
							detail="The range value #rangeValue# is not valid." />
			<cfelse>
				<cfif ListLen(rangeValue, "-") eq 2>
					<cfloop index="i" from="#ListGetAt(rangeValue, 1, '-')#" to="#ListGetAt(rangeValue, 2, '-')#">
						<cfset deleteRow(i) />
					</cfloop>
				<cfelse>
					<cfset deleteRow(rangeValue) />
				</cfif>
			</cfif>
		</cfloop>
	</cffunction>
	
	<cffunction name="shiftRows" access="public" output="false" returntype="void" 
			hint="Shifts rows up (negative integer) or down (positive integer)">
		<cfargument name="startRow" type="numeric" required="true" />
		<cfargument name="endRow" type="numeric" required="false" />
		<cfargument name="offset" type="numeric" required="false" default="1" />
		
		<cfif not StructKeyExists(arguments, "endRow")>
			<cfset arguments.endRow = arguments.startRow />
		</cfif>
		
		<cfset getActiveSheet().shiftRows(JavaCast("int", arguments.startRow - 1), 
											JavaCast("int", arguments.endRow - 1), 
											JavaCast("int", arguments.offset)) />
	</cffunction>
	
	<!--- TODO: for some reason setRowStyle() formats the empty cells but leaves the populated cells 
				alone, which is exactly opposite of what we want, so looping over each populated 
				cell and setting the cell format individually instead. Better way to do this? --->
	<cffunction name="formatRow" access="public" output="false" returntype="void" 
			hint="Sets various formatting values on a row">
		<cfargument name="format" type="struct" required="true" />
		<cfargument name="rowNum" type="numeric" required="true" />
		
		<cfset var cellIterator = getActiveSheet().getRow(arguments.rowNum - 1).cellIterator() />
		
		<cfloop condition="#cellIterator.hasNext()#">
			<cfset formatCell(arguments.format, arguments.rowNum, cellIterator.next().getColumnIndex() + 1) />
		</cfloop>
	</cffunction>
	
	<cffunction name="formatRows" access="public" output="false" returntype="void" 
			hint="Sets various formatting values on multiple rows">
		<cfargument name="format" type="struct" required="true" />
		<cfargument name="range" type="string" required="true" />

		<cfset var rangeValue = 0 />
		<cfset var rangeTest = "^[0-9]{1,}(-[0-9]{1,})?$">
		<cfset var i = 0 />
		
		<!--- Range is a comma-delimited list of ranges, and each value can be either 
				a single number or a range of numbers with a hyphen. --->
		<cfloop list="#arguments.range#" index="rangeValue">
			<cfif REFind(rangeTest, rangeValue) eq 0>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Invalid Range Value" 
							detail="The range value #rangeValue# is not valid." />
			<cfelse>
				<cfif ListLen(rangeValue, "-") eq 2>
					<cfloop index="i" from="#ListGetAt(rangeValue, 1, '-')#" to="#ListGetAt(rangeValue, 2, '-')#">
						<cfset formatRow(arguments.format, i) />
					</cfloop>
				<cfelse>
					<cfset formatRow(arguments.format, rangeValue) />
				</cfif>
			</cfif>
		</cfloop>
	</cffunction>
	
	<cffunction name="setRowHeight" access="public" output="false" returntype="void" 
			hint="Sets the height of a row in points">
		<cfargument name="row" type="numeric" required="true" />
		<cfargument name="height" type="numeric" required="true" />
		
		<cfset getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).setHeightInPoints(JavaCast("int", arguments.height)) />
	</cffunction>
	
	<!--- column functions --->
	<cffunction name="addColumn" access="public" output="false" returntype="void" 
			hint="Adds a column and inserts the data provided into the new column.">
		<cfargument name="data" type="string" required="true" />
		<cfargument name="delimiter" type="string" required="true" />
		<cfargument name="column" type="numeric" required="false" />
		<cfargument name="startRow" type="numeric" required="false" />
		<cfargument name="insert" type="boolean" required="false" default="true" 
			hint="If false, will overwrite data in an existing column if one exists" />
		
		<cfset var row = 0 />
		<cfset var cell = 0 />
		<cfset var rowNum = 0 />
		<cfset var cellNum = 0 />
		<cfset var lastCellNum = 0 />
		<cfset var i = 0 />
		<cfset var tempCell = 0 />
		<cfset var temp = 0 />
		<cfset var cellValue = 0 />
		
		<cfif StructKeyExists(arguments, "startRow") and arguments.startRow lte 0>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Start Row Value" 
						detail="The value for start row must be greater than or equal to 1." />
		</cfif>
		
		<cfif StructKeyExists(arguments, "column") and arguments.column lte 0>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Column Value" 
						detail="The value for column must be greater than or equal to 1." />
		</cfif>
		
		<cfif StructKeyExists(arguments, "startRow")>
			<cfset rowNum = arguments.startRow - 1 />
		</cfif>
		
		<cfif StructKeyExists(arguments, "column")>
			<cfset cellNum = arguments.column - 1 />
		</cfif>


		<cfloop list="#arguments.data#" index="cellValue" delimiters="#arguments.delimiter#">
			<!--- if rowNum is greater than the last row of the sheet, need 
					to create a new row --->
				
			<!---
				REMOVED as second part of logic is not simple value 
				<cfif rowNum GT getActiveSheet().getLastRowNum() OR getActiveSheet().getRow(rowNum) EQ "">
			--->
			<cfif rowNum GT getActiveSheet().getLastRowNum() OR isNull(getActiveSheet().getRow(rowNum).getCell(1))>

				<cfset row = createRow(rowNum) />

			<cfelse>
				<cfset row = getActiveSheet().getRow(rowNum) />
			</cfif>
			
			<!--- POI doesn't have any 'shift column' functionality akin to shiftRows() 
					so inserts get interesting ... --->
			<cfif arguments.insert and cellNum lt row.getLastCellNum()>
				<!--- need to get the last populated column number in the row, figure out which 
						cells are impacted, and shift the impacted cells to the right to make 
						room for the new data --->
				<cfset lastCellNum = row.getLastCellNum() + 1 />

				<cfloop index="i" from="#lastCellNum#" to="#cellNum + 1#" step="-1">
					<cfset tempCell = row.getCell(JavaCast("int", i - 1)) />
					
					<!--- getLastCellNum() apparently returns the max cell number in ANY row (?), 
							so we need to check if this is null --->
					<cfif isDefined( "tempCell" ) AND tempCell.toString() neq "">
						<cfset temp = tempCell.getStringCellValue() />
						<cfset cell = createCell(row, i) />
						<cfset cell.setCellValue(JavaCast("string", temp)) />
					</cfif>
				</cfloop>
			</cfif>

			<cfset cell = createCell(row, cellNum) />
			
			<cfset cell.setCellValue(JavaCast("string", cellValue)) />
			
			<cfset rowNum = rowNum + 1 />
		</cfloop>
	</cffunction>
	
	<cffunction name="deleteColumn" access="public" output="false" returntype="void" 
			hint="Deletes the data from a column. Does not physically remove the column.">
		<cfargument name="columnNum" type="numeric" required="true" />
		
		<cfset var rowIterator = getActiveSheet().rowIterator() />
		<cfset var row = 0 />
		<cfset var cell = 0 />
		
		<cfif arguments.columnNum lte 0>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Column Value" 
						detail="The value for column must be greater than or equal to 1." />
		</cfif>
		
		<!--- POI doesn't have remove column functionality, so iterate over all the rows 
				and remove the column indicated --->
		<cfloop condition="#rowIterator.hasNext()#">
			<cfset row = rowIterator.next() />
			<cfset cell = row.getCell(JavaCast("int", arguments.columnNum - 1)) />
			
			<cfif cell neq "">
				<cfset row.removeCell(cell) />
			</cfif>
		</cfloop>
	</cffunction>
	
	<cffunction name="deleteColumns" access="public" output="false" returntype="void" 
			hint="Deletes a range of columns">
		<cfargument name="range" type="string" required="true" />
		
		<cfset var rangeValue = 0 />
		<cfset var rangeTest = "^[0-9]{1,}(-[0-9]{1,})?$">
		<cfset var i = 0 />
		
		<!--- Range is a comma-delimited list of ranges, and each value can be either 
				a single number or a range of numbers with a hyphen. --->
		<cfloop list="#arguments.range#" index="rangeValue">
			<cfif REFind(rangeTest, rangeValue) eq 0>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Invalid Range Value" 
							detail="The range value #rangeValue# is not valid." />
			<cfelse>
				<cfif ListLen(rangeValue, "-") eq 2>
					<cfloop index="i" from="#ListGetAt(rangeValue, 1, '-')#" to="#ListGetAt(rangeValue, 2, '-')#">
						<cfset deleteColumn(i) />
					</cfloop>
				<cfelse>
					<cfset deleteColumn(rangeValue) />
				</cfif>
			</cfif>
		</cfloop>
	</cffunction>
	
	<cffunction name="shiftColumns" access="public" output="false" returntype="void" 
			hint="Shifts columns left (negative integer) or right (positive integer)">
		<cfargument name="startColumn" type="numeric" required="true" />
		<cfargument name="endColumn" type="numeric" required="false" />
		<cfargument name="offset" type="numeric" required="false" default="1" />
				
		<cfset var rowIterator = getActiveSheet().rowIterator() />
		<cfset var row = 0 />
		<cfset var tempCell = 0 />
		<cfset var cell = 0 />
		<cfset var i = 0 />
		<cfset var numColsShifted = 0 />
		<cfset var numColsToDelete = 0 />
		
		<cfif arguments.startColumn lte 0>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Start Column Value" 
						detail="The value for start column must be greater than or equal to 1." />
		</cfif>

		<cfif StructKeyExists(arguments, "endColumn") and 
				(arguments.endColumn lte 0 or arguments.endColumn lt arguments.startColumn)>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid End Column Value" 
						detail="The value of end column must be greater than or equal to the value of start column." />
		</cfif>

		<cfset arguments.startColumn = arguments.startColumn - 1 />
		
		<cfif not StructKeyExists(arguments, "endColumn")>
			<cfset arguments.endColumn = arguments.startColumn />
		<cfelse>
			<cfset arguments.endColumn = arguments.endColumn - 1 />
		</cfif>
		
		<cfloop condition="#rowIterator.hasNext()#">
			<cfset row = rowIterator.next() />
			
			<cfif arguments.offset gt 0>
				<cfloop index="i" from="#arguments.endColumn#" to="#arguments.startColumn#" step="-1">
					<cfset tempCell = row.getCell(JavaCast("int", i)) />
					<cfset cell = createCell(row, i + arguments.offset) />
					
					<cfif tempCell neq "">
						<cfset cell.setCellValue(JavaCast("string", tempCell.getStringCellValue())) />
					</cfif>
				</cfloop>
			<cfelse>
				<cfloop index="i" from="#arguments.startColumn#" to="#arguments.endColumn#" step="1">
					<cfset tempCell = row.getCell(JavaCast("int", i)) />
					<cfset cell = createCell(row, i + arguments.offset) />
					
					<cfif tempCell neq "">
						<cfset cell.setCellValue(JavaCast("string", tempCell.getStringCellValue())) />
					</cfif>
				</cfloop>
			</cfif>
		</cfloop>

		<!--- clean up any columns that need to be deleted after the shift --->
		<cfset numColsShifted = arguments.endColumn - arguments.startColumn + 1 />
		
		<cfset numColsToDelete = Abs(arguments.offset) />
		
		<cfif numColsToDelete gt numColsShifted>
			<cfset numColsToDelete = numColsShifted />
		</cfif>
		
		<cfif arguments.offset gt 0>
			<cfloop index="i" from="#arguments.startColumn#" to="#arguments.startColumn + numColsToDelete - 1#">
				<cfset deleteColumn(i + 1) />
			</cfloop>
		<cfelse>
			<cfloop index="i" from="#arguments.endColumn#" to="#arguments.endColumn - numColsToDelete + 1#" step="-1">
				<cfset deleteColumn(i + 1) />
			</cfloop>
		</cfif>
	</cffunction>
	
	<cffunction name="formatCell" access="public" output="false" returntype="void" 
			hint="Sets various formatting values on a single cell">
		<cfargument name="format" type="struct" required="true" />
		<cfargument name="row" type="numeric" required="true" />
		<cfargument name="column" type="numeric" required="true" />
		
		<cfset var cell = getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)) />
		
		<cfif cell neq "">
			<cfset cell.setCellStyle(buildCellStyle(arguments.format)) />
		</cfif>
	</cffunction>
	
	<cffunction name="formatColumn" access="public" output="false" returntype="void" 
			hint="Sets various formatting values on a column">
		<cfargument name="format" type="struct" required="true" />
		<cfargument name="column" type="numeric" required="true" />
		
		<cfset var rowIterator = getActiveSheet().rowIterator() />
		
		<cfif arguments.column lte 0>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Column Value" 
						detail="The column value must be greater than 0." />
		</cfif>
		
		<cfloop condition="#rowIterator.hasNext()#">
			<cfset formatCell(arguments.format, rowIterator.next().getRowNum() + 1, arguments.column) />
		</cfloop>
	</cffunction>
	
	<cffunction name="formatColumns" access="public" output="false" returntype="void" 
			hint="Sets various formatting values on multiple columns">
		<cfargument name="format" type="struct" required="true" />
		<cfargument name="range" type="string" required="true" />
		
		<cfset var rangeValue = 0 />
		<cfset var rangeTest = "^[0-9]{1,}(-[0-9]{1,})?$">
		<cfset var i = 0 />
		
		<cfloop list="#arguments.range#" index="rangeValue">
			<cfif REFind(rangeTest, rangeValue) eq 0>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Invalid Range Value" 
							detail="The range value #rangeValue# is not valid." />
			<cfelse>
				<cfif ListLen(rangeValue, "-") eq 2>
					<cfloop index="i" from="#ListGetAt(rangeValue, 1, '-')#" to="#ListGetAt(rangeValue, 2, '-')#">
						<cfset formatColumn(arguments.format, i) />
					</cfloop>
				<cfelse>
					<cfset formatColumn(arguments.format, rangeValue) />
				</cfif>
			</cfif>
		</cfloop>
	</cffunction>
	
	<cffunction name="getCellComment" access="public" output="false" returntype="any" 
			hint="Returns a struct containing comment info (author, column, row, and comment) for a specific cell, or an array of structs containing the comments for the entire sheet">
		<cfargument name="row" type="numeric" required="false" />
		<cfargument name="column" type="numeric" required="false" />
		
		<cfset var comment = 0 />
		<cfset var theComment = 0 />
		<cfset var comments = StructNew() />
		<cfset var rowIterator = 0 />
		<cfset var cellIterator = 0 />
		
		<cfif (StructKeyExists(arguments, "row") and not StructKeyExists(arguments, "column")) 
				or (StructKeyExists(arguments, "column") and not StructKeyExists(arguments, "row"))>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Invalid Argument Combination" 
						detail="If row or column is passed to getCellComment, both row and column must be provided." />
		</cfif>
		
		<cfif StructKeyExists(arguments, "row")>
			<cfset comment = getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)).getCellComment() />
			
			<cfif comment neq "">
				<cfset comments.author = comment.getAuthor() />
				<cfset comments.column = arguments.column />
				<cfset comments.comment = comment.getString().getString() />
				<cfset comments.row = arguments.row />
			</cfif>
		<cfelse>
			<!--- row and column weren't provided so loop over the whole shooting match and get all the comments --->
			<cfset comments = ArrayNew(1) />
			<cfset rowIterator = getActiveSheet().rowIterator() />
			
			<cfloop condition="#rowIterator.hasNext()#">
				<cfset cellIterator = rowIterator.next().cellIterator() />
				
				<cfloop condition="#cellIterator.hasNext()#">
					<cfset comment = cellIterator.next().getCellComment() />
					
					<cfif comment neq "">
						<cfset theComment = StructNew() />
						<cfset theComment.author = comment.getAuthor() />
						<cfset theComment.column = comment.getColumn() + 1 />
						<cfset theComment.comment = comment.getString().getString() />
						<cfset theComment.row = comment.getRow() + 1 />
						
						<cfset ArrayAppend(comments, theComment) />
					</cfif>
				</cfloop>
			</cfloop>
		</cfif>
		
		<cfreturn comments />
	</cffunction>
	
	<cffunction name="setCellComment" access="public" output="false" returntype="void" 
			hint="Sets a cell comment">
		<cfargument name="comment" type="struct" required="true" />
		<cfargument name="row" type="numeric" required="true" />
		<cfargument name="column" type="numeric" required="true" />
		
		<!--- 
			The comment struct may contain the following keys: 
			* anchor
			* author
			* bold
			* color
			* comment
			* fillcolor
			* font
			* horizontalalignment
			* italic
			* linestyle
			* linestylecolor
			* size
			* strikeout
			* underline
			* verticalalignment
			* visible
		--->
		
		<!--- <cfset var creationHelper = getWorkbook().getCreationHelper() /> --->
		<cfset var drawingPatriarch = getActiveSheet().createDrawingPatriarch() />
		<cfset var clientAnchor = 0 />
		<cfset var commentObj = 0 />
		<cfset var commentString = CreateObject("java", "org.apache.poi.hssf.usermodel.HSSFRichTextString").init(JavaCast("string", arguments.comment.comment)) />
		<cfset var font = 0 />
		<cfset var javaColorRGB = 0 />
		
		<!--- make sure the cell exists before proceeding --->
		<cfif getActiveSheet().getRow(JavaCast("int", arguments.row - 1)) eq "" 
				or getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)) eq "">
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Cell Does Not Exist" 
						detail="The cell on which a comment is attempting to be set does not exist." />
		<cfelse>
			<cfif StructKeyExists(arguments.comment, "anchor")>
				<cfset clientAnchor = CreateObject("java", "org.apache.poi.hssf.usermodel.HSSFClientAnchor").init(JavaCast("int", 0), 
																													JavaCast("int", 0), 
																													JavaCast("int", 0), 
																													JavaCast("int", 0), 
																													JavaCast("int", ListGetAt(arguments.comment.anchor, 1)), 
																													JavaCast("int", ListGetAt(arguments.comment.anchor, 2)), 
																													JavaCast("int", ListGetAt(arguments.comment.anchor, 3)), 
																													JavaCast("int", ListGetAt(arguments.comment.anchor, 4))) />
			<cfelse>
				<!--- if no anchor is provided, just use + 2 --->
				<cfset clientAnchor = CreateObject("java", "org.apache.poi.hssf.usermodel.HSSFClientAnchor").init(JavaCast("int", 0), 
																													JavaCast("int", 0), 
																													JavaCast("int", 0), 
																													JavaCast("int", 0), 
																													JavaCast("int", arguments.column), 
																													JavaCast("int", arguments.row), 
																													JavaCast("int", arguments.column + 2), 
																													JavaCast("int", arguments.row + 2)) />
			</cfif>
			
			<cfset commentObj = drawingPatriarch.createComment(clientAnchor) />
			
			<cfif StructKeyExists(arguments.comment, "author")>
				<cfset commentObj.setAuthor(JavaCast("string", arguments.comment.author)) />
			</cfif>
			
			<!--- If we're going to do anything font related, need to create a font. 
					Didn't really want to create it above since it might not be needed. --->
			<cfif StructKeyExists(arguments.comment, "bold") 
					or StructKeyExists(arguments.comment, "color") 
					or StructKeyExists(arguments.comment, "font")
					or StructKeyExists(arguments.comment, "italic")
					or StructKeyExists(arguments.comment, "size") 
					or StructKeyExists(arguments.comment, "strikeout") 
					or StructKeyExists(arguments.comment, "underline")>
				<cfset font = getWorkbook().createFont() />
				
				<cfif StructKeyExists(arguments.comment, "bold")>
					<cfif arguments.comment.bold>
						<cfset font.setBoldweight(font.BOLDWEIGHT_BOLD) />
					<cfelse>
						<cfset font.setBoldweight(font.BOLDWEIGHT_NORMAL) />
					</cfif>
				</cfif>
				
				<cfif StructKeyExists(arguments.comment, "color")>
					<cfset font.setColor(JavaCast("int", getColorIndex(arguments.comment.color))) />
				</cfif>
				
				<cfif StructKeyExists(arguments.comment, "font")>
					<cfset font.setFontName(JavaCast("string", arguments.comment.font)) />
				</cfif>
				
				<cfif StructKeyExists(arguments.comment, "italic")>
					<cfset font.setItalic(JavaCast("boolean", arguments.comment.italic)) />
				</cfif>
				
				<cfif StructKeyExists(arguments.comment, "size")>
					<cfset font.setFontHeightInPoints(JavaCast("int", arguments.comment.size)) />
				</cfif>
				
				<cfif StructKeyExists(arguments.comment, "strikeout")>
					<cfset font.setStrikeout(JavaCast("boolean", arguments.comment.strikeout)) />
				</cfif>
				
				<cfif StructKeyExists(arguments.comment, "underline")>
					<cfset font.setUnderline(JavaCast("boolean", arguments.comment.underline)) />
				</cfif>
				
				<cfset commentString.applyFont(font) />
			</cfif>
			
			<cfif StructKeyExists(arguments.comment, "fillcolor")>
				<cfset javaColorRGB = getJavaColorRGB(arguments.comment.fillcolor) />
				<cfset commentObj.setFillColor(JavaCast("int", javaColorRGB.red), 
												JavaCast("int", javaColorRGB.green), 
												JavaCast("int", javaColorRGB.blue)) />
			</cfif>
			
			<!---- Horizontal alignment can be left, center, right, justify, or distributed. 
					Note that the constants on the Java class are slightly different in some cases:
					'center' = CENTERED
					'justify' = JUSTIFIED --->
			<cfif StructKeyExists(arguments.comment, "horizontalalignment")>
				<cfif UCase(arguments.comment.horizontalalignment) eq "CENTER">
					<cfset arguments.comment.horizontalalignment = "CENTERED" />
				</cfif>
				
				<cfif UCase(arguments.comment.horizontalalignment) eq "JUSTIFY">
					<cfset arguments.comment.horizontalalignment = "JUSTIFIED" />
				</cfif>
				
				<cfset commentObj.setHorizontalAlignment(JavaCast("int", Evaluate("commentObj.HORIZONTAL_ALIGNMENT_#UCase(arguments.comment.horizontalalignment)#"))) />
			</cfif>
			
			<!--- Valid values for linestyle are:
					* solid
					* dashsys
					* dashdotsys
					* dashdotdotsys
					* dotgel
					* dashgel
					* longdashgel
					* dashdotgel
					* longdashdotgel
					* longdashdotdotgel
			--->
			<cfif StructKeyExists(arguments.comment, "linestyle")>
				<cfset commentObj.setLineStyle(JavaCast("int", Evaluate("commentObj.LINESTYLE_#UCase(arguments.comment.linestyle)#"))) />
			</cfif>
			
			<!--- TODO: This doesn't seem to be working (no error, but doesn't do anything).
						Saw reference on the POI mailing list to this not working but it was
						from over a year ago; maybe it's just still broken.  --->
			<cfif StructKeyExists(arguments.comment, "linestylecolor")>
				<cfset javaColorRGB = getJavaColorRGB(arguments.comment.fillcolor) />
				<cfset commentObj.setLineStyleColor(JavaCast("int", javaColorRGB.red), 
													JavaCast("int", javaColorRGB.green), 
													JavaCast("int", javaColorRGB.blue)) />
			</cfif>
			
			<!--- Vertical alignment can be top, center, bottom, justify, and distributed. 
					Note that center and justify are DIFFERENT than the constants for 
					horizontal alignment, which are CENTERED and JUSTIFIED. --->
			<cfif StructKeyExists(arguments.comment, "verticalalignment")>
				<cfset commentObj.setVerticalAlignment(JavaCast("int", Evaluate("commentObj.VERTICAL_ALIGNMENT_#UCase(arguments.comment.verticalalignment)#"))) />
			</cfif>
			
			<cfif StructKeyExists(arguments.comment, "visible")>
				<cfset commentObj.setVisible(JavaCast("boolean", arguments.comment.visible)) />
			</cfif>
			
			<cfset commentObj.setString(commentString) />
	
			<cfset getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)).setCellComment(commentObj) />
		</cfif>
	</cffunction>
	
	<cffunction name="getCellFormula" access="public" output="false" returntype="any" 
			hint="Returns the formula for a cell or for the entire spreadsheet">
		<cfargument name="row" type="numeric" required="false" />
		<cfargument name="column" type="numeric" required="false" />
		
		<cfset var formulaStruct = 0 />
		<cfset var formulas = 0 />
		<cfset var rowIterator = 0 />
		<cfset var cellIterator = 0 />
		<cfset var cell = 0 />
		
		<!--- if row and column are passed in, return the formula for a single cell as a string --->
		<cfif StructKeyExists(arguments, "row") and StructKeyExists(arguments, "column")>
			<cfreturn getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)).getCellFormula() />
		<cfelse>
			<!--- no row and column provided so return an array of structs containing formulas 
					for the entire sheet --->
			<cfset rowIterator = getActiveSheet().rowIterator() />
			<cfset formulas = ArrayNew(1) />
			
			<cfloop condition="#rowIterator.hasNext()#">
				<cfset cellIterator = rowIterator.next().cellIterator() />
				
				<cfloop condition="#cellIterator.hasNext()#">
					<cfset cell = cellIterator.next() />
					
					<cfset formulaStruct = StructNew() />
					<cfset formulaStruct.row = cell.getRowIndex() + 1 />
					<cfset formulaStruct.column = cell.getColumnIndex() + 1 />
					
					<cftry>
						<cfset formulaStruct.formula = cell.getCellFormula() />
						<cfcatch type="any">
							<cfset formulaStruct.formula = "" />
						</cfcatch>
					</cftry>
					
					<cfif formulaStruct.formula neq "">
						<cfset ArrayAppend(formulas, formulaStruct) />
					</cfif>
				</cfloop>
			</cfloop>
			
			<cfreturn formulas />
		</cfif>
	</cffunction>
	
	<cffunction name="setCellFormula" access="public" output="false" returntype="void" 
			hint="Sets the formula for a cell">
		<cfargument name="formula" type="string" required="true" />
		<cfargument name="row" type="numeric" required="true" />
		<cfargument name="column" type="numeric" required="true" />
		
		<cfset getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)).setCellFormula(JavaCast("string", arguments.formula)) />
	</cffunction>
	
	<cffunction name="getCellValue" access="public" output="false" returntype="string" 
			hint="Returns the value of a single cell">
		<cfargument name="row" type="numeric" required="true" />
		<cfargument name="column" type="numeric" required="true" />
		
		<cfset var returnVal = "" />
		
		<!--- TODO: need to worry about additional cell types? --->
		<cfswitch expression="#getActiveSheet().getRow(JavaCast('int', arguments.row - 1)).getCell(JavaCast('int', arguments.column - 1)).getCellType()#">
			<!--- numeric or formula --->
			<cfcase value="0,2" delimiters=",">
				<cfset returnVal = getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)).getNumericCellValue() />
			</cfcase>
			
			<!--- string --->
			<cfcase value="1">
				<cfset returnVal = getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)).getStringCellValue() />
			</cfcase>
		</cfswitch>
		
		<cfreturn returnVal />
	</cffunction>
	
	<cffunction name="setCellValue" access="public" output="false" returntype="void" 
			hint="Sets the value of a single cell">
		<cfargument name="cellValue" type="string" required="true" />
		<cfargument name="row" type="numeric" required="true" />
		<cfargument name="column" type="numeric" required="true" />
		
		<!--- TODO: need to worry about data types? doing everything as a string for now --->
		<cfset getActiveSheet().getRow(JavaCast("int", arguments.row - 1)).getCell(JavaCast("int", arguments.column - 1)).setCellValue(JavaCast("string", arguments.cellValue)) />
	</cffunction>
	
	<cffunction name="setColumnWidth" access="public" output="false" returntype="void" 
			hint="Sets the width of a column">
		<cfargument name="column" type="numeric" required="true" />
		<cfargument name="width" type="numeric" required="true" />
		
		<cfset getActiveSheet().setColumnWidth(JavaCast("int", arguments.column - 1), JavaCast("int", arguments.width * 256)) />
	</cffunction>
	
	<cffunction name="mergeCells" access="public" output="false" returntype="void" 
			hint="Merges two or more cells">
		<cfargument name="startRow" type="numeric" required="true" />
		<cfargument name="startColumn" type="numeric" required="true" />
		<cfargument name="endRow" type="numeric" required="true" />
		<cfargument name="endColumn" type="numeric" required="true" />
		
		<cfset var cellRangeAddress = CreateObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(JavaCast("int", arguments.startRow - 1), 
																											JavaCast("int", arguments.endRow - 1), 
																											JavaCast("int", arguments.startColumn - 1), 
																											JavaCast("int", arguments.endColumn - 1)) />
		
		<cfset getActiveSheet().addMergedRegion(cellRangeAddress) />
	</cffunction>

	<!--- LOWER-LEVEL SPREADSHEET MANIPULATION FUNCTIONS --->
	<cffunction name="createRow" access="public" output="false" returntype="any" 
			hint="Creates a new row in the sheet and returns the row">
		<cfargument name="rowNum" type="numeric" required="false" />
		
		<!--- if rowNum is provided and is lte the last row number, 
				need to shift existing rows down by 1 --->
		<cfif not StructKeyExists(arguments, "rowNum")>
			<cfset arguments.rowNum = getActiveSheet().getLastRowNum() />
		<!--- TODO: need to revisit this; this isn't quite the behavior necessary, but 
					leaving it out for now is fine
		 <cfelse>
			<cfif arguments.rowNum lte getActiveSheet().getLastRowNum()>
				<cfset shiftRows(arguments.rowNum, getActiveSheet().getLastRowNum()) />
			</cfif> --->
		</cfif>
		
		<cfreturn getActiveSheet().createRow(JavaCast("int", arguments.rowNum)) />
	</cffunction>
	
	<!--- TODO: POI supports setting the cell type when the cell is created. Need to worry about this? --->
	<cffunction name="createCell" access="public" output="false" returntype="any" 
		hint="Creates a new cell in a row and returns the cell">
		<cfargument name="row" type="any" required="true" />
		<cfargument name="cellNum" type="numeric" required="false" />
		
		<cfif not StructKeyExists(arguments, "cellNum")>
			<cfset arguments.cellNum = arguments.row.getLastCellNum() />
		</cfif>
		
		<cfreturn arguments.row.createCell(JavaCast("int", arguments.cellNum)) />
	</cffunction>
	
	<!--- GET/SET FUNCTIONS FOR INTERNAL USE AND USING THIS CFC WITHOUT THE CORRESPONDING CUSTOM TAG --->
	<cffunction name="setWorkbook" access="public" output="false" returntype="void">
		<cfargument name="workbook" type="any" required="true" />
		<cfset variables.workbook = arguments.workbook />
	</cffunction>
	
	<cffunction name="getWorkbook" access="public" output="false" returntype="any">
		<cfreturn variables.workbook />
	</cffunction>
	
	<cffunction name="setActiveSheet" access="public" output="false" returntype="void" 
			hint="Sets the active sheet within the workbook, either by name or by index">
		<cfargument name="sheetName" type="string" required="false" />
		<cfargument name="sheetIndex" type="numeric" required="false" />
		
		<cfif not StructKeyExists(arguments, "sheetName") and not StructKeyExists(arguments, "sheetIndex")>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="Either sheetName or sheetIndex Must Be Provided" 
						detail="Either sheetName or sheetIndex must be provided as an argument" />
		</cfif>
		
		<cfif StructKeyExists(arguments, "sheetName")>
			<cfset getWorkbook().setActiveSheet(JavaCast("int", getWorkbook().getSheetIndex(JavaCast("string", arguments.sheetName)))) />
		<cfelse>
			<cfset getWorkbook().setActiveSheet(JavaCast("int", arguments.sheetIndex - 1)) />
		</cfif>
	</cffunction>

	<cffunction name="getActiveSheet" access="public" output="false" returntype="any">
		<cfreturn getWorkbook().getSheetAt(JavaCast("int", getWorkbook().getActiveSheetIndex())) />
	</cffunction>
	
	<!--- PRIVATE FUNCTIONS --->
	<cffunction name="readFromFile" access="private" output="false" returntype="any" 
			hint="Reads a workbook file from disk and returns a POI HSSFWorkbook object.">
		<!--- TODO: need to make sure this handles XSSF format; works with HSSF for now --->
		<cfargument name="src" type="string" required="true" hint="The full file path to the spreadsheet" />
		<cfargument name="sheet" type="numeric" required="false" hint="Used to set the active sheet" />
		<cfargument name="sheetname" type="string" required="false" hint="Used to set the active sheet" />
		
		<cfset var inputStream = CreateObject("java", "java.io.FileInputStream").init(arguments.src) />
		<!--- <cfset var workbookFactory = CreateObject("java", "org.apache.poi.ss.usermodel.WorkbookFactory").init() /> --->
		<cfset var workbookFactory = loadPoi("org.apache.poi.ss.usermodel.WorkbookFactory") />

		<cfset var workbook = workbookFactory.create(inputStream) />
		
		<cfif StructKeyExists(arguments, "sheet") and StructKeyExists(arguments, "sheetname")>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
					message="Cannot Provide Both Sheet and SheetName Attributes" 
					detail="Only one of either 'sheet' or 'sheetname' attributes may be provided.">
		</cfif>
		
		<cfset inputStream.close() />
		
		<cfif StructKeyExists(arguments, "sheet")>
			<cfset workbook.setActiveSheet(JavaCast("int", arguments.sheet - 1)) />
		<cfelseif StructKeyExists(arguments, "sheetname")>
			<cfset workbook.setActiveSheet(JavaCast("int", workbook.getSheetIndex(JavaCast("string", arguments.sheetname)))) />
		<cfelse>
			<cfset workbook.setActiveSheet(JavaCast("int", 0)) />
		</cfif>
		
		<cfreturn workbook />
	</cffunction>
	
	<cffunction name="writeToFile" access="private" output="false" returntype="void" 
			hint="Writes a spreadsheet file to disk">
		<cfargument name="filepath" type="string" required="true" />
		<cfargument name="workbook" type="any" required="true" />
		<cfargument name="overwrite" type="boolean" required="false" default="false" />
		
		<cfset var fos = 0 />
		
		<cfif not arguments.overwrite and FileExists(arguments.filepath)>
			<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
						message="File Exists" 
						detail="The file attempting to be written to already exists. Either use the update action or pass an overwrite argument of true to this function." />
		</cfif>
		
		<cfset fos = CreateObject("java", "java.io.FileOutputStream").init(arguments.filepath) />
		<cfset arguments.workbook.write(fos) />
		<cfset fos.close() />
	</cffunction>
	
	<cffunction name="cloneFont" access="private" output="false" returntype="any" 
			hint="Returns a new Font object with the same settings as the Font object passed in">
		<cfargument name="fontToClone" type="any" required="true" />

		<cfset var newFont = getWorkbook().createFont() />
		
		<!--- copy the existing cell's font settings to the new font --->
		<cfset newFont.setBoldweight(arguments.fontToClone.getBoldweight()) />
		<cfset newFont.setCharSet(arguments.fontToClone.getCharSet()) />
		<cfset newFont.setColor(arguments.fontToClone.getColor()) />
		<cfset newFont.setFontHeight(arguments.fontToClone.getFontHeight()) />
		<cfset newFont.setFontName(arguments.fontToClone.getFontName()) />
		<cfset newFont.setItalic(arguments.fontToClone.getItalic()) />
		<cfset newFont.setStrikeout(arguments.fontToClone.getStrikeout()) />
		<cfset newFont.setTypeOffset(arguments.fontToClone.getTypeOffset()) />
		<cfset newFont.setUnderline(arguments.fontToClone.getUnderline()) />
		
		<cfreturn newFont />
	</cffunction>
	
	<cffunction name="buildCellStyle" access="private" output="false" returntype="any" 
			hint="Builds an HSSFCellStyle with settings provided in a struct">
		<cfargument name="format" type="struct" required="true" />
		
		<!---Only some alignment types require the word "ALIGN" concatenated to them--->
		<cfset var alignList = "left, right, center, justify, general, fill, center_selection" />
		<cfset var nonAlignList = "vertical_top, vertical_bottom, vertical_center, vertical_justify" />
		
		<cfset var cellStyle = getWorkbook().createCellStyle() />
		<cfset var font = 0 />
		<cfset var setting = 0 />
		
		<!---
			Valid values of the format struct are:
			* alignment
			* bold
			* bottomborder
			* bottombordercolor
			* color
			* dataformat
			* fgcolor
			* fillpattern
			* font
			* fontsize
			* hidden
			* indent
			* italic
			* leftborder
			* leftbordercolor
			* locked
			* rightborder
			* rightbordercolor
			* rotation
			* strikeout
			* textwrap
			* topborder
			* topbordercolor
			* underline
		--->
		
		<cfloop collection="#arguments.format#" item="setting">
			<cfswitch expression="#setting#">
				<cfcase value="alignment">
					<cfif listFindNoCase(alignList,StructFind(arguments.format, setting))>
						<cfset cellStyle.setAlignment(Evaluate("cellStyle." & "ALIGN_" & UCase(StructFind(arguments.format, setting)))) />
					<cfelse>
						<cfset cellStyle.setVerticalAlignment(Evaluate("cellStyle." & UCase(StructFind(arguments.format, setting)))) />
					</cfif>
				</cfcase>
				
				<cfcase value="bold">
					<cfset font = cloneFont(getWorkbook().getFontAt(cellStyle.getFontIndex())) />
					
					<cfif StructFind(arguments.format, setting)>
						<cfset font.setBoldweight(font.BOLDWEIGHT_BOLD) />
					<cfelse>
						<cfset font.setBoldweight(font.BOLDWEIGHT_NORMAL)>
					</cfif>
					
					<cfset cellStyle.setFont(font) />
				</cfcase>
				
				<cfcase value="bottomborder">
					<cfset cellStyle.setBorderBottom(Evaluate("cellStyle." & "BORDER_" & UCase(StructFind(arguments.format, setting)))) />
				</cfcase>
				
				<cfcase value="bottombordercolor">
					<cfset cellStyle.setBottomBorderColor(JavaCast("int", getColorIndex(StructFind(arguments.format, setting)))) />
				</cfcase>
				
				<cfcase value="color">
					<cfset font = cloneFont(getWorkbook().getFontAt(cellStyle.getFontIndex())) />
					<cfset font.setColor(getColorIndex(StructFind(arguments.format, setting))) />
					<cfset cellStyle.setFont(font) />
				</cfcase>
				
				<!--- TODO: this is returning the correct data format index from HSSFDataFormat but 
							doesn't seem to have any effect on the cell. Could be that I'm testing 
							with OpenOffice so I'll have to check things in MS Excel --->
				<cfcase value="dataformat">
					<cfset cellStyle.setDataFormat(CreateObject("java", "org.apache.poi.hssf.usermodel.HSSFDataFormat").getBuiltinFormat(JavaCast("string", StructFind(arguments.format, setting)))) />
				</cfcase>

				<cfcase value="fgcolor">
					<cfset cellStyle.setFillForegroundColor(getColorIndex(StructFind(arguments.format, setting))) />
				</cfcase>
				
				<!--- TODO: CF 9 docs list "nofill" as opposed to "no_fill"; docs wrong? The rest match POI 
							settings exactly.If it really is nofill instead of no_fill, just change to no_fill 
							before calling setFillPattern --->
				<cfcase value="fillpattern">
					<cfset cellStyle.setFillPattern(Evaluate("cellStyle." & UCase(StructFind(arguments.format, setting)))) />
				</cfcase>

				<cfcase value="font">
					<cfset font = cloneFont(getWorkbook().getFontAt(cellStyle.getFontIndex())) />					
					<cfset font.setFontName(JavaCast("string", StructFind(arguments.format, setting))) />
					<cfset cellStyle.setFont(font) />
				</cfcase>

				<cfcase value="fontsize">
					<cfset font = cloneFont(getWorkbook().getFontAt(cellStyle.getFontIndex())) />					
					<cfset font.setFontHeightInPoints(JavaCast("int", StructFind(arguments.format, setting))) />
					<cfset cellStyle.setFont(font) />
				</cfcase>

				<!--- TODO: I may just not understand what's supposed to be happening here, 
							but this doesn't seem to do anything--->
				<cfcase value="hidden">
					<cfset cellStyle.setHidden(JavaCast("boolean", StructFind(arguments.format, setting))) />
				</cfcase>

				<!--- TODO: I may just not understand what's supposed to be happening here, 
							but this doesn't seem to do anything--->
				<cfcase value="indent">
					<cfset cellStyle.setIndention(JavaCast("int", StructFind(arguments.format, setting))) />
				</cfcase>

				<cfcase value="italic">
					<cfset font = cloneFont(getWorkbook().getFontAt(cellStyle.getFontIndex())) />
					
					<cfif StructFind(arguments.format, setting)>
						<cfset font.setItalic(JavaCast("boolean", true)) />
					<cfelse>
						<cfset font.setItalic(JavaCast("boolean", false)) />
					</cfif>
					
					<cfset cellStyle.setFont(font) />
				</cfcase>

				<cfcase value="leftborder">
					<cfset cellStyle.setBorderLeft(Evaluate("cellStyle." & "BORDER_" & UCase(StructFind(arguments.format, setting)))) />
				</cfcase>

				<cfcase value="leftbordercolor">
					<cfset cellStyle.setLeftBorderColor(getColorIndex(StructFind(arguments.format, setting))) />
				</cfcase>
				
				<!--- TODO: I may just not understand what's supposed to be happening here, 
							but this doesn't seem to do anything--->
				<cfcase value="locked">
					<cfset cellStyle.setLocked(JavaCast("boolean", StructFind(arguments.format, setting))) />
				</cfcase>

				<cfcase value="rightborder">
					<cfset cellStyle.setBorderRight(Evaluate("cellStyle." & "BORDER_" & UCase(StructFind(arguments.format, setting)))) />
				</cfcase>

				<cfcase value="rightbordercolor">
					<cfset cellStyle.setRightBorderColor(getColorIndex(StructFind(arguments.format, setting))) />
				</cfcase>

				<cfcase value="rotation">
					<cfset cellStyle.setRotation(JavaCast("int", StructFind(arguments.format, setting))) />
				</cfcase>

				<cfcase value="strikeout">
					<cfset font = cloneFont(getWorkbook().getFontAt(cellStyle.getFontIndex())) />
					
					<cfif StructFind(arguments.format, setting)>
						<cfset font.setStrikeout(JavaCast("boolean", true)) />
					<cfelse>
						<cfset font.setStrikeout(JavaCast("boolean", false)) />
					</cfif>
					
					<cfset cellStyle.setFont(font) />
				</cfcase>

				<cfcase value="textwrap">
					<cfset cellStyle.setWrapText(JavaCast("boolean", StructFind(arguments.format, setting))) />
				</cfcase>

				<cfcase value="topborder">
					<cfset cellStyle.setBorderTop(Evaluate("cellStyle." & "BORDER_" & UCase(StructFind(arguments.format, setting)))) />
				</cfcase>

				<cfcase value="topbordercolor">
					<cfset cellStyle.setTopBorderColor(getColorIndex(StructFind(arguments.format, setting))) />
				</cfcase>
				
				<cfcase value="underline">
					<cfset font = cloneFont(getWorkbook().getFontAt(cellStyle.getFontIndex())) />
					
					<cfif StructFind(arguments.format, setting)>
						<cfset font.setUnderline(JavaCast("boolean", true)) />
					<cfelse>
						<cfset font.setUnderline(JavaCast("boolean", false)) />
					</cfif>
					
					<cfset cellStyle.setFont(font) />
				</cfcase>
			</cfswitch>
		</cfloop>
		
		<cfreturn cellStyle />
	</cffunction>
	
	<cffunction name="getColorIndex" access="private" output="false" returntype="numeric" 
			hint="Returns the color index of a color string">
		<cfargument name="colorName" type="string" required="true" />
		
		<cfset var colorIndex = 0 />
		
		<!--- Evaluate doesn't seem to work with instantiating nested java classes, hence the switch. 
				And yes, each individual color is implemented as a nested class in HSSFColor. Joy. --->
		<cfswitch expression="#UCase(arguments.colorName)#">
			<cfcase value="AQUA">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$AQUA").index />
			</cfcase>
			
			<cfcase value="AUTOMATIC">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$AUTOMATIC").index />
			</cfcase>
			
			<cfcase value="BLACK">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$BLACK").index />
			</cfcase>
			
			<cfcase value="BLUE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$BLUE").index />
			</cfcase>
			
			<cfcase value="BLUE_GREY">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$BLUE_GREY").index />
			</cfcase>
			
			<cfcase value="BRIGHT_GREEN">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$BRIGHT_GREEN").index />
			</cfcase>
			
			<cfcase value="BROWN">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$BROWN").index />
			</cfcase>
			
			<cfcase value="CORAL">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$CORAL").index />
			</cfcase>
			
			<cfcase value="CORNFLOWER_BLUE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$CORNFLOWER_BLUE").index />
			</cfcase>
			
			<cfcase value="DARK_BLUE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$DARK_BLUE").index />
			</cfcase>
			
			<cfcase value="DARK_GREEN">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$DARK_GREEN").index />
			</cfcase>
			
			<cfcase value="DARK_RED">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$DARK_RED").index />
			</cfcase>
			
			<cfcase value="DARK_TEAL">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$DARK_TEAL").index />
			</cfcase>
			
			<cfcase value="DARK_YELLOW">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$DARK_YELLOW").index />
			</cfcase>
			
			<cfcase value="GOLD">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$GOLD").index />
			</cfcase>
			
			<cfcase value="GREEN">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$GREEN").index />
			</cfcase>
			
			<cfcase value="GREY_25_PERCENT">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$GREY_25_PERCENT").index />
			</cfcase>
			
			<cfcase value="GREY_40_PERCENT">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$GREY_40_PERCENT").index />
			</cfcase>
			
			<cfcase value="GREY_50_PERCENT">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$GREY_50_PERCENT").index />
			</cfcase>
			
			<cfcase value="GREY_80_PERCENT">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$GREY_80_PERCENT").index />
			</cfcase>
			
			<cfcase value="INDIGO">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$INDIGO").index />
			</cfcase>
			
			<cfcase value="LAVENDER">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LAVENDER").index />
			</cfcase>
			
			<cfcase value="LEMON_CHIFFON">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LEMON_CHIFFON").index />
			</cfcase>
			
			<cfcase value="LIGHT_BLUE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LIGHT_BLUE").index />
			</cfcase>
			
			<cfcase value="LIGHT_CORNFLOWER_BLUE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LIGHT_CORNFLOWER_BLUE").index />
			</cfcase>
			
			<cfcase value="LIGHT_GREEN">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LIGHT_GREEN").index />
			</cfcase>
			
			<cfcase value="LIGHT_ORANGE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LIGHT_ORANGE").index />
			</cfcase>
			
			<cfcase value="LIGHT_TURQUOISE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LIGHT_TURQUOISE").index />
			</cfcase>
			
			<cfcase value="LIGHT_YELLOW">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LIGHT_YELLOW").index />
			</cfcase>
			
			<cfcase value="LIME">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$LIME").index />
			</cfcase>
			
			<cfcase value="MAROON">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$MAROON").index />
			</cfcase>
			
			<cfcase value="OLIVE_GREEN">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$OLIVE_GREEN").index />
			</cfcase>
			
			<cfcase value="ORANGE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$ORANGE").index />
			</cfcase>
			
			<cfcase value="ORCHID">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$ORCHID").index />
			</cfcase>
			
			<cfcase value="PALE_BLUE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$PALE_BLUE").index />
			</cfcase>
			
			<cfcase value="PINK">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$PINK").index />
			</cfcase>
			
			<cfcase value="PLUM">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$PLUM").index />
			</cfcase>
			
			<cfcase value="RED">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$RED").index />
			</cfcase>
			
			<cfcase value="ROSE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$ROSE").index />
			</cfcase>
			
			<cfcase value="ROYAL_BLUE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$ROYAL_BLUE").index />
			</cfcase>
			
			<cfcase value="SEA_GREEN">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$SEA_GREEN").index />
			</cfcase>
			
			<cfcase value="SKY_BLUE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$SKY_BLUE").index />
			</cfcase>
			
			<cfcase value="TAN">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$TAN").index />
			</cfcase>
			
			<cfcase value="TEAL">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$TEAL").index />
			</cfcase>
			
			<cfcase value="TURQUOISE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$TURQUOISE").index />
			</cfcase>
			
			<cfcase value="VIOLET">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$VIOLET").index />
			</cfcase>
			
			<cfcase value="WHITE">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$WHITE").index />
			</cfcase>
			
			<cfcase value="YELLOW">
				<cfset colorIndex = CreateObject("java", "org.apache.poi.hssf.util.HSSFColor$YELLOW").index />
			</cfcase>
			
			<cfdefaultcase>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Invalid Color" 
							detail="The color provided (#arguments.colorName#) is not valid." />
			</cfdefaultcase>
		</cfswitch>
		
		<cfreturn colorIndex />
	</cffunction>
	
	<cffunction name="getJavaColorRGB" access="private" output="false" returntype="struct" 
			hint="Returns a struct containing RGB values from java.awt.Color for the color name passed in">
		<cfargument name="colorName" type="string" required="true" />
		
		<cfset var color = 0 />
		<cfset var colorRGB = StructNew() />
		
		<cfswitch expression="#arguments.colorName#">
			<cfcase value="black">
				<cfset color = CreateObject("java", "java.awt.Color").BLACK />
			</cfcase>
			
			<cfcase value="blue">
				<cfset color = CreateObject("java", "java.awt.Color").BLUE />
			</cfcase>
			
			<cfcase value="cyan">
				<cfset color = CreateObject("java", "java.awt.Color").CYAN />
			</cfcase>
			
			<cfcase value="dark_gray,darkGray" delimiters=",">
				<cfset color = CreateObject("java", "java.awt.Color").DARK_GRAY />
			</cfcase>
			
			<cfcase value="gray">
				<cfset color = CreateObject("java", "java.awt.Color").GRAY />
			</cfcase>

			<cfcase value="green">
				<cfset color = CreateObject("java", "java.awt.Color").GREEN />
			</cfcase>

			<cfcase value="light_gray,lightGray" delimiters=",">
				<cfset color = CreateObject("java", "java.awt.Color").LIGHT_GRAY />
			</cfcase>

			<cfcase value="magenta">
				<cfset color = CreateObject("java", "java.awt.Color").MAGENTA />
			</cfcase>

			<cfcase value="orange">
				<cfset color = CreateObject("java", "java.awt.Color").ORANGE />
			</cfcase>

			<cfcase value="pink">
				<cfset color = CreateObject("java", "java.awt.Color").PINK />
			</cfcase>

			<cfcase value="red">
				<cfset color = CreateObject("java", "java.awt.Color").RED />
			</cfcase>

			<cfcase value="white">
				<cfset color = CreateObject("java", "java.awt.Color").WHITE />
			</cfcase>

			<cfcase value="yellow">
				<cfset color = CreateObject("java", "java.awt.Color").YELLOW />
			</cfcase>

			<cfdefaultcase>
				<cfthrow type="org.cfpoi.spreadsheet.Spreadsheet" 
							message="Invalid Color" 
							detail="The color provided (#arguments.colorName#) is not valid." />
			</cfdefaultcase>
		</cfswitch>
		
		<cfset colorRGB.red = color.getRed() />
		<cfset colorRGB.green = color.getGreen() />
		<cfset colorRGB.blue = color.getBlue() />

		<cfreturn colorRGB />
	</cffunction>
</cfcomponent>