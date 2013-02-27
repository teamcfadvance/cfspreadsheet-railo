component displayname="extension provider" output="false" {
	
	instance = {
		thisAddress = "http://#cgi.SERVER_NAME#:#cgi.SERVER_PORT#/RailoExtensionProvider/"
	};
	
	
	remote struct function getInfo(){
		var info = {
			title="AndyJarrett.co.uk",
			description="",
			image="http://www.andyjarrett.co.uk/andy_jarrett_logo.png",
			url="http://www.andyjarrett.co.uk/blog",
			mode="develop"
		};
		return info;
	}
	
	remote query function listApplications(){
		var apps = queryNew('type,id,name,label,description,version,category,image,download,author,codename,video,support,documentation,forum,mailinglist,network,created');
		var rootURL=getInfo().url;
		var desc = "cfspreadhsheet tag &amp; functions";
		QueryAddRow(apps);
		QuerySetCell(apps,'id','10EEC23A-0779-4068-9507A9C5ED4A8641');
		QuerySetCell(apps,'name','CFPOI');
		QuerySetCell(apps,'type','web');
		QuerySetCell(apps,'label','&lt;cfspreadhsheet/&gt; tag &amp; functions');
		QuerySetCell(apps,'description',desc);
		QuerySetCell(apps,'author','Ext by Andy Jarrett.<br/>CFPOI by Matt Woodward');
		QuerySetCell(apps,'image','http://www.gstatic.com/codesite/ph/images/defaultlogo.png');
		QuerySetCell(apps,'support','https://github.com/andyj/RailoExtensionProvider/issues');
		QuerySetCell(apps,'documentation','http://code.google.com/p/cfpoi/w/list');
		QuerySetCell(apps,'created',CreateDate(2009,2,24));
		QuerySetCell(apps,'version',"v34");
		QuerySetCell(apps,'category',"Application");
		QuerySetCell(apps,'download','#instance.thisAddress#/cfpoi.zip');
		//QuerySetCell(apps,'download','https://github.com/andyj/RailoExtensionProvider/blob/master/cfpoi.zip');
		
		return apps;			
	}
}

