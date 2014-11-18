&lt;cfspreadsheet&gt; for Railo

This is a fork of what used to be at http://code.google.com/p/cfpoi/ (no longer active) for Railo so a massive thank you must go out to Matt Woodward (www.mattwoodward.com / @mpwoodward) as he has done all the hard work for this project I am just forking it to make it workable with Railo as an extension.

At the moment this should be considered alpha/beta as it needs some full testing to be done. It is currently working for my limited needs but it requires feedback before being considered as stable.



## Installation

The installation process is in a bit of disrepair, but here are the options:

## Latest Version Installation

### Automatic Installation - ZIP Upload via Railo Admin GUI

Normally, you'd be able to install via the Web administrator, by uploading a ZIP; however, this doesn't work at present, as there seems to be a problem with this plugin's ZIP and/or with [Railo](https://issues.jboss.org/browse/RAILO-2502).

### Manual Installation

The zip contains nested zips, so the installation procedure is best explained as this executable (Linux) script. (This also makes a decent Vagrant provisioning step, if you're into that sort of thing.)

``` bash
# set serverPath to your railo web folder
serverPath=/path/to/your/web/root/WEB-INF/railo # USE YOUR OWN PATH!

# remove temp dir if this is not the first run
rm -rf /tmp/cfspreadsheet
# create temp dir
mkdir -p /tmp/cfspreadsheet
# download the zip (change to https://raw.githubusercontent.com/teamcfadvance/cfspreadsheet-railo/master/cfspreadsheetInstaller.zip if pull request is accepted)
cd /tmp/cfspreadsheet && curl -C - -O https://raw.githubusercontent.com/jamiejackson/cfspreadsheet-railo/installation_instructions_tweak/cfspreadsheetInstaller.zip
# unzip the installer
unzip -o cfspreadsheetInstaller.zip
# unzip the cfpoi.zip, which, itself, contains another cfpoi.zip, so let's jump through a couple hoops
mv cfpoi.zip foo.zip && unzip -o foo.zip && rm foo.zip

# install the extension by extracting some zips into the proper locations
unzip -o -d ${serverPath}/components/org/ cfpoi.zip
unzip -o -d ${serverPath}/library/function functions.zip
unzip -o -d ${serverPath}/library/tag tags.zip
unzip -o -d ${serverPath}/lib poiLib.zip

# restart railo to activate extension.
# you may need to restart your servlet container instead, depending on your railo installation
sudo service railo_ctl restart

```

### Semi-Automatic "Spoofed Provider" Installation via Railo Admin GUI

In this method, you will create your own temporary extension provider:

1. Download https://raw.githubusercontent.com/teamcfadvance/cfspreadsheet-railo/master/cfspreadsheetInstaller.zip
1. Extract it as `/path/to/your/wwwroot/RailoExtensionProvider` (if it has ExtensionProvider.cfc directly under that directory, you've extracted properly)
1. Add a file `/path/to/your/wwwroot/RailoExtensionProvider/Application.cfc` with the contents<br>`component {}`
1. Railo *Web* Admin > Extensions > Providers: Add:  http://&lt;your_local_site_host&gt;/RailoExtensionProvider/ExtensionProvider.cfc
1. Railo *Web* Admin > Extensions > Applications:  Click on "&lt;cfspreadhsheet/&gt; tag &amp;a... Free" (sic)
1. Click "Install" Button
1. Restart Railo or Tomcat (depending on installation type)

## Outdated Version Installation

### (Outdated) Automatic Installation via Railo Admin GUI

This provider URL is no longer maintained, and will supply an outdated version of &lt;cfspreadsheet&gt;:

log in to your Web Administrator and add http://www.andyjarrett.co.uk/RailoExtensionProvider/ExtensionProvider.cfc to your Providers (under Extension)
Then go to Applications (still under Extension) and install "CFPOI, a wrapper for the Apache POI Project."
Once done you will need to restart your Railo server.
With your server restarted you should be able to use the cfspreadsheet functions and tags. 

All the code os open source so to have a look around or log issues please check out the following links: 
Source: https://github.com/andyj/RailoExtensionProvider
Report issues: https://github.com/andyj/RailoExtensionProvider/issues
