#!/usr/bin/env ruby
#
# These are the ruby gems that we require to be installed
# you would do 'gem install openssl', etc. from the command prompt for a new installation

# This gem probably not required in Ruby > 1.9 because it's included by default
require 'rubygems'
# watir is a Ruby gem that automates web browser interactions
require 'watir-webdriver'
# openssl gem required because we're browsing HTTPS websites
require 'openssl'
# Selenium is a gem that we use to configure the Firefox profile settings
require 'selenium-webdriver'


### This function sets up the browser and logs in to the extranetdev / cognos site
def login
  # Pick download directory and write it to the screen
  @download_directory = "C:\\Code\\downloads"
  puts "Download dir: #{@download_directory}"
  
  # Configuring Firefox to use the download directory and save files to disk instead of prompting
  @profile = Selenium::WebDriver::Firefox::Profile.new
  @profile['browser.download.dir'] = @download_directory
  @profile['browser.download.folderList'] = 2
  @profile['browser.helperApps.neverAsk.saveToDisk'] = "application/pdf,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  @profile['pdfjs.disabled'] = true # Disables pdf.js in newer Firefox
  @profile['pdfjs.firstRun'] = false # More disabling pdfjs
  @profile['pdfjs.previousHandler.preferredAction'] = 0 # And more disabling
  @profile['plugin.scan.Acrobat'] = "999" # Disables Acrobat PDF Viewer plugin
  
  # Assigns this new profile to the Firefox browser object we create here
  @browser = Watir::Browser.new :firefox, :profile => @profile

  ##### Primary Logon Page
  @browser.goto 'https://extranetdev.marriott.com/pkmslogin.form'
  sleep 1 # Give the page time to load so that we don't start trying to interact with a page that may not be there yet
# Find a text field where <INPUT NAME="username" and set the value to "eabarr218"
# Got this information from reading the source for the extranetdev site
  @browser.text_field(:name, 'username').set @user
# Same for this one
  @browser.text_field(:name, 'password').set @pass
  # The browser is normally pretty quick at processing the login click and sending the request to the server
  # but if it ever breaks, put a 1 second sleep after this statement
  @browser.button(:value => 'Login').click
  ##### Cognos Logon Page
  # Now we log onto the cognos site within the Marriott extranetdev site
  @browser.goto 'https://extranetdev.marriott.com/portal/cgi-bin/cognosisapi.dll'
  # Again so we don't clobber the page loading
  sleep 1 
  # Again, got these values from reading the HTML source for the cognos logon page
  @browser.text_field(:name, 'CAMUsername').set @user
  @browser.text_field(:name, 'CAMPassword').set @pass
  # Click OK to log in and get the session cookies we need to continues
  @browser.button(:value => 'OK').click
end


### Build the hash of reports
# Key = Report Page
# Value = Report
## Had to do it this way because keys needed to be unique but had to track both values
# so that we could pass both in the HTTP GET request for the Cognos page
def buildList
	@reports = {
###### ATR1
    "ATR1 H Property Detail" => "ATR1 Account Production Summary",
    "ATR1 H Summary" => "ATR1 Account Production Summary",
    "ATR1 P Property Detail" => "ATR1 Account Production Summary",
    "ATR1 P Summary" => "ATR1 Account Production Summary",
    "ATR1 X Property Detail" => "ATR1 Account Production Summary",
    "ATR1 X Summary" => "ATR1 Account Production Summary",
###### ATR2
		"ATR2 H Booking Source" => "ATR2 Account Production Analysis",
		"ATR2 H Freq Program" => "ATR2 Account Production Analysis",
		"ATR2 H GDS Detail" => "ATR2 Account Production Analysis",
		"ATR2 H Market Category" => "ATR2 Account Production Analysis",
		"ATR2 H Market Code" => "ATR2 Account Production Analysis",
		"ATR2 H Market Prefix" => "ATR2 Account Production Analysis",
		"ATR2 H Property Detail" => "ATR2 Account Production Analysis",
		"ATR2 H Room Pool" => "ATR2 Account Production Analysis",
		"ATR2 H Room Pool Mkt Pfx" => "ATR2 Account Production Analysis",
		"ATR2 H Summary" => "ATR2 Account Production Analysis",
		"ATR2 P Booking Source" => "ATR2 Account Production Analysis",
		"ATR2 P Freq Program" => "ATR2 Account Production Analysis",
		"ATR2 P GDS Detail" => "ATR2 Account Production Analysis",
		"ATR2 P Market Category" => "ATR2 Account Production Analysis",
		"ATR2 P Market Code" => "ATR2 Account Production Analysis",
		"ATR2 P Market Prefix" => "ATR2 Account Production Analysis",
		"ATR2 P Property Detail" => "ATR2 Account Production Analysis",
		"ATR2 P Room Pool" => "ATR2 Account Production Analysis",
		"ATR2 P Room Pool Mkt Pfx" => "ATR2 Account Production Analysis",
		"ATR2 P Summary" => "ATR2 Account Production Analysis",
		"ATR2 X Booking Source" => "ATR2 Account Production Analysis",
		"ATR2 X Freq Program" => "ATR2 Account Production Analysis",
		"ATR2 X GDS Detail" => "ATR2 Account Production Analysis",
		"ATR2 X Market Category" => "ATR2 Account Production Analysis",
		"ATR2 X Market Code" => "ATR2 Account Production Analysis",
		"ATR2 X Market Prefix" => "ATR2 Account Production Analysis",
		"ATR2 X Property Detail" => "ATR2 Account Production Analysis",
		"ATR2 X Room Pool" => "ATR2 Account Production Analysis",
		"ATR2 X Room Pool Mkt Pfx" => "ATR2 Account Production Analysis",
		"ATR2 X Summary" => "ATR2 Account Production Analysis",
###### ATR3
    "ATR3 H Summary" => "ATR3 Account Production By Day Of Week",
    "ATR3 P Summary" => "ATR3 Account Production By Day Of Week",
    "ATR3 X Summary" => "ATR3 Account Production By Day Of Week",
###### ATR4
    "ATR4 H Room Pool" => "ATR4 Account Production By Length Of Stay",
    "ATR4 H Summary" => "ATR4 Account Production By Length Of Stay",
    "ATR4 P Room Pool" => "ATR4 Account Production By Length Of Stay",
    "ATR4 P Summary" => "ATR4 Account Production By Length Of Stay",
    "ATR4 X Room Pool" => "ATR4 Account Production By Length Of Stay",
    "ATR4 X Summary" => "ATR4 Account Production By Length Of Stay",
###### ATR5
    "ATR5 H Marriott Rewards" => "ATR5 Account Production By Rate Mix",
    "ATR5 H Property Detail" => "ATR5 Account Production By Rate Mix",
    "ATR5 H Summary" => "ATR5 Account Production By Rate Mix",
    "ATR5 P Marriott Rewards" => "ATR5 Account Production By Rate Mix",
    "ATR5 P Property Detail" => "ATR5 Account Production By Rate Mix",
    "ATR5 P Summary" => "ATR5 Account Production By Rate Mix",
    "ATR5 X Marriott Rewards" => "ATR5 Account Production By Rate Mix",
    "ATR5 X Property Detail" => "ATR5 Account Production By Rate Mix",
    "ATR5 X Summary" => "ATR5 Account Production By Rate Mix"
	}
end



## Connect to the report URL
# URL = query URL
# report = Pretty name of the URL [i.e.: ATR2 H Blah]
# index = which query we're on [1-5]
def connect(url,report,index)

### Tells the browser to go to the URL we want, based upon the query URL we pass in
  @browser.goto url
  ## Change this depending on type of report, 20 seconds may not be enough and we won't grab the output
  sleep(20) # Wait X seconds for the report to generate
  # Report name
#  name = "#{report}-#{index}."
#  con = 0

### There was going to be more code down here about how to save it based on the name and being able to write out HTML files
# But ran into issues w/ the watir gem and being able to accurately save the page
# Also, watir used to have the ability to do full page screencaptures (as in everything on the page from top to bottom, not just what
# The browser can see at once. However, that feature was removed a few months ago and despite asking, it's not being added back in yet.
# So there's no good way to capture the HTML reports.
#  if( report =~ /ATR[0-9] H/ )
#		# If it's the HTML, we don't know how to handle this
#  elsif( report =~ /ATR[0-9] P/ ) # If it's the PDF
#		# If it's PDF, it automatically saves file to disk  
#  elsif( report =~ /ATR[0-9] X/ )
		# If it's the Excel file, it automatically saves to disk
#  end
end


### Prompts to determine whether or not we want to run queries on a certain report
def prompt
	print "Type 'GO' to continue or 'SKIP' to bypass this report.\n"
	input = gets.strip
### If the value of 'input' variable is case insensitive like the letters "go" then return 1
	if( input =~ /go/i )
		return 1
	elsif( input =~ /skip/i )
### If the value of 'input' is case insensitive like 'skip', return -1
		return -1
	else
		return 0
	end
end


### This is the real meat of the program - work happens here
def doList
##### For each key (report page) in our hash of reports (key = report page, value = report), we do the following
  @reports.keys.each do |report|
# The base Cognos report URL that we are querying, this remains constant for everything
  baseurl = "https://extranetdev.marriott.com/cognos10/portal/cgi-bin/cognosisapi.dll?b_action=cognosViewer&ui.action=run&ui.object=CAMID('marreds:u:cn=ebarr218,ou=people')/folder[@name='My Folders']/folder[@name='Account Tracking']/"
# This is the specific customization where we pull the report and report page and put into the web request
  specific = "folder[@name='#{@reports[report]}']/report[@name='#{report}']"
# Additional options 
  tfname = "Period-%20TY%20%26%20LY"

  # If the report is of type ATR4 or ATR5, we use the period time frame
	if( report =~ /ATR[45] / )
		tfname = "PD"
	elsif( report =~ /ATR3 / )
	# Otherwise if report is of type ATR3, we use PD and YTD
		tfname = "PD%20%26%20YTD"
	end

# Prompting the user if we want to run the queries for this report page	
	con = 0
	print "Do we want to run the four queries for report #{report}\n"
	while( con == 0 )
		con = prompt
	end
	if( con == 1 )
		# We want to run this report because the user entered GO/go
		# Build the query string for this report
		q1 = "&cv.toolbar=false&cv.header=false&run.prompt=false&p_security=1&p_PorM=P&p_AsOfDate=201109&p_TFName=#{tfname}&p_Currency=Local&p_NetGross=Net&p_NumAct=350&p_AccType=&p_AcctHier=H&p_Prop=NYCMQ&p_Prop=ATLMQ&p_PropComp=&p_MgmtType=&p_USumm=1&p_Brand=&p_PropCountry=&p_reportType=P&p_AVP=&p_MyC=&p_PubC=&p_GblReg=&p_GblDiv=&p_ActId=1&p_TF1Rmnts=1&p_TF1Rev=1&p_TF1ADR=1&p_TF1LOS=1&p_TF1PCT=1&p_TF2Rmnts=1&p_TF2Rev=1&p_TF2ADR=1&p_TF2LOS=1&p_TF2PCT=1&p_PCRmnts=1&p_PCRev=1&p_PCADR=1&p_PCLOS=1&p_PCPCT=1&p_MktCat=&p_RptContent=D"
		query = baseurl + specific + q1
		## Call the connect function for this query/report and track that it is query #1
		connect(query,report,1)
		
		q2 = "&cv.toolbar=false&cv.header=false&run.prompt=false&p_security=1&p_PorM=P&p_AsOfDate=201109&p_TFName=#{tfname}&p_Currency=Local&p_NetGross=Net&p_NumAct=350&p_AccType=&p_AcctHier=H&p_Prop=NYCMQ&p_PropComp=&p_MgmtType=&p_USumm=1&p_Brand=&p_PropCountry=&p_reportType=P&p_AVP=&p_MyC=&p_PubC=&p_GblReg=&p_GblDiv=&p_ActId=1&p_TF1Rmnts=1&p_TF1Rev=1&p_TF1ADR=1&p_TF1LOS=1&p_TF1PCT=1&p_TF2Rmnts=1&p_TF2Rev=1&p_TF2ADR=1&p_TF2LOS=1&p_TF2PCT=1&p_PCRmnts=1&p_PCRev=1&p_PCADR=1&p_PCLOS=1&p_PCPCT=1&p_MktCat=&p_RptContent=S"
		query = baseurl + specific + q2
		connect(query,report,2)

		q3 = "&cv.toolbar=false&cv.header=false&run.prompt=false&p_security=1&p_PorM=P&p_AsOfDate=201109&p_TFName=#{tfname}&p_Currency=Local&p_NetGross=Net&p_NumAct=350&p_AccType=&p_AcctHier=H&p_Prop=&p_PropComp=&p_MgmtType=&p_USumm=1&p_Brand=AK&p_PropCountry=US&p_reportType=P&p_AVP=&p_MyC=&p_PubC=&p_GblReg=&p_GblDiv=&p_ActId=1&p_TF1Rmnts=1&p_TF1Rev=1&p_TF1ADR=1&p_TF1LOS=1&p_TF1PCT=1&p_TF2Rmnts=1&p_TF2Rev=1&p_TF2ADR=1&p_TF2LOS=1&p_TF2PCT=1&p_PCRmnts=1&p_PCRev=1&p_PCADR=1&p_PCLOS=1&p_PCPCT=1&p_MktCat=&p_RptContent=S"
		query = baseurl + specific + q3
		connect(query,report,3)

		q4 = "&cv.toolbar=false&cv.header=false&run.prompt=false&p_security=1&p_PorM=P&p_AsOfDate=201109&p_TFName=#{tfname}&p_Currency=Local&p_NetGross=Net&p_NumAct=350&p_AccType=&p_AcctHier=H&p_Prop=&p_PropComp=&p_MgmtType=&p_USumm=1&p_Brand=AK&p_PropCountry=US&p_reportType=P&p_AVP=&p_MyC=&p_PubC=&p_GblReg=&p_GblDiv=&p_ActId=1&p_TF1Rmnts=1&p_TF1Rev=1&p_TF1ADR=1&p_TF1LOS=1&p_TF1PCT=1&p_TF2Rmnts=1&p_TF2Rev=1&p_TF2ADR=1&p_TF2LOS=1&p_TF2PCT=1&p_PCRmnts=1&p_PCRev=1&p_PCADR=1&p_PCLOS=1&p_PCPCT=1&p_MktCat=&p_RptContent=D"
		query = baseurl + specific + q4
		connect(query,report,4)
	end
  end
end

# Get out credentials
def getcredentials
	@user = ""
	@pass = ""
	file = File.new( "creds.txt", "r" )
	num = 0
	while( line = file.gets )
		line = line.chomp
		if( num == 0 )
			@user = line # First line of creds.txt is username
		elsif( num == 1 )
			@pass = line # Second line of creds.txt is password
		end
		num += 1
	end
	print "Using credentials for #{@user}\n"
end

# The script runs these functions, sequentially, in order to produce the desired output
getcredentials
login
buildList
doList
