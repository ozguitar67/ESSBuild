#SingleInstance Force

blank := 0
DACIP := ""
EIACOL := ""
MPGCOL := ""
SVCCOL := ""
SRCIDCOL := ""
DEVCOL := ""
PROVCOL := ""
APEX_NAME := ""
HEADER := ""
SVCi := 1
XMLEND := 0
REMOVED := 0
ADDED := 0

IfExist, Added.txt
  FileDelete, Added.txt
IfExist, Removed.txt
  FileDelete, Removed.txt
IfExist, Issues.txt
  FileDelete, Issues.txt
IfExist, dacConfiguration_new.xml
  FileDelete, dacConfiguration_new.xml

Gui, Add, Text, X10 Y40, Select the exported dacConfiguration.xml file
Gui, Add, Edit, X10 Y60 W300 vXMLFILE
Gui, Add, Button, X315 Y59 W55 vXMLf gGetFile, Browse
Gui, Add, Text, X10 Y20, DAC IP: 
Gui, Add, Edit, X50 Y17 W90 vDACIP
Gui, Add, Button, X373 Y59 Default, OK
Gui, Add, Text, X10 Y88, EIA Column:
Gui, Add, Edit, Uppercase X70 Y85 W25 vEIACOL
Gui, Add, Text, X100 Y88, MPEG Column:
Gui, Add, Edit, Uppercase X175 Y85 W25 vMPGCOL
Gui, Add, Text, X205 Y88, Service Name Column:
Gui, Add, Edit, Uppercase X315 Y85 W25 vSVCCOL
Gui, Add, Text, X10 Y113, Source ID Column:
Gui, Add, Edit, Uppercase X100 Y110 W25 vSRCIDCOL
Gui, Add, Text, X130 Y113, Device Column:
Gui, Add, Edit, Uppercase X207 Y110 W25 vDEVCOL
Gui, Add, Text, X237 Y113, Provider Name Column:
Gui, Add, Edit, Uppercase X350 Y110 W25 vPROVCOL
Gui, Show, W415, ESS Build
return

ButtonOK:
Gui, Submit
Gui, Destroy
SVCCOLe := RegExMatch(SVCCOL,"[a-zA-Z]{1,}")
SRCIDCOLe := RegExMatch(SRCIDCOL,"[a-zA-Z]{1,}")
PROVCOLe := RegExMatch(PROVCOL,"[a-zA-Z]{1,}")
EIACOLe := RegExMatch(EIACOL,"[a-zA-Z]{1,}")
MPGCOLe := RegExMatch(MPGCOL,"[a-zA-Z]{1,}")
DEVCOLe := RegExMatch(DEVCOL,"[a-zA-Z]{1,}")
DACIPe := RegExMatch(DACIP,"[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}")
XMLFILEe := RegExMatch(XMLFILE,"i).*\.xml$")
if ((EIACOLe < 1) || (MPGCOLe < 1) || (DEVCOLe < 1) || (XMLFILEe < 1))
{
 msgbox, Necessary information is missing.`nExiting...
 ExitApp
}
if ((SVCCOLe < 1) || (PROVCOLe < 1))
{
 if ((SRCIDCOLe < 1) || (DACIPe < 1))
   {
	msgbox, You need the Source ID column and DAC IP if you do not have the Source and Provider Names.`nExiting...
	ExitApp
   }
}
FileRead, DACexport, %XMLFILE%
LINE := 1
Loop
{
 Loop, Parse, DACexport, `n, `r
 {
  if (LINE > A_Index)
 	continue
  FoundPos := RegExMatch(A_LoopField,"<externalServiceSet externalServiceSetName=""(.*)"">",ImportESSname)
  if (FoundPos <> 0)
    {
 	 ESSNAME := ImportESSname1 . ".xml"
 	 ESSLIST := ESSLIST . "," ESSNAME
 	 LINE := A_Index
 	 break
    }
  FoundPos := RegExMatch(A_LoopField,"</externalServiceSets>")
  if (FoundPos > 0)
    {
     XMLEND := 1
 	 break
    }
 }
 if (XMLEND = 1)
   break
 ESSLIST := RegExReplace(ESSLIST,"^,","")
 COUNT := 0
 Loop, Parse, DACexport, `n, `r
 {
  if (LINE > A_Index)
 	continue
  FoundPos := RegExMatch(A_LoopField,"</externalServiceSetEntries>")
  if (FoundPos > 0)
    {
 	 LINE := A_Index
 	 break
    }
  FoundPos := RegExMatch(A_LoopField,"<externalServiceSetEntry ")
  if (FoundPos <> 0)
    {
      COUNT++
 	 if (COUNT = 1)
 	 	FileAppend, %A_LoopField%, %ESSNAME%
 	 else
 	 	FileAppend, `n%A_LoopField%, %ESSNAME%
    }
 }
}
if (SVCCOLe < 1)
{
  IE := ComObjCreate("InternetExplorer.Application")
  URL := "http://" . DACIP . ":8081/cgi-bin/service_list.pl"
  IE.Navigate(URL)
  while IE.ReadyState <> 4
   {
  	sleep, 250
  	continue
   }
  HTML := IE.Document.All[0].innerHTML
  IE.Quit
  Services := []
  Loop, Parse, HTML, `n, `r
   {
  	StringGetPos, FoundPos, A_LoopField, <option value=
  	  if ((FoundPos <> -1))
  		{
  		 FoundPos := RegExMatch(A_LoopField,"value=""[0-9]*"" alt=""(.*)""",SRCID)
  		 FoundPos := RegExMatch(A_LoopField,">([^ ]*) ",SOURCE)
  		 FoundPos := RegExMatch(A_LoopField,">[^ ]* (.*)<",PROVIDER)  		 
  		 Services[SVCi] := new Service
  		 Services[SVCi].SourceID := SRCID1 
  		 Services[SVCi].SourceName := SOURCE1 
  		 Services[SVCi].ProviderName := PROVIDER1
  		 SVCi++
  		}
  	  else
  		continue
   }
}
xlApp := ComObjActive("Excel.Application")
xl := xlApp.ActiveSheet
Loop
{
	if (blank > 17)
		break
	EIA := xl.Range(EIACOL . A_Index).Value
	SetFormat, float, 6.0
	EIA+=0
	EIA := "A" . EIA
	EIA := RegExReplace(EIA," ","")
	MPEG := xl.Range(MPGCOL . A_Index).Value
	SetFormat, float, 6.0
	MPEG+=0
	if (SRCIDCOLe <> 0)
	  {
		SOURCEID := xl.Range(SRCIDCOL . A_Index).Value
		SetFormat, float, 6.0
		SOURCEID+=0
	  }
	else 
	  {
		SOURCEID := 0
		SetFormat, float, 6.0
		SOURCEID+=0
	  }
	if (SVCCOLe < 1)
		Gosub, FindSourceName
	else
	  {
	   SVCNAME := xl.Range(SVCCOL . A_Index).Value
	   PROVNAME := xl.Range(PROVCOL . A_Index).Value
	  }
	ANDSN := RegExReplace(SVCNAME,"&","&amp;",ANDCOUNT)
	if ((ANDCOUNT > 0) && (SVCCOLe <> 0))
		SVCNAME := ANDSN	   
	APEX_NAME := xl.Range(DEVCOL . A_Index).Value
	StringUpper, APEX_NAME, APEX_NAME
	if ((SVCCOLe > 0) && (PROVCOLe > 0)) 
	{
	  if ((MPEG < 1) || (PROVNAME = "") || (SVCNAME = "") || (EIA = "A0") || (APEX_NAME = ""))
	  {
	   blank++
	   MPEG := PROVNAME := SVCNAME := EIA := SOURCEID := ""
	   continue
	  }
	}
	else if ((SRCIDCOLe > 0) && (DACIPe > 0))
	{
	  if ((MPEG < 1) || (SRCID < 1) || (EIA = "A0") || (APEX_NAME = ""))
	  {
	   blank++
	   MPEG := PROVNAME := SVCNAME := EIA := SOURCEID := ""
	   continue
	  }
	}
	else
	 blank := 0
	FILENAME := APEX_NAME . ".ess"
	FILENAME := RegExReplace(FILENAME, " ", "")
	FILELIST := FILENAME . "," . FILELIST
	ESS_ENTRY := "<externalServiceSetEntry modulationMode=""QAM_256"" mpegServiceNumber=""" . MPEG . """ serviceProviderName=""" . PROVNAME . """ sourceName=""" . SVCNAME . """ sourceUserId="""" tunedChannel=""" . EIA . """/>"
	Loop, Parse, ESSLIST, `,
	{
	 X := CleanESS(A_LoopField, SVCNAME, PROVNAME)
	 if (X = 1)
		REMOVED++
	}
	MPEG := PROVNAME := SVCNAME := EIA := SOURCEID := ""
	FileAppend, %ESS_ENTRY%`n, %FILENAME%
	FileAppend, %ESS_ENTRY%`n,Added.txt
	ADDED++
	blank := 0
}
Sort, FILELIST, U D,
guiESSLIST := RegExReplace(FILELIST,",$","")
guiESSLIST := RegExReplace(guiESSLIST,".ess","")
guiESSLIST := RegExReplace(guiESSLIST,",","|")
Loop, Parse, ESSLIST, `,
{
	ESSNAME := RegExReplace(A_LoopField,".xml","")
	if (ESSNAME = "")
		continue	
	FileRead, ESSINFO, %A_LoopField%
	Gui, 2: Add, Text, X10 Y20, What new name would you like for the following ESS: %ESSNAME%?
	Gui, 2: Add, Edit, X10 Y38 vESSNAME W160, ESS Name
	Gui, 2: Add, Text, X10 Y65, Would like like to merge entries from a given device (Not Required)?
	Gui, 2: Add, DropDownList, X10 Y90 vmESS W175, %guiESSLIST%
	Gui, 2: Add, Button, X173 Y36 gSubmit, OK
	Gui, 2: Show, W415, Name ESS
	WinWaitClose, Name ESS
	mESSe := RegExMatch(mESS,"[a-zA-Z]{1,}")
	if (mESSe <> 0)
	  {
		MERGEESSNAME := mESS . ".ess"
		FileRead, mESSINFO, %MERGEESSNAME%
	  }
	if (A_Index = 1)
	{
FileAppend,
(
<?xml version="1.0" encoding="iso-8859-15" standalone="no"?>
<ccadie:dacConfiguration xmlns:ccadie="http://com/ccadllc/dac/importexport/ccadie" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" schemaVersion="0.1" xsi:schemaLocation="http://com/ccadllc/dac/importexport/ccadie schemas/dacConfiguration.xsd">
<externalServiceSets type="ExternalServiceSet">`n
), dacConfiguration_new.xml
	}

FileAppend,
(
<externalServiceSet externalServiceSetName="%ESSNAME%">
<externalServiceSetEntries>`n
), dacConfiguration_new.xml
FileAppend, %ESSINFO%, dacConfiguration_new.xml
if (mESSe <> 0)
  FileAppend, `n%mESSINFO%, dacConfiguration_new.xml
FileAppend,
(
</externalServiceSetEntries>
<distributions />
</externalServiceSet>`n
), dacConfiguration_new.xml

FileDelete, %A_LoopField%
}

FileAppend,
(
</externalServiceSets>
</ccadie:dacConfiguration>
), dacConfiguration_new.xml

FileRead, ESSINFO, %XMLFILE%
NewStr := RegExReplace(ESSINFO, "<externalServiceSetEntry.*?/>","",oCOUNT)
FileRead, ESSINFO, dacConfiguration_new.xml
NewStr := RegExReplace(ESSINFO, "<externalServiceSetEntry.*?/>","",nCOUNT)

FileDelete, *.ess
diff("Added.txt","Removed.txt")
msgbox, ,ESS Entries, External Service Set Entries in each file`n`ndacConfiguration.xml: %oCount%`ndacConfiguration_new.xml: %nCount%`nEntries Removed: %REMOVED%`nEntries Added: %ADDED%`n`n4 files have been created`nAdded.txt - All entries that were added to the ESS`ndacConfiguration_new.xml - The new file to import`nIssues.txt - Entries that need investigation`nRemoved.txt - All the entries that were removed from the ESS
ExitApp
return

Submit:
Gui, 2: Submit
Gui, 2: Destroy
return

GetFile:
FileSelectFile, XMLFILE, 1, , Select the current dacConfiguration.xml, XML Files (*.xml)
GuiControl, ,XMLFILE, %XMLFILE%
return

class Service
{
	SourceID := ""
	SourceName := ""
	ProviderName := ""
}

CleanESS(ESSIN, SOURCEN, PROVN)
{
 SKIP := 0
 ESSTEMP := RegExReplace(ESSIN,"\.ess","")
 ESSTEMP := ESSTEMP . ".tmp"
 ENTRY := "<externalServiceSetEntry modulationMode=""QAM_256"" mpegServiceNumber=""[0-9]*"" serviceProviderName=""" . PROVN . """ sourceName=""\Q" . SOURCEN . "\E"" sourceUserId="""" tunedChannel=""A[0-9]*""/>"
 Loop, Read, %ESSIN%, %ESSTEMP%
 {
  FoundPos := RegExMatch(A_LoopReadLine,ENTRY)
  if (FoundPos = 0)
   {	  
	if (A_Index = 1)
	  FileAppend, %A_LoopReadLine%
	else
	  FileAppend, `n%A_LoopReadLine%
   }
  else
   {
	SKIP := 1
	FileAppend, %A_LoopReadLine%`n,Removed.txt
	continue
   }
 }
 FileMove, %ESSTEMP%, %ESSIN%, 1
 return SKIP
}

diff(FILE1, FILE2)
{
 FileRead, FILEIN, %FILE2%
 Loop, Read, %FILE1%, Issues.txt
 {
	FoundPos := RegExMatch(A_LoopReadLine,"sourceName=""(.*?)""",SOURCE)
	FoundPos := RegExMatch(A_LoopReadLine,"serviceProviderName=""(.*?)""",PROVIDER)
	SourceName := SOURCE1 	
	ProviderName := PROVIDER1
	ENTRY := "<externalServiceSetEntry modulationMode=""QAM_256"" mpegServiceNumber=""[0-9]*"" serviceProviderName=""" . ProviderName . """ sourceName=""\Q" . SourceName . "\E"" sourceUserId="""" tunedChannel=""A[0-9]*""/>"
	FoundPos := RegExMatch(FILEIN, ENTRY)
	if (FoundPos = 0)
		FileAppend, %A_LoopReadLine%`n
	else
		continue
 }
 return
}

FindSourceName:
Loop % SVCi + 1
{
 if (Services[A_Index].SourceID = SOURCEID)
   {
	SVCNAME := Services[A_Index].SourceName
	PROVNAME := Services[A_Index].ProviderName
	break
   }
 else
	continue
}
return