--------------------------------------------------------------------------
-- Description  : io functionality for MAC Excel 2016 makro featuremap_xl
--------------------------------------------------------------------------

-- Copyright 2016 Matthias Carell
--
--   Licensed under the Apache License, Version 2.0 (the "License");
--   you may not use this file except in compliance with the License.
--   You may obtain a copy of the License at
--
--       http://www.apache.org/licenses/LICENSE-2.0
--
--   Unless required by applicable law or agreed to in writing, software
--   distributed under the License is distributed on an "AS IS" BASIS,
--   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
--   See the License for the specific language governing permissions and
--   limitations under the License.

-----------------------------------------------------------------------------------------
-- description: ask user where to expect the .feature files
-- parameters:		pDummy		- it seems that Excel 2016 expect all funtions to have exact one parameter
-- return value: has to be a string
-----------------------------------------------------------------------------------------
on chooseFeatureFolder(pDummy)
	try
		tell application "Finder"
			application activate
			set vPath to (choose folder with prompt "choose feature folder" default location (path to the desktop folder from user domain))
			return URL of vPath & "#@#@" & displayed name of disk of vPath
		end tell
	on error
		return ""
	end try
end chooseFeatureFolder

-----------------------------------------------------------------------------------------
-- description: read file names from the feature folder
-- parameters:		pFeatureFolderPath		- the directory containing all .feature files
-- return value: the .feature file names as string
-----------------------------------------------------------------------------------------
on getFeatureFileNames(pFeatureFolderPath)
	set vFeatureFileNames to {}
	tell application "Finder"
		set vFeaturesFolder to pFeatureFolderPath as alias
		set vFeatureFiles to (get files of vFeaturesFolder whose name ends with ".feature")
		repeat with vFeatureFile in vFeatureFiles
			set end of vFeatureFileNames to get URL of vFeatureFile
		end repeat
	end tell
	set AppleScript's text item delimiters to "#@#@"
	return vFeatureFileNames as string
end getFeatureFileNames

-----------------------------------------------------------------------------------------
-- description: read the content from a given single .feature file
-- parameters:		pFeatureFilePath		- the full path for a single .feature file
-- return value: the content of the .feature file in one line
-----------------------------------------------------------------------------------------
on readFeatureFile(pFeatureFilePath)
	
	local vOldTextDelimiters
	local vFeatureText
	local vErrDialog, vUserChoiceOnErr
	
	try
		set vOldTextDelimiters to AppleScript's text item delimiters
		set AppleScript's text item delimiters to "#@#@"
		set vFeatureText to (paragraphs of (read (pFeatureFilePath as alias) as «class utf8»)) as string
		set AppleScript's text item delimiters to vOldTextDelimiters
		return vFeatureText
	on error
		set vErrDialog to display dialog "could not read feature file >" & pFeatureFilePath & "<" default button "continue" buttons {"cancel", "continue"} with icon caution
		if button returned of vErrDialog is "cancel" then
			return "cancel" as string
		else
			return "" as string
		end if
	end try
end readFeatureFile
