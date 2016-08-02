Attribute VB_Name = "basConfig"
'------------------------------------------------------------------------
' Description  : contains script configuration options as global constants
'------------------------------------------------------------------------

' Copyright 2016 Matthias Carell
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

'Options
Option Explicit

'if true, script assumes feature names contain aggregates
' expected format is using this scheme <aggregate name> - <feature name>
Global Const cblnGetAggregatesFromFeatureName = True

'if true aggregates are hidden from the drawing unless user decides otherwise
'may be changed by user on the options form
Global Const cblnHideAggregatesDefault = False

'if true every domain box will be drawn on a dedicate worksheet
'may be changed by user on the options form
Global Const cblnDrawDomainsOnSeparatePagesDefault = False

'setting names used in collection objects
Global Const cstrOptionNameHideAggregates = "hideAggregates"
Global Const cstrOptionNameDrawDomainsOnSeparatePages = "drawDomainsOnSeparatePages"

Global Const cstrDefaultDrawingPageName = "feature map"

'reserved tag names
Global Const cstrDomainTag = "d"    'domain tag example for a car domain: @d-car

'distance between drawing and document border
Global Const clngDocPaddingX = 20
Global Const clngDocPaddingY = 20
'distance between cDocPaddingX and domain box (e.g. to place user icons)
Global Const clngDomainPaddingX = 20
'white space around any item (e.g. feature, scenario or aggregate)
Global Const clngItemPaddingX = 20
Global Const clngItemPaddingY = 20
'item size
Global Const clngItemWidth = 200
Global Const clngItemHeight = 90
