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
