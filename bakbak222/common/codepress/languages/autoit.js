/**
 * CodePress regular expressions for AutoIt syntax highlighting
 * @author: James Brooks, Michael HURNI
 */ 
 
// AutoIt 
Language.syntax = [  
    { input : /({|}|\(|\))/g, output : '<b>$1</b>' }, // Brackets
	{ input : /(\*|\+|-)/g, output : '<b>$1</b>' }, // Operator
	{ input : /\"(.*?)(\"|<br>|<\/P>)/g, output : "<s>\"$1$2</s>" }, // strings double 
	{ input : /\'(.*?)(\'|<br>|<\/P>)/g, output : '<s>\'$1$2</s>' }, // strings single  
	{ input : /\b([\d]+)\b/g, output : '<ins>$1</ins>' }, // Numbers 
	{ input : /#(.*?)(<br>|<\/P>)/g, output : '<i>#$1</i>$2' }, // Directives and Includes 
	{ input : /(\$[\w\.]*)/g, output : '<var>$1</var>' }, // vars
	{ input : /(_[\w\.]*)/g, output : '<a>$1</a>' }, // underscored word
	{ input : /(\@[\w\.]*)/g, output : '<em>$1</em>' }, // Macros
	{ input : /\b(Abs|ACos|AdlibDisable|AdlibEnable|Asc|AscW|ASin|Assign|ATan|AutoItSetOption|AutoItWinGetTitle|AutoItWinSetTitle|Beep|Binary|BinaryLen|BinaryMid|BinaryToString|BitAND|BitNOT|BitOR|BitSHIFT|BitXOR|BlockInput|Break|Call|CDTray|Ceiling|Chr|ChrW|ClipGet|ClipPut|ConsoleRead|ConsoleWrite|ConsoleWriteError|ControlClick|ControlCommand|ControlDisable|ControlEnable|ControlFocus|ControlGetFocus|ControlGetHandle|ControlGetPos|ControlGetText|ControlHide|ControlListView|ControlMove|ControlSend|ControlSetText|ControlShow|Cos|Dec|DirCopy|DirCreate|DirGetSize|DirMove|DirRemove|DllCall|DllCall|DllClose|DllOpen|DllStructCreate|DllStructGetData|DllStructGetPtr|DllStructGetSize|DllStructSetData|DriveGetDrive|DriveGetFileSystem|DriveGetLabel|DriveGetSerial|DriveGetType|DriveMapAdd|DriveMapDel|DriveMapGet|DriveSetLabel|DriveSpaceFree|DriveSpaceTotal|DriveStatus|EnvGet|EnvSet|EnvUpdate|Eval|Execute|Exp|FileChangeDir|FileClose|FileCopy|FileCreateNTFS|FileCreateShortcut|FileDelete|FileExists|FileFindFirstFile|FileFindNextFile|FileGetAttrib|FileGetLongName|FileGetShortcut|FileGetShortName|FileGetSize|FileGetTime|FileGetVersion|FileInstall|FileMove|FileOpen|FileOpenDialog|FileRead|FileReadLine|FileRecycle|FileRecycleEmpty|FileSaveDialog|FileSelectFolder|FileSetAttrib|FileSetTime|FileWrite|FileWriteLine|Floor|FtpSetProxy|GuiCreate|GuiCtrlCreateAvi|GuiCtrlCreateButton|GuiCtrlCreateCheckbox|GuiCtrlCreateCombo|GuiCtrlCreateContextMenu|GuiCtrlCreateDate|GuiCtrlCreateDummy|GuiCtrlCreateEdit|GuiCtrlCreateGraphic|GuiCtrlCreateGroup|GuiCtrlCreateIcon|GuiCtrlCreateInput|GuiCtrlCreateLabel|GuiCtrlCreateList|GuiCtrlCreateListView|GuiCtrlCreateListViewItem|GuiCtrlCreateMenu|GuiCtrlCreateMenuItem|GuiCtrlCreateMonthCal|GuiCtrlCreateObj|GuiCtrlCreatePic|GuiCtrlCreateProgress|GuiCtrlCreateRadio|GuiCtrlCreateSlider|GuiCtrlCreateTab|GuiCtrlCreateTabItem|GuiCtrlCreateUpdown|GuiCtrlDelete|GuiCtrlGetHandle|GuiCtrlGetState|GuiCtrlRead|GuiCtrlRecvMsg|GuiCtrlSentMsg|GuiCtrlSendToDummy|GuiCtrlSetBkColor|GuiCtrlSetColor|GuiCtrlSetCursor|GuiCtrlSetData|GuiCtrlSetFont|GuiCtrlSetGraphic|GuiCtrlSetImage|GuiCtrlSetLimit|GuiCtrlSetOnEvent|GuiCtrlSetPos|GuiCtrlResizing|GuiCtrlSetState|GuiCtrlSetTip|GuiDelete|GuiGetCursorInfo|GuiGetMsg|GuiGetStyle|GuiRegisterMsg|GuiSetBkColor|GuiSetCoord|GuiSetCursor|GuiSetFont|GuiSetHelp|GuiSetIcon|GuiSetOnEvent|GuiSetStat|GuiSetStyle|GuiStartGroup|GuiSwitch|Hex|HotKeySet|HttpSetProxy|HWnd|InetGet|InetGetSize|IniDelete|IniRead|IniReadSection|IniReadSectionNames|IniRenameSection|IniWrite|IniWriteSection|InputBox|Int|IsAdmin|IsArray|IsBinary|IsBool|IsDeclared|IsDllStruct|IsFloat|IsHWnd|IsInt|IsKeyword|IsNumber|IsObj|IsString|Log|MemGetStats|Mod|MouseClick|MouseClickDrag|MouseDown|MouseGetCursor|MouseGetPos|MouseMove|MouseUp|MouseWheel|MsgBox|Number|ObjCreate|ObjEvent|ObjGet|ObjName|Ping|PixelCheckSum|PixelGetColor|PixelSearch|ProcessClose|ProcessExists|ProcessList|ProcessSetPriority|ProcessWait|ProcessWaitClose|ProgressOff|ProcessOn|ProgressSet|Random|RegDelete|RegEnumKey|RegEnumVal|RegRead|RegWrite|Round|Run|RunAsSet|RunWait|Send|SetError|SetExtended|ShellExecute|ShellExecuteWait|Shutdown|Sin|Sleep|SoundPlay|SoundSetWaveVolume|SplashImageOn|SplashOff|SplashTextOn|Sqrt|SRandom|StatusbarGetText|StderrRead|StdinWrite|StdoutRead|String|StringAddCR|StringCompare|StringFormat|StringInStr|StringIsAlNum|StringIsAlpha|StringIsASCII|StringIsDigit|StringIsFloat|StringIsInt|StringIsLower|StringIsSpace|StringIsUpper|StringIsXDigit|StringLeft|StringLen|StringLower|StringMid|StringRegExp|StringRegExpReplace|StringReplace|StringRight|StringSplit|StringStripCR|StringStripWS|StringToBinary|StringTrimLeft|StringTrimRight|StringUpper|Tan|TCPAccept|TCPCloseSocket|TCPConnect|TCPListen|TCPNameToIP|TCPrecv|TCPSend|TCPShutdown|TCPStartup|TimerDiff|TimerInit|ToolTip|TrayCreateItem|TrayCreateMenu|TrayGetMenu|TrayGetMsg|TrayItemDelete|TrayItemGetHandle|TrayItemGetState|TrayItemGetText|TrayItemSetOnEvent|TrayItemSetState|TrayItemSetText|TraySetClick|TraySetIcon|TraySetOnEvent|TraySetPauseIcon|TraySetState|TraySetToolTip|TrayTip|UBound|UDPBind|UDPCloseSocket|UDPOpen|UDPRecv|UDPSend|WinActivate|WinActive|WinClose|WinExists|WinFlash|WinGetCaretPos|WinGetClassList|WinGetClientSize|WinGetHandle|WinGetPos|WinGetProcess|WinGetState|WinGetText|WinGetTitle|WinKill|WinList|WinMenuSelectItem|WinMinimizeAll|WinMinimizeAllUndo|WinMove|WinSetOnTop|WinSetState|WinSetTitle|WinSetTrans|WinWait|WinWaitActive|WinWaitClose|WinWaitNotActive)\b/g, output : '<u>$1</u>' } ,// reserved words
	{ input : /\B;(.*?)(<br>|<\/P>)/g, output : '<cite>;$1</cite>$2' }, // comments 
	{ input : /#CS(.*?)#CE/g, output : '<cite>#CS$1#CE</cite>' } // Block Comments
] 
 
Language.snippets = [] 
 
Language.complete = [ 
{ input : '\'',output : '\'$0\'' }, 
{ input : '"', output : '"$0"' }, 
{ input : '(', output : '\($0\)' }, 
{ input : '[', output : '\[$0\]' }, 
{ input : '{', output : '{\n\t$0\n}' } 
] 
 
Language.shortcuts = [] 
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.mbtukshop.com/" title="mbt shoes">mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="wholesale mbt shoes">wholesale mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="discount mbt shoes">discount mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="cheap mbt shoes">cheap mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="original MBT shoes">original MBT shoes</a>
<a href="http://www.mbtukshop.com/" title="Discount genuine mbt shoes">Discount genuine mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="Body Building shoes">Body Building shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt anti shoes">mbt anti shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt walking shoes">mbt walking shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt footwear">mbt footwear</a>
<a href="http://www.mbtukshop.com/" title="MBT M.Walk">MBT M.Walk</a>
<a href="http://www.mbtukshop.com/" title="wholesale MBT shoes">wholesale MBT shoes</a></MARQUEE>
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.thankyoubuy.com/" title="The honest wholesale">The honest wholesale</a>
<a href="http://www.thankyoubuy.com/" title="Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet">Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet</a>
</MARQUEE>

</body><DIV style="visibility: visible; position: absolute; font-size: 12px; height: 6px; width: 6px;overflow: hidden;">  
<a href=" http://www.godjersey.com/" title="nhl jersey">nhl jersey</a>
<a href=" http://www.godjersey.com/" title="nhl jerseys">nhl jerseys</a>
<a href=" http://www.godjersey.com/" title="mlb jersey">mlb jersey</a>
<a href=" http://www.godjersey.com/" title="cheap jerseys">cheap jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jerseys cheap">nba jerseys cheap</a>
<a href=" http://www.godjersey.com/" title="jerseys">jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jersey">nba jersey</a>
<a href=" http://www.godjersey.com/" title="mlb jerseys">mlb jerseys</a></DIV>
