# A friendly helper DLL that will make you smile.
**No installation**, **no ActiveX**, **no Admin-Rights.** 
Just add this Dll to your VBA project folder and have some cool UI features. Have only tested in MS Access but it should work in all VBA environment. Works with ACCDE too.

The main goal is to bring some .NET function to your VBA project. Make your project stand-out visually and functionally!
Since this plugin does not require admin rights nor installation, you can distribute your app without having to worry about your client's admin policy.

And of course with minimal code!

```diff
- NOTE:
- If you get an error regarding unable to load find/load dlls. Please make sure you have the bin folder where your project is AND right click on vba_tools.dll => property => unblock. Do the same for all dlls in the bin folder.
- this is an evolving project. Function names from one version to another might varry, please test your wrappers before updating to the newest one.

```

# Be safe
Use following sites to check for any malware for any files you download from online.
https://virusdesk.kaspersky.com/
https://www.virustotal.com/

![OnlineScanner](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/vbatoolsIsSafe.png)

# Unblock the Dlls if necessary.
![UnblockADllPicture](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/unblockADll1.png)

## Current progress / bugfixings
```VBA
	+ All functions are grouped under appropriate class names.  i.e Dll.String.AllStringRelatedFunctions
	- Make sure, you have update all your previous codings
	
	+ Make any Window transparent: Make either entire window as transparent or set your key colour and only make those pixels transparent.
	+ IsKeyDown(keyCode): returns true or false, whether a given key is pressed or not. This can be used in any event and useful to check whether CTRL or Shift key is pressed.
	
	+ object ExecuteScalar : (DatabaseType dbType, string connectionString, string sql) :Executes and SQL and returns first row first column of the result set or an empty string
	+ bool ExecuteNonQuery(DatabaseType dbType, string connectionString, string sql)	:Returns true or false if the nonQuery sql command was success
	+ string MySqlGetAvailableServerFromList(string[] connectionString)					:Takes array of connetionstrings and returns first reachable connectionstring. or ""
	+ bool MySqlServerIsReachable(string connectionString)								:Connects to a MySql server using the connectionstring and returns true or false
```


## What it does?
Helps you to make your application more user-friendly by providing some .NET components and functions that you can use within your VBA application. Visually and functionally cooler than VBA!

## How to use?
Some basic VBA skills are required! 
Just download the <a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> Dll + bin folder </a> from the sample folder and add them to your project folder. The ACCDB contains samples where you can copy and paste it to your VBA application.

**Keep the VBA_TOOLS.Dll & Bin folder where your vba project is to enjoy all functions.**



# Interesting List of features
(Update hasn't been released with these new functions)
<ul>
	<li>ContextMenu
		<ul>
			<li>Show Cool Context Menu</li>
		</ul>
	</li>
	<li>Barcode Control
		<ul>
			<li>Currently supports Code39, Code128, QrCode</li>
			<li>Currently under testing</li>
		</ul>
	</li>
	<li>Colour
		<ul>
			<li>HexToAccess</li>
			<li>AccessToHex</li>
		</ul>
	</li>
	<li>DB Class: (Only supported MySQL)
		<ul>
			<li>ExecuteScalar	: executes a MySql select query and returns first row first column object</li>
			<li>ExecuteNonQuery	: executes a MySql update/insert query and returns true or false if success.</li>
			<li>string MySqlGetAvailableServerFromList(string[] connectionString)</li>
				<ul>Takes array of connetionstrings and returns first reachable connectionstring. or "".<br/>Handy if you have multiple back-end servers and would like to know which one is reachable. Uses threads so it's fast!</ul>
			<li>bool MySqlServerIsReachable(string connectionString)</li>
				<ul>:Connects to a MySql server using the connectionstring and returns true or false</ul>
		</ul>
	</li>
	<li>Dialog Boxes
		<ul>
			<li>Cool DialogBox</li>
			<li>Extended Cool DialogBox</li>
			<li>Cool Simple DialogBox</li>
			<li>Simple Are you sure?</li>
			<li>Drag and Drop OpenFileDialog
			<ul>
				<li>A simple open file dialog box that supports drag </li>
			</ul>
			</li>
		</ul>
	</li>
	<li>Display
	<ul>
		<li>GetNumberOfMonitors
		<ul>
			<li>Returns number of monitors </li>
		</ul>
		</li>
	</ul>
	</li>
	<li>
	<ul>
		<li>GetPrimaryMonitorHandle
		<ul>
			<li>Returns handle to primary monitor</li>
		</ul>
		</li>
		<li>GetPrimaryMonitorBounds
		<ul>
			<li>Get information about the primary monitor</li>
		</ul>
		</li>
		<li>GetMonitorBoundsByHandle
		<ul>
			<li>Returns monitor information from a window handle.</li>
		</ul>
		</li>
		<li>GetCursorPosition</li>
	</ul>
	</li>
	<li>Form
	<ul>
		<li>Make Transparent Window</li>
		<li>Change Access background colour</li>
		<li>Drage Me
		<ul>
			<li>Allows to move a borderless form</li>
		</ul>
		</li>
	</ul>
	</li>
	<li>FTP
	<ul>
		<li>FTPS_UPLOAD
		<ul>
			<li>uploads file to given ftp server</li>
		</ul>
		</li>
		<li>FTPDeleteFile
		<ul>
			<li>deletes a file from given ftp location</li>
		</ul>
		</li>
		<li>FTPFileExists
		<ul>
			<li>Checks if a given file exists in given ftp location</li>
		</ul>
		</li>
	</ul>
	</li>
	<li></li>
	<li></li>
	<li>Graphics
	<ul>
		<li>Load Picture from URL to ImageControl without saving</li>
		<li>Load web urls</li>
		<li>Load local pictures</li>
		<li>Convert blob fields to pictures</li>
		<li>MeasureText
		<ul>
			<li>Calculates height and width in pixels for a given text</li>
		</ul>
		</li>
		<li>ByteToImage
		<ul>
			<li>Converts ByteArray to an Image and saves in a provided location.</li>
		</ul>
		</li>
		<li>ByteToBitmap
		<ul>
			<li>Converts Byte Array to a bitmap</li>
		</ul>
		</li>
		<li>TakeScreenShotFromHwnd
		<ul>
			<li>Takes a screenshot of a window by provided handle. Returns byte array </li>
		</ul>
		</li>
		<li>TakeScreenShot
		<ul>
			<li>Take screenshot of entire desktop. Returns byte array</li>
		</ul>
		</li>
		<li>TakeScreenShot1
		<ul>
			<li>Takes screenshot and saves in a given path</li>
		</ul>
		</li>
		<li>PictureFromUrl
		<ul>
			<li>Loads picture from an url and returns byte array</li>
		</ul>
		</li>
		<li>SaveClipboardToImage
		<ul>
			<li>Saves clipboard picture to a given path and format</li>
		</ul>
		</li>
	</ul>
	</li>
	<li>InputBoxes
	<ul>
		<li>Show Cool InputBox</li>
		<li>Email with validation</li>
		<li>Password</li>
		<li>Multiline / single line</li>
		<li>Number only</li>
		<li>Dates with validation</li>
		<li>Show DropDown box</li>
	</ul>
	</li>
	<li>JSON
	<ul>
		<li>Uses Newtonsoft.Json</li>
		<li>ExportToJSON
		<ul>
			<li>Allows MS Access users you to export queries, tables SQL results as JSON string</li>
		</ul>
		</li>
		<li>ImportJSON
		<ul>
			<li>Allows MS Access users to import records to table using JSON string arrays</li>
		</ul>
		</li>
		<li>JSONString
		<ul>
			<li>String expression of a Newtonsoft.Json object</li>
		</ul>
		</li>
		<li>JSONGetObject
		<ul>
			<li>retireve Json object by property name</li>
		</ul>
		</li>
		<li>JSONSetObject
		<ul>
			<li>Adds a property to Json object</li>
		</ul>
		</li>
		<li>JSONGetValue
		<ul>
			<li>Retrieves a value from json object</li>
		</ul>
		</li>
		<li>JSONToObject
		<ul>
			<li>Converts string json to a dynamic object</li>
		</ul>
		</li>
		<li>JSONSerialize
		<ul>
			<li>Returns string expression of a dynamic json object</li>
		</ul>
		</li>
	</ul>
	</li>
	<li>NetClass
	<ul>
		<li>UrlIsReachable
		<ul>
			<li>Returns true or false whether the url is reachable</li>
		</ul>
		</li>
		<li>UrlIsValid
		<ul>
			<li>Returns true or false whether the url is well formatted</li>
		</ul>
		</li>
		<li>UrlIsLocalPath
		<ul>
			<li>Is the given url a local file path?</li>
		</ul>
		</li>
		<li>GetExternalIP</li>
	</ul>
	</li>
	<li>Notification
	<ul>
		<li>Show Non-Blocking Notifications</li>
		<li>Show Success</li>
		<li>Show Warnings</li>
		<li>Show Error</li>
	</ul>
	</li>
	<li>ProgressBar
	<ul>
		<li>Show Cool ProgressBar</li>
	</ul>
	</li>
	<li>RegEx
	<ul>
		<li>IsMatch</li>
		<li>GetFirstMatch</li>
		<li>Replace</li>
	</ul>
	</li>
	<li>String
	<ul>
		<li>PadLeft</li>
		<li>PadRight</li>
		<li>TrimEnd</li>
		<li>TrimStart</li>
		<li>StartsWith</li>
		<li>EndsWith</li>
		<li>.NET string.format</li>
	</ul>
	</li>
	<li></li>
</ul>


### [Show non-blocking notifications]
Inspired from Toastr (https://github.com/CodeSeven/toastr).
Allowing VBA users to show simple notifications without having to wait or stress their VBA application.
With a simple command a little colourful notification pops up with a message without taking any focus or disturbing the user.
I mainly use it to show messages that do not require action. I.e. A mail has arrived or a task has been completed.

![just a notification](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/information.png)

## customise your notification like you want:
following customisations are possible now.
```
1.Message   : can contain <a href="">text</a> for hyperlinks (any other html tags are ignored, hyperlink must begin with www or http or https (basically well formatted links only?)
```diff
+ It is now possible to open local files. hyperlinks must be a local file format. I.e. <a href="F:\folderName\picture.png">
+ Testing: A call back command (Docmd.OpenForm only) can be embedded into the hyperlink
```
2.Duration in Milli-Seconds (default 2000. 0 will keep the notification for long time.  int.max)
3.Background colour (html colour code)
4.Font colour (html colour code)
5.X,Y position on the desktop
```



![picture of 3 notifications](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-collections.png)
```VBA
'used commands
Toastr.Toast "Ups something went wrong!",vberror,0
Toastr.Toast "Yellow weather warning!",vbexclamation,0
Toastr.Toast "You've just received a notification",vbinformation,0
```

in Action
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InAction.gif)
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InAction1.gif)

## how about little interaction with user and show some hyperlinks?
You can have html ```<a href="">text</a>``` tags in your message which will be translated into hyperlinks.
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/Hyperlink.png)

Local files as hyperlinks: ```<a href="F:\folderName\picture.png">ViewThisImage</a>```

Callback command:
1. OpenForm hyperlink: ```<a href="DoCmd.OpenForm frmImageView,acNormal,,wherecondition:=id=2">OpenForm</a>```
```VBA
	Note: the docmd command does not contains any " or '
	Filter, WhereCondition, DataMode, WindowMode must be named parameters. I.e. Filter:=FilterCondition or WhereCondition:=id=2
	
	'Similarly you may also pass a function name which will be executed to the host application
	<a href="ExecuteMe()"> Execute a function in the host application </a>
	
	
		
	[Toast Notifications]
	1. Toast notification parsing hyperlink function corrected
	2. Toast can now open Docmd.OpenForm
	3. Toast can now execute local functions. I.e. <a href="ExecuteMe()"> Execute a function in the host application </a> will execute "ExecuteMe()" when clicked the link.
	4. Toast / Simple Dialogboxes can now execute local functions with parameters i.e. <a href="ExecuteMe('ParameterA','ParameterB')"> ExecuteMe </a>
	4. Toast / Simple  Dialogboxes can now close itself after clicking a hyperlink. Use closeme="true" attribute. i.e. <a href="ExecuteMe('ParameterA','ParameterB')" closeme="true"> Execute and close Me </a>
	
	
```



## Download 
Download the sample and test it in your project. Please leave comment how you feel.
<a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> Samples</a>


<hr>
<hr>

# Show Cool DialogBox
Standard Message boxes are great but sometimes you want little more than standard features.
I.e
<ul>
  <li>Be able to have some colours</li>
  <li>Be able to have more than 3 buttons</li>
  <li>Be able auto-close</li>
  <li>Be able to use HTML tags </li>
  <li>not stressing your vba app with a loop?</li>
</ul>
Meet the new simplified DialogBox for VBA users. This dialogbox will allow above listed features and should help you to keep your application colourful. :) This feature is still under development and could some feedback from testers.











<HR>

![Cool DialogBox](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-DialogboxGreen.png)
![Cool DialogBox1](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-DIALOG-BOX.png)

There is vba wrapper in the sample accdb which can be extended as per your need. It uses the 3rd party JSON Converter plugin with some miner fixes from my side.

```
  'usign the wrapper it would be as simple as 
  Debug.Print gDll.DialogRich("This is a title", "Some content", (vbExclamation + vbYesNo))
```

a simplified version is also avilable (without HTML rendering)
# Cool Simple MessageBox
![Cool DialogBox](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-CoolSimpleMessageBox.png)

Allows one to show simple message box


# Show cool Progressbar
Progressbars 

crucial element when informing users about a progress. Meet the cool progressbar which can pop up on top of your application at any time with a simple code as such as.

```
  Dim ProgressBarID As Long
  ProgressBarID = gDll.ShowProgressBar(100, "Executing your query", "Please wait. We are preparing printer drivers")
    
  ProgressBarID = gDll.SetProgressBar(ProgressBarID, 10, "Waiting for driver..")
  
  gdll.CloseProgressbar ProgressbarId 'Will close the progressbar
```
![Cool ProgressbarGreen](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-ProgressBar.png)

As usual, you are allowed to change theme colours as per your taste.
![Cool ProgressbarRed](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-ProgressBarRed.png)

### note:
```ShowProgressBar and SetProgressBar``` returns an ID which you can refer your progressbar to. This also allows VBA users to have multiple progressbars at the same time. 

# Show Cool InputBoxes
InputBox another heavily used component. Some like the plain system looking InputBox but we love the modern UI colours :)
What would you chose from these tables?

![InputBoxCollection](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/InputBoxDefault.png)  ![InputBoxCollection](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-InputBoxMultiline.png) 

## Nice colours! but what's the point?
The new InputBoxes comes with some inbuilt functions and can be configured accordingly.
Following types are supported now.
```
'        Password        = 1, : Masked using systempassword mask
'        Text            = 2, : Single line text:
'        MultilineText   = 32, : Multi line text box
'        Number          = 4, : Numbers only
'        ShortDate       = 8, : Masked dd/mm/yyyy. Dates are validated upon exit
'        LongDate        = 16,  : masked using dd/Month/yyyy
'        DateTime        = 48,  : masked using dd/mm/yyyy hh:mm:ss

and following parameters are accepted: 
  Except Type, all others are optional
  
  InputBoxType Type,    : number
  string Title,         : Tile for the input box
  string Message,       : optional text for the input box
  int PosX,             : x coordinate relative to the screen to positon this box to
  int PosY,             : y coordinate relative to the screen to position this box to
  string ThemeBg,       : html colour code
  string ThemeForeColour: html colour code

' With the dll in place, use it as

  result = gDll.DLL.showinputbox(Type:=32, Title:="", Message:="Tell us what happened on that day!", ThemeBg:="", ThemeForeColour:="")
```
#### check out the getCursorPosition function which returns x,y position of the cursor!


in action:

![InputBoxCollection](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/InputBox.png)

as always we can change theme colours:)

![purple input box](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-InputBoxPurple.png)

Download <a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> sample</a>

# [Show DropDown Box]
Requested from a VBA_TOOLS user. Like other user inputboxes, you can now let your user to select from a cool dropdown box.
![purple input box](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/DropDownBox.png)

When we decided to make a cool dropdown box, we thought about reasons for not using existing Ms Access DropDown box.
I personally think, showing icons in a dropdown box would be an amazing idea! :) In addition, standard DropDownBox does not allow one to search partially within the content. That is, being able to search any part of the dropdown selection.
Like traditional DropDown, we want to show a list of existing query or table.
So we decided to cover those points at the moment. Surely, going forward we are planing to add following functions.

1. Grouped entries: DropDown entries will be grouped with a group header.
2. Having multiple image column?
3. Change the DropDown style completely: Maybe like a menu with sub menus..
4. Fix any errors


Here is how to use it:
```VBA
 ?gDll.ShowDropDown("Select an item", "Some inner message?", "qryDropDown", 2, Array(50, 50))
 
 
 Param list:
 /// <summary>
  /// Shows a Dialogbox with a dropdown for selection. Returns a string value
  /// </summary>
  /// <param name="title">Title for the inputbox</param>
  /// <param name="message">Inner message for the input box</param>
  /// <param name="dbSource">Database path</param>
  /// <param name="tableSource">Table name or SQL. If SQL is used, use isRawSql=true</param>
  /// <param name="boundColumn">Column Index to get the value from</param>
  /// <param name="columnWidths">an Array of integers</param>
  /// <param name="isRawSql">Specifies whether the tablesource is a plain SQL command </param>
  /// <param name="posX"></param>
  /// <param name="posY"></param>
  /// <param name="themeColour"></param>
  /// <param name="themeForeColour"></param>
  /// <returns>String value</returns>
```


### Note:
If your datasource contains "icon" as the first column and it is a hyperlink (web or local file), by default those links will be converted to image.
Usse Array(column0_width, column1_width ...) to set up the column widths


<hr>

# Drag and drop OpenFileDialog
WHAT!! Drag and drop function for vba??? Yes you've read it correct but don't get too excited though:) It's just a file-drop function. Allowing users to select/open/get files using drag and drop method. Direct alternative to an existing FileOpenDialog method. 
<hr>
  
### returns a string of JSON Array with all selected files. (if you wish to have string array see below)
What you do with those file paths is up to you. Maybe at some point later, we might link this with our existing FTP component.


Currently following parameters are accepted:

```c#
  All of below are optional.
  
  string Message,         : A message for the dialog box.
  bool AllowMulti,        : Should it allow multiple files?
  string[] Filters,       : An array of string => (Description |*.png). Used for file extention filters
  int PosX,               : X Position relative to the monitor where this box should appear
  int PosY,               : Y position relative to your monitor where this box should appear
  string ThemeBg,         : HtmlColourCode
  string ThemeForeColour  : HtmlColourCode
```
(Assuming the Dll part is already done:) Use in VBA like this:
or just download the sample file and look what functions you would copy to your application.

```
    Dim FilePaths As String
    FilePaths = gDll.DLL.ShowDialogForFile("No multiple files allowed", False)
```

or customised one:
```VBA
    Dim Filters(2) As String
    
    Filters(0) = "Png Pictures only |*.png"
    Filters(1) = "All files |*.*"
    
    Dim FilePaths As String
    FilePaths = gDll.DLL.ShowDialogForFile(Message:="Feel free to drop many files", allowmulti:=False, Filters:=Filters, PosX:=0, PosY:=0, ThemeBg:="", ThemeForeColour:="")
    
```
If you wish to have a string array result
```
    dim Files() as string
    Files = gDll.ShowDialogForFileArray(Message:="Feel free to drop many files", allowmulti:=False, Filters:=Filters, PosX:=0, PosY:=0, ThemeBg:="", ThemeForeColour:="")
    'Will return a string array of selected files
```

View in action:
![File drag and drop gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/FileDropInAction.gif)

Errors
![File drag and drop error gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/VBA-RICH-UI-DRAG-DROP.gif)
<hr>






# Load Picture from URL to ImageControl without saving
Oh wow! how many people wished this was possible out-of-the-box? Many of us spent good amount of time searching for good tutorials and most the results are simple wayarounds than solutions. Pages after pages of codes with APIs and classes or use web-browser control, buy third-party image control or download the picture and load again.

No offence to the web-browser control. It is great for what it is but surely not designed for showing images(IMHO). Functions like, zooming, streching aren't available via web-browser control. Of course you can use HTML tags but that would be a "way around" to another "way around" problem. isn't it?

Don't want to buy third party controls because they need to be installed! (no-go for many)
Don't want to download and load either. Too much footprint/mess to clean up with.

Let's meet our simple one liner which can load images into an Image control. No download, no too much code, no nonsense

```VBA
  'Dll function
  'PictureFromUrl(
    string URL,             :  Image url. web url or local path
    bool ShowError = false, : Show error notification when url cannot be loaded
    long sender = 0         : Sender HWND, not used now.
    )
  
  'VBA Wrapper (used for simplicity)
  'ImageControlGetImage(ImagePath as string, optional ShowError=true)
  
  
'Loading web url
Private Sub Command147_Click()
    Dim WebPicture As String
    WebPicture = "https://avatars2.githubusercontent.com/u/1001697?s=460&v=4"
    
    Me.Image113.PictureData = gDll.ImageControlGetImage(WebPicture, ShowError:=True)
End Sub

'Same function used to load local file path
Private Sub Command149_Click()
    Dim WebPicture As String
    WebPicture = "F:\Projects\VBA_DLL\dialogboxgreen.png"
    
    Me.Image113.PictureData = gDll.ImageControlGetImage(WebPicture, ShowError:=True)
    
End Sub

```
See it in action:
![Image from web url](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/ImageControlInAction.gif)

### If you would like to read urls from your table
instead using the `control source` property, use the `on current` event in your form to load the pictures.
```VBA
Private Sub Form_Current()
  'Load pictures 
    Me.Image8.PictureData = gDll.ImageControlGetImage([url], True)
End Sub
```
Enjoy and let us know what you think!.


# Barcode Control for vba
Another request from Vba_tools user to be able to show barcodes. I have no idea about barcodes but found a great source in google (https://sourceforge.net/projects/zintnet/). Thanks for the zintnet owner.
I've adapted few classes and added to our VBA_TOOLS plugin.

Unlike other components barcode-control will be embedded into forms and reports so the control cannot be a standalone form so when printing reports or invoices the barcode is visible. To achieve this, we create a barcode in .NET environment and pass the barcode back to Access as an Image. This way an Image control on a form or report can show barcodes.

Again this is beta version. Have a look and inform us about your thoughts.

How to use it?

```VBA
  Me.imgBarcode.PictureData = gDll.CreateBarcode(Val(Me.BrcodeType.value), Me.txtBarcodeData.value, Val(BarcodeSizeMultiplier.value))
  
  Parameter list:
  '    /// <summary>
'    ///
'    /// </summary>
'    /// <param name="symbology">Type of the barcode</param>
'    /// <param name="barcodeData">Data value for barcode</param>
'    /// <param name="width">Width of the graphics / Image</param>
'    /// <param name="height">height of the grapics/ Image</param>
'    /// <param name="multiplier">Multiply the size by this value.</param>
'    /// <returns>Picture Data</returns>
'    CreateBarcode(Symbology symbology, string barcodeData, int width, int height, float multiplier )

```
![qrBarcode.png](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/qrBarcode.png)
![Code39Barcode.png](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/Code39barcode.png)

# Cool Context Menu for vba
[Under testing]: testers needed. This function is currently not wrapped. Means more parameters will be added once ready for publish

What can I say? Probably most vba users wished this control existed out of the box. Similar to right click context menu, we developed a left mouse clickable context menu.
Of course styleable, moveable and iconed menu items.

Enough said. lets see in action.
![ContextMenuPicture](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/VBA-RICH-UI-CoolContextMenu.gif)

Since we've got your attention now, let's see how this control works. There are two version
1>Simple: Just takes an array of strings as menu items and returns the menu item text as the selected value.
2> Extended: Extended menu can show icons as well as take a dataValue for each menu item. That being said, each menu item is constructed as an array(string:IconPath, string:DataReturnValue, string:MenuItem) ie. array("c:\email.png","1","Send Email")
add each menu entry within another array to use the extended menu. i.e. Array(array("c:\email.png","1","Send Email"), array("c:\door.png","2","Exit"), array("http://someweblink.png","3","Some menu item with web icon") )

in code that would be:
```VBA
	'Simple Menu
	'Create an array of menu items
    Dim MenuItems() As String
	Dim result As String

    MenuItems = VBA.Split("Do Something,I'm so cool, Send Email, Print, Settings, Save, Save As, Make Pdf", ",")
    FnArrayAddItem MenuItems, "Exit"
    FnArrayAddItem MenuItems, "Exit Application"
    
    result = gDll.DLL.ShowContextMenu(MenuItems)
	gDll.Toast result, , , Me.hwnd
    If (result = "Exit") Then
        DoCmd.Close acForm, Me.Name, acSaveYes
    ElseIf (result = "Exit Application") Then
        Application.Quit
    End If
	
```

```VBA
	'Extended menu with icons
	Dim result As String
    result = (gDll.DLL.ShowContextMenuA(Array(Array("", "0", "Web loading takes time"), Array("F:\PROJECT_SUPPORT\Images\csharp.png", "1", ".NET is cool"), Array("https://static.thinkster.io/topics/node_icon.png", "2", "Loading icon from web"), Array("glyphicons-389-exit", "3", "Exit"))))
    
    
    gDll.Toast result, vbInformation
    
    If (Val(result)) = 3 Then
        DoCmd.Close acForm, Me.Name, acSaveYes
		
```






<hr>
<hr>



# Other Features that are interesting

# DragMe
A simple function that allows one to drag a borderless form.
have a look here. 
![DragME](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/VBA-RICH-UI-DragMe.gif)

How to use?
```VBA
	'simply use the mouseDown event
	
	Private Sub Label251_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
		Call gDll.DLL.DragMe(Me.hwnd)
	End Sub

```

### AreYouSure?
a simple yes no popup returns true or false. Sometims you just want to confirm the user for yes or no action.
This might be a silly but cool yes no box:) 
```VBA
? gDll.AreYouSure

```
It is also possible to colour theme the AreYouSureBox by providing Hex Colour code or if you are using bootstrap class
```VBA
? gDll.AreYouSureE(Me, "#aa66cc", "#000000", "#aa66cc", "#F65656") or 
?gDll.AreYouSureE(, gBootstrap.default_color_dark, gBootstrap.WHITE, gBootstrap.AMBER, gBootstrap.TEAL_LIGHTEN_3)
```
![AreYouSureCollection](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/VBA-RICH-UI-AreYouSureCollection.png)

### Download a file and show progressbar for vba
Another cool feature. This function allows you to download a file from the internet and shows the download progress using above cool progressbar.

```DownloadedFile = DLL.DownloadAFile(Url, [Destination], [OverWrite = true], [ShowProgress = true])```
Except the Url, all other parameters are optional. If destination is not provided. File will be saved in application.path

### Save Clipboard images to local file
Sometimes, simple things can be very dificult in VBA. If you are after saving clipboard image to a local path. Check this function.

``` SaveClipboardToImage(string PathToSave, string FileName, string ImageType) ``` All parameters are optional and by default Jpeg image type is used. If the clipbord object contains any images, it will be saved wherever you want and the file path is returned.
in the sample accdb, there is a wrapper ```SaveClipboardToImage``` check it out.

### PadLeft and PadRight
Uses .NET padleft and padRight function.
``` gdll.DLL.PadLeft("1",10,"0") => 0000000001
    ?gdll.DLL.PadRight("1",10,"0") = > 1000000000
```     
### check out the getCursorPosition function which returns x,y position of the cursor!

### [Receive SignalR Messages]
This works for me because I do have own signalR server but generally is under development or say not ready yet!
It's like google push messages or any other push message service. You can send notification to all of your logged in yours from one place.
Expanding this, you could also use as a chat server where all logged in participants could send and receive messages among them.
Again without stressing VBA apps.


### ByteToImage
ByteToImage(byte[] byteArraym string TemporaryPath, bool useCache) is a function for MS Access users. Basically you can convert a byteARray received from database into a pictures.
Will return the path of the image file. Use the path as image location for your image property.
something like Me.Image32.Picture = gDll.ByteToImage(ByteArray, "SaveLocationPath")

### FTP(S) UPLOAD
simple tool which uses WinScp to upload files securely to your host. Handy if you want to upload some files without doing too much VBA or having activeX components.
```VBA
	'Simply use as 
	Debug.Print gDll.FTPUploadFile(ServerName, port, Username, Password, "F:\Projects\VBA_DLL\Modern Inputbox for vba purple.png", "/screenshots/", SSHFingerprintOfTheRemoteServer, Ftp, Explicit, False)
	
	Parameter list
```
```C#
		/// <summary>
        /// Uoloads a file to the given ftp server
        /// </summary>
        /// <param name="host">Host server</param>
        /// <param name="port">Port number</param>
        /// <param name="username">FTP UserName</param>
        /// <param name="password">FTP password</param>
        /// <param name="localFileName">Path to local file</param>
        /// <param name="remoteLocation">Location in remote server</param>
        /// <param name="hostCertificateFingerprint">Remote Server FingerPrint</param>
        /// <param name="protocol">Ftp Protocol, ftp, sftp...</param>
        /// <param name="ftpSecure">Type of connection, implicit, explicit</param>
        /// <param name="giveUpSecurityAndAcceptAnyTlsHostCertificate">debug use only</param>
        /// <returns>string representation of true or false or error message</returns>
````

### FTP Delete Remote File
Simply delete a file from your remote server. Returns true or false + error message as string.
```VBA
  'Server as string
  'Port as number
  'Username as string
  'Password as string
  'RemoteFile as string
  'SSHFingerprint as string
? DLL.FTPDeleteFile(ServerName, Port, Username, Password, RemoteFile, TLSHostFingerprint)
```

### ImportJSON
Somewhere similar to Application.ImportXML, you can create access tables from JSON array strings.
I haven't done extensive test but works for my needs.

Simply call
```VBA
  gdll.ImportJSON(YourJSonArrayString, "Target Table name", ImportOptions[append,structureOnly,structureAndData], recreate)
 'Recreate will delete and recreate the table. If ApendOnly requested, recreate is ineffective
 
 'Here is a sample working command. Which will create a new table called tblJsonTest and import all the content from the array.
 gdll.ImportJson("[{""id"":10,""name"":""User"",""add"":false,""edit"":true,""authorize"":true,""view"":true},    {""id"":11,""name"":""Group"",""add"":true,""edit"":false,""authorize"":false,""view"":true},    {""id"":12,""name"":""Permission"",""add"":true,""edit"":true,""authorize"":true,""view"":true}]","tblJsonTest",acStructureAndData,True)
  '
```

### ExportToJson
It is now possible to export table contents as JSON string. 
Method1:
```VBA
  'Eecute the SQL SELECT command and saves the result set as JSON formatted string.
  gdll.ExportToJSON("select * from tbljsontest where authorize = true;","MyJson.Txt",overwrite:=false,isRawSql:=true)
```

Method2:
```VBA
  'Export everything from the table/query
  gdll.ExportToJSON("tbljsontest ",SaveAs:= "MyJson.Txt",overwrite:=false,isRawSql:=false)
```
In this method, we have passed a table name/query name to the export function and set isRawSql = false. The export function will then generate SQL statement similar to “SELECT * FROM givenTableName/QueryName;” and perform the JSON Export.

If the SaveAs (target file name) is empty, no file will be exported but the conversion will still happen and converted string will be returned as result.

Download the sample project and have a play.


# [Upcoming functions]
many... :) 
if you want a specific function email or leave a comment :)

# Can't wait? Just download! and enjoy
<a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> sample</a>

# [Copyrights, Licence, Credits]

Copyright © 2018 Krish

You are free to use the dll for non-commercial purposes. Commercial users, you can use the dll with one condition, please let us know who you are. We are very happy to have your/company name in out clients list.

Would appreciate your credits and links to my GitHub page.




<hr>
<hr>
<hr>
# Raw methods from class
<hr>
<hr>

```C#
/// <summary>
/// Shows toasts on the desktop
/// </summary>
public async void FN_SHOW_TOAST(string iMessage, int iDuration, string iBG_COLOR, long iANIME_DURATION, string iFONT_COLOR, int iX, int iY, int iANIM_DIRECTION, bool iAUTO_CLOSE = true)
{
}
 
 

/// <summary>
/// Converts ByteArray to an Image and saves in a provided location.
/// </summary>
/// <returns>The path of the image saved locally</returns>
public string ByteToImage(byte[] byteArrayIn, string iTempPath, bool useCache)
{
}
 
/// <summary>
/// Converts Byte Array to a bitmap
/// </summary>
public Bitmap ByteToBitmap(byte[] byteArr)
{
}
 
/// <summary>
/// Returns a ByteArray of the image
/// </summary>
/// <param name="hWND"></param>
public byte[] TakeScreenShotFromHwnd(long hWND)
{
}
 
/// <summary>
/// Take screen-shot of entire desktop. Returns byteArray
/// </summary>
public byte[] TakeScreenShot()
{
}
 
/// <summary>
/// Take screen-shot of entire desktop. Saves in a location and returns the location
/// </summary>
public string TakeScreenShot1(string SavePath)
{
}
 
/// <summary>
/// Returns ByteArray containing the picture received from the url
/// </summary>
/// <param name="URL"></param>
public byte[] PictureFromUrl(string URL, bool ShowError = false, long sender = 0)
{
}
/// <summary>
/// uses winScp. securely uploads files to the given host
/// </summary>
public string FTPS_UPLOAD(string iHost, int iPort, string iUsername, string iPassword, string iLocalFileName, string iRemoteLocation, string iHostCertificateFingerprint = "")
{
}
 
/// <summary>
/// Returns a formated string using C# string.format()
/// </summary>
public string FN_STRING_FORMAT(string iString, params object[] iParams)
{
}
 
 
public string FN_SERIALIZE(dynamic iObject)
{
}
 
/// <summary>
/// Returns Cursor position Relative to the screen
/// </summary>
public string getCursorPosition()
{
}
 
/// <summary>
/// Shows a dialog-form for parent window.. non customizable
/// </summary>
/// <param name="iHWND"></param>
public int AreYouSure(int iHWND)
{
}
 
/// <summary>
/// Shows confirm dialog, customizable
/// </summary>
public int ShowDialog(string caption, string message, string buttonTextForYes, string buttonTextForNo)
{
}
 
/// <summary>
/// Shows confirm dialog, customizable
/// </summary>
public int ShowDialogRich(string caption, string message, string buttonTextForYes, string buttonTextForNo)
{
}
 
/// <summary>
/// Shows rich dialog form using JSON configuration
/// </summary>
public int ShowDialogJSON(string JSONConfig)
{
}
 
/// <summary>
/// Shows Input-box form
/// </summary>
public string ShowInputBox(InputBoxType Type = InputBoxType.Text, string Title = "", string Message = "", int PosX=0, int PosY=0, string ThemeBg = "", string ThemeForeColour = "")
{
}
 
/// <summary>
/// Shows progressbar
/// </summary>
public long OpenProgressBar(string Title, string Message, int Total, bool AutoClose, string ThemeBg, string TitleForeColour)
{
}
 
/// <summary>
/// Sets value for an existing progressbar or show error
/// </summary>
public long SetProgressBar(long Handle, int CurrentValue, string Message, int NewMaxValue, bool AutoClose = false)
{
}
 
/// <summary>
/// Closes an already open progressbar.
/// </summary>
/// <param name="Handle"></param>
public void CloseProgressBar(long Handle)
{
}
 
/// <summary>
/// If clipboard contains an Image, save in temp location and return the file path
/// </summary>
public string SaveClipboardToImage(string path, string FileName, string ImageType)
{
}
 
/// <summary>
/// Download a file from web and save it to local path. Returns saved file path
/// </summary>
public string DownloadAFile(string url, string destination, bool overWrite, bool ShowProgress)
{
}
 
 

public string PadLeft(string Input, int Length, string PaddingChar="")
{
}

public string PadRight(string Input, int Length, string PaddingChar="")
{
}
 
 
/// <summary>
/// De-Serializes a JSON string to a dynamic type. Returns the dynamic object
/// </summary>
public object JSONToObject(string json)
{
}
 
/// <summary>
/// Reads a property from JSON dynamic object and returns the property value.
/// </summary>
public string JSONGetValue(object iObject, string propertyName)
{
}
 
/// <summary>
/// Extracts a JSON property from given JSON object and returns the value as JSON object.
/// </summary>
public object JSONGetObject(object jsonParsedObject, string propertyName)
{
}
 
/// <summary>
/// Show modal modern UI calendar for vBA users
/// </summary>
/// <returns></returns>
public DateTime ShowCalendar()
{
}
 
/// <summary>
/// Shows custom open file dialog. Allows drag and drop too.
/// </summary>
/// <returns>Json formatted string</returns>
public string ShowDialogForFile(string Message = "", bool AllowMulti = true, string[] Filters = null, int PosX =0, int PosY =0, string ThemeBg="", string ThemeForeColour="", bool closeAfterFileDrop = true)
{
}
 
/// <summary>
/// Shows custom open file dialog. Allows drag and drop too.
/// </summary>
/// <returns>String[] array</returns>
public string[] ShowDialogForFileArray(string Message = "", bool AllowMulti = true, string[] Filters = null, int PosX = 0, int PosY = 0, string ThemeBg = "", string ThemeForeColour = "", bool closeAfterFileDrop = true)
{
}
 
/// <summary>
/// Converts HTML color to access color code
/// </summary>
public int ColorHexToAccess(string HTMLColor)
{
}
 
/// <summary>
/// Converts MS ACCESS color to HTML colour code
/// </summary>
public string ColorAccessToHex(long AccessColor)
{
}
 
 
/// <summary>
/// Returns true or false whether the url is reachable
/// </summary>
public bool UrlIsReachable(string url)
{
}
 
/// <summary>
/// Returns true or false whether the url is well formatted
/// </summary>
public bool UrlIsValid(string url)
{
}
/// <summary>
/// Is the given url a local file path?
/// </summary>
public bool UrlIsLocalPath(string p)
{
}
 
/// <summary>
/// Is the given url a local file path?
/// </summary>
public bool UriIsLocalPath(string p)
{
}
 
// ------------------  Dell specific functions------------------------
 
/// <summary>
/// Returns App version
/// </summary>
public string version()
{
}
 
public string copyright()
{
}
```
