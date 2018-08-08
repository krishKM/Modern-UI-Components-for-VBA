# A friendly helper DLL that will make you smile.
**No installation**, **no ActiveX**, **no Admin-Rights.** 
Just add this Dll to your VBA project folder and have some cool UI features. Have only tested in MS Access but it should work in all VBA environment. Works with ACCDE too.

The main goal is to bring some .NET function to your VBA project to make your project stand-out visually and functionally.
Of course with minimal code!

```diff
- NOTE:
- this is an evolving project. Function names from one version to another might varry, please test your wrappers before updating to the newest one.

- If you have any new suggestions, please feel free to let me know.

+ Simple ShowDialog can be colour themed
+ Import data to a table using JsonImport
+ AreYouSure can be colour themed
+ JsonImport and ExportToJson added.

```

# upcoming
```diff
+ Barcode control: Generate barcode from your strings in forms, reports. Works are nearly finished. Need to make it userfriendly
``` 




## History
it all started with this question <a href="https://stackoverflow.com/questions/39224308/non-blocking-toast-like-notifications-for-microsoft-access-vba">SO Question</a> 

## What it does?
Helps you to make your application more user-friendly by providing some .NET components and functions that you can use within your VBA application. Visually and functionally cooler than VBA!

## How to use?
Some basic VBA skills are required! Just download the <a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> sample</a> ACCDB from sample folder where you can find the Dll. Copy and paste the functions you require to your VBA application. 


# Interesting features
<ul>
  <li>Show non-blocking notifications</li>
  <li><a href="https://github.com/krishKM/VBA_TOOLS/blob/master/README.md#show-cool-dialogbox"> Show Cool DialogBox</a></li>
  <li><a href="https://github.com/krishKM/VBA_TOOLS/blob/master/README.md#show-cool-progressbar"> Show Cool Progressbar</a></li>
  <li><a href="https://github.com/krishKM/VBA_TOOLS/blob/master/README.md#other-Features-that-are-interesting"> Download a file with progressbar</a></li>
  <li><a href="https://github.com/krishKM/VBA_TOOLS/blob/master/README.md#show-cool-inputboxes">Show Cool InputBoxes</a></li>
  <li><a href="https://github.com/krishKM/VBA_TOOLS/blob/master/README.md#drag-and-drop-openfiledialog">Drag and drop OpenFileDialog</a></li>
  <li><a href="https://github.com/krishKM/VBA_TOOLS/blob/master/README.md#load-picture-from-url-to-imagecontrol-without-saving">Load Picture from URL to ImageControl without saving</a></li>
  <li><a href="https://github.com/krishKM/VBA_TOOLS/blob/master/README.md#Other-Features-that-are-interesting">Other Features that are interesting</a></li>
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
1.Message   : can contain <a href="">text</a> for hyperlinks (any other html tags are ignored, hyperlink must begin with www or http or https (basically web links only?)
2.Duration in Milli-Seconds (default 2000. 0 will keep the notification for long time.  int.max)
3.Background colour (html colour code)
4.Font colour (html colour code)
5.X,Y position on the desktop
```



![picture of 3 notifications](https://raw.githubusercontent.com/krishKM/VBA_TOOLS/master/screenshots/collections.png)
```VBA
'used commands
Toastr.Toast "Ups something went wrong!",vberror,0
Toastr.Toast "Yellow weather warning!",vbexclamation,0
Toastr.Toast "You've just received a notification",vbinformation,0
```

in Action
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InAction.gif)
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InAction1.gif)

#### how about little interaction with user and show some hyperlinks?
You can have html ```<a href="">text</a>``` tags in your message which will be translated into hyperlinks.
![Notification in action gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/Hyperlink.png)

## Download 
Download the sample and test it in your project. Please leave comment how you feel.
<a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> Samples</a>

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


![Cool DialogBox](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/DialogboxGreen.png)
![Cool DialogBox1](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/4Buttons.png)

There is vba wrapper in the sample accdb which can be extended as per your need. It uses the 3rd party JSON Converter plugin with some miner fixes from my side.

```
  'usign the wrapper it would be as simple as 
  Debug.Print gDll.DialogRich("This is a title", "Some content", (vbExclamation + vbYesNo))
```

# Show cool Progressbar
Progressbars 

crucial element when informing users about a progress. Meet the cool progressbar which can pop up on top of your application at any time with a simple code as such as.

```
  Dim ProgressBarID As Long
  ProgressBarID = gDll.ShowProgressBar(100, "Executing your query", "Please wait. We are preparing printer drivers")
    
  ProgressBarID = gDll.SetProgressBar(ProgressBarID, 10, "Waiting for driver..")
  
  gdll.CloseProgressbar ProgressbarId 'Will close the progressbar
```
![Cool ProgressbarGreen](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/ProgressBar.png)

As usual, you are allowed to change theme colours as per your taste.
![Cool ProgressbarRed](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/ProgressBarRed.png)

### note:
```ShowProgressBar and SetProgressBar``` returns an ID which you can refer your progressbar to. This also allows VBA users to have multiple progressbars at the same time. 

# Show Cool InputBoxes
InputBox another heavily used component. Some like the plain system looking InputBox but we love the modern UI colours :)
What would you chose from these tables?

![InputBoxCollection](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InputBoxDefault.png)  ![InputBoxCollection](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InputBoxMultiline.png) 

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

![InputBoxCollection](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InputBox.png)

as always we can change theme colours:)

![purple input box](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/InputBoxPurple.png)

Download <a href="https://github.com/krishKM/VBA_TOOLS/tree/master/samples"> sample</a>


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
![File drag and drop error gif](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/FileDropErrorInAction.gif)
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





<hr>
<hr>


# [Other Features that are interesting]

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
![AreYouSureCollection](https://github.com/krishKM/VBA_TOOLS/blob/master/screenshots/AreYouSureCollection.png)

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
/// <summary>
/// Shows toasts on the desktop
/// </summary>
public async void FN_SHOW_TOAST(string iMessage, int iDuration, string iBG_COLOR, long iANIME_DURATION, string iFONT_COLOR, int iX, int iY, int iANIM_DIRECTION, bool iAUTO_CLOSE = true)
{
}
 
 

/// <summary>
/// Converts ByteArray to an Image and saves in a provided location.
/// </summary>
/// <returns>The path of the image saved locally</returns>
public string ByteToImage(byte[] byteArrayIn, string iTempPath, bool useCache)
{
}
 
/// <summary>
/// Converts Byte Array to a bitmap
/// </summary>
public Bitmap ByteToBitmap(byte[] byteArr)
{
}
 
/// <summary>
/// Returns a ByteArray of the image
/// </summary>
/// <param name="hWND"></param>
public byte[] TakeScreenShotFromHwnd(long hWND)
{
}
 
/// <summary>
/// Take screen-shot of entire desktop. Returns byteArray
/// </summary>
public byte[] TakeScreenShot()
{
}
 
/// <summary>
/// Take screen-shot of entire desktop. Saves in a location and returns the location
/// </summary>
public string TakeScreenShot1(string SavePath)
{
}
 
/// <summary>
/// Returns ByteArray containing the picture received from the url
/// </summary>
/// <param name="URL"></param>
public byte[] PictureFromUrl(string URL, bool ShowError = false, long sender = 0)
{
}
/// <summary>
/// uses winScp. securely uploads files to the given host
/// </summary>
public string FTPS_UPLOAD(string iHost, int iPort, string iUsername, string iPassword, string iLocalFileName, string iRemoteLocation, string iHostCertificateFingerprint = "")
{
}
 
/// <summary>
/// Returns a formated string using C# string.format()
/// </summary>
public string FN_STRING_FORMAT(string iString, params object[] iParams)
{
}
 
 
public string FN_SERIALIZE(dynamic iObject)
{
}
 
/// <summary>
/// Returns Cursor position Relative to the screen
/// </summary>
public string getCursorPosition()
{
}
 
/// <summary>
/// Shows a dialog-form for parent window.. non customizable
/// </summary>
/// <param name="iHWND"></param>
public int AreYouSure(int iHWND)
{
}
 
/// <summary>
/// Shows confirm dialog, customizable
/// </summary>
public int ShowDialog(string caption, string message, string buttonTextForYes, string buttonTextForNo)
{
}
 
/// <summary>
/// Shows confirm dialog, customizable
/// </summary>
public int ShowDialogRich(string caption, string message, string buttonTextForYes, string buttonTextForNo)
{
}
 
/// <summary>
/// Shows rich dialog form using JSON configuration
/// </summary>
public int ShowDialogJSON(string JSONConfig)
{
}
 
/// <summary>
/// Shows Input-box form
/// </summary>
public string ShowInputBox(InputBoxType Type = InputBoxType.Text, string Title = "", string Message = "", int PosX=0, int PosY=0, string ThemeBg = "", string ThemeForeColour = "")
{
}
 
/// <summary>
/// Shows progressbar
/// </summary>
public long OpenProgressBar(string Title, string Message, int Total, bool AutoClose, string ThemeBg, string TitleForeColour)
{
}
 
/// <summary>
/// Sets value for an existing progressbar or show error
/// </summary>
public long SetProgressBar(long Handle, int CurrentValue, string Message, int NewMaxValue, bool AutoClose = false)
{
}
 
/// <summary>
/// Closes an already open progressbar.
/// </summary>
/// <param name="Handle"></param>
public void CloseProgressBar(long Handle)
{
}
 
/// <summary>
/// If clipboard contains an Image, save in temp location and return the file path
/// </summary>
public string SaveClipboardToImage(string path, string FileName, string ImageType)
{
}
 
/// <summary>
/// Download a file from web and save it to local path. Returns saved file path
/// </summary>
public string DownloadAFile(string url, string destination, bool overWrite, bool ShowProgress)
{
}
 
 

public string PadLeft(string Input, int Length, string PaddingChar="")
{
}

public string PadRight(string Input, int Length, string PaddingChar="")
{
}
 
 
/// <summary>
/// De-Serializes a JSON string to a dynamic type. Returns the dynamic object
/// </summary>
public object JSONToObject(string json)
{
}
 
/// <summary>
/// Reads a property from JSON dynamic object and returns the property value.
/// </summary>
public string JSONGetValue(object iObject, string propertyName)
{
}
 
/// <summary>
/// Extracts a JSON property from given JSON object and returns the value as JSON object.
/// </summary>
public object JSONGetObject(object jsonParsedObject, string propertyName)
{
}
 
/// <summary>
/// Show modal modern UI calendar for vBA users
/// </summary>
/// <returns></returns>
public DateTime ShowCalendar()
{
}
 
/// <summary>
/// Shows custom open file dialog. Allows drag and drop too.
/// </summary>
/// <returns>Json formatted string</returns>
public string ShowDialogForFile(string Message = "", bool AllowMulti = true, string[] Filters = null, int PosX =0, int PosY =0, string ThemeBg="", string ThemeForeColour="", bool closeAfterFileDrop = true)
{
}
 
/// <summary>
/// Shows custom open file dialog. Allows drag and drop too.
/// </summary>
/// <returns>String[] array</returns>
public string[] ShowDialogForFileArray(string Message = "", bool AllowMulti = true, string[] Filters = null, int PosX = 0, int PosY = 0, string ThemeBg = "", string ThemeForeColour = "", bool closeAfterFileDrop = true)
{
}
 
/// <summary>
/// Converts HTML color to access color code
/// </summary>
public int ColorHexToAccess(string HTMLColor)
{
}
 
/// <summary>
/// Converts MS ACCESS color to HTML colour code
/// </summary>
public string ColorAccessToHex(long AccessColor)
{
}
 
 
/// <summary>
/// Returns true or false whether the url is reachable
/// </summary>
public bool UrlIsReachable(string url)
{
}
 
/// <summary>
/// Returns true or false whether the url is well formatted
/// </summary>
public bool UrlIsValid(string url)
{
}
/// <summary>
/// Is the given url a local file path?
/// </summary>
public bool UrlIsLocalPath(string p)
{
}
 
/// <summary>
/// Is the given url a local file path?
/// </summary>
public bool UriIsLocalPath(string p)
{
}
 
// ------------------  Dell specific functions------------------------
 
/// <summary>
/// Returns App version
/// </summary>
public string version()
{
}
 
public string copyright()
{
}
```
