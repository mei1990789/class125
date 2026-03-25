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
		alert("pos_main");
		window.close();
	}
	catch (e)
	{
		window.close();
	}
	
	return;
}

//</script>
/////////////////////////////////////////////////////////////////////////////////////////////
