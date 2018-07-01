(function() {
	var wsh = {
		sh: new ActiveXObject('WScript.Shell'),
		fs: new ActiveXObject('Scripting.FileSystemObject')
	};
	var ForReading = 1, ForWriting = 2, ForAppending = 8;
	var TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0;

	var script_path = WScript.ScriptFullName;
	var script_dir = wsh.fs.GetParentFolderName(script_path);
	var submodule_path = wsh.fs.BuildPath(script_dir, 'submodule.js');
	var submodule_content = '';

	// ‚±‚±‚ÅSubModuleObject‚ð’è‹`‚µ‚Ä‚¨‚­‚Æ‰º‚Ìtypeof‚ÌŒ‹‰Ê‚ªobject‚É‚È‚é
	// ‚±‚±‚ÅSubModuleObject‚ð’è‹`‚µ‚È‚¢‚Æ‰º‚Ìtypeof‚ÌŒ‹‰Ê‚ªundefined‚É‚È‚é
	// var SubModuleObject;
	if (wsh.fs.FileExists(submodule_path)) {
		if (wsh.fs.GetFile(submodule_path).Size > 0) {
			var reader = wsh.fs.OpenTextFile(submodule_path, ForReading, false, TristateFalse);
			var submodule_content = reader.ReadAll();
			reader.Close();
			eval(submodule_content);
		}
	}

	WScript.StdOut.WriteLine('typeof SubModuleObject: ' + typeof SubModuleObject);
	WScript.StdOut.WriteLine('typeof submodule_function: ' + typeof submodule_function);
})();

//// submodule.js
//var SubModuleObject;
//function submodule_setup() { SubModuleObject = {}; }
//submodule_setup();
//function submodule_function() {}
