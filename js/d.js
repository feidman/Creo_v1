$(document).ready(function(){
    //Fades in menus
    $('div').hide().fadeIn(1000);
    $('#howto').hide();
    $('#pdfApprove').hide();

    //Fades out How-To menu after a while. Doesn't fade it out if How-To is open.
    setTimeout(function() {
	if($('#howto_button').text() === "How-To")
	{$('#howto_button').fadeOut('fast');}
    }, 10000);

    //Toggles the appearance of the list items - hover(on action, off action)
    $('li').hover(
	function(){
	    $(this).css('color', 'gold');
	},
	function(){
	    $(this).css('color', '');
	}
    );

    $('#howto_button').css('cursor','pointer');
    $('li').css('cursor','pointer');

    $('#howto_button').click(function(){
	if($('#howto_button').text() != "Hide"){
	    $('#howto_button').text("Hide").css("background-color", "rgba(0,0,0,0)");
	}else{
	    $('#howto_button').text("How-To").css("background-color", "rgba(0,0,0,0.5)");
	}
	$('#howto').slideToggle('fast');
    });

    //jqGrid Definition
    $("#ProEOutput").jqGrid({
	datatype: "local",
	height: 'auto',
	width: $(window).width()-$('#howto_button').width()-25,
	forceFit: true,
	rowNum: 250,      //This sets the max number of rows possible, if this wasn't here sorting the files shrinks it down to the default 20 vis
	colNames:['ID','Status','Part Number', 'Description', 'Target Directory', 'Original Directory','Short Directory'],
	colModel:[
	    {name:'id',index:'id', width:15, sorttype:"int", title: false, hidden:true},
	    {name:'fileExists',index:'fileExists', width:15, sortable: false,title: false,
	     cellattr: function(rowId, cValue, rawObject, cm, rdata) {
		 //The below correctly shows the jquery icon, but it also shows all the jqeuery icons after it too! No good.
//		 if (cValue === "true") {return '<span class="ui-icon ui-icon-refresh"></span>'; }
//		 if (cValue === "false") {return '<span class="ui-icon ui-icon-check" title="New Drawing"></span>'; }
	     }
	    },
	    {name:'partNumber',index:'partNumber', width:50, title: false},
	    {name:'description',index:'description', title: false},
	    {name:'directory',index:'directory', hidden:true, title: false},
	    {name:'origDir',index:'origDir', hidden:true, title: false},
	    {name:'shortDir',index:'shortDir', width:50, title: false,
	    	     cellattr: function(rowId, cValue, rawObject, cm, rdata) {
			 if (rawObject.origDir !== dirTarget('Desktop')){
//			     return 'title= "' + rawObject.origDir + " and " + rawObject.directory +'"';
//			     return 'style = "font-weight:bold"';
			     if (rawObject.directory !== rawObject.origDir) {return 'style = "font-weight:bold"';}
			     else {return 'style = "font-weight:normal"';}
			 }
		     }
	    }
	],
	multiselect: true,
	caption: " ",
	hiddengrid:true,
	deselectAfterSort:false,
	onCellSelect: function(rowId,iCol,cellContent){
	    /*This part is probably really slow, but it makes it so that the order of the columns doesn't matter. iCol returns a number.
	     The below returns instead triggers the logic if the correct column name is triggered.*/
	    var $grid=$(this);    
	    var cm = $grid.jqGrid("getGridParam", "colModel");
	    var colName = cm[iCol].name;

	    if(colName === 'shortDir'){
		var origDirectory = $grid.getCell(rowId,'origDir');
		var shortOrigDirectory = ShortDirName(origDirectory);
		if(cellContent === shortOrigDirectory){
		    $grid.setCell(rowId,'directory',dirTarget('Desktop'));
		    $grid.setCell(rowId,iCol,ShortDirName(dirTarget('Desktop')));
		}
		else{
		    $grid.setCell(rowId,'directory',origDirectory);
		    $grid.setCell(rowId,iCol,shortOrigDirectory);
		}
	    }
	    if(colName === 'fileExists'){
		if(cellContent === 'Exists'){
		    $grid.setCell(rowId,iCol,'Overwrite');
		}
		else if(cellContent == "Overwrite" || cellContent == "Released"){
		    $grid.setCell(rowId,iCol,'Exists');
		}
	    }
	    if(iCol !== 0){  //This part makes it so that only the first column, the checkboxes is able to select columns.I could have used colName !== 'cb', because cb is apparently the name of the checkbox column.
		$grid.setSelection(rowId,false);
	    }
	}
    });

    $("#ProEOutput").parents('div.ui-jqgrid-bdiv').css("max-height","540px");

//INITIALIZE CONNECTION TO PRO-E (The window. is javascript syntax to make mGlob and oSession global for use in the DescFromPart function since it's a function decleration its intepreted before this line is ran, so oSession would otherwise be out of scope of DescFromPart.
window.mGlob = new ActiveXObject("pfc.MpfcCOMGlobal"); //Makes connection to Pro-E
window.oSession = mGlob.GetProESession();  //Returns reference to current session to oSession
window.fso = new ActiveXObject("Scripting.FileSystemObject"); //This needed to be global so I could populate the FileExists Column in one function and move/delete files in another function.
//MAIN PROGRAM

//OPENS THE DRW OF THE CURRENTLY OPENED PRT.
    $('#OpenDrw').click(function(){
	try{
	   var CopyObjectFullName = oSession.CurrentModel.FullName;
	   var ModelDescriptor = new ActiveXObject("pfc.pfcModelDescriptor");
	   var newModDescriptor = ModelDescriptor.Create (new ActiveXObject("pfc.pfcModelType").MDL_DRAWING, CopyObjectFullName  , null);
	   var DrwWin = oSession.OpenFile(newModDescriptor);
	   oSession.CurrentWindow.SetBrowserSize(0.0); //Originally this was first, but I don't want the window to hide unless it works
	   DrwWin.Activate();
	}
	catch(err){
	    if (err.number == "-2147352567")  //this number is the error code for "file not found"
	    {
		// var Options = new ActiveXObject("pfc.pfcDrawingCreateOptions");
		// Options.Append(new ActiveXObject("pfc.pfcDrawingCreateOptions").DRAWINGCREATE_DISPLAY_DRAWING);
		// Options.Append(new ActiveXObject("pfc.pfcDrawingCreateOptions").DRAWINGCREATE_WRITE_ERROR_FILE);
		// var DrawingModel = oSession.GetModel(CopyObjectFullName, new ActiveXObject("pfc.pfcModelType").MDL_ASSEMBLY);
		// var NewDwg = oSession.CreateDrawingFromTemplate("TEST_DRW","a0_drawing",DrawingModel.Descr, Options);
		alert("Do you want to create a DRW for this part?");
	    }
	}
    });

//CREATES A DRW OF THE CURRENTLY OPENED PRT
    $('#ExportDrw').click(function(){

       var ServerHandle = oSession.GetActiveServer();
       var OpenOptions = new ActiveXObject("pfc.pfcFileListOpt").FILE_LIST_LATEST;
       var Dir = "wtws://" + ServerHandle.Alias + "/" + ServerHandle.ActiveWorkspace;
       var DrwSeq = oSession.ListFiles("*.drw",OpenOptions,Dir); //This gives you a Pro-E "sequence" of strings. Each item, accessed using .Item(i) is the string.

       //This section constructs an array full of each row of the selection table.
       var TableData = [];
       var numDrws = DrwSeq.Count;
       for (var i = 0; i < numDrws; i++) {
	   var currentDrw = ShortName(DrwSeq.Item(i));
	   targetDir = dirTarget(currentDrw);
	   TableData.push(
	       {
		   id: i+1,
		   fileExists: fso.FileExists(targetDir + currentDrw + ".pdf") ? "Exists" : "New",
		   partNumber: currentDrw,
		   description: DescFromPart(currentDrw),
		   directory: targetDir,
		   origDir: targetDir,
		   shortDir: ShortDirName(targetDir)
	       }
	   );
       }

       jQuery("#ProEOutput").jqGrid('setCaption',"Drawings in Current Workspace");
       $('#ProEOutput').jqGrid('clearGridData');   //Clears the grid before repopulating it.
       for(var i=0;i<TableData.length;i++)
	   $("#ProEOutput").jqGrid('addRowData',TableData[i].id,TableData[i]);
       $("#ProEOutput").jqGrid('setGridState','visible');
       $("#pdfApprove").show();
    });

    $('#pdfApprove').click(function(){
	var selectedRows = $("#ProEOutput").jqGrid('getGridParam','selarrrow');
	var numSelected = selectedRows.length;

	//Only creates this stuff if anything was selected AKA selection count is greater than 0.
	if (numSelected > 0){

	    //This section before the for loop creates all the possible options. They have to be here to be within scope.
	    var ColorPDF = new ActiveXObject("pfc.pfcPDFExportInstructions").Create();
	    var MonoPDF = new ActiveXObject("pfc.pfcPDFExportInstructions").Create();
	    var SettingsMono = new ActiveXObject("pfc.pfcPDFOptions");
	    var SettingsColor = new ActiveXObject("pfc.pfcPDFOptions");
	    var OptionFactory = new ActiveXObject("pfc.pfcPDFOption");

	    //The below creates the setting that stops Pro-E from opening the PDF viewer when it saves a PDF.
	    var SettingViewerOFF = OptionFactory.Create();
	    SettingViewerOFF.OptionType = new ActiveXObject("pfc.pfcPDFOptionType").PDFOPT_LAUNCH_VIEWER;
	    SettingViewerOFF.OptionValue = new ActiveXObject("pfc.MpfcArgument").CreateBoolArgValue(false);   //The documentation mentions the method .setOptionValue, leaving out set made it work for me.
	    SettingsColor.Append(SettingViewerOFF);
	    SettingsMono.Append(SettingViewerOFF);

	    //The below creates the setting to make the colors black and white. By default it's color, so I only have to add the Mono seting to the SettingsMono object.
	    var ColorSetting = OptionFactory.Create();
	    ColorSetting.OptionType = new ActiveXObject("pfc.pfcPDFOptionType").PDFOPT_COLOR_DEPTH;
	    ColorSetting.OptionValue = new ActiveXObject("pfc.MpfcArgument").CreateIntArgValue(new ActiveXObject("pfc.pfcPDFColorDepth").PDF_CD_MONO);
	    SettingsMono.Append(ColorSetting);

	    MonoPDF.Options = SettingsMono;
	    ColorPDF.Options = SettingsColor;

	    //This stores the original window size to restore it later.
	    var windowSize = oSession.CurrentWindow.GetBrowserSize();	    
	    var DescriptorFactory = new ActiveXObject("pfc.pfcModelDescriptor");

	    for(var i=0;i<numSelected;i++){
		var $grid =  $("#ProEOutput");
		var actualRowNum = selectedRows[i];
		var curDrw = $grid.jqGrid('getRowData', actualRowNum);
		var targetPN = curDrw.partNumber;
		var targetOverwrite = curDrw.fileExists;
		var fullNetworkFile = curDrw.directory + targetPN + ".pdf";
		var fullNewFile = oSession.GetCurrentDirectory() + targetPN+ ".pdf";
		
		var targetDescript = DescriptorFactory.Create (new ActiveXObject("pfc.pfcModelType").MDL_DRAWING, targetPN , null);
		var target = oSession.RetrieveModel(targetDescript);
		target.Display();
		target.Export(targetPN,ColorPDF);
		target.Erase();     //This is commented because it might erase drawings they already have opened and haven't saved.

		if (targetOverwrite === "New"){
		    fso.MoveFile(fullNewFile, fullNetworkFile);
		    $grid.setCell(actualRowNum,'fileExists','Released');
		}
		else if (targetOverwrite === "Overwrite"){
		    fso.deleteFile(fullNetworkFile);
		    fso.MoveFile(fullNewFile, fullNetworkFile);
		    $grid.setCell(actualRowNum,'fileExists','Released');
		}
		else if (targetOverwrite === "Exists" || targetOverwrite === "Released"){
		    fso.deleteFile(fullNewFile);    //deletes the exported file instead of moving it.
		}
	    }
	    oSession.CurrentWindow.SetBrowserSize(windowSize);	    
	}
    });

});

//Our Drawings don't utilize the drawing description parameter, instead it's drawn from the part that's used in the drw.
//This function takes the partnumber, such as "PR13WT45C111" and returns the description from the .prt with that name as a string.
//If it can't find the part to open, then it will return "UNAVAILABLE"

function DescFromPart (DrawingName) {
    try{
	var DescriptorFactory = new ActiveXObject("pfc.pfcModelDescriptor");
	var PartDescript = DescriptorFactory.Create (new ActiveXObject("pfc.pfcModelType").MDL_PART, DrawingName , null);
	var DrwPart = oSession.RetrieveModel(PartDescript);

	var Desc = DrwPart.GetParam("DESCRIPTION");
	return Desc.Value.StringValue;
    }
    catch(err){
	return "NOT AVAILABLE";
    }
}

//This function is intended to parse the filename string returned by the ListFile function. Example "wtws://pdm10/Workspace/test_part.drw" would be trimmed to "test_part"
function ShortName (fullname) {
    return fullname.slice(fullname.lastIndexOf("/")+1,fullname.length-4);
}

//This function parses the full target directory down to the final folder that everyone knows.
function ShortDirName (fullDirectory) {
    var fullDirName =  fullDirectory.substr(0,fullDirectory.length-1);
    return fullDirName.slice(fullDirName.lastIndexOf("\\")+1,fullDirName.length);
}

//This function returns the directory the file should be move to, based on it's part number (example: PRS-P-... would be placed in the P-Pedals folder). If it doesn't recognize it as a model part or FS part, it puts it on te desktop.
function dirTarget (partNumber) {

    //This defines the root locations for the parts (FS, Model, or Desktop)
    var targetFS = "\\\\thea\\DavWWWRoot\\Engineering\\PRS\\aero\\Drawings\\";
    var targetModel = "X:\\StockCar\\NASCAR Engineering\\Design\\Design Database\\Scale Model\\Drawings- WTM\\";
    var wshell = new ActiveXObject("WScript.Shell");
    var targetDesktop = wshell.SpecialFolders("Desktop");

    //This defines the critical character in the filename (zero based, so 0 is the first character).
    var critFSLetter = partNumber.substr(4,1).toUpperCase();
    var critModelLetter = partNumber.substr(8,1).toUpperCase();
    var firstThreeChars = partNumber.substr(0,3).toUpperCase();

    //These are calculated here to save the multiple calculations it took to include them in the if structure.
    //The crit*Ascii determines the ascii codes for later checks that it's an actual letter.
    var critFSAscii = critFSLetter.charCodeAt(0);
    var critModelAscii = critModelLetter.charCodeAt(0);
    var partLength = partNumber.length;

    //These are bjects used instead of a switch case to determine the appropriate target folder based on the critical identified letter in the partnumber. They will return the target using modelDirs["A"] for example.
    var modelDirs = {
	'A':'A Aero',
	'B':'B Brakes',
	'C':'C Chassis',
	'D':'D Dampers',
	'E':'E Electronics and Instrumentation',
	'F':'F Front Suspension',
	'G':'G Gear Train',
	'H':'H Hardware',
	'I':'I Engine and GearBox',
	'J':'J Jigs and Hardware',
	'M':'M Data Acquisiton',
	'O':'O Oil System',
	'P':'P Pedals',
	'R':'R Rear Suspension',
	'S':'S Steering',
	'T':'T Tools',
	'U':'U Fuel',
	'V':'V Vendor Parts',
	'W':'W Cooling System',
	'X':'X Research and Development',
	'Y':'Y Template Fixtures',
	'Z':'Z Hardware'
    };

    var fsDirs = {
	'A':'A Aero',
	'B':'B Brakes',
	'C':'C Chassis',
	'D':'D Dampers',
	'E':'E Electronics and Instrumentation',
	'F':'F Front Suspension',
	'H':'H Hardware',
	'J':'J Jigs and Hardware',
	'M':'M Data Acquisiton',
	'O':'O Oil System',
	'P':'P Pedals',
	'Q':'Q MFG',
	'R':'R Rear Suspension',
	'S':'S Steering',
	'T':'T Tools',
	'U':'U Fuel',
	'V':'V Vendor Parts',
	'W':'W Cooling System',
	'X':'X Research and Development',
	'Y':'Y Template Fixtures',
	'Z':'Z Schemes'
    };

    //The first check is for FS parts and ensures the crit character is a letter and the
    if (partNumber === 'Desktop'){
	return targetDesktop + "\\";
    }
    else if (firstThreeChars==="PRS" && critFSAscii<91 && critFSAscii>64 && partLength<18 && partLength>15) {
	//THEN PARSE critFSLetter USING A FULL-SCALE PART SPECIFIC METHOD
	return targetFS + fsDirs[critFSLetter] + "\\";
    }
    else if (firstThreeChars==="PR1" && critModelAscii<91 && critModelAscii>64 && partLength<14 && partLength>11) {
	//THEN PARSE critModelLetter USING A MODEL PART SPECIFIC METHOD
	return targetModel + modelDirs[critModelLetter] + "\\";
    }
    else {
	//IT IS NOT PARSABLE AND RETURN "NOT EXPORTING"
	return targetDesktop + "\\";
    }

}
