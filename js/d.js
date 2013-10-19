$(document).ready(function(){
    //Fades in menus
    $('div').hide().fadeIn(1000);
    $('#howto').hide();

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

//INITIALIZE CONNECTION TO PRO-E (The window. is javascript syntax to make mGlob and oSession global for use in the DescFromPart function since it's a function decleration its intepreted before this line is ran, so oSession would otherwise be out of scope of DescFromPart.
window.mGlob = new ActiveXObject("pfc.MpfcCOMGlobal"); //Makes connection to Pro-E
window.oSession = mGlob.GetProESession();  //Returns reference to current session to oSession

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

       //This section constructs an array full of each row of the selection table.
       var TableData = [];
       var numDrws = DrwSeq.Count;
       for (var i = 0; i < numDrws; i++) {
	   var currentDrw = ShortName(DrwSeq.Item(i));
	   TableData.push(
	       {
		   id: i,
		   partNumber: currentDrw,
		   description: DescFromPart(currentDrw),
		   directory: dirTarget(currentDrw) 
	       }
	   );
       }

       //Combine the below into just an openfile using DrwSeq.item(i) and then Erase().
//    var target_drw = "delete.drw";
//    var targetDescriptor = new ActiveXObject("pfc.pfcModelDescriptor");
//    var target_descr = targetDescriptor.Create(new ActiveXObject("pfc.pfcModelType").MDL_DRAWING, target_drw, null);
//    var target = oSession.RetrieveModel(target_descr);   //There are other ways to get models, possibly easier using GetModelByName("delete.drw"); etc. that probably make Pro-E do some of the work.
//    target.Display();   //You have to display the drawing to actually export the PDF.
//    target.Export("test.pdf",ExportPDF);
//    target.Erase();     //This clears it afterwards out of memory. Could use EraseWithDependencies() too.

   });

    $('div').css('color', 'black'); //This is just so that text is readable when I have the background turned off. Delete for final product.
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

//This function returns the directory the file should be move to, based on it's part number (example: PRS-P-... would be placed in the P-Pedals folder). If it doesn't recognize it as a model part or FS part, it puts it on te desktop.
function dirTarget (partNumber) {

    //This defines the root locations for the parts (FS, Model, or Desktop)
    var targetFS = "\\\\thea\\DavWWWRoot\\Engineering\\PRS\\aero\\Drawings\\";
    var targetModel = "X:\\StockCar\\NASCAR Engineering\\Design\\Design Database\\Scale Model\\Drawings- WTM\\";
    var targetDesktop = "DesktopDirectory";

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
    if (firstThreeChars==="PRS" && critFSAscii<91 && critFSAscii>64 && partLength<18 && partLength>15) {
	//THEN PARSE critFSLetter USING A FULL-SCALE PART SPECIFIC METHOD
	return targetFS + fsDirs[critFSLetter];
    }
    else if (firstThreeChars==="PR1" && critModelAscii<91 && critModelAscii>64 && partLength<14 && partLength>11) {
	//THEN PARSE critModelLetter USING A MODEL PART SPECIFIC METHOD
	return targetModel + modelDirs[critModelLetter];
    }
    else {
	//IT IS NOT PARSABLE AND RETURN "NOT EXPORTING"
	return targetDesktop;
    }

}
