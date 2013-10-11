$(document).ready(function(){
    //Fades in menus
    $('div').hide().fadeIn(1000);
    $('#howto').hide();
    //Fades out How-To menu after a while. Doesn't fade it out if How-To is open.
    setTimeout(function() {
        if($('#howto_button').text() === "How-To")
            {$('#howto_button').fadeOut('fast')}
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
   
//INITIALIZE CONNECTION TO PRO-E
var mGlob = new ActiveXObject("pfc.MpfcCOMGlobal"); //Makes connection to Pro-E
var oSession = mGlob.GetProESession();  //Returns reference to current session to oSession
   
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
       var DrwSeq = oSession.ListFiles("*.drw",OpenOptions,Dir); //This gives you a Pro-E "sequence" of pfcModel objects, easily accessed using the .item(i) method.
       
       var ExportPDF = new ActiveXObject("pfc.pfcPDFExportInstructions").Create();           
       var SettingsPDF = new ActiveXObject("pfc.pfcPDFOptions");
       var SettingFactory = new ActiveXObject("pfc.pfcPDFOption");

       //The below creates the setting that stops Pro-E from opening the PDF viewer when it saves a PDF.
       var SettingViewerOFF = SettingFactory.Create();
       SettingViewerOFF.OptionType = new ActiveXObject("pfc.pfcPDFOptionType").PDFOPT_LAUNCH_VIEWER;
       SettingViewerOFF.OptionValue = new ActiveXObject("pfc.MpfcArgument").CreateBoolArgValue(false);   //The documentation mentions the method .setOptionValue, leaving out set made it work for me.
       SettingsPDF.Append(SettingViewerOFF); 

       //The below creates the setting to make the colors black and white, by default it's color.
       var ColorSetting = SettingFactory.Create();
       ColorSetting.OptionType = new ActiveXObject("pfc.pfcPDFOptionType").PDFOPT_COLOR_DEPTH;
       ColorSetting.OptionValue = new ActiveXObject("pfc.MpfcArgument").CreateIntArgValue(new ActiveXObject("pfc.pfcPDFColorDepth").PDF_CD_MONO);
       SettingsPDF.Append(ColorSetting); 

       ExportPDF.Options = SettingsPDF;

alert('Running');

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
