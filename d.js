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

    var server_handle = oSession.GetActiveServer(); //This works with the methods (result of method) .ActiveWorkspace (REPORT), .Location (http://pdm.prnet.us/Windchill),  .isActive (true), .Version (10.1), CollectWorkspaces().Count
    var open_options = new ActiveXObject("pfc.pfcFileListOpt").FILE_LIST_LATEST;
    var dir = "wtws://" + server_handle.Alias + "/" + server_handle.ActiveWorkspace;
    var files_seq = oSession.ListFiles("*.drw",open_options,dir);   //Loop through the actual part names using .Count to determine how many and .item(i) to return the full filename for each one. example returned is wtws://pdm10/BodyFix/delete.drw
    document.getElementById("num_models").innerHTML = "Directory: " + dir;
    document.getElementById("num_parts").innerHTML = "Number of DRWs: " + files_seq.Count;
    document.getElementById("num_drw").innerHTML = "Empty: ";
    document.getElementById("part_select").innerHTML = "Test Drw: wtws://pdm10/BodyFix/delete.drw";
    });

//JUST SPITS OUT THE DXF OF THE A HARDCODED DXF. WILL SOON BE APPLIED TO DYNAMICALLY EXPORT AN ENTIRE LIST.
    $('div').css('color', 'black'); //This is just so that text is readable when I have the background turned off. Delete for final product.

    var export_pdf = new ActiveXObject("pfc.pfcPDFExportInstructions").Create();    

    var pdf_settings = new ActiveXObject("pfc.pfcPDFOptions");
    var setting = new ActiveXObject("pfc.pfcPDFOption");

    var test_setting = setting.Create();
    test_setting.OptionType = new ActiveXObject("pfc.pfcPDFOptionType").PDFOPT_LAUNCH_VIEWER;
    test_setting.OptionValue = new ActiveXObject("pfc.MpfcArgument").CreateBoolArgValue(false);   //The documentation mentions the method .setOptionValue, leaving out set made it work for me.

//    var color_setting = setting.Create();
//    color_setting.OptionType = new ActiveXObject("pfc.pfcPDFOptionType").PDFOPT_COLOR_DEPTH;
//    color_setting.OptionValue = new ActiveXObject("pfc.pfcPDFColorDepth").PDF_CD_MONO;   //The documentation mentions the method .setOptionValue, leaving out set made it work for me.

    pdf_settings.Append(test_setting); 
//    pdf_settings.Append(color_setting); 
    export_pdf.Options = pdf_settings;
 


//I HAVE BEEN ABLE TO PRINT A PDF OF AN ACTIVE DRAWING, THE BELOW ATTEMPTS TO RETRIEVE A DRAWING USING IT'S NAME.
    var target_drw = "delete.drw";
    var targetDescriptor = new ActiveXObject("pfc.pfcModelDescriptor");
    var target_descr = targetDescriptor.Create(new ActiveXObject("pfc.pfcModelType").MDL_DRAWING, target_drw, null);   
    var target = oSession.RetrieveModel(target_descr);   //There are other ways to get models, possibly easier using GetModelByName("delete.drw"); etc. that probably make Pro-E do some of the work.

    target.Display();   //You have to display the drawing to actually export the PDF.
    target.Export("test.pdf",export_pdf);
    target.Erase();     //This clears it afterwards out of memory. Could use EraseWithDependencies() too.
});
