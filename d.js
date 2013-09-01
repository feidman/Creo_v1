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
   $('#CreateDrw').click(function(){
        var  CAD_Models= oSession.ListModelsByType(new ActiveXObject("pfc.pfcModelType").MDL_PART);  //Returns all the CAD models in-session... not the ones in the workspace.
        alert(CAD_Models.Count);
        var  Parts = oSession.ListModels();  //returns all the models (both prts and asms in session.)
        alert(Parts.Count);
         var  Drawings= oSession.ListModelsByType(new ActiveXObject("pfc.pfcModelType").MDL_DRAWING);  //Returns only the drawings in session.
        alert(Drawings.Count);
        // Below is trying to get the sequence selection to work.
        //var dwg_one = Drawings.Item(1);
        // alert(dwg_one);
   });
   
       
   //This is to try and get a list of all the parts in the workspace:
    var  parts_seq= oSession.ListModelsByType(new ActiveXObject("pfc.pfcModelType").MDL_PART);  //Returns all the CAD models in-session... not the ones in the workspace.
    document.getElementById("num_models").innerHTML="Model Count: " + parts_seq.Count;
    var  models_seq = oSession.ListModels();  //returns all the models (both prts and asms in session.)
    document.getElementById("num_parts").innerHTML="Part Count: " + models_seq.Count;
    var  drw_seq = oSession.ListModelsByType(new ActiveXObject("pfc.pfcModelType").MDL_DRAWING);
    document.getElementById("num_drw").innerHTML="Drawing Count: " + drw_seq.Count;
    document.getElementById("part_select").innerHTML=parts_seq.Item(1);
});