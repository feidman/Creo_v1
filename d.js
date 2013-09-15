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
    var server_name = oSession.GetServerByUrl("http://pdm.prnet.us/Windchill");    
    document.getElementById("part_select").innerHTML = "Server Name 1: " + pfcServer.Alias;
    });
   
});
