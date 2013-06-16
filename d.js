$(document).ready(function(){
    //Fades in menus
    $('div').hide().fadeIn(1000);
    $('#howto').hide();
    //Fades out How-To menu after a while. Doesn't fade it out if How-To is open.
    setTimeout(function() {
        if($('#howto_button').text() == "How-To")
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
   
//INITIALIZATION 
var mGlob = new ActiveXObject("pfc.MpfcCOMGlobal"); //Makes connection to Pro-E
var oSession = mGlob.GetProESession();  //Returns reference to current session to oSession
   
//MAIN PROGRAM
   $('#OpenDrw').click(function(){
       try{
            var CurModel = oSession.CurrentModel;
            var CopyObjectFullName	= CurModel.FullName;
            var ModelDescriptor = new ActiveXObject("pfc.pfcModelDescriptor");
            var newModDescriptor = ModelDescriptor.Create (new ActiveXObject("pfc.pfcModelType").MDL_DRAWING, CopyObjectFullName  , null); 
            var DrwWin = oSession.OpenFile(newModDescriptor);
            DrwWin.Activate ();
            oSession.CurrentWindow.SetBrowserSize(0.0); //Originally this like was first, but I don't want the window to hide unless it works
		}
        catch(e){
            //Do nothing if the *.prt has no *.drw (that's what would cause it to error out)
        }
   });
   
   $('#CreateDrw').click(function(){
       alert("In Development");
   });
   
});