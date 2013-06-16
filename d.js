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
var mGlob = pfcCreate("MpfcCOMGlobal"); //Makes connection to Pro-E
var oSession = mGlob.GetProESession();  //Returns reference to current session to oSession
   
//MAIN PROGRAM
   $('#OpenDrw').click(function(){
        var CurModel = oSession.CurrentModel;
        var CopyObjectFullName	= CurModel.FullName;
        var ModelDescriptor = pfcCreate("pfcModelDescriptor");
		var newModDescriptor = ModelDescriptor.Create (pfcCreate("pfcModelType").MDL_DRAWING, CopyObjectFullName  , null); 
		try{
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
  
  
  
  
  
  
  
  
  
//FUNCTIONS 
//Checks if using windows   
function pfcIsWindows () 
{ 
  if (navigator.appName.indexOf ("Microsoft") != -1) 
    return true; 
  else 
    return false; 
} 
 
//Function creates Pro/Web.Link objects for UNIX and Windows 
function pfcCreate (className) 
{ 
 if (!pfcIsWindows()) 
   netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect"); 
 
 if (pfcIsWindows()) 
   return new ActiveXObject ("pfc."+className); 
 else 
 { 
  ret = Components.classes ["@ptc.com/pfc/" +  
            className + ";1"].createInstance(); 
  return ret; 
 } 
} 
   
});