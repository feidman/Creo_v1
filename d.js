$(document).ready(function(){
    //Fades in menus
    $('div').hide().fadeIn(2000); 
    $('#howto').hide();
    //Fades out How-To menu after a while
    // setTimeout(function() {
    // $('#howto_button').fadeOut('fast');
    // }, 10000);
    
    //Toggles the appearance of the list items - hover(on action, off action)
    $('li').hover(
        function(){
          $(this).css('color', 'gold');
        },
        function(){
          $(this).css('color', '');
        }
    );
   
   //Changes the cursor to a pointer on hover over howto. I previously used a hover
   $('#howto_button').css('cursor','hand').css('cursor','pointer');
   
   $('#howto_button').click(function(){
       $('#howto').slideToggle('slow', function(){      
           if($('#howto_button').text() != "Hide"){
           $('#howto_button').text("Hide");
           $('#howto_button').css("background-color", "rgba(0,0,0,0)");
           }else{
           $('#howto_button').text("How-To");
           $('#howto_button').css("background-color", "rgba(0,0,0,0.5)");
           }
       });
   });
   
});