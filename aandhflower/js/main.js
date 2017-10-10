// JavaScript Document

$(document).ready(function(){ 






});

  var bauhaus = {
    src: 'perpetua.swf'    
  };

  sIFR.debugMode = false;

  
  sIFR.activate(bauhaus);

  sIFR.replace(bauhaus, {
    selector: 'h1',
	wmode:"transparent",
	css: {
      '.sIFR-root': { 'color': '#FFFFFF' }
	}
	
  });