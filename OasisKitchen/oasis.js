// JavaScript Document
var current = "";

function loader() {
	if (location.search) {
		current = location.search.substring(1);
	} else {
		current = "home";
	}
	document.images[current].src = "images/"+document.images[current].src.split("images/")[1].slice(0,-4)+"c."+document.images[current].src.split("images/")[1].slice(-3);
	window.frames['iframe'].location = current+".html";
}

function imgSwap(imgName) {
	var theImage = document.images[imgName];
	var theSource = theImage.src.split("images/")[1];
	if (theSource.slice(-5,-4) == "b") {
		if (current == imgName) {
			theImage.src = "images/"+theSource.slice(0,-5)+"c."+theSource.slice(-3);
		} else {
			theImage.src = "images/"+theSource.slice(0,-5)+"."+theSource.slice(-3);
		}
	} else if (theSource.slice(-5,-4) == "c") {
		theImage.src = "images/"+theSource.slice(0,-5)+"b."+theSource.slice(-3);
	} else {
		theImage.src = "images/"+theSource.slice(0,-4)+"b."+theSource.slice(-3);
	}
}

function goTo(where) {
	document.location = document.location.pathname + '?' + where;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}