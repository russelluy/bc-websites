
// *************** GLOBALS ******************
var gsCurrentWin = null;	//holds handle to current pop up window
//*******************************************

// open URL in a window of specific size and name
function openWindow(xsURL, xsWidth, xsHeight, xsName) {
	var vsName;
	
	if(xsName) { 
		vsName	= xsName;
	} else { 
		vsName = "NewWin";
	}
	if (xsName == 'LittlePopWin')
		var vsNewWindow = window.open(xsURL, vsName,'width='+xsWidth+',height='+xsHeight+',scrollbars=yes,resizable=yes,alwaysRaised,screenX=250,screenY=250,top=250,left=250');
	else
		var vsNewWindow = window.open(xsURL, vsName,'width='+xsWidth+',height='+xsHeight+',scrollbars=yes,resizable=yes,alwaysRaised,screenX=0,screenY=0,top=0,left=0');
	
	vsNewWindow.focus();
	
	if ( gsCurrentWin && !gsCurrentWin.closed ) {
		if ( gsCurrentWin.name != vsName ) { gsCurrentWin.close(); }
	}
	
	gsCurrentWin = vsNewWindow;
}

// NAV BAR FUNCTIONS

function openCertificate() {
	openWindow('../cert/cert_010_pop.html', '500', '400', 'NavWin');
}  //end function

function openJobAid() {
	openWindow('../job_aid/job_aid_010_pop.html', '500', '400', 'NavWin');
}  //end function

function openToolBox() {
	openWindow('../toolbox/toolbox_010_pop.html', '500', '400', 'NavWin');
}  //end function

function openHelp() {
	openWindow('../help/help_010_pop.html', '500', '400', 'NavWin');
}  //end function

function openLittlePop(url) {
	openWindow(url, '340', '350', 'LittlePopWin');
}  //end function

function openBigPop(url) {
	openWindow(url, '770', '500', 'BigPopWin');
}  //end function

function openSurvey() {
	openWindow('overview_050_pop_010.html', '770', '500', 'SurveyWin');
}  //end function

function openCaseStudies(url) {
	openWindow(url, '770', '500', 'CSWin');
}  //end function

function openDefPop(url) {
	openWindow(url, '400', '400', 'DefWin');
}  //end function

//function for going forward a page
function goNext() {
	document.location.href = nextPage;
}  //end function

function exitProgram() {
	top.close();
}  //end function

//function for going back a page
function goBack() {
		document.location.href = backPage;
}  //end function


currReveal = 0;
bgOn = 'url(../images/click_2_reveal_bg.gif)';
//reveals the correct text for the click to reveal pieces
function revealText(num) {
	//ignore if we are already showing the clicked on piece
	if (currReveal == num)
		return;
	
	//alert(num);
	
	if (document.getElementById("reveal" + num)) {
		document.getElementById("reveal" + num).style.display = 'block';
		
		//do we need to hide the text that is currently showing?
		//also need to switch to other bg
		if (currReveal != 0)
			document.getElementById("reveal" + currReveal).style.display = 'none';
		else
			document.getElementById("rightcol-c2r").style.backgroundImage = bgOn;
			
		currReveal = num;
	}
}  //end function

//saves scores during survey comparison
function saveScores() {
	sStr = new String();
	df = document.forms[0];
	for (i = 1; i <=10; i++) {
		dfq = eval("df.q" + i)
		for (j = 0; j <=4; j++) {
			if (dfq[j].checked)
				sStr += j;
		}  //end inner for
	}  //end outer for
	//alert(sStr);	
	document.location.href = 'overview_050_pop_020.html?' + sStr;
}


function parseScore() {
	if (sStr == '')
		return;
	df = document.forms[0];
	for (i = 0; i <=9; i++) {
		y = (i+1);
		y = y.toString();
		dfq = eval("df.q" + y);
		//alert(dfq);
		var num = sStr.charAt(i);
		//alert(num);
		dfq[num].checked = true;
	} // end i

}  //end function

function writeScore(n) {
	if (sStr == '')
		return '<td>&nbsp;</td>';

	var t = parseInt(sStr.charAt(n-1));
	return '<td align="center"><strong>' + (t+1) + '</strong></td>';

} //end function

function writeTotalScore() {
	if (sStr == '')
		return '<td>&nbsp;</td>';
	var tot = 0;
			
	for (i = 0; i <=9; i++) {
		var num = parseInt(sStr.charAt(i)) + 1;
		tot += num;
	}
	
	var ave = tot/10;
	return '<td align="center" class="no-q"><strong>' + ave + '</strong></td>';
}

//funciton for toggling elements in an expand collapse list
var currOpen = 0;
function toggleList(num){
	element = document.getElementById("expando" + num);

	if ((currOpen != 0) && (currOpen != num)) {
		closeEle = document.getElementById("expando" + currOpen).style;
		closeEle.display = 'none';
	}
	
	if ((element.style.display == 'none') || (element.style.display == '')) {
			//document.getElementById("img" + num).src = '../images/arrow_down.gif';
			displayDiv(element);
	} else {
			//document.getElementById("img" + num).src = '../images/arrow.gif';
			hideDiv(element);
	}
	currOpen = num;
}

//generic function for displaying a DIV
function displayDiv(ele) {
	ele.style.display = 'block';
}

//generic function for hiding a DIV
function hideDiv(ele) {
	ele.style.display = 'none';
}


//a global variable that will hold the current answer value
currentAnswer = '';
////////////////////////////////////////////////////////////////////////
//this function records the answer for a multiple choice question with multiple answers
function RecordMulti(ans)
{

if (currentAnswer == '')
		currentAnswer = ans;
	else if (currentAnswer.toString().indexOf(ans) != -1)
	{
		splitStr = currentAnswer.toString().split(eval(ans) + ',');
		if (splitStr.length == 2)		
			currentAnswer = splitStr[0] + splitStr[1];
		else
		{
			splitStr = currentAnswer.toString().split(',' + eval(ans));
			if (splitStr.length == 2)
			  currentAnswer = splitStr[0] + splitStr[1];
			else
			{
	  		    splitStr = currentAnswer.toString().split(eval(ans));
				currentAnswer = splitStr[0] + splitStr[1];
			}  //end else
		}  //end else
	}  //end else if
	else
		currentAnswer += ',' + ans;  

} //end function

/////////////////////////////////////////////////////////////////////////////
//this function checks the answer for a multiple choice question with multiple answers
function CheckAnswerMulti()
{

     userArray = currentAnswer.toString().split(',');
	 userArray.sort();
	 correctArray = correctAnswer.toString().split(',');
	 correctArray.sort();
	 correct = true;
	 if (correctArray.length != userArray.length)
	   correct = false;
	 else
	 {
	 	for (i = 0; i < correctArray.length; i++)	 
	 	{
	 		if (correctArray[i] != userArray[i])
			  correct = false;
	 	}  //end for 
	 }
	 
	if (correct)
	{
		revealText(2);
		//MM_swapImage('feedback','','../images/correctfeedback.gif',1)
	}
	else
	{
		revealText(1);
		//MM_swapImage('feedback','','../images/incorrectfeedback.gif',1)
	}

}  //end function

//Swap and image.
function imgSwap(imgName, imgPath) {
	var theImage = document.images[imgName];
	theImage.src = imgPath;
}
