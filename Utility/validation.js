// WORD FILTER JAVASCRIPT 
<!-- BEGIN WORD FILTER JAVASCRIPT 
//<script language="JavaScript"> 
// Word Filter 2.0 
// (c) 2002 Premshree Pillai 
// Created : 29 September 2002 
// http://www.qiksearch.com 
// http://javascript.qik.cjb.net 

var swear_words_arr=new Array("fuck","shit","phuck","asshole","a$$hole","pussy","cunt","penis","dick","bitch","whore","cock"); 
var swear_alert_arr=new Array(); 
var swear_alert_count=0; 

function reset_alert_count() 
{ 
swear_alert_count=0; 
} 

function wordFilter(form,fields) 
{ 
reset_alert_count(); 
var compare_text; 
var fieldErrArr=new Array(); 
var fieldErrIndex=0; 
for(var i=0; i<fields.length; i++) 
{ 
eval('compare_text=document.' + form + '.' + fields[i] + '.value;'); 
for(var j=0; j<swear_words_arr.length; j++) 
{ 
for(var k=0; k<(compare_text.length); k++) 
{ 
if(swear_words_arr[j]==compare_text.substring(k,(k+swear_words_arr[j].length)).toLowerCase()) 
{ 
swear_alert_arr[swear_alert_count]=compare_text.substring(k,(k+swear_words_arr[j].length)); 
swear_alert_count++; 
fieldErrArr[fieldErrIndex]=i; 
fieldErrIndex++; 
} 
} 
} 
} 
var alert_text=""; 
for(var k=1; k<=swear_alert_count; k++) 
{ 
alert_text+="\n" + "(" + k + ") " + swear_alert_arr[k-1]; 
eval('compare_text=document.' + form + '.' + fields[fieldErrArr[0]] + '.focus();'); 
eval('compare_text=document.' + form + '.' + fields[fieldErrArr[0]] + '.select();'); 
} 
if(swear_alert_count>0) 
{ 
alert("The form cannot be submitted.\nThe following illegal words were found:\n_______________________________\n" + alert_text + "\n_______________________________"); 
return false; 
} 
else 
{ 
return true; 
} 
} 


/* Check Empty */
function CheckEmpty(who) {   
	if (who == "") {
          alert("The form cannot be submitted.\nA required field is empty.\n_______________________________"); 
           return false;		
	} else {
	 return true;	
	} 
}


/* Check Empty2 */
function CheckEmpty2(who,val1) {   
	if (who == "") {
          alert(val1 + ".\n_______________________________"); 
           return false;		
	} else {
	 return true;	
	} 
}


/* Check Choose */
function CheckChoose(who,val1,descrip1) {   
	if (who == val1) {
          alert("" + descrip1 + "\n_______________________________"); 
           return false;		
	} else {
	 return true;	
	} 
}


/* CheckOneOrTheOther */
function CheckOneOrTheOther(who1,descrip1,who2,descrip2) {   
	if (who1 == "" && who2 == "") {
          alert("The form cannot be submitted.\nA You must enter one of the following fields.\n_______________________________\n" + descrip1 + "\n" + descrip2); 
           return false;		
	} else {
	 return true;	
	} 
}

//////////////////////////////////////////////////
//	<Email Validator>			//
// 	(c) 2003 Premshree Pillai		//
//	Written on: 29/04/03 (dd/mm/yy)		//
//	http://www.qiksearch.com		//
//	http://premshree.resource-locator.com	//
//	Email : qiksearch@rediffmail.com	//
//////////////////////////////////////////////////

/* Without RegExps */
function isEmail(who) {
	function isEmpty(who) {
		var testArr=who.split("");
		if(testArr.length==0)
			return true;
		var toggle=0;
		for(var i=0; i<testArr.length; i++) {
			if(testArr[i]==" ") {
				toggle=1;
				break;
			}
		}
		if(toggle)
			return true;
		alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
       return false;
	}

	function isValid(who) {
		var invalidChars=new Array("~","!","@","#","$","%","^","&","*","(",")","+","=","[","]",":",";",",","\"","'","|","{","}","\\","/","<",">","?");
		var testArr=who.split("");
		for(var i=0; i<testArr.length; i++) {
			for(var j=0; j<invalidChars.length; j++) {
				if(testArr[i]==invalidChars[j]) {
					alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
       return false;
				}
			}
		}
		return true;
	}

	function isfl(who) {
		var invalidChars=new Array("-","_",".");
		var testArr=who.split("");
		which=0;
		for(var i=0; i<2; i++) {
			for(var j=0; j<invalidChars.length; j++) {
				if(testArr[which]==invalidChars[j]) {
					alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
       return false;
				}
			}
			which=testArr.length-1;
		}
		return true;
	}

	function isDomain(who) {
		var invalidChars=new Array("-","_",".");
		var testArr=who.split("");
		if(testArr.length<2||testArr.length>4) {
			alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
       return false;
		}
		for(var i=0; i<testArr.length; i++) {
			for(var j=0; j<invalidChars.length; j++) {
				if(testArr[i]==invalidChars[j]) {
					alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
       return false;
				}
			}
		}
		return true;
	}


	var testArr=who.split("@");
	if(testArr.length<=1||testArr.length>2) {
		alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
       return false;
	}
	else {
		if(isValid(testArr[0])&&isfl(testArr[0])&&isValid(testArr[1])) {
			if(!isEmpty(testArr[testArr.length-1])&&!isEmpty(testArr[0])) {
				var testArr2=testArr[testArr.length-1].split(".");
				if(testArr2.length>=2) {
					var toggle=1;
					for(var i=0; i<testArr2.length; i++) {
						if(isEmpty(testArr2[i])||!isfl(testArr2[i])) {
							toggle=0;
							break;
						}
					}
					if(toggle&&isDomain(testArr2[testArr2.length-1]))
						return true;
					alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
       return false;
				}
				alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
       return false;
			}
		}
	}
}

/* With RegExp */
function isEmail2(who) {
	var email=/^[A-Za-z0-9]+([_\.-][A-Za-z0-9]+)*@[A-Za-z0-9]+([_\.-][A-Za-z0-9]+)*\.([A-Za-z]){2,4}$/i;

	   
	if (who == "") {
          return true;		
	} else if (email.test(who) == false) {
	 alert("The form cannot be submitted.\nThe following email is invalid:\n_______________________________\n" + who + "\n_______________________________"); 
         return false;
	}  else {
	 return true;	
	} 
}


/* With RegExp and suffix filter */
function IsEmail3(str){

var validsuffix=new Array()
validsuffix[0]="com"
validsuffix[1]="org"
validsuffix[2]="net"
validsuffix[3]="biz"
validsuffix[4]="edu"
validsuffix[5]="name"
validsuffix[6]="info"
validsuffix[7]="tv"
validsuffix[8]="gov"
validsuffix[9]="usa"
validsuffix[10]="uk"

var invalidcheck=0;
var testresults

var filter=/^(\w+(?:\.\w+)*)@((?:\w+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i
if (filter.test(str)){
 
var suffix
suffix = str.substring(str.lastIndexOf(".") + 1,str.length);
 
invalidcheck=1
for (i=0;i<validsuffix.length;i++){
if (suffix == validsuffix[i])
invalidcheck=0
}
if (invalidcheck!=1)
testresults=true
else{
alert("Please input a more official email address!")
testresults=false
}
}
else{
alert("Please input a valid email address!")
testresults=false
}
return (testresults)
}



function fn_CheckCancelOrDelete(val,CancelVal,DeleteVal){
   if ( (val == CancelVal) || (val == DeleteVal) ) {
     alert ("cancel true "+val); 	
     return true;
   } else { 
     alert ("cancel false "+val);
     return false;
   }
}


function fn_CheckMaxRow (textarea1, maxrow) 
{
var cantidad = textarea1.value.match(/\n+/g);
var cuenta = cantidad?cantidad.length:0;
   	if (cuenta < maxrow) 
    	{ return true;}
  	else
	{
	//textarea1.selected = false;
	alert ("Comment maximum rows of "+maxrow+" was exceeded!");
	textarea1.value = textarea1.value.substring(0,textarea1.value.length - 2);
	return false;
	}
}



function fn_CheckMaxLength(textarea1,max){
   if (textarea1.value.length >= max) {
      alert   ("Comment maximum length of "+max+" was exceeded!");
      textarea1.value = textarea1.value.substring(0,max);
      return false;
   } else { 
     return true;
   }
}


function textareaValidation(theField,theCharCounter,theLineCounter,maxChars,maxLines,maxPerLine)
{

var strTemp = "";
var strLineCounter = 0;
var strCharCounter = 0;

for (var i = 0; i < theField.value.length; i++)
{
	var strChar = theField.value.substring(i, i + 1);

	if (strChar == '\n')
	{
		strTemp += strChar;
		strCharCounter = 1;
		strLineCounter += 1;
	}
	else if (strCharCounter == maxPerLine)
	{
		strTemp += '\n' + strChar;
		strCharCounter = 1;
		strLineCounter += 1;
	}
	else
	{
		strTemp += strChar;
		strCharCounter ++;
	}
}


if (maxChars - strTemp.length < 0) {
alert ("You have exceeded the maximum characters allowed!")
theField.value = theField.value.substring(0,theField.value.length - (theField.value.length - maxChars));
theCharCounter.value = 0;
return false;
}
if (maxLines - strLineCounter < 0) {
alert ("You have exceeded the maximum lines allowed!")
theField.value = theField.value.substring(0,theField.value.length - 2*(strLineCounter - maxLines));
return false;
}

theCharCounter.value = maxChars - strTemp.length;
theLineCounter.value = maxLines - strLineCounter;
return true;
}


function checkEmptyLine(val){
var lines=val.split('\n');
var empty=new Array(), lns=new Array();
for(i=0;i<lines.length;i++){
empty[i]=true;
for(j=0;j<lines[i].length;j++){
  if(lines[i].charCodeAt(j)!=13&&lines[i].charCodeAt(j)!=32){empty[i]=false;}}}
for(i=0;i<empty.length;i++){if(empty[i]){lns[lns.length]=i;}}
if(lns.length==1){alert('Line '+lns[0]+' is empty.');}
else if(lns.length==2){alert('Lines '+lns[0]+' and '+lns[1]+' are empty.');}
else if(lns.length>2){
 var msg='Lines '+lns[0]; for(i=1;i<lns.length-1;i++){msg+=', '+lns[i];}
 msg+=' and '+lns[lns.length-1]+' are empty.';alert(msg);}
}


// Confirm Deletion
function confirmDelete()
{
var agree=confirm("Click \"OK\" to confirm the DELETE!\nClick \"Cancel\" if you DO NOT want to DELETE!");
if (agree)
	return true ;
else
	return false ;
}


// Confirm Deletion 2
function confirmDelete2(msg)
{
var agree=confirm(msg);
if (agree)
	return true ;
else
	return false ;
}

//</script>
// --> 
