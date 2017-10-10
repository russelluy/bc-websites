<?php
/*
 * @className   : Validator
 * @purpose     : Performs error checking from the given inputs..
 * @version 	: 1.0 10-Sept-2005 2:22pm
 * @author 	    : mindh4c]k[
 * @todaysQuotes:  Vestibulum porta mauris accumsan mauris. In luctus malesuada dui. Praesent faucibus, purus eget sagittis mollis, leo nunc egestas sapien, ut molestie nisl magna id leo
*/
class Validator  {
	var $errors;
	var $allowBlanks = false;
	
	// Validate email address  
	function validateEmail($email, $description =''){
	    if ($this->allowBlanks && $email == ''){
			return true;
		}
		$result = eregi("^[a-z0-9]+[a-z0-9_-]*(\.[a-z0-9_-]+)*@[a-z0-9_-]+(\.[a-z0-9_-]+)*\.([a-z]+){2,}$", $email); 
		if ($result) {
		   return true;	
		}  								
	    else {
		   return false;
		}
	}
	// Validate text only, with spaces
	function validateTextOnlyWithSpaces($str){
		if ($this->allowBlanks && $str == ''){ 
			return true; 
		}
		$result = ereg("^[A-Za-z0-9 ]+$",$str);
		if ($result) {
		   return true;	
		}  								
	    else {
		   return false;
		}
	}
	// Validate text only, no spaces
	function validateTextOnlyNoSpaces($str){
		if ($this->allowBlanks && $str == ''){ 
			return false; 
		}
		$result = ereg("^[A-Za-z]+$",$str);
		if ($result) {
		   return true;	
		}  								
	    else {
		   return false;
		}
	}

	function validateFieldIsEmpty($field){
		if ((trim($field) == "")) {
			return true;
		}
		else{
			return false;
		}
	}
	
}
?>