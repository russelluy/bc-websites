<?php
$mailto='info@hitechhardwoodfloor.com';
$subject_header='Contact US info from '.$_POST['name'];
$headers = 'To: info@hitechhardwoodfloor.com \r\n';
$headers .= 'From: '.$_POST['email']."\r\n";
$headers .= 'Subject: '.$subject_header."\r\n";
$message .= $_POST['FomedData'];
$sent=mail($mailto, $subject_header , $message, $headers);
if($sent)
{
	echo'<script language="javascript" type="text/javascript">alert("Contact Request Successfully Sent");</script>';
}else{
	echo'<script language="javascript" type="text/javascript">alert("Error In Sending Contact Request");</script>';
}
?>