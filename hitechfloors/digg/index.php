<?

error_reporting(E_ALL);

/**
 *           RiSearch PHP
 * 
 * web search engine, version 0.2
 * (c) Sergej Tarasov, 2000-2004
 * 
 * Homepage: http://risearch.org/
 * email: risearch@risearch.org
 */


include "config.php";
include "common_lib.php";


print "Start indexing<BR>\n";


print "\$base_dir: $base_dir<BR>\n";
print "\$base_url: $base_url<BR>\n";

#DEFINE CONSTANTS
$cfn = 0;
$cwn = 0;
$kbcount = 0;

if(!is_dir("db")) {

	mkdir("db",0755) or die("Can't create directory DB!!!");
	echo"Directory 'db' has been created";

}



$fp_FINFO = fopen ("$FINFO", "wb") or die("Can't open index file!");
fwrite($fp_FINFO, "\x0A");
$fp_SITEWORDS = fopen ("$SITEWORDS", "wb") or die("Can't open index file!");
$fp_WORD_IND = fopen ("$WORD_IND", "wb") or die("Can't open index file!");

$time1 = getmicrotime();

scan_files($base_dir);

$time2 = getmicrotime();
$time = $time2-$time1;
print "<BR>Scan took $time sec.<BR>";


if ($cfn == 0) {
    print "No files are indexed\n\n";
    die;
}


print "Writing SITEWORDS\n";
    $pos_sitewords = ftell($fp_SITEWORDS);
    $pos_word_ind  = ftell($fp_WORD_IND);
    $to_print_sitewords = "";
    $to_print_word_ind  = "";
    foreach($words as $word=>$value) {
        $cwn++;
        $words_word_dum = pack("NN",$pos_sitewords+strlen($to_print_sitewords),
    	                        $pos_word_ind+strlen($to_print_word_ind));
    	$to_print_sitewords .= "$word\x0A";
    	$to_print_word_ind .= pack("N",strlen($value)/4).$value;
    	$words[$word] = $words_word_dum;
    	if (strlen($to_print_word_ind) > 32000) {
    	    fwrite($fp_SITEWORDS, $to_print_sitewords);
    	    fwrite($fp_WORD_IND, $to_print_word_ind);
    	    $to_print_sitewords = "";
    	    $to_print_word_ind  = "";
    	    $pos_sitewords = ftell($fp_SITEWORDS);
    	    $pos_word_ind  = ftell($fp_WORD_IND);
    	}

    }
    fwrite($fp_SITEWORDS, $to_print_sitewords);
    fwrite($fp_WORD_IND, $to_print_word_ind);
fclose($fp_SITEWORDS);
fclose($fp_WORD_IND);

print "Build hash\n";

build_hash();

print "$cfn files are indexed\n";

#=====================================================================
#
#    Function scan_files ($dir)
#    Last modified: 05.04.2005 16:41
#
#=====================================================================

function  scan_files ($dir) {

    global $base_dir, $base_url, $cfn;
    global $no_index_dir, $file_ext, $cut_default_filenames, $default_filenames, $url_to_lower_case, $no_index_files;

    $dir_h = opendir($dir) or die("Can't open $dir");
    
    while (false !== ($file = readdir($dir_h))) { 
        if ($file != "." && $file != "..") {
            $new_dir = $dir."/".$file;
            if ( is_dir($new_dir)) {
                if (preg_match ("'$no_index_dir'i", $new_dir)) { continue; }
                scan_files($new_dir);
            } else {
                if (preg_match ("'$file_ext'i", $new_dir)) {
                    $url = preg_replace ("'^$base_dir/'", "$base_url", $new_dir);
                    if (preg_match ("'$no_index_files'i", $url)) { continue; };
                    if ($cut_default_filenames == "YES") {
                        $url = preg_replace ("'$default_filenames'i", "/", $url);
                    }
                    if ($url_to_lower_case == "YES") {
                        $url = strtolower($url);
                    }
                    if ($fd = fopen ($new_dir, "rb") or print "Can't open file: $new_dir<BR>\n") {
                        $size = filesize($new_dir);
                        $html_text = @fread ($fd, $size);
                        fclose ($fd);
                        index_file($html_text,$url);
                    }
                }
                
            }
        }
    }
    closedir($dir_h);

}
#=====================================================================


?>