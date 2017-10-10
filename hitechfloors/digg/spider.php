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

start_spidering();

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
#    Function start_spidering()
#    Last modified: 16.04.2004 18:26
#
#=====================================================================

function start_spidering() {

    global $start_url, $allow_url;

foreach ($start_url as $v) {
    $to_visit[$v] = 1;
}
$visited = array();

do {

    if (count($to_visit) == 0) {
        break;
    } else {
        list ($url,) = each($to_visit);
    }
    
    




    $fp = @fopen($url,"r");
    $visited[$url] = 1;

    if ( $fp == FALSE ) {
        print "Error in opening file: $url<BR>\n";
        unset($to_visit[$url]);
    } else {
        $text = "";
        while (!feof ($fp)) {
            $text .= fgets($fp, 4096);
        }
        print "URL: $url - ".strlen($text)." bytes<BR>\n";
        
        $base = $url;
        if (preg_match_all("/<base\\s+href=([\"']?)([^\\s\"'>]+)\\1/is", $text, $matches,PREG_SET_ORDER)) {
            $base = $matches[0][2];
        }
        
        $links = get_link($text);
        foreach ($links as $k => $v) {
            $new_link = get_absolute_url($base,$k);
            $new_link = preg_replace("/#.*/","",$new_link);
#            $new_link_stripped = preg_replace("/\?.*/","",$new_link);
            if ( check_url($new_link)) {
                if ( ! array_key_exists($new_link,$visited)) {
                    $to_visit[$new_link] = 1;
                }
            }
        }

        index_file($text,$url);
        
        unset($to_visit[$url]);
    }


} while (1);


}
#=====================================================================
#
#    Function get_link($text)
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function get_link($text) {
    
    $links = array();
    $count = preg_match_all("/<a[^>]+href=([\"']?)([^\\s\"'>]+)\\1/is", $text, $matches, PREG_SET_ORDER);
    for($i=0; $i < count($matches); $i++) {
        $links[$matches[$i][2]] = 1;
    }

    $count = preg_match_all("/<frame[^>]+src=([\"']?)([^\\s\"'>]+)\\1/is", $text, $matches, PREG_SET_ORDER);
    for($i=0; $i < count($matches); $i++) {
        $links[$matches[$i][2]] = 1;
    }

    $count = preg_match_all("/<area[^>]+href=([\"']?)([^\\s\"'>]+)\\1/is", $text, $matches, PREG_SET_ORDER);
    for($i=0; $i < count($matches); $i++) {
        $links[$matches[$i][2]] = 1;
    }


    return $links;
}

#=====================================================================
#
#    Function get_absolute_url($base,$url)
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function get_absolute_url($base,$url) {

    $url_arr = parse_url($url);
    if (isset($url_arr["scheme"])) {
        return($url);
    }
    
    $base_arr = parse_url($base);
    $base_base = strtolower($base_arr["scheme"])."://";
    if (isset($base_arr["user"])) {
        $base_base .= $base_arr["user"].":".$base_arr["pass"]."@";
    }
    $base_base .= strtolower($base_arr["host"]);
    if (isset($base_arr["port"])) {
        $base_base .= ":".$base_arr["port"];
    }
    $base_path = @$base_arr["path"];
    if ($base_path == "") { $base_path = "/"; }
    $base_path = preg_replace("/(.*\/).*/","\\1",$base_path);
    
    if (@$url_arr["path"][0] == "/") {
        return $base_base.$url;
    }
    
    if (preg_match("'^\./'",$url)) {
        $url = preg_replace("'^\./'","",$url);
        return $base_base.$base_path.$url;
    }
    
    while (preg_match("'^\.\./'",$url)) {
        $url = preg_replace("'^\.\./'","",$url);
        $base_path = preg_replace("/(.*\/).*\//","\\1",$base_path);
    }
    return $base_base.$base_path.$url;    
}
#=====================================================================
#
#    Function check_url($url)
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function check_url($url) {
    
    global $file_ext, $no_index_files, $no_index_dir, $allow_url;

    if ( ! preg_match("'^http://'",$url)) { return FALSE; }
    if ( ! preg_match ("'$file_ext'i", $url)) { return FALSE; }
    if ( preg_match ("'$no_index_files'i", $url)) { return FALSE; }
    if ( preg_match ("'$no_index_dir'i", $url)) { return FALSE; }
    
    $allow = 0;
    foreach ($allow_url as $v) {
        if ( preg_match("'$v'i", $url)) {
            $allow = 1;
            break;
        }
    }
    if ($allow == 0) { return FALSE; }
    
    return TRUE;
}
#=====================================================================



?>