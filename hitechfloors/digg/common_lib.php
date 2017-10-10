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



#=====================================================================
#
#    Function risearch_hash($key)
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function risearch_hash($key) {

    $chars = preg_split("//",$key);
    for($i=1;$i<count($chars)-1;$i++) {
        $chars2[$i] = ord($chars[$i]);
    }
        
    $h = hexdec("00000000");
    $f = hexdec("0F000000");
    
    for($i=1;$i<count($chars)-1;$i++) {
        $h = ($h << 4) + $chars2[$i];
        if ($g = $h & $f) { $h ^= $g >> 24; };
        $h &= ~$g;
    }
    
    return $h;
    
}
#=====================================================================
#
#    Function getmicrotime()
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function getmicrotime(){ 
    list($usec, $sec) = explode(" ",microtime()); 
    return ((float)$usec + (float)$sec); 
}
#=====================================================================
#
#    Function get_META_info($html)
#    Last modified: 07.05.2005 0:03
#
#=====================================================================

function get_META_info($html) {

    preg_match("/<\s*[Mm][Ee][Tt][Aa]\s*[Nn][Aa][Mm][Ee]=\"?[Kk][Ee][Yy][Ww][Oo][Rr][Dd][Ss]\"?\s*[Cc][Oo][Nn][Tt][Ee][Nn][Tt]=\"?([^\"]*)\"?\s*\/?>/s",$html,$matches);
    $res[0] = @$matches[1];
    preg_match("/<\s*[Mm][Ee][Tt][Aa]\s*[Nn][Aa][Mm][Ee]=\"?[Dd][Ee][Ss][Cc][Rr][Ii][Pp][Tt][Ii][Oo][Nn]\"?\s*[Cc][Oo][Nn][Tt][Ee][Nn][Tt]=\"?([^\"]*)\"?\s*\/?>/s",$html,$matches);
    $res[1] = @$matches[1];

    return $res;
}
#=====================================================================
#
#    Function index_file($html_text,$url)
#    Last modified: 15.07.2004 11:35
#
#=====================================================================

function index_file($html_text,$url) {

    global $cfn, $kbcount, $descr_size, $min_length, $stop_words_array, $use_esc;
    global $use_selective_indexing, $no_index_strings;
    global $use_META, $use_META_descr;
    global $fp_FINFO;
    global $words;
    global $numbers;
    
    
    $cfn++;
    $size = strlen($html_text);
    $kbcount += intval($size/1024);
    print "$cfn -> $url; totalsize -> $kbcount kb<BR>\n";
    

    # Delete parts of document, which should not be indexed
    if ($use_selective_indexing == "YES") {
        foreach ($no_index_strings as $k => $v) {
    	    $html_text = preg_replace("/$k.*?$v/s"," ",$html_text);
    	}
    }
    
    
    $title = "";
    if (preg_match("/<title>\s*(.*?)\s*<\/title>/is",$html_text,$matches)) {
        $title = $matches[1];
    }
    $title = preg_replace("/\s+/"," ",$title);
    
    $keywords = "";
    $description = "";
    if ($use_META == "YES") { 
        $res = get_META_info($html_text);
        $keywords = $res[0];
        $description = $res[1];
    }

    $html_text = preg_replace("/<title>\s*(.*?)\s*<\/title>/is"," ",$html_text);
    $html_text = preg_replace("/<!--.*?-->/s"," ",$html_text);
    $html_text = preg_replace("/<[Ss][Cc][Rr][Ii][Pp][Tt].*?<\/[Ss][Cc][Rr][Ii][Pp][Tt]>/s"," ",$html_text);
    $html_text = preg_replace("/<[Ss][Tt][Yy][Ll][Ee].*?<\/[Ss][Tt][Yy][Ll][Ee]>/s"," ",$html_text);
    $html_text = preg_replace("/<[^>]*>/s"," ",$html_text);
    if ($use_esc == "YES") { $html_text = preg_replace_callback("/&[a-zA-Z0-9#]*?;/", 'esc2char', $html_text); }

    if (($use_META_descr == "YES") & ($description != "")) {
        $descript = substr($description,0,$descr_size);
    } else {
        $html_text = preg_replace("/\s+/s"," ",$html_text);
        $descript = substr($html_text,0,$descr_size);
    }

    $html_text = $html_text." ".$keywords." ".$description." ".$title;

    $html_text = preg_replace("/[^a-zA-Zà-ÿÀ-ß$numbers -]/"," ",$html_text);
    $html_text = preg_replace("/\s+/s"," ",$html_text);
    $html_text = strtolower($html_text);
    
    $words_temp = array();
    
    $pos = 0;
    do  {
        $new_pos = strpos($html_text," ",$pos);
        if ($new_pos === FALSE) {
            $word = substr($html_text,$pos);
            $words_temp[$word] = 1;
            break;
        };
        $word = substr($html_text,$pos,$new_pos-$pos);
        $words_temp[$word] = 1;
        $pos = $new_pos+1;
    } while (1>0);

    

    $title = preg_replace("/:+/",":",$title);
    $descript = preg_replace("/:+/",":",$descript);
    if ($title == "") { $title = "No title"; }
    $pos = ftell($fp_FINFO);
    $pos = pack("N",$pos);
    fwrite($fp_FINFO, "$url::$size::$title::$descript\x0A");
    
    foreach($words_temp as $word => $val) {
        if (strlen($word) < $min_length) { continue; }
        if (array_key_exists($word,$stop_words_array)) { continue; }
        @$words[$word] .= $pos;
    }    
    
    
    unset($words_temp);
    unset($words_temp2);
    
}
#=====================================================================
#
#    Function build_hash()
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function build_hash() {

    global $words;
    global $HASHSIZE, $INDEXING_SCHEME, $HASH, $HASHWORDS;

    
    for ($i=0; $i<$HASHSIZE; $i++) {$hash_array[$i] = "";};

    foreach($words as $word=>$value) {
        if ($INDEXING_SCHEME == 3) { $subbound = strlen($word)-3; }
        else { $subbound = 1; }
        if (strlen($word)==3) {$subbound = 1;}
        $substring_length = 4;
        if ($INDEXING_SCHEME == 1) { $substring_length = strlen($word); }

        for ($i=0; $i<$subbound; $i++){
            $hash_value = abs(risearch_hash(substr($word,$i,$substring_length)) % $HASHSIZE);
    	    $hash_array[$hash_value] .= $value;
    	};   
        
    }



    $fp_HASH = fopen ("$HASH", "wb") or die("Can't open index file!");
    $fp_HASHWORDS = fopen ("$HASHWORDS", "wb") or die("Can't open index file!");

    $zzz = pack("N", 0);
    fwrite($fp_HASHWORDS, $zzz);
    $pos_hashwords = ftell($fp_HASHWORDS);
    $to_print_hash = "";
    $to_print_hashwords = "";

    for ($i=0; $i<$HASHSIZE; $i++){
    	
        if ($hash_array[$i] == "") {$to_print_hash .= $zzz;};
        if ($hash_array[$i] != "") {
            $to_print_hash .= pack("N",$pos_hashwords + strlen($to_print_hashwords));
            $to_print_hashwords .= pack("N", strlen($hash_array[$i])/8).$hash_array[$i];
        };   
        if (strlen($to_print_hashwords) > 64000) {
            fwrite($fp_HASH,$to_print_hash);
            fwrite($fp_HASHWORDS,$to_print_hashwords);
            $to_print_hash = "";
            $to_print_hashwords = "";
            $pos_hashwords  = ftell($fp_HASHWORDS);
        }
    }; # for $i
    fwrite($fp_HASH,$to_print_hash);
    fwrite($fp_HASHWORDS,$to_print_hashwords);
    
fclose($fp_HASH);
fclose($fp_HASHWORDS);


}
#=====================================================================




?>