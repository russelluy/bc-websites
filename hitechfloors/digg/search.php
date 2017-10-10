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




$time1 = getmicrotime();

read_template("template.html");



$stpos = 0;
$stype = "AND";
$query = "";

$abort = 0;

get_query();
if (count($query_arr) > 0) {
    get_results();
    $time3 = getmicrotime();
    $time = $time3-$time1;
#    print "<BR>get_results() took $time sec.<BR>";
        
    boolean();
    $time4 = getmicrotime();
    $time = $time4-$time3;
    $search_time = $time4-$time1;
    $search_time = sprintf("%2.4f", $search_time);
#    print "<BR>boolean() took $time sec.<BR>";
}


print print_template("header");

if (count($query_arr) > 0) {
    if ($rescount>0) {
        print print_template("results_header");
        print_results();
        print print_template("results_footer");
    } else {
        print print_template("no_results");
    }
} else {
    print print_template("empty_query");
}


print print_template("footer");


#=====================================================================
#
#    Function get_query()
#    Last modified: 25.08.2005 22:04
#
#=====================================================================

function get_query() {

    global $HTTP_GET_VARS;
    global $query, $stpos, $stype, $query_arr, $wholeword, $querymode, $stop_words_array;
    global $min_length;
    global $numbers;
    
    $query = isset($HTTP_GET_VARS["query"])?$HTTP_GET_VARS["query"]:$query;
    $stpos = isset($HTTP_GET_VARS["stpos"])?$HTTP_GET_VARS["stpos"]:$stpos;
    $stype = isset($HTTP_GET_VARS["stype"])?$HTTP_GET_VARS["stype"]:$stype;
    
    
    $query = strtolower($query);
    $query = preg_replace("/[^a-zà-ÿ$numbers +!-]/"," ",$query);
    $query_arr_dum = preg_split("/\s+/",$query);

    foreach($query_arr_dum as $word) {
        if (strlen($word) < $min_length) { continue; }
        if (array_key_exists($word,$stop_words_array)) { continue; }
        $query_arr[] = $word;
    }    

    for ($i=0; $i<count($query_arr); $i++) {
        $wholeword[$i] = 0;
        $querymode[$i] = 0;
        if (preg_match("/\!/", $query_arr[$i]))   { $wholeword[$i] = 1;} # WholeWord
        $query_arr[$i] = preg_replace("/[\! ]/","",$query_arr[$i]);
        if ($stype == "AND")     { $querymode[$i] = 2;} # AND
        if (preg_match ("/^\-/", $query_arr[$i])) { $querymode[$i] = 1;} # NOT
        if (preg_match ("/^\+/", $query_arr[$i])) { $querymode[$i] = 2;} # AND
        $query_arr[$i] = preg_replace("/^[\+\- ]/","",$query_arr[$i]);
    }
    
    if ($stpos <0) {$stpos = 0;};    
}
#=====================================================================
#
#    Function get_results()
#    Last modified: 10.05.2004 18:43
#
#=====================================================================

function get_results() {

    global $HASHSIZE, $INDEXING_SCHEME, $HASH, $HASHWORDS, $FINFO, $SITEWORDS, $WORD_IND;
    
    global $query_arr, $wholeword, $querymode;
    global $res, $allres, $rescount, $query_statistics;

    
    $fp_HASH = fopen ("$HASH", "rb") or die("No index file is found! Please run indexing script again.");
    $fp_HASHWORDS = fopen ("$HASHWORDS", "rb") or die("No index file is found! Please run indexing script again.");
    $fp_SITEWORDS = fopen ("$SITEWORDS", "rb") or die("No index file is found! Please run indexing script again.");
    $fp_WORD_IND = fopen ("$WORD_IND", "rb") or die("No index file is found! Please run indexing script again.");



for ($j=0; $j<count($query_arr); $j++) {
    $query = $query_arr[$j];
    $allres[$j] = array();

    if ($INDEXING_SCHEME == 1) {
    	$substring_length = strlen($query);
    } else {
    	$substring_length = 4;
    }
    $hash_value = abs(risearch_hash(substr($query,0,$substring_length)) % $HASHSIZE);
    
    fseek($fp_HASH,$hash_value*4,0);
    $dum = fread($fp_HASH,4);
    $dum = unpack("Ndum", $dum);
    fseek($fp_HASHWORDS,$dum['dum'],0);
    $dum = fread($fp_HASHWORDS,4);
    $dum1 = unpack("Ndum", $dum);
    
    for ($i=0; $i<$dum1['dum']; $i++) {
        $dum = fread($fp_HASHWORDS,8);
        $arr_dum = unpack("Nwordpos/Nfilepos",$dum);
        fseek($fp_SITEWORDS,$arr_dum['wordpos'],0);
        $word = fgets($fp_SITEWORDS,1024);
        $word = preg_replace("/\x0A/","",$word);
        $word = preg_replace("/\x0D/","",$word);
        
        if ( ($wholeword[$j]==1) && ($word != $query) ) {$word = "";};
        $pos = strpos($word, $query);
        if ($pos !== false) {
            fseek($fp_WORD_IND,$arr_dum['filepos'],0);
            $dum = fread($fp_WORD_IND,4);
            $dum2 = unpack("Ndum",$dum);
            $dum = fread($fp_WORD_IND,$dum2['dum']*4);
            for($k=0; $k<$dum2['dum']; $k++){
                $zzz = unpack("Ndum",substr($dum,$k*4,4));
                $allres[$j][$zzz['dum']] = 1;
            }
        }
            
    };   


}


    for ($j=0; $j<count($query_arr); $j++) {
    	$found_number = count($allres[$j]);
        $query_statistics .= " $query_arr[$j]-$found_number\n";
    }


}
#=====================================================================
#
#    Function boolean()
#    Last modified: 10.05.2004 18:43
#
#=====================================================================

function boolean() {

    global $query_arr, $querymode, $stype;
    global $res, $allres, $rescount;


if (count($query_arr) == 1) {
    foreach ($allres[0] as $k => $v) {
        if ($k) {
            $res .= pack("N",$k);
        }
    }
    $rescount = intval(strlen($res)/4);
    unset($allres);
    return;
} else {

    if ($stype == "AND") {
        for ($i=0; $i<count($query_arr); $i++) {
            if ($querymode[$i] == 2) {
                $min = $i;
                break;
            }
        }
        for ($i=$min+1; $i<count($query_arr); $i++) {
            if (count($allres[$i]) < count($allres[$min]) && $querymode[$i] == 2) {
                $min = $i;
            }
        }
        for ($i=0; $i<count($query_arr); $i++) {
            if ($i == $min) {
                continue;
            }
            if ($querymode[$i] == 2) {
                foreach ($allres[$min] as $k => $v) {
                    if (array_key_exists($k,$allres[$i])) {
                    } else {
                        unset($allres[$min][$k]);
                    }
                }
            } else {
                foreach ($allres[$min] as $k => $v) {
                    if (array_key_exists($k,$allres[$i])) {
                        unset($allres[$min][$k]);
                    }
                }
            }
        }
        foreach ($allres[$min] as $k => $v) {
            if ($k) {
                $res .= pack("N",$k);
            }
        }
        $rescount = intval(strlen($res)/4);
        return;
    }
    
    
    if ($stype == "OR") {
        for ($i=0; $i<count($query_arr); $i++) {
            if ($querymode[$i] != 1) {
                $max = $i;
                break;
            }
        }
        for ($i=$max+1; $i<count($query_arr); $i++) {
            if (count($allres[$i]) > count($allres[$max]) && $querymode[$i] != 1) {
                $max = $i;
            }
        }
        for ($i=0; $i<count($query_arr); $i++) {
            if ($i == $max) {
                continue;
            }
            if ($querymode[$i] != 1) {
                foreach ($allres[$i] as $k => $v) {
                    $allres[$max][$k] = 1;
                }
            } else {
                foreach ($allres[$i] as $k => $v) {
                    if (array_key_exists($k,$allres[$max])) {
                        unset($allres[$max][$k]);
                    }
                }
            }
        }
        foreach ($allres[$max] as $k => $v) {
            if ($k) {
                $res .= pack("N",$k);
            }
        }
        $rescount = intval(strlen($res)/4);
        return;
    }
    
}
    

}
#=====================================================================
#
#    Function print_results()
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function print_results() {

    global $FINFO, $FINFO_IND, $query, $stpos, $stype, $res_num, $res;
    global $url, $title, $size, $description, $rescount, $next_results;
    global $query_arr;

    $time1 = getmicrotime();

    $fp_FINFO = fopen ("$FINFO", "rb") or die("No index file is found! Please run indexing script again.");

    for ($i=$stpos; $i<$stpos+$res_num; $i++) {
        if ($i >= strlen($res)/4) {break;};
        $strpos = unpack("Npos",substr($res,$i*4,4));
        fseek($fp_FINFO,$strpos['pos'],0);
        $dum = fgets($fp_FINFO,4024);
        list($url, $size, $title, $description) = explode("::",$dum);
        for ($j=0; $j<count($query_arr); $j++) {
            $tquery = $query_arr[$j];
            $description = preg_replace ("'\b($tquery)\b'i", "<b style='color:black;background-color:#ffff66'>$1</b>", $description);
        }
        print print_template("results");
    };  # for



    if ($rescount <= $res_num) {$next_results = ""; return 1;}
    

    $mhits = 20 * $res_num;
    $pos2 = $stpos - $stpos % $mhits;
    $pos1 = $pos2 - $mhits;
    $pos3 = $pos2 + $mhits;

    if ($pos1 < 0) { $prev = ""; }
    else {
        $prev = " <A HREF=search.php?query=".urlencode($query)."&stpos=".$pos1."&stype=".$stype;
        $prev .= ">PREV</A> \n";
    }

    if ($pos3 > $rescount) { $next = ""; }
    else {
        $next = " <A HREF=search.php?query=".urlencode($query)."&stpos=".$pos3."&stype=".$stype;
        $next .= ">NEXT</A> \n";
    }

    $next_results .= $prev;
    $next_results .=  " |\n";
    for ($i=$pos2; $i<$pos3; $i += $res_num) {
       if ($i >= $rescount) {break;}
       $page_number = $i/$res_num+1;
       if ( $i != $stpos ) {
           $next_results .=  "<A HREF=search.php?query=".urlencode($query)."&stpos=".$i."&stype=".$stype;
           $next_results .=  ">".$page_number."</A> |\n";
       } else {
           $next_results .=  $page_number." |\n";
       }
    }
    $next_results .=  $next;
    


}
#=====================================================================
#
#    Function read_template($filename)
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function read_template($filename) {

$size = filesize($filename);
$fd = @fopen ($filename, "rb") or die("Template file is not found!");
$template = fread ($fd, $size);
fclose ($fd);

global $templates;

    $count = preg_match_all("/<!-- RiSearch::([^:]+?)::start -->(.*?)<!-- RiSearch::\\1::end -->/s", $template, $matches, PREG_SET_ORDER);
    for($i=0; $i < count($matches); $i++) {
        $templates[$matches[$i][1]] = $matches[$i][2];
    }
    
}
#=====================================================================
#
#    Function print_template($part)
#    Last modified: 16.04.2004 17:54
#
#=====================================================================

function print_template($part) {

    global $templates;
    global $query, $search_time, $query_statistics, $stpos, $url, $title, $size, $description, $rescount, $next_results;
    $template = $templates[$part];      
    
    
    $template = preg_replace("/%query%/s","$query",$template);
    $template = preg_replace("/%search_time%/s","$search_time",$template);
    $template = preg_replace("/%query_statistics%/s","$query_statistics",$template);
    $template = preg_replace("/%stpos%/s",$stpos+1,$template);
    $template = preg_replace("/%url%/s","$url",$template);
    $template = preg_replace("/%title%/s","$title",$template);
    $template = preg_replace("/%size%/s","$size",$template);
    $template = preg_replace("/%description%/s","$description",$template);
    $template = preg_replace("/%rescount%/s","$rescount",$template);
    $template = preg_replace("/%next_results%/s","$next_results",$template);
    
    return $template;
}
#===================================================================


?>