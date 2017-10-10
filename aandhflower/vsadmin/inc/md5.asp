<SCRIPT LANGUAGE=JScript RUNAT=SERVER> 
/*
 * A JavaScript implementation of the RSA Data Security, Inc. MD5 Message
 * Digest Algorithm, as defined in RFC 1321.
 * Version 1.1 Copyright (C) Paul Johnston 1999 - 2002.
 * Code also contributed by Greg Holt
 * See http://pajhome.org.uk/site/legal.html for details.
 */
/*
 * Add integers, wrapping at 2^32. This uses 16-bit operations internally
 * to work around bugs in some JS interpreters.
 */
function safe_add(x, y){
  var lsw = (x & 0xFFFF) + (y & 0xFFFF)
  var msw = (x >> 16) + (y >> 16) + (lsw >> 16)
  return (msw << 16) | (lsw & 0xFFFF)
}
/*
 * Bitwise rotate a 32-bit number to the left.
 */
function rol(num, cnt){
  return (num << cnt) | (num >>> (32 - cnt))
}
/*
 * These functions implement the four basic operations the algorithm uses.
 */
function cmn(q, a, b, x, s, t){
  return safe_add(rol(safe_add(safe_add(a, q), safe_add(x, t)), s), b)
}
function ffxx(a, b, c, d, x, s, t){
  return cmn((b & c) | ((~b) & d), a, b, x, s, t)
}
function ggxx(a, b, c, d, x, s, t){
  return cmn((b & d) | (c & (~d)), a, b, x, s, t)
}
function hhxx(a, b, c, d, x, s, t){
  return cmn(b ^ c ^ d, a, b, x, s, t)
}
function iixx(a, b, c, d, x, s, t){
  return cmn(c ^ (b | (~d)), a, b, x, s, t)
}
/*
 * Calculate the MD5 of an array of little-endian words, producing an array
 * of little-endian words.
 */
function coreMD5(x){
  var a =  1732584193
  var b = -271733879
  var c = -1732584194
  var d =  271733878

  for(i = 0; i < x.length; i += 16){
    var olda = a
    var oldb = b
    var oldc = c
    var oldd = d

    a = ffxx(a, b, c, d, x[i+ 0], 7 , -680876936)
    d = ffxx(d, a, b, c, x[i+ 1], 12, -389564586)
    c = ffxx(c, d, a, b, x[i+ 2], 17,  606105819)
    b = ffxx(b, c, d, a, x[i+ 3], 22, -1044525330)
    a = ffxx(a, b, c, d, x[i+ 4], 7 , -176418897)
    d = ffxx(d, a, b, c, x[i+ 5], 12,  1200080426)
    c = ffxx(c, d, a, b, x[i+ 6], 17, -1473231341)
    b = ffxx(b, c, d, a, x[i+ 7], 22, -45705983)
    a = ffxx(a, b, c, d, x[i+ 8], 7 ,  1770035416)
    d = ffxx(d, a, b, c, x[i+ 9], 12, -1958414417)
    c = ffxx(c, d, a, b, x[i+10], 17, -42063)
    b = ffxx(b, c, d, a, x[i+11], 22, -1990404162)
    a = ffxx(a, b, c, d, x[i+12], 7 ,  1804603682)
    d = ffxx(d, a, b, c, x[i+13], 12, -40341101)
    c = ffxx(c, d, a, b, x[i+14], 17, -1502002290)
    b = ffxx(b, c, d, a, x[i+15], 22,  1236535329)

    a = ggxx(a, b, c, d, x[i+ 1], 5 , -165796510)
    d = ggxx(d, a, b, c, x[i+ 6], 9 , -1069501632)
    c = ggxx(c, d, a, b, x[i+11], 14,  643717713)
    b = ggxx(b, c, d, a, x[i+ 0], 20, -373897302)
    a = ggxx(a, b, c, d, x[i+ 5], 5 , -701558691)
    d = ggxx(d, a, b, c, x[i+10], 9 ,  38016083)
    c = ggxx(c, d, a, b, x[i+15], 14, -660478335)
    b = ggxx(b, c, d, a, x[i+ 4], 20, -405537848)
    a = ggxx(a, b, c, d, x[i+ 9], 5 ,  568446438)
    d = ggxx(d, a, b, c, x[i+14], 9 , -1019803690)
    c = ggxx(c, d, a, b, x[i+ 3], 14, -187363961)
    b = ggxx(b, c, d, a, x[i+ 8], 20,  1163531501)
    a = ggxx(a, b, c, d, x[i+13], 5 , -1444681467)
    d = ggxx(d, a, b, c, x[i+ 2], 9 , -51403784)
    c = ggxx(c, d, a, b, x[i+ 7], 14,  1735328473)
    b = ggxx(b, c, d, a, x[i+12], 20, -1926607734)

    a = hhxx(a, b, c, d, x[i+ 5], 4 , -378558)
    d = hhxx(d, a, b, c, x[i+ 8], 11, -2022574463)
    c = hhxx(c, d, a, b, x[i+11], 16,  1839030562)
    b = hhxx(b, c, d, a, x[i+14], 23, -35309556)
    a = hhxx(a, b, c, d, x[i+ 1], 4 , -1530992060)
    d = hhxx(d, a, b, c, x[i+ 4], 11,  1272893353)
    c = hhxx(c, d, a, b, x[i+ 7], 16, -155497632)
    b = hhxx(b, c, d, a, x[i+10], 23, -1094730640)
    a = hhxx(a, b, c, d, x[i+13], 4 ,  681279174)
    d = hhxx(d, a, b, c, x[i+ 0], 11, -358537222)
    c = hhxx(c, d, a, b, x[i+ 3], 16, -722521979)
    b = hhxx(b, c, d, a, x[i+ 6], 23,  76029189)
    a = hhxx(a, b, c, d, x[i+ 9], 4 , -640364487)
    d = hhxx(d, a, b, c, x[i+12], 11, -421815835)
    c = hhxx(c, d, a, b, x[i+15], 16,  530742520)
    b = hhxx(b, c, d, a, x[i+ 2], 23, -995338651)

    a = iixx(a, b, c, d, x[i+ 0], 6 , -198630844)
    d = iixx(d, a, b, c, x[i+ 7], 10,  1126891415)
    c = iixx(c, d, a, b, x[i+14], 15, -1416354905)
    b = iixx(b, c, d, a, x[i+ 5], 21, -57434055)
    a = iixx(a, b, c, d, x[i+12], 6 ,  1700485571)
    d = iixx(d, a, b, c, x[i+ 3], 10, -1894986606)
    c = iixx(c, d, a, b, x[i+10], 15, -1051523)
    b = iixx(b, c, d, a, x[i+ 1], 21, -2054922799)
    a = iixx(a, b, c, d, x[i+ 8], 6 ,  1873313359)
    d = iixx(d, a, b, c, x[i+15], 10, -30611744)
    c = iixx(c, d, a, b, x[i+ 6], 15, -1560198380)
    b = iixx(b, c, d, a, x[i+13], 21,  1309151649)
    a = iixx(a, b, c, d, x[i+ 4], 6 , -145523070)
    d = iixx(d, a, b, c, x[i+11], 10, -1120210379)
    c = iixx(c, d, a, b, x[i+ 2], 15,  718787259)
    b = iixx(b, c, d, a, x[i+ 9], 21, -343485551)

    a = safe_add(a, olda)
    b = safe_add(b, oldb)
    c = safe_add(c, oldc)
    d = safe_add(d, oldd)
  }
  return [a, b, c, d]
}
/*
 * Convert an array of little-endian words to a hex string.
 */
function binl2hex(binarray){
  var hex_tab = "0123456789abcdef"
  var str = ""
  for(var i = 0; i < binarray.length * 4; i++)
  {
    str += hex_tab.charAt((binarray[i>>2] >> ((i%4)*8+4)) & 0xF) +
           hex_tab.charAt((binarray[i>>2] >> ((i%4)*8)) & 0xF)
  }
  return str
}
/*
 * Convert an array of little-endian words to a base64 encoded string.
 */
function binl2b64(binarray){
  var tab = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  var str = ""
  for(var i = 0; i < binarray.length * 32; i += 6){
    str += tab.charAt(((binarray[i>>5] << (i%32)) & 0x3F) |
                      ((binarray[i>>5+1] >> (32-i%32)) & 0x3F))
  }
  return str
}
/*
 * Convert an 8-bit character string to a sequence of 16-word blocks, stored
 * as an array, and append appropriate padding for MD4/5 calculation.
 * If any of the characters are >255, the high byte is silently ignored.
 */
function str2binl(str){
  var nblk = ((str.length + 8) >> 6) + 1 // number of 16-word blocks
  var blks = new Array(nblk * 16)
  for(var i = 0; i < nblk * 16; i++) blks[i] = 0
  for(var i = 0; i < str.length; i++)
    blks[i>>2] |= (str.charCodeAt(i) & 0xFF) << ((i%4) * 8)
  blks[i>>2] |= 0x80 << ((i%4) * 8)
  blks[nblk*16-2] = str.length * 8
  return blks
}
/*
 * Convert a wide-character string to a sequence of 16-word blocks, stored as
 * an array, and append appropriate padding for MD4/5 calculation.
 */
function strw2binl(str){
  var nblk = ((str.length + 4) >> 5) + 1 // number of 16-word blocks
  var blks = new Array(nblk * 16)
  for(var i = 0; i < nblk * 16; i++) blks[i] = 0
  for(var i = 0; i < str.length; i++)
    blks[i>>1] |= str.charCodeAt(i) << ((i%2) * 16)
  blks[i>>1] |= 0x80 << ((i%2) * 16)
  blks[nblk*16-2] = str.length * 16
  return blks
}
/*
 * External interface
 */
function hexMD5 (str) { return binl2hex(coreMD5( str2binl(str))) }
function hexMD5w(str) { return binl2hex(coreMD5(strw2binl(str))) }
function b64MD5 (str) { return binl2b64(coreMD5( str2binl(str))) }
function b64MD5w(str) { return binl2b64(coreMD5(strw2binl(str))) }
/* Backward compatibility */
function calcMD5(str) { return binl2hex(coreMD5( str2binl(str))) }

function binl2byt(binarray){
var hex_tab = "0123456789abcdef";
var  bytarray = new Array(binarray.length * 4);
var str = "";
for(var i = 0; i < binarray.length * 4; i++){
bytarray[i] = (binarray[i>>2] >> ((i%4)*8+4) & 0xF) << 4 | binarray[i>>2] >> ((i%4)*8) & 0xF;
}
return bytarray;
}
function bytarray2word (barray){
var blks = new Array(barray.length / 4);
for(var i = 0; i < blks.length; i++) blks[i] = 0
for(var i = 0; i < barray.length; i++)
blks[i>>2] |= (barray[i] & 0xFF) << ((i%4) * 8)
//blks[i>>2] |= 0x80 << ((i%4) * 8)
//blks[nblk*16-2] = barray.length * 8
return blks
}
function bytarray2binl (barray){
var nblk = ((barray.length + 8) >> 6) + 1 // number of 16-word blocks
var blks = new Array(nblk * 16)
for(var i = 0; i < nblk * 16; i++) blks[i] = 0
for(var i = 0; i < barray.length; i++)
blks[i>>2] |= (barray[i] & 0xFF) << ((i%4) * 8)
blks[i>>2] |= 0x80 << ((i%4) * 8)
blks[nblk*16-2] = barray.length * 8
return blks
}
function b_calcMD5(barray) { return coreMD5(bytarray2binl(barray)) }
function HMAC(key, text){
var hkey,idata,odata;
var ipad= new Array(64);
var opad= new Array (64);
idata = new Array (64 + text.length);
odata = new Array (64 + 16);
if (key.length > 64){
	hkey = calcMD5(key);
}
else
	hkey = key;
for (i=0;i<64;i++){
	idata[i] = ipad[i] = 0x36;
	odata[i] = opad[i] = 0x5C;
}
for(i=0;i<hkey.length; i++){
	ipad[i] ^= hkey.charCodeAt(i);
	opad[i] ^= hkey.charCodeAt(i);
	idata[i]= ipad[i] & 0xFF;
	odata[i] = opad[i] & 0xFF;
}
for (i=0;i<text.length;i++) {
	idata[64+i] = text.charCodeAt(i) & 0xFF;
}
var innerhashout = binl2byt(b_calcMD5(idata));
for (i=0;i<16;i++) {
odata[64+i] = innerhashout[i];
}
return binl2hex(b_calcMD5(odata));
}
function GetSecondsSince1970(){
var d = new Date();
var secs= Math.floor(d.getTime() / 1000);
return (secs);
}
function hmac_sha1(key, text){
var hkey,idata,odata;
var ipad= new Array(64);
var opad= new Array (64);
idata = new Array (64 + text.length);
odata = new Array (64 + 16);
if (key.length > 64){
	hkey = sha1(key);
}
else
	hkey = key;
for (i=0;i<64;i++){
	idata[i] = ipad[i] = 0x36;
	odata[i] = opad[i] = 0x5C;
}
for(i=0;i<hkey.length; i++){
	ipad[i] ^= hkey.charCodeAt(i);
	opad[i] ^= hkey.charCodeAt(i);
	idata[i]= ipad[i] & 0xFF;
	odata[i] = opad[i] & 0xFF;
}
for (i=0;i<text.length;i++) {
	idata[64+i] = text.charCodeAt(i) & 0xFF;
}
var innerhashout = sha1(bytarray2binl(idata));
for (i=0;i<16;i++) {
odata[64+i] = innerhashout[i];
}
return binl2hex(b_calcMD5(odata));
}
/*
 * A JavaScript implementation of the Secure Hash Algorithm, SHA-1, as defined
 * in FIPS PUB 180-1
 * Version 2.1a Copyright Paul Johnston 2000 - 2002.
 * Other contributors: Greg Holt, Andrew Kepert, Ydnar, Lostinet
 * Distributed under the BSD License
 * See http://pajhome.org.uk/crypt/md5 for details.
 */

/*
 * Configurable variables. You may need to tweak these to be compatible with
 * the server-side, but the defaults work in most cases.
 */
var hexcase = 0;  /* hex output format. 0 - lowercase; 1 - uppercase        */
var b64pad  = ""; /* base-64 pad character. "=" for strict RFC compliance   */
var chrsz   = 8;  /* bits per input character. 8 - ASCII; 16 - Unicode      */

/*
 * These are the functions you'll usually want to call
 * They take string arguments and return either hex or base-64 encoded strings
 */
function hex_sha1(s){return binb2hex(core_sha1(str2binb(s),s.length * chrsz));}
function b64_sha1(s){return binb2b64(core_sha1(str2binb(s),s.length * chrsz));}
function str_sha1(s){return binb2str(core_sha1(str2binb(s),s.length * chrsz));}
function hex_hmac_sha1(key, data){ return binb2hex(core_hmac_sha1(key, data));}
function b64_hmac_sha1(key, data){ return binb2b64(core_hmac_sha1(key, data));}
function str_hmac_sha1(key, data){ return binb2str(core_hmac_sha1(key, data));}

/*
 * Perform a simple self-test to see if the VM is working
 */
function sha1_vm_test()
{
  return hex_sha1("abc") == "a9993e364706816aba3e25717850c26c9cd0d89d";
}

/*
 * Calculate the SHA-1 of an array of big-endian words, and a bit length
 */
function core_sha1(x, len)
{
  /* append padding */
  x[len >> 5] |= 0x80 << (24 - len % 32);
  x[((len + 64 >> 9) << 4) + 15] = len;

  var w = Array(80);
  var a =  1732584193;
  var b = -271733879;
  var c = -1732584194;
  var d =  271733878;
  var e = -1009589776;

  for(var i = 0; i < x.length; i += 16)
  {
    var olda = a;
    var oldb = b;
    var oldc = c;
    var oldd = d;
    var olde = e;

    for(var j = 0; j < 80; j++)
    {
      if(j < 16) w[j] = x[i + j];
      else w[j] = rol(w[j-3] ^ w[j-8] ^ w[j-14] ^ w[j-16], 1);
      var t = safe_add(safe_add(rol(a, 5), sha1_ft(j, b, c, d)),
                       safe_add(safe_add(e, w[j]), sha1_kt(j)));
      e = d;
      d = c;
      c = rol(b, 30);
      b = a;
      a = t;
    }

    a = safe_add(a, olda);
    b = safe_add(b, oldb);
    c = safe_add(c, oldc);
    d = safe_add(d, oldd);
    e = safe_add(e, olde);
  }
  return Array(a, b, c, d, e);

}

/*
 * Perform the appropriate triplet combination function for the current
 * iteration
 */
function sha1_ft(t, b, c, d)
{
  if(t < 20) return (b & c) | ((~b) & d);
  if(t < 40) return b ^ c ^ d;
  if(t < 60) return (b & c) | (b & d) | (c & d);
  return b ^ c ^ d;
}

/*
 * Determine the appropriate additive constant for the current iteration
 */
function sha1_kt(t)
{
  return (t < 20) ?  1518500249 : (t < 40) ?  1859775393 :
         (t < 60) ? -1894007588 : -899497514;
}

/*
 * Calculate the HMAC-SHA1 of a key and some data
 */
function core_hmac_sha1(key, data)
{
  var bkey = str2binb(key);
  if(bkey.length > 16) bkey = core_sha1(bkey, key.length * chrsz);

  var ipad = Array(16), opad = Array(16);
  for(var i = 0; i < 16; i++)
  {
    ipad[i] = bkey[i] ^ 0x36363636;
    opad[i] = bkey[i] ^ 0x5C5C5C5C;
  }

  var hash = core_sha1(ipad.concat(str2binb(data)), 512 + data.length * chrsz);
  return core_sha1(opad.concat(hash), 512 + 160);
}

/*
 * Convert an 8-bit or 16-bit string to an array of big-endian words
 * In 8-bit function, characters >255 have their hi-byte silently ignored.
 */
function str2binb(str)
{
  var bin = Array();
  var mask = (1 << chrsz) - 1;
  for(var i = 0; i < str.length * chrsz; i += chrsz)
    bin[i>>5] |= (str.charCodeAt(i / chrsz) & mask) << (32 - chrsz - i%32);
  return bin;
}

/*
 * Convert an array of big-endian words to a string
 */
function binb2str(bin)
{
  var str = "";
  var mask = (1 << chrsz) - 1;
  for(var i = 0; i < bin.length * 32; i += chrsz)
    str += String.fromCharCode((bin[i>>5] >>> (32 - chrsz - i%32)) & mask);
  return str;
}

/*
 * Convert an array of big-endian words to a hex string.
 */
function binb2hex(binarray)
{
  var hex_tab = hexcase ? "0123456789ABCDEF" : "0123456789abcdef";
  var str = "";
  for(var i = 0; i < binarray.length * 4; i++)
  {
    str += hex_tab.charAt((binarray[i>>2] >> ((3 - i%4)*8+4)) & 0xF) +
           hex_tab.charAt((binarray[i>>2] >> ((3 - i%4)*8  )) & 0xF);
  }
  return str;
}

/*
 * Convert an array of big-endian words to a base-64 string
 */
function binb2b64(binarray)
{
  var tab = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  var str = "";
  for(var i = 0; i < binarray.length * 4; i += 3)
  {
    var triplet = (((binarray[i   >> 2] >> 8 * (3 -  i   %4)) & 0xFF) << 16)
                | (((binarray[i+1 >> 2] >> 8 * (3 - (i+1)%4)) & 0xFF) << 8 )
                |  ((binarray[i+2 >> 2] >> 8 * (3 - (i+2)%4)) & 0xFF);
    for(var j = 0; j < 4; j++)
    {
      if(i * 8 + j * 6 > binarray.length * 32) str += b64pad;
      else str += tab.charAt((triplet >> 6*(3-j)) & 0x3F);
    }
  }
  return str;
}
</script>