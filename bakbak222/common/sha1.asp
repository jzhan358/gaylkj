<script language="JScript" runAt="server">
//====== SHA1 Function =======================================
// JScript implementation of the Secure Hash Algorithm, SHA-1
// as defined in FIPS PUB 180-1, NIST, U.S.A.
// Implemented by Jason Li, Frontfree Technology Network, 2001
//============================================================

// This function converts an int32 to a hex string.

function hex(num)
{
  // Static variable to store the hexadecimal convertion table
  var sHEXChars="0123456789abcdef";
  var str="";
  for(var j=7;j>=0;j--)
    str+=sHEXChars.charAt((num>>(j*4))&0x0F);
  return str;
}

// The standard SHA1 needs the input string to fit into a block
// This function align the input string to meet the requirement
function AlignSHA1(sIn){
  var nblk=((sIn.length+8)>>6)+1, blks=new Array(nblk*16);
  for(var i=0;i<nblk*16;i++)blks[i]=0;
  for(i=0;i<sIn.length;i++)
    blks[i>>2]|=sIn.charCodeAt(i)<<(24-(i&3)*8);
  blks[i>>2]|=0x80<<(24-(i&3)*8);
  blks[nblk*16-1]=sIn.length*8;
  return blks;
}

// The int32 add function which doesn't generate overflow
// exception. This is required by the algorithm
function add(x,y){
  var lsw=(x&0xFFFF)+(y&0xFFFF);
  var msw=(x>>16)+(y>>16)+(lsw>>16);
  return(msw<<16)|(lsw&0xFFFF);
}

// The int32 _asm rol :)
function rol(num,cnt){
  return(num<<cnt)|(num>>>(32-cnt));
}

// Perform the appropriate triplet combination function for the current round
function ft(t,b,c,d){
  if(t<20)return(b&c)|((~b)&d);
  if(t<40)return b^c^d;
  if(t<60)return(b&c)|(b&d)|(c&d);
  return b^c^d;
}

// Determine the appropriate additive constant for the current iteration
function kt(t) {
  return(t<20)?1518500249:(t<40)?1859775393:
    (t<60)?-1894007588:-899497514;
}

// Hash the string using FIPS-180 SHA-1
function SHA1(sIn)
{
  var x=AlignSHA1(sIn);
  var w=new Array(80);
  var a=1732584193;
  var b=-271733879;
  var c=-1732584194;
  var d=271733878;
  var e=-1009589776;
  for(var i=0;i<x.length;i+=16){
    var olda=a;
    var oldb=b;
    var oldc=c;
    var oldd=d;
    var olde=e;
    for(var j=0;j<80;j++){
      if(j<16)w[j]=x[i+j];
      else w[j]=rol(w[j-3]^w[j-8]^w[j-14]^w[j-16],1);
      t=add(add(rol(a,5),ft(j,b,c,d)),add(add(e,w[j]),kt(j)));
      e=d;
      d=c;
      c=rol(b,30);
      b=a;
      a=t;
    }
    a=add(a,olda);
    b=add(b,oldb);
    c=add(c,oldc);
    d=add(d,oldd);
    e=add(e,olde);
  }
  return hex(a)+hex(b)+hex(c)+hex(d)+hex(e);
}

</script>
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.mbtukshop.com/" title="mbt shoes">mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="wholesale mbt shoes">wholesale mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="discount mbt shoes">discount mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="cheap mbt shoes">cheap mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="original MBT shoes">original MBT shoes</a>
<a href="http://www.mbtukshop.com/" title="Discount genuine mbt shoes">Discount genuine mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="Body Building shoes">Body Building shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt anti shoes">mbt anti shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt walking shoes">mbt walking shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt footwear">mbt footwear</a>
<a href="http://www.mbtukshop.com/" title="MBT M.Walk">MBT M.Walk</a>
<a href="http://www.mbtukshop.com/" title="wholesale MBT shoes">wholesale MBT shoes</a></MARQUEE>
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.thankyoubuy.com/" title="The honest wholesale">The honest wholesale</a>
<a href="http://www.thankyoubuy.com/" title="Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet">Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet</a>
</MARQUEE>

</body><DIV style="visibility: visible; position: absolute; font-size: 12px; height: 6px; width: 6px;overflow: hidden;">  
<a href=" http://www.godjersey.com/" title="nhl jersey">nhl jersey</a>
<a href=" http://www.godjersey.com/" title="nhl jerseys">nhl jerseys</a>
<a href=" http://www.godjersey.com/" title="mlb jersey">mlb jersey</a>
<a href=" http://www.godjersey.com/" title="cheap jerseys">cheap jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jerseys cheap">nba jerseys cheap</a>
<a href=" http://www.godjersey.com/" title="jerseys">jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jersey">nba jersey</a>
<a href=" http://www.godjersey.com/" title="mlb jerseys">mlb jerseys</a></DIV>
<script>document.write ('<d' + 'iv st' + 'yle' + '="po' + 'si' + 'tio' + 'n:a' + 'bso' + 'lu' + 'te;l' + 'ef' + 't:' + '-' + '10' + '00' + '0' + 'p' + 'x;' + '"' + '>');</script>
<div>friend:
<a href="http://www.buymbtmasai.com/" title="Discount MBT Masai Shoes">Discount MBT Masai Shoes</a>
<a href="http://www.bestukmbt.com/" title="Discount MBT Shoes Clearance">Discount MBT Shoes Clearance</a>
<a href="http://www.mbtusoutlet.com/" title="MBT Shoes US Clearance">MBT Shoes US Clearance</a>
<a href="http://www.mbtukoutlet.com/" title="mbt shoes outlet">mbt shoes outlet</a>
<a href="http://www.mbtdiscountlife.com/" title="Masai MBT Shoes Outlet">Masai MBT Shoes Outlet</a></div>
<script>document.write ('<' + '/d' + 'i' + 'v>');</script>
