// xPageGrey1
// cross-browser.com

window.onload = function()
{
  //alert('in onload');
};
window.onunload = function() // assigning an unload listener disables bfcache (fastback) in Opera and Firefox
{
  //if (pg) pg.hide();
};
var pg=null;
xAddEventListener(window, 'load',
  function() {
    if (document.getElementById && document.createElement && document.body.appendChild) {
      pg = new xPageGrey1('clsPageGreyDiv', '../../images/circle-thick-orange.gif', 'clsPageGreyImg', 'Please wait... (click anywhere to cancel)', 'clsPageGreyMsg');
      pg.ele.onclick = function(){pg.hide();};/////////////
      xGetElementById('demoBtn').onclick = function() {
        pg.show();
      };
      xGetElementById('demoForm').onsubmit = function() {
        pg.show();
        return true;
      };
    }
  }, false
);
function xPageGrey1(sDivClass, sImgUrl, sImgClass, sMsg, sMsgClass)
{
  /*@cc_on
  @if (@_jscript_version < 5.5) //  not supported in IE until v5.5
  this.ele = null;
  @else @*/
  this.ele = document.createElement('div');
  this.ele.className = sDivClass;
  if (sImgUrl) {
    var img = document.createElement('img');
    img.src = sImgUrl;
    img.className = sImgClass;
    this.msg = document.createElement('p');
    this.msg.className = sMsgClass;
    this.msg.appendChild(img);
    this.msg.appendChild(document.createTextNode(sMsg));
    document.body.appendChild(this.msg);
  }
  document.body.appendChild(this.ele);
  /*@end @*/
  this.show = function()
  {
    if (this.ele) {
      var ds = xDocSize();
      xMoveTo(this.ele, 0, 0);
      xResizeTo(this.ele, ds.w, ds.h);
      if (this.msg) {
        xMoveTo(this.msg, xScrollLeft()+(xClientWidth()-xWidth(this.msg))/2, xScrollTop()+(xClientHeight()-xHeight(this.msg))/2);
      }
    }
  };
  this.hide = function()
  {
    if (this.ele) {
      xResizeTo(this.ele, 10, 10);
      xMoveTo(this.ele, -10, -10);
      if (this.msg) {
        xMoveTo(this.msg, -xWidth(this.msg), 0);
      }
    }
  };
}
