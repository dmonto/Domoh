/////////////////////////////////
// File Name: mBanner.js       //
// By: Manish Kumar Namdeo     //
/////////////////////////////////
// BANNER OBJECT
function Banner(objName){this.obj = objName; this.aNodes = []; this.currentBanner = 0;};

// ADD NEW BANNER
Banner.prototype.add = function(bannerType, bannerPath, bannerDuration, height, width, hyperlink) {
	this.aNodes[this.aNodes.length] = new Node(this.obj +"_"+ this.aNodes.length, bannerType, bannerPath, bannerDuration, height, width, hyperlink);};

// Node object
function Node(name, bannerType, bannerPath, bannerDuration, height, width, hyperlink) {
    this.name = name; this.bannerType = bannerType; this.bannerPath = bannerPath; this.bannerDuration = bannerDuration; this.height = height; this.width = width; this.hyperlink= hyperlink;
    //	alert (name +"|" + bannerType +"|" + bannerPath +"|" + bannerDuration +"|" + height +"|" + width + "|" + hyperlink);
};

// Outputs the banner to the page
Banner.prototype.toString = function() {
    var str = "";
	for (var iCtr=0; iCtr < this.aNodes.length; iCtr++) {
		str = str + '<span name="'+this.aNodes[iCtr].name+'" ' + 'id="'+this.aNodes[iCtr].name+'" ' + 'class=m_banner_hide >\n';
		if (this.aNodes[iCtr].hyperlink != ""){str = str + '<a target=_new href="'+this.aNodes[iCtr].hyperlink+'">';}
			
		if ( this.aNodes[iCtr].bannerType == "FLASH" ) {
			str = str + '<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" '
			str = str + 'width="'+this.aNodes[iCtr].width+'" height="'+this.aNodes[iCtr].height+'" id="bnr_'+this.aNodes[iCtr].name+'" align="" viewastext>'
			str = str + '<param name=movie value="'+ this.aNodes[iCtr].bannerPath + '"><param name=quality value=high><param name=bgcolor value=#FFFCDA>'
			str = str + '<embed src="'+this.aNodes[iCtr].bannerPath+'" quality=high '
//			str = str + 'bgcolor=#FFFCDA '
			str = str + 'width="'+this.aNodes[iCtr].width+'" height="'+this.aNodes[iCtr].height+'" name="bnr_'+this.aNodes[iCtr].name+'" ' + 'align=center type="application/x-shockwave-flash" '
			str = str + 'pluginspage="http://www.macromedia.com/go/getflashplayer"></embed></object>'
		} else if ( this.aNodes[iCtr].bannerType == "IMAGE" ){
			str = str + '<img src="'+this.aNodes[iCtr].bannerPath+'" style = "width:100%">';	}

		if (this.aNodes[iCtr].hyperlink != "") {str = str + '</a>';}
		str += '</span>';
	}
	return str;
};

// START THE BANNER ROTATION
Banner.prototype.start = function(){
	this.changeBanner(); var thisBannerObj = this.obj;
	// CURRENT BANNER IS ALREADY INCREMENTED IN cahngeBanner() FUNCTION
	setTimeout(thisBannerObj+".start()", this.aNodes[this.currentBanner].bannerDuration * 1000);
}

// CHANGE BANNER
Banner.prototype.changeBanner = function() {
	var thisBanner;
	var prevBanner = -1;

	if (this.currentBanner < this.aNodes.length ) {
		thisBanner = this.currentBanner;
		if (this.aNodes.length > 1) {if (thisBanner > 0 ) {prevBanner = thisBanner - 1;} else {prevBanner = this.aNodes.length-1;}}
		if (this.currentBanner < this.aNodes.length - 1) {this.currentBanner = this.currentBanner + 1;} else {this.currentBanner = 0;}
	}
	
	if (prevBanner >= 0){document.getElementById(this.aNodes[prevBanner].name).className = "m_banner_hide";	}
	document.getElementById(this.aNodes[thisBanner].name).className = "m_banner_show";
}