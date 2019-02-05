

<!-- Begin POP-UP SIZES AND OPTIONS CODE


function popUp(URL) {
var view_width = 550
var view_height = 450
var look='toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=1,width='+view_width+',height='+view_height+','
popwin=window.open("","",look)
popwin.document.open()
popwin.document.write('<title>Close Up</title><head>')
popwin.document.write('<link rel=StyleSheet href="corporatestyle.css" type="text/css" media="screen"></head>')
popwin.document.write('<body bgcolor="#FFFFFF" leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0 marginheight=0 marginwidth=0>')
popwin.document.write('<TABLE cellpadding=0 cellspacing=0 border=0 width="100%" height="100%" ><tr><td align="center">')
popwin.document.write('<center><br>')
popwin.document.write('<TABLE cellpadding=0 cellspacing=0 border=1 bordercolor="000000"><tr><td>')
popwin.document.write('<img src="'+URL+'">')
popwin.document.write('</td></tr></table>')
popwin.document.write('</td></tr><tr><td valign="bottom" align="center">')
popwin.document.write('<br><form><input type=submit value="Close" class="button-popups" onClick=\'self.close()\'></form><br>')
popwin.document.write('</center>')
popwin.document.write('</td></tr></table>')
popwin.document.write('</body>')
popwin.document.close()
}


// End -->