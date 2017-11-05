# ------------------------------------------------------------------------------
# This is data for the breakdownemail.py script
# ------------------------------------------------------------------------------

rowOpenTag = "<tr style='height:15.0pt'>"
rowCloseTag = "</tr>"

agentNameOpenTag = "<td width=167 nowrap valign=bottom style='width:20.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
agentNameCloseTag = "<o:p></o:p></span></p></td>"

acctNumOpenTag = "<td width=75 nowrap valign=bottom style='width:56.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
acctNumCloseTag = "<o:p></o:p></span></p></td>"

orderNumOpenTag = "<td width=75 nowrap valign=bottom style='width:56.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
orderNumCloseTag = "<o:p></o:p></span></p></td>"

orderStatusOpenTag = "<td width=173 nowrap valign=bottom style='width:148.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
orderStatusCloseTag = "<o:p></o:p></span></p></td>"

DEPPNameOpenTag = "<td width=231 nowrap valign=bottom style='width:173.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
DEPPNameCloseTag = "<o:p></o:p></span></p></td>"

fcpAgentNameOpenTag = "<td width=167 nowrap valign=bottom style='width:125.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
fcpAgentNameCloseTag = "<o:p></o:p></span></p></td>"

fcpAcctNumOneOpenTag = "<td width=159 nowrap valign=bottom style='width:119.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
fcpAcctNumOneCloseTag = "<o:p></o:p></span></p></td>"

fcpAcctNumTwoOpenTag = "<td width=128 nowrap valign=bottom style='width:96.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
fcpAcctNumTwoCloseTag = "<o:p></o:p></span></p></td>"

fcpOrderStatusOpenTag = "<td width=151 nowrap valign=bottom style='width:113.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'><p class=MsoNormal><span style='color:black'>"
fcpOrderStatusCloseTag = "<o:p></o:p></span></p></td>"

tableCloseTag = "</table>"

tableGutter = "<td width=19 nowrap valign=bottom style='width:14.0pt;background:#AFB2B5;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>"
fcpTableGutter = "<td width=19 nowrap valign=bottom style='width:14.0pt;background:#AFB2B5;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>"

emailStartHtml = """
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta http-equiv=Content-Type content="text/html; charset=us-ascii">
    <meta name=Generator content="Microsoft Word 15 (filtered medium)">
    <style><!--
      /* Font Definitions */
      @font-face
      	{font-family:"Cambria Math";
      	panose-1:2 4 5 3 5 4 6 3 2 4;}
      @font-face
      	{font-family:Calibri;
      	panose-1:2 15 5 2 2 2 4 3 2 4;}
      /* Style Definitions */
      p.MsoNormal, li.MsoNormal, div.MsoNormal
      	{margin:0in;
      	margin-bottom:.0001pt;
      	font-size:11.0pt;
      	font-family:"Calibri",sans-serif;}
      a:link, span.MsoHyperlink
      	{mso-style-priority:99;
      	color:#0563C1;
      	text-decoration:underline;}
      a:visited, span.MsoHyperlinkFollowed
      	{mso-style-priority:99;
      	color:#954F72;
      	text-decoration:underline;}
      span.EmailStyle17
      	{mso-style-type:personal-compose;
      	font-family:"Calibri",sans-serif;
      	color:windowtext;}
      .MsoChpDefault
      	{mso-style-type:export-only;
      	font-family:"Calibri",sans-serif;}
      @page WordSection1
      	{size:8.5in 11.0in;
      	margin:1.0in 1.0in 1.0in 1.0in;}
      div.WordSection1
      	{page:WordSection1;}
  --></style><!--[if gte mso 9]><xml>
  <o:shapedefaults v:ext="edit" spidmax="1026" />
  </xml><![endif]--><!--[if gte mso 9]><xml>
  <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1" />
  </o:shapelayout></xml><![endif]-->
  </head>
  <body lang=EN-US link="#0563C1" vlink="#954F72">
    <div class=WordSection1>
        <br>
"""

salesDEPPTableOpenTag = """
      <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=721 style='width:600.0pt;border-collapse:collapse'>
        <tr style='height:33.75pt'>   
          <td width=721 colspan=5 valign=bottom style='width:600.0pt;padding:0in 5.4pt 0in  5.4pt;height:33.75pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:26.0pt;font-family:"Arial",sans-serif;color:black'>DEPP Breakdown:<o:p></o:p></span></b></p></td>
        </tr>
        <tr style='height:56.25pt'>
		  <td width=200 valign=bottom style='width:200.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent Name<o:p></o:p></span></b></p></td>
		  <td width=75 valign=bottom style='width:20.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Account Number<o:p></o:p></span></b></p></td>
		  <td width=75 valign=bottom style='width:20.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Order Number<o:p></o:p></span></b></p></td>
		  <td width=231 valign=bottom style='width:50.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>DEPP Name<o:p></o:p></span></b></p></td>
		  <td width=173 valign=bottom style='width:20.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Order Status<o:p></o:p></span></b></p></td>
	   </tr>
"""

FCPTableOpenTag = """
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=851 style='width:638.0pt;border-collapse:collapse'>
  <tr style='height:33.75pt'>
	<td width=325 nowrap colspan=2 valign=bottom style='width:244.0pt;padding:0in 5.4pt 0in 5.4pt;height:33.75pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:26.0pt;font-family:"Arial",sans-serif;color:black'>FCP Sales<o:p></o:p></span></b></p></td>
	<td width=19 nowrap valign=bottom style='width:14.0pt;padding:0in 5.4pt 0in 5.4pt;height:33.75pt'></td>
	<td width=445 nowrap colspan=3 valign=bottom style='width:334.0pt;padding:0in 5.4pt 0in 5.4pt;height:33.75pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:26.0pt;font-family:"Arial",sans-serif;color:black'>FCP Opportunities<o:p></o:p></span></b></p></td>
  </tr>
  <tr style='height:56.25pt'>
	<td width=167 valign=bottom style='width:125.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent Name<o:p></o:p></span></b></p></td>
	<td width=159 valign=bottom style='width:119.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>First Choice Power Account Number<o:p></o:p></span></b></p></td>
	<td width=19 nowrap valign=bottom style='width:14.0pt;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'></td>
	<td width=167 valign=bottom style='width:125.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Agent Name<o:p></o:p></span></b></p></td>
	<td width=128 valign=bottom style='width:96.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>POGO Account Number<o:p></o:p></span></b></p></td>
	<td width=151 valign=bottom style='width:113.0pt;background:#BDD7EE;padding:0in 5.4pt 0in 5.4pt;height:56.25pt'><p class=MsoNormal><b><span style='font-size:14.0pt;color:black'>Pogo Status<o:p></o:p></span></b></p></td>
  </tr>
"""

emailEndHtml = """
    </div>
  </body>
</html>
"""
