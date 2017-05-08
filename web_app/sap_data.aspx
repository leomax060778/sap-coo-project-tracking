<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_data.aspx.vb" Inherits="Default2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>SAP - Deliver AI</title>

<!-- Meta -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0, initial-scale=1.0" />
<meta name="robots" content="index, follow" />
<meta name="description" content="" />
<meta name="keywords" content="" />
<meta name="author" content="" />
<meta name="copyright" content="" />

<!-- Favicon -->
<link rel="shortcut icon" type="image/x-icon" href="images/favicon.ico" />

<!-- CSS -->
<link href="css/main.css" rel="stylesheet" type="text/css" />
<link href="css/jquery.datetimepicker.css" rel="stylesheet" type="text/css" />
<link href="css/nivo-slider.css" rel="stylesheet" type="text/css" />
<link href="css/nivotheme.css" rel="stylesheet" type="text/css" />
<link href="css/prettyPhoto.css" rel="stylesheet" type="text/css" />

<!-- JS -->
<script type="text/javascript" src="js/jquery-1.7.2.js"></script>
<script type="text/javascript" src="js/jquery.datetimepicker.js"></script>
<script type="text/javascript" src="js/jquery.nivo.slider.pack.js"></script>
<script type="text/javascript" src="js/jquery.quicksand.js"></script>
<script type="text/javascript" src="js/jquery.prettyPhoto.js"></script>
<script type="text/javascript" src="js/jquery.easing.1.3.js"></script>
<script type="text/javascript" src="js/script.js"></script>
<script type="text/javascript" src="js/custom.js"></script>
<!--[if lte IE 7]><script src="lte-ie7.js"></script><![endif]-->

<!--[if IE 9]>
	<link rel="stylesheet" type="text/css" href="css/ie9.css">
<![endif]-->

</head>

<body>

	<form id="form1" runat="server">

	<!-- Header
		============================= -->
	<div id="header">
        <div class="inner">

            <!-- Logo -->
            <h1 class="logo left"><a href="./sap_main.aspx">SAP</a></h1><!-- .logo-->

            <!-- Nav Menu -->
            <ul class="nav-menu right">
                <li class="current"><a href="#">home</a></li>
                <li><a href="#">requests</a></li>
                <li><a href="#">archive</a></li>
                <li><a href="#">support</a></li>
                <li><a href="#">users</a></li>
            </ul><!-- .nav-menu-->

        </div><!-- .inner -->
	</div><!-- #header -->
	<!-- End Header -->


	<!-- Title
		============================= -->
	<div id="title" class="theme-default">
        <div class="inner">
            <h4>Request creation</h4>
            <h2>Information needed</h2>
        </div><!-- .inner -->
	</div><!-- #title -->
	<!-- End Title -->

	<!-- Content
		============================= -->
	<div id="content">
        <div class="inner">
            <!-- <h4 class="subtitle">Action Item</h4> -->
            <table id="ai_list" cellpadding="10" cellspacing="0" style="margin-top: 20px;">
                <tr>
                    <th class="ai-id">Name <input type="checkbox" name="check_name" value="#FFFC6E"/></th>
                    <th class="ai-desc">Description <input type="checkbox" name="check_desc" value="#FFFC6E"/></th>
                    <th class="ai-other">Due Date <input type="checkbox" name="check_due" value="#FFFC6E"/></th>
                </tr>
                <tr>
                    <td id="req_td_name"><b><asp:Literal ID="req_tbl_name" runat="server"></asp:Literal></b></td>
                    <td id="req_td_desc"><b><asp:Literal ID="req_tbl_desc" runat="server"></asp:Literal></b></td>
                    <td id="req_td_due"><b><asp:Literal ID="req_tbl_due" runat="server"></asp:Literal></b></td>
                </tr>
                <tr>
                    <td colspan="3">
                        Detail the information needed to create the request:<br/>
                        <textarea rows="4" cols="50" id="detail" name="detail" runat="server"></textarea>
                        <input type="hidden" name="form_req_id" id="form_req_id" runat="server"/>
                        <input type="hidden" name="form_req_name" id="form_req_name" runat="server"/>
                        <input type="hidden" name="form_req_descr" id="form_req_descr" runat="server"/>
                        <input type="hidden" name="form_req_duedate" id="form_req_duedate" runat="server"/>
                        <input type="hidden" name="requestor_mail" id="requestor_mail" runat="server"/>
                    </td>
                </tr>
                <tr><td colspan="3" class="submitform"><input type="submit" value="SEND TO REQUESTOR" /></td></tr>
            </table>

        </div><!-- .inner -->
	</div><!-- #content -->
	<!-- End Content -->


	<!-- Footer
		============================= -->
	<div id="footer">
        <div class="inner">

        	<p class="left"><a href="http://www.sap.com/">SAP</a> © Copyright 2015</p>

            <ul>
            	<li>
                    <a href="https://www.facebook.com/SAP" target="_blank"><span class="icon-facebook"></span></a>
                    <a href="https://plus.google.com/+SAP/posts" target="_blank"><span class="icon-gplus"></span></a>
                    <a href="https://twitter.com/sap" target="_blank"><span class="icon-twitter"></span></a>
                    <a href="https://www.linkedin.com/company/sap" target="_blank"><span class="icon-linkedin"></span></a>
                </li>
            </ul>

        </div><!-- .inner -->
	</div><!-- #footer -->
	<!-- End Footer -->

    <a href="" class="go-top">Top</a>
    <script type="text/javascript">
        jQuery('.datetimepicker').datetimepicker({ timepicker: false, format: 'd-m-Y', minDate: 0 });
    </script>
    </form>
</body>
</html>
