<%@ Page Language="VB" AutoEventWireup="false" %>
<%
    Dim su As SapUser = New SapUser()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>SAP - New User</title>

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
            <h4>Please contact the support center</h4>
            <h2>Are you a new user?</h2>
        </div><!-- .inner -->
	</div><!-- #title -->
	<!-- End Title -->

	<!-- Content
		============================= -->
	<div id="content">
        <div class="inner">
            <ul>
                <!-- Short Description -->
				<div style="  text-align: center;  padding: 100px;">
            
            <div class="description" style="
    /* text-align: center; */
    /* display: inline; */
    margin: 0 auto;
"><img src="/images/bloqueado.png"><div>Your user id '<b><%= su.getId() %></b>' is not associated with this service.</div></div>
        </div>
            </ul>

        </div><!-- .inner -->
	</div><!-- #content -->
	<!-- End Content -->


	<!-- Footer
		============================= -->
	<div id="footer">
        <div class="inner">

        	<p class="left"><a href="http://www.sap.com/">SAP</a> © Copyright 2015</p>

            

        </div><!-- .inner -->
	</div><!-- #footer -->
	<!-- End Footer -->

    <a href="" class="go-top">Top</a>
    <script type="text/javascript">
        jQuery('.datetimepicker').datetimepicker({ timepicker: false, format: 'd-m-Y', minDate: 0 });
    </script>
</body>
</html>
