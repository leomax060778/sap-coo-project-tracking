﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_owner.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>SAP - Action Items Panel</title>

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
<link href="css/nivo-slider.css" rel="stylesheet" type="text/css" />
<link href="css/nivotheme.css" rel="stylesheet" type="text/css" />
<link href="css/prettyPhoto.css" rel="stylesheet" type="text/css" />
<link href="css/jquery.datetimepicker.css" rel="stylesheet" type="text/css" />

<!-- JS -->
<script type="text/javascript" src="js/jquery-1.7.2.js"></script>
<script type="text/javascript" src="js/jquery.nivo.slider.pack.js"></script>
<script type="text/javascript" src="js/jquery.quicksand.js"></script>
<script type="text/javascript" src="js/jquery.prettyPhoto.js"></script>
<script type="text/javascript" src="js/jquery.easing.1.3.js"></script>
<script type="text/javascript" src="js/script.js"></script>
<script type="text/javascript" src="js/custom.js"></script>
<script type="text/javascript" src="js/jquery.datetimepicker.js"></script>

<!--[if lte IE 7]><script src="lte-ie7.js"></script><![endif]-->

<!--[if IE 9]>
    <link rel="stylesheet" type="text/css" href="css/ie9.css">
<![endif]-->

<!-- code section -->
<script type="text/javascript">
    jQuery(function () {

        $("#ai_list tr").on("click", function(){
            window.location =  "sap_ai_view.aspx?id=" + $(this).attr("id");
        });



    });

</script>
</head>
<body>
    <form id="form1" method="post" action="" runat="server">
    <div>

    <!-- Header
        ============================= -->
    <div id="header">
        <div class="inner">

            <!-- Logo -->
            <h1 class="logo left"><a ID="current_url" runat="server" href="./sap_main.aspx">SAP</a></h1><!-- .logo-->

            <!-- Nav Menu -->

            <ul class="nav-menu right">
                <li class="current"><a href="./sap_main.aspx">home</a></li>
                <li><a href="./sap_main.aspx">requests</a></li>
                <li><a href="./sap_archive.aspx">archive</a></li>
                <li><a href="#">support</a></li>
                <li><a href="./sap_crud.aspx">users</a></li>
                <li class="right non-clickeable"><asp:Literal ID="current_user" runat="server"></asp:Literal></li>
            </ul><!-- .nav-menu-->

        </div><!-- .inner -->
    </div><!-- #header -->
    <!-- End Header -->


    <!-- Title
        ============================= -->
    <div id="title" class="theme-default">
        <div class="inner">
            <h2 style="display: none"><asp:Literal ID="req_id" runat="server"></asp:Literal></h2>
            <h4>&nbsp;</h4>
            <h2><a style="color:white" href="sap_owner.aspx">SAP COO Project Tracking Tool</a>
            <br />
        </div><!-- .inner -->
    </div><!-- #title -->
    <!-- End Title -->

    <!-- Content
        ============================= -->
    <div id="content">
        <div class="inner">
            <p class="subtitle">Welcome to the SAP COO Project Tracking Tool</p>
                <span class="text">A solution developed for the benefit of improving our collaboration as a team. This tool provides us with the opportunity to simplify the way we process our many action items with increased accuracy, predictability, full transparency and visibility.  All the elements in one single place [project, owner, objective, and due date] and the workflow streamlined to make our interaction significantly more efficient.   
                </br>
                </br>
                Please, see below what it is in your portfolio of actions today:
            </span>
        </div>
        
        <div class="inner">
            <span class="projects">Latest Projects</span>
            <hr class="separator" />
        </div>

        <div class="inner">
            <table id="ai_list" cellpadding="10" cellspacing="0" runat="server">
                <tr class="req-titles">
                    <th class="ai-id">Request Number</th>
                    <th class="ai-other">Description</th>
                    <th class="ai-other">Created</th>
                    <th class="ai-other">Due Date</th>
                    <th class="ai-other">Status</th>
                    <th class="ai-actions" style="width:220px;">Action</th>
                </tr>
            </table>
            <div align="center"><br><br><h2 style="color: #CCCCCC;"><asp:Literal ID="empty_inbox" runat="server"></asp:Literal></h2></div>
        </div><!-- .inner -->
    </div><!-- #content -->
    <!-- End Content -->


    <!-- Footer
        ============================= -->
    <div id="footer">
        <div class="inner">

            <p class="left"><a href="http://www.sap.com/">SAP</a> © Copyright <script type="text/javascript">document.write(new Date().getFullYear())</script></p>


        </div><!-- .inner -->
    </div><!-- #footer -->
    <!-- End Footer -->

    <a href="" class="go-top">Top</a>
    </div>
    </form>
</body>
</html>
