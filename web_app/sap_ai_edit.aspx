﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_ai_edit.aspx.vb" Inherits="_Default" %>

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
<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />
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
<script type="text/javascript" src="js/utils.js"></script>
<script type="text/javascript" src="js/jquery.datetimepicker.js"></script>
<script type="text/javascript" src="js/autosize.min.js"></script>

<!--[if lte IE 7]><script src="lte-ie7.js"></script><![endif]-->

<!--[if IE 9]>
    <link rel="stylesheet" type="text/css" href="css/ie9.css">
<![endif]-->

<!-- code section -->
<script type="text/javascript">
    var max_date = <asp:Literal ID="max_date" runat="server"></asp:Literal>;
    jQuery(function () {

        jQuery('#duedate').datetimepicker({ timepicker: false, format: 'd/M/Y', closeOnDateSelect: true, minDate: new Date(), maxDate: max_date });

        jQuery('.add-ai').on('click', function () {
            $("#new_item_dialog").alert();
            $("#new_item_dialog").show();
        });

        jQuery('#view_mail').live('click', function () {
            $("#mail_detail").alert();
            $("#mail_detail").show();
        });
		
		autosize($('#descr'));

    });

    function formValid() {
        var owner, descr, duedate, emailExp;
        descr = document.getElementById("descr").value;
        owner = document.getElementById("owner_list").value;
        duedate = document.getElementById("duedate").value;
        emailExp = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([com\co\.\in])+$/; // to validate owner
        dateExp = /^(19|20)\d\d[\/](0\d|1[012])[\/](0\d|1\d|2\d|3[01])$/; // to validate duedate
        if (descr == null || descr == "") {
            alert("Description is missing!");
            return false;
        }

        var endDate = Date.parse(duedate);
        var currentDate = Date.now();
        if (isNaN(endDate) == false) {
            if (endDate <= currentDate) {
                alert("Invalid past duedate!");
                return false;
            }
        }
        else {
            alert("Invalid duedate DD/MMM/YYYY!");
			return false;
        }

        var maxd = Date.parse(max_date);

        if (endDate > maxd) {
            alert("Invalid duedate! Cannot be greater than Req duedate");
            return false;
        }

    }
</script>
</head>
<body>

    <div>
    <!-- Header
        ============================= -->
    <div id="header">
        <div class="inner">

            <!-- Logo -->
            <h1 class="logo left"><a href="./sap_main.aspx">SAP</a></h1><!-- .logo-->

            <!-- Nav Menu -->
            <ul class="nav-menu right">
                <li class="current"><a href="#">home</a></li>
                <li><a href="./sap_main.aspx">requests</a></li>
                <li><a href="#">archive</a></li>
                <li><a href="#">support</a></li>
                <li><a href="./sap_crud.aspx">users</a></li>
                <li><a href="./sap_owner.aspx"><asp:Literal ID="current_user" runat="server"></asp:Literal></a></li>
                <li style='margin–right:13px'></li>
            </ul><!-- .nav-menu-->

        </div><!-- .inner -->
    </div><!-- #header -->
    <!-- End Header -->_


    <!-- Title
        ============================= -->
    <div id="title" class="theme-default">
        <div class="inner">
            <h4>AI# <asp:Literal ID="ai_id" runat="server"></asp:Literal></h4>
            <h2>Edit Action Item</h2>
        </div><!-- .inner -->
    </div><!-- #title -->
    <!-- End Title -->

    <!-- Content
        ============================= -->
    <div id="content">
        <div class="inner">
            <form id="form1" method="post" action="" runat="server">
            <table cellpadding="10" cellspacing="0" style="width: 100%;" class="newai">
                <tbody>
                    <tr>
                        <td><h4>Owner</h4>
                            <select name="owner" id="owner" runat="server" class="form-control" style=" width: 200px;">
                            </select>
                        </td>
                        <td><div style="float: left;"><h4>Due Date<span style="color: #AAA; font-weight: normal;font-size: 15px;font-family: arial;margin-left: 9px;"><asp:Literal ID="request_due" runat="server"></asp:Literal></span></h4><input type="text" name="duedate" id="duedate" class="datetimepicker form-control" runat="server" /></div></td>
                    </tr>
                    <tr>
                        <td><h4>Status</h4>
                            <select name="status" id="status" runat="server" class="form-control" style="width: 200px;">
                                <option value="PD">Pending</option>
                                <option value="IP">In Progress</option>
                                <option value="DL">Delivered</option>
                            </select>
                        </td>
                        <td><h4>Description</h4><textarea rows="4" cols="30" name="descr" id="descr" class="fw fs form-control" runat="server" ></textarea></td>
                    </tr>
                    <tr>
						<td style="text-align: left;">
							<a runat="server" onclick="return confirm('Clicking ACCEPT you will DELETE this Action Item. Are you sure?');" class="trash" href="#" id="link_del_ai" name="link_del_ai" style="float: left;">
								<img src="/images/trash.png" style="float: left;">
								<div style="float: left;padding: 3px 1px;font-size: 15px;">Delete AI</div>
							</a>                    
						</td>
						<td colspan="2" style="text-align: right;">
                            <a href="#" name="link_req_id" id="link_req_id" runat="server">
                                <input type="button" value="Cancel" class="submitbtn"/>
                            </a>
							<input type="submit" name="btnSave" value="Save changes" onclick="return formValid();" id="Submit1" class="submitbtn" style=""/>
						</td>
                    </tr>
                </tbody>
            </table>
            </form>


        </div><!-- .inner -->
    </div><!-- #content -->
    <!-- End Content -->


    <!-- Footer
        ============================= -->
    <div id="footer">
        <div class="inner">
            <input type="hidden" id="max_due_input" runat="server"/>
            <p class="left"><a href="http://www.sap.com/">SAP</a> © Copyright <script type="text/javascript">document.write(new Date().getFullYear())</script></p>

        </div><!-- .inner -->
    </div><!-- #footer -->
    <!-- End Footer -->

    <a href="" class="go-top">Top</a>
    </div>

</body>
</html>
