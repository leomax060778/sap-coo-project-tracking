<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_user_edit.aspx.vb" Inherits="_Default" %>

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
jQuery(function () {


});

    function formValid() {
        var owner, descr, duedate, emailExp;
        descr = document.getElementById("descr").value;
        owner = document.getElementById("owner").value;
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
        	    <li><a href="./sap_main.aspx"><asp:Literal ID="current_user" runat="server"></asp:Literal></a></li>
                <li style='margin–right:13px'></li>            
            </ul><!-- .nav-menu-->

        </div><!-- .inner -->
    </div><!-- #header -->
    <!-- End Header -->_


    <!-- Title
        ============================= -->
    <div id="title" class="theme-default">
        <div class="inner">
            <h4>Edit User <asp:Literal ID="req_id" runat="server"></asp:Literal></h4>
            <h2><asp:Literal ID="req_detail" runat="server"></asp:Literal></h2>
        </div><!-- .inner -->
    </div><!-- #title -->
    <!-- End Title -->

    <!-- Content
        ============================= -->
    <div id="content">
        <div class="inner">
            <div id="new_item_dialog" style="display:block;padding:10px;">
            <form id="form1" method="post" action="" runat="server">
            <table cellpadding="10" cellspacing="0" style="width: 100%;" class="newai">
                <tbody>
                    <tr>
                        <td><h4>Exchange Account User Name</h4><textarea rows="1" cols="30" name="fullname" id="fullname" class="form-control" runat="server" style="width:80%" ></textarea></td>
                        <td><div style="float: left;"><h4>ID</h4><input disabled="disabled" type="text" name="id" id="id" class="form-control" runat="server" /></div></td>
                    </tr>
					<tr>
                        <td><h4>Role</h4>
                            <select name="role" id="role" runat="server" class="form-control" style="width: 200px;">
                                <option value="OW">Owner</option>
                                <option value="AD">Admin</option>
                                <option value="AO">Admin & Owner</option>
								<option value="RQ">Requestor</option>
                            </select>
							</td>
					
							<td><h4>Email</h4><textarea rows="1" cols="30" name="email" id="email" class="form-control" runat="server" style="width:80%" ></textarea></td>
					
                    </tr>
					<tr>
					<td style="text-align: left;">
						<a runat="server" onclick="return confirm('Clicking ACCEPT you will DELETE this user. Are you sure?');" class="trash" href="#" id="link_del_user" name="link_del_user" style="float: left;">
							<img src="/images/trash.png" style="float: left;">
							<div style="float: left;padding: 3px 1px;font-size: 15px;">Delete User</div>
						</a>                    
						
                    </td>
                    <td style="text-align: right;">
                    <a href="./sap_crud.aspx" name="link_req_id" id="link_req_id" runat="server">
                        <input type="button" value="Cancel" class="submitbtn"/>
                    </a>
                    <input type="submit" name="btnSave" value="Save changes" onclick="return formValid();" id="Submit1" class="submitbtn" style=""/>
                    </td>
                    </tr>
                </tbody>
            </table>
            </form>
            </div>

            <div id="mail_detail" style="display:none;padding:10px;" runat="server">
            </div>

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

</body>
</html>
