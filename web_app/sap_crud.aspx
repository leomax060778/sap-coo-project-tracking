<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_crud.aspx.vb" Inherits="Default2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>SAP - Users</title>

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
    <script type="text/javascript" src="js/utils.js"></script>
    <script type="text/javascript" src="js/jquery.datetimepicker.js"></script>
    <!--[if lte IE 7]><script src="lte-ie7.js"></script><![endif]-->

    <!--[if IE 9]>
        <link rel="stylesheet" type="text/css" href="css/ie9.css">
    <![endif]-->

<!-- code section -->
<script type="text/javascript">
    jQuery(function () {

        jQuery('.add-ai').on('click', function () {
            $("#new_item_dialog").alert("New User");
            $("#new_item_dialog").show();
        });

    });

    function formValid() {
        return true;
    }

</script>

</head>
<body>
    <!-- Header
        ============================= -->
    <div id="header">
        <div class="inner">

            <!-- Logo -->
            <h1 class="logo left"><a ID="current_url" runat="server" href="sap_main.aspx">SAP</a></h1><!-- .logo-->

            <!-- Nav Menu -->

            <ul class="nav-menu right">
                <li class="current"><a href="./sap_main.aspx">home</a></li>
                <li><a href="./sap_main.aspx">requests</a></li>
                <li><a href="./sap_archive.aspx">archive</a></li>
                <li><a href="#">support</a></li>
                <li><a href="./sap_crud.aspx">users</a></li>
        	    <li class="right"><a href="./sap_owner.aspx"><asp:Literal ID="current_user" runat="server"></asp:Literal></a></li>
            </ul><!-- .nav-menu-->

        </div><!-- .inner -->
    </div><!-- #header -->
    <!-- End Header -->


    <!-- Title
        ============================= -->
    <div id="title" class="theme-default">
        <div class="inner">
            <h4>&nbsp;</h4>
            <h2>Users list</h2>
        </div><!-- .inner -->
    </div><!-- #title -->
    <!-- End Title -->

    <!-- Content
        ============================= -->
    <div id="content">
        <div class="inner">
            <p class="subtitle">Welcome to the SAP COO Project Tracking Tool</p>
            <!--<span class="text">Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.</span>-->
        </div>
        
        <div class="inner">
            <span class="projects">Latest Projects</span>
            <hr class="separator" />
        </div>
        
        <div class="inner">

            <table id="ai_list" cellpadding="10" cellspacing="0" runat="server">
                <tr class="ai-maintitle">
                    <th class="add-ai"><a href="#" id="NAI" title="Add New User">+</a></th>
                    <th colspan="6" style="  background-color: #0563C1;"><h4>Add New User</h4></th>
                </tr>
                <tr class="req-titles">
                    <th class="ai-id">ID</th>
                    <th class="ai-other">Exchange Name</th>
                    <th class="ai-other">Email</th>
                    <th class="ai-other">Role</th>
                    <th class="ai-actions">Action</th>
                </tr>
            </table>

            <div align="center"><br/><br/><h2 style="color: #CCCCCC;"><asp:Literal ID="empty_inbox" runat="server"></asp:Literal></h2></div>
             
            <div id="new_item_dialog" style="display:none;padding:10px;">
                <form id="form1" method="post" action="" runat="server">
				    <table cellpadding="10" cellspacing="0" style="width: 100%;" class="newai">
					    <tbody>
						    <tr>
							    <td colspan="2"><h4>ID</h4><input maxlength="10" style="width: 85%;" name="user_id" id="user_id" class="fw fs" runat="server" /></td>
							    <td colspan="2"><h4>Exchange Name</h4><input style="width: 92%;" name="user_name" id="user_name" class="fw fs" runat="server" /></td>
						    </tr>
						    <tr>
							    <td colspan="2"><h4>Role</h4>
                                    <select name="user_role" id="user_role" runat="server" class="fs" style="background: #ddd; width: 200px;  border: 0;  padding: 15px;  margin: 0 0 16px 0;  outline: none;  overflow: auto;  resize: none;  font-size: 12px;  color: #555;  font-family: Arial, Helvetica, sans-serif;  -moz-border-radius: 5px;  -webkit-border-radius: 5px;  -khtml-border-radius: 5px;  border-radius: 5px;">
                                        <option value="OW">Owner</option>
                                        <option value="AD">Admin</option>
                                        <option value="RQ">Requestor</option>
                                    </select>
                                </td>
							    <td colspan="2"><h4>Email</h4><input style="width: 92%;" name="user_email" id="user_email" class="fw fs" runat="server" /></td>
						    </tr>
					    </tbody>
				    </table>
				    <div class="height: 50px;" style="/* float: right; */">
					    <input type="submit" name="btnSave" value="Save" onclick="return formValid();" id="btnSave" class="submitbtn" style="float: right;"/>
					    <a href="#" onclick="$('.modal_dialog').hide();">
                            <input type="button" name="btnCancel" value="Cancel" id="btnCancel" class="submitbtn" style="float: right; margin-right: 10px"/>
					    </a>
				    </div>
                </form>
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
</body>
</html>
