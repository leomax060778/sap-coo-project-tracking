<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_req.aspx.vb" Inherits="_Default" %>

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
<link href="css/timesheet.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" type="text/css" href="//fonts.googleapis.com/css?family=Signika+Negative" />

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
<script type="text/javascript" src="js/timesheet.js"></script>

<!--[if lte IE 7]><script src="lte-ie7.js"></script><![endif]-->

<!--[if IE 9]>
    <link rel="stylesheet" type="text/css" href="css/ie9.css">
<![endif]-->

<!-- code section -->
<script type="text/javascript">
jQuery(function () {

    jQuery('.add-ai').on('click', function(){
        $("#new_item_dialog").alert("New Action Item");
        jQuery('#duedate').datetimepicker({ timepicker: false, format: 'd/M/Y', closeOnDateSelect: true, minDate: <asp:Literal ID="min_date" runat="server"></asp:Literal>, maxDate: <asp:Literal ID="max_date" runat="server"></asp:Literal> });
        $("#new_item_dialog").show();
    });

    jQuery('#view_mail').live('click', function () {
        $("#mail_detail").alert();
        $("#mail_detail").show();
    });


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
        /*
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
            */
        }
       
</script>

<asp:Literal ID="section_width" runat="server"></asp:Literal>

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
                <li class="current"><a href="./sap_main.aspx">home</a></li>
                <li><a href="./sap_main.aspx">requests</a></li>
                <li><a href="#">archive</a></li>
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
            <h4>Request Number <asp:Literal ID="req_id" runat="server"></asp:Literal></h4>
            <h2><asp:Literal ID="req_detail" runat="server"></asp:Literal></h2>
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
            <ul>
                <li class="description" style="width: 70%;"><h4><asp:Literal ID="req_description" runat="server"></asp:Literal> <a href="#" id="download_link" title="View mail..." runat="server">Show E-mail Documentation</a></h4></li>
                <li>
                    <div class="created" title="25/01/2015" id="req_created">
                        <div class="daymonth">
                            <div id="cm"><asp:Literal ID="req_created_month" runat="server"></asp:Literal></div>
                            <div id="cy"><asp:Literal ID="req_created_year" runat="server"></asp:Literal></div>
                        </div>
                        <div id="cd"><asp:Literal ID="req_created_day" runat="server"></asp:Literal></div>
                    </div>
                    <div class="duedate" title="14/Feb/2015" id="req_duedate">
                        <div class="daymonth">
                            <div id="dm"><asp:Literal ID="req_duedate_month" runat="server"></asp:Literal></div>
                            <div id="dy"><asp:Literal ID="req_duedate_year" runat="server"></asp:Literal></div>
                        </div>
                        <div id="dd"><asp:Literal ID="req_duedate_day" runat="server"></asp:Literal></div>
                    </div>
                </li>
            </ul>
            <p></p>
            <table id="ai_list" cellpadding="10" cellspacing="0" runat="server">
                <tr class="ai-maintitle">
                    <th class="add-ai"><a href="#" id="NAI" title="Add new AI" style="text-decoration: none;color:white;padding-top:5px">+</a></th>
                    <th colspan="6" style="  background-color: #0563C1;"><h4>Action Items</h4></th>
                </tr>
                <tr>
                    <th class="ai-id">AI#</th>
                    <th class="ai-other">Description</th>
                    <th class="ai-other">Owner</th>
                    <th class="ai-other">Created</th>
                    <th class="ai-other">Due Date</th>
                    <th class="ai-other">Status</th>
                    <th class="ai-actions" style="width: 25%">Action</th>
                </tr>
            </table>

            <!--div id="timesheet"></div-->

            <div id="req_timesheet" class="timesheet color-scheme-default">
                <asp:Literal ID="req_timesheet_data" runat="server"></asp:Literal>
            </div>

            <div align="center"><br/><br/><h2 style="color: #CCCCCC;"><asp:Literal ID="empty_inbox" runat="server"></asp:Literal></h2></div>

            <div id="new_item_dialog" style="display:none;padding:10px;">
                <form id="form1" method="post" action="" runat="server">
                <table cellpadding="10" cellspacing="0" style="width: 100%;" class="newai">
                    <tbody>
                        <tr>
                            <td colspan="2"><h4>Description</h4><textarea style="width: 92%;" rows="4" cols="50" name="descr" id="descr" class="fw fs" runat="server" ></textarea></td>
                        </tr>
                        <tr>
                            <td><h4>Owner</h4>
                                <select name="owner" id="owner" runat="server" class="fs" style="background: #ddd; width: 200px;  border: 0;  padding: 15px;  margin: 0 0 16px 0;  outline: none;  overflow: auto;  resize: none;  font-size: 12px;  color: #555;  font-family: Arial, Helvetica, sans-serif;  -moz-border-radius: 5px;  -webkit-border-radius: 5px;  -khtml-border-radius: 5px;  border-radius: 5px;">
                                </select>
                            </td>
                            <td><div style="float: left;"><h4>Due Date</h4><input style="  height: 46px;" type="text" name="duedate" id="duedate" class="datetimepicker" runat="server" /></div></td>
                        </tr>
                    </tbody>
                </table>
                <div class="height: 50px;" style="/* float: right; */">
                    <input type="button" name="btnCancel" value="Cancel" onclick="$('.modal_dialog').hide();" id="btnCancel" class="submitbtn" style="float: right;margin-left: 10px;"/>
                    <input type="submit" name="btnSave" value="Save" onclick="return formValid();" id="btnSave" class="submitbtn" style="float: right;"/>
                </div>
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
