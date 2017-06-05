<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_ext.aspx.vb" Inherits="Default2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">

<title>SAP - AI Request Extension</title>

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

<!-- code section -->
<script type="text/javascript">
    jQuery(function () { jQuery('#duedate').datetimepicker({ timepicker: false, format: 'd/M/Y', defaultDate: <asp:Literal ID="set_date" runat="server"></asp:Literal>, minDate: <asp:Literal ID="min_date" runat="server"></asp:Literal> }); });
    function formValid() {
        var descr, duedate, emailExp;
        descr = document.getElementById("descr").value;
        duedate = document.getElementById("duedate").value;
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
            alert("Invalid duedate YYYY/MM/DD!");
        }
    }
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
            <h1 class="logo left"><a href="./sap_main.aspx">SAP</a></h1><!-- .logo-->

            <!-- Nav Menu -->
            <ul class="nav-menu right">
                <li class="current"><a href="./sap_main.aspx">home</a></li>
            </ul><!-- .nav-menu-->

        </div><!-- .inner -->
    </div><!-- #header -->
    <!-- End Header -->


    <!-- Title
        ============================= -->
    <div id="title" class="theme-default">
        <div class="inner">
            <h4>AI# <asp:Literal ID="label_ai_id" runat="server"></asp:Literal></h4>
            <h2>Request Extension</h2>
        </div><!-- .inner -->
    </div><!-- #title -->
    <!-- End Title -->

    <!-- Content
        ============================= -->
    <div id="content">
        <div class="inner">
            <ul>
                <!-- Short Description -->
                <li class="description">You are requesting an <b>extension</b> for:</li>
                <!-- <li>
                    <div class="created" title="25/01/2015">
                        <div class="daymonth">
                            <div id="cm">Jan</div>
                            <div id="cy">2015</div>
                        </div>
                        <div id="cd">25</div>
                    </div>
                    <div class="duedate" title="14/02/2015">
                        <div class="daymonth">
                            <div id="dm">Feb</div>
                            <div id="dy">2015</div>
                        </div>
                        <div id="dd">14</div>
                    </div>
                </li> -->
            </ul>

            <!-- <h4 class="subtitle">Action Item</h4> -->

            <table id="ai_list" class="white-table" cellpadding="10" cellspacing="0" style="margin-top: 20px;">
<!--                 <tr class="ai-maintitle">
                    <th class="add-ai"><a href="#" title="Add new AI">+</a></th>
                    <th colspan="5"><h4>Action Items</h4></th>
                </tr> -->
                <tr>
                    <th class="ai-id">AI#</th>
                    <th class="ai-desc">Description</th>
                    <th class="ai-other">Owner</th>
                    <th class="ai-other">Created</th>
                    <th class="ai-other">Due Date</th>
                    <th class="ai-other">Status</th>
                </tr>
                <tr>
                    <td><asp:Literal ID="ai_tbl_id" runat="server"></asp:Literal></td>
                    <td class="ai-desc"><asp:Literal ID="ai_tbl_desc" runat="server"></asp:Literal></td>
                    <td><asp:Literal ID="ai_tbl_owner" runat="server"></asp:Literal></td>
                    <td><asp:Literal ID="ai_tbl_created" runat="server"></asp:Literal></td>
                    <td><asp:Literal ID="ai_tbl_due" runat="server"></asp:Literal></td>
                    <td><asp:Literal ID="ai_tbl_status" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <td colspan="4" style="background: white">
                        Enter reason here:<br>
                        <textarea rows="4" cols="50" name="reason" id="reason" runat="server"></textarea>
                    </td>
                    <td colspan="2" style="background: white">
                        New Due Date:<br>
                        <input type="text" name="duedate" id="duedate"  class="datetimepicker" style="height: 42px; background-color: lightgray" runat="server"/>
                        <input type="hidden" name="form_ai_id" id="form_ai_id" runat="server"/>
                    </td>
                </tr>
                <tr><td colspan="6" class="submitform"><asp:Button ID="btnSave" name="subscribe" type="submit" value="submit" class="submitbtn" Text="SEND TO REQUESTOR FOR APPROVAL" OnClientClick="return formValid();" style="width: 360px;" runat="server" /></td></tr>
            </table>

            <ul style="display: none;">
                <li>
                    <form action="" method="get">

                        <input name="Name" class="" type="text" value="Name (Required)" onfocus="if(this.value == 'Name (Required)') { this.value = ''; }" onblur="if(this.value == '') { this.value = 'Name (Required)'; }" />

                        <input name="Email" class="" type="text" value="Email (Required)" onfocus="if(this.value == 'Email (Required)') { this.value = ''; }" onblur="if(this.value == '') { this.value = 'Email (Required)'; }" />

                        <input name="Subject" class="" type="text" value="Subject" onfocus="if(this.value == 'Subject') { this.value = ''; }" onblur="if(this.value == '') { this.value = 'Subject'; }" />

                        <textarea name="Detail" cols="" rows="" onfocus="if(this.value == 'Describe your project in detail...') { this.value = ''; }" onblur="if(this.value == '') { this.value = 'Describe your project in detail...'; }">Describe your project in detail...</textarea>

                        <input type="submit" value="submit" name="subscribe" class="submitbtn" />

                    </form>
                </li>

                <li>
                    <h4>Contact Information</h4>
                    <p>Quisque hendrerit purus dapibus, ornare nibh vitae, viverra nibh. Fusce vitae aliquam tellus. Proin sit amet volutpat libero. Nulla sed nunc et tortor luctus faucibus. Morbi at aliquet turpis, et consequat felis.</p>

                    <span><i class="li_location"></i>Elm St. 14/05 Lost City </span>
                    <span><i class="icon-phone"></i>lll</span>
                    <span><i class="icon-mail"></i>info@singolo.com</span>
                </li>
            </ul>


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
