<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_main.aspx.vb" Inherits="Default2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>SAP - Requests</title>

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

        $("tr.req-edit").on("click", function(){
            window.location = $("h1.logo a").attr("href") + "sap_req.aspx?id=" + $(this).attr("id");
        });

        jQuery('#new_due_date').datetimepicker({ timepicker: false, format: 'Y-m-d' });

        jQuery('.new_due').live('click', function () {
            $("#new_due_date")[0].value = "";
            $(this).attr("id");
            $("#set_due_dialog").alert('Set new date');
            $("#set_due_dialog").show();
        });

        jQuery('.html_mail').live('click', function () {
            $(this).attr("id");
            $("#set_due_dialog").alert();
            $("#set_due_dialog").show();
        });

    });

    function formValid() {
            alert("Invalid duedate YYYY/MM/DD!");
    }

</script>
<style>
.afilter {
    font-size: 13px;
    margin-left: 10px;
    font-weight: normal;
    color: silver;
}
.sel {
background-color: white;
  padding: 3px 7px;
  color: rgb(86, 86, 86);
  border-radius: 5px;
  font-weight: normal;
}
</style>
</head>
<body>
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
                <li><a href="./sap_owner.aspx"><asp:Literal ID="current_user" runat="server"></asp:Literal></a></li>
                <li style='margin–right:13px'></li>
            </ul><!-- .nav-menu-->

        </div><!-- .inner -->
    </div><!-- #header -->
    <!-- End Header -->


    <!-- Title
        ============================= -->
    <div id="title" class="theme-default">
        <div class="inner">
            <h4>&nbsp;</h4>
            <h2><a style="color:white" href="sap_main.aspx">Requests list</a>
            <br />
            <% Dim f = Request.Querystring("f")%>
            <a class="afilter <%=iif(f="","sel","") %>" href="sap_main.aspx">All</a>
            <a class="afilter <%=iif(f="nd","sel","") %>" href="sap_main.aspx?f=nd">Need data</a>
            <a class="afilter <%=IIf(f = "ur", "sel", "") %>" href="sap_main.aspx?f=ur">Unsassigned</a>
            <a class="afilter <%=IIf(f = "du", "sel", "") %>" href="sap_main.aspx?f=du">Mult Owner</a>
            <a class="afilter <%=IIf(f = "ov", "sel", "") %>" href="sap_main.aspx?f=ov">OverDue</a>
            <a class="afilter <%=IIf(f = "dw", "sel", "") %>" href="sap_main.aspx?f=dw">This Week DD</a>
            <a class="afilter <%=iif(f="ex","sel","") %>" href="sap_main.aspx?f=ex">Extension Pend</a>
            <a class="afilter <%=IIf(f = "ap", "sel", "") %>" href="sap_main.aspx?f=ap">Accept Pend (OW)</a>
            <a class="afilter <%=IIf(f = "rq", "sel", "") %>" href="sap_main.aspx?f=rq">Accept Pend(RQ)</a>           

            </h2>
        </div><!-- .inner -->
    </div><!-- #title -->
    <!-- End Title -->

    <!-- Content
        ============================= -->
    <div id="content">
        <div class="inner">
            <form id="form1" method="post" action="" runat="server">
                <table id="ai_list" cellpadding="10" cellspacing="0" runat="server">
                    <tr class="req-titles">
                        <th class="ai-id">REQ#</th>
                        <th class="ai-desc">Description</th>
                        <th class="ai-other">Created</th>
                        <th class="ai-other">Next Due Date</th>
                        <th class="ai-other">Status</th>
                        <th class="ai-actions">Action</th>
                    </tr>
                </table>

            <div align="center"><br/><br/><h2 style="color: #CCCCCC;"><asp:Literal ID="empty_inbox" runat="server"></asp:Literal></h2></div>

            <div id="set_due_dialog" style="display:none;padding:10px;">

            <h4 class="subtitle">Set Due Date</h4>
            <table cellpadding="10" cellspacing="0" style="width: 100%;" class="newai">
                <tbody>
                    <tr>
                        <td><div style="float: left;"><h4>Due Date</h4><input type="text" name="new_due_date" id="new_due_date" class="datetimepicker" runat="server" /></div></td>
                    </tr>
                </tbody>
            </table>
            <div class="height: 50px;" style="/* float: right; */"><input type="submit" name="btnSave" value="Submit" id="btnSave" class="submitbtn" style="float: right;" /></div>

            </div>
          </form>
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
