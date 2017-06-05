<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_dlvr.aspx.vb" Inherits="Default2" %>

<script runat="server">
</script>

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

<script type = "text/javascript">
    function DisableButton() {
        document.getElementById("<%=Button1.ClientID %>").disabled = true;
    }
    window.onbeforeunload = DisableButton;
</script>

<style>
#Label1 {color:red !important}
</style>
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
                <li class="current"><a href="sap_main.aspx">home</a></li>
            </ul><!-- .nav-menu-->

        </div><!-- .inner -->
	</div>_<!-- #header -->
	<!-- End Header -->


	<!-- Title
		============================= -->
	<div id="title" class="theme-default" style="margin-top: -1em">
        <div class="inner">
            <h4>AI# <asp:Literal ID="label_ai_id" runat="server"></asp:Literal></h4>
            <h2>AI Completion</h2>
        </div><!-- .inner -->
	</div><!-- #title -->
	<!-- End Title -->

	<!-- Content
		============================= -->
	<div id="content">
        <div class="inner">
            <form id="form1" runat="server">
                <table id="ai_list" cellpadding="10" cellspacing="0" style="margin-top: 20px;">
                    <tr>
                        <th class="ai-id">AI#</th>
                        <th class="ai-desc">Description</th>
                        <th class="ai-other"></th>
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
                        <td colspan="6" style="background: white;">
                            Describe completion status here:<br />
                            <textarea rows="4" cols="50" id="reason" name="reason" runat="server"></textarea>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="6">
                            Attach documents:<br />
                            <asp:FileUpload ID="FileUpload1" runat="server" /><br />
                            <asp:Label ID="Label1" runat="server"></asp:Label>
							<div id="file2" style="display:none">
							<asp:FileUpload ID="FileUpload2" runat="server" /><br />
                            <asp:Label ID="Label2" runat="server"></asp:Label>
							</div>
							<div id="file3" style="display:none">
							<asp:FileUpload ID="FileUpload3" runat="server" /><br />
                            <asp:Label ID="Label3" runat="server"></asp:Label>
							</div>
							<div id="file4" style="display:none">
							<asp:FileUpload ID="FileUpload4" runat="server" /><br />
                            <asp:Label ID="Label4" runat="server"></asp:Label>
							</div>
							<div id="file5" style="display:none">
							<asp:FileUpload ID="FileUpload5" runat="server" /><br />
                            <asp:Label ID="Label5" runat="server"></asp:Label>
							</div>
                        </td>
                    </tr>
					
                    <tr><td colspan="6" class="submitform" style="background: white;"><asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="SEND TO REQUESTOR FOR APPROVAL" /><br /></td></tr>
             </table>
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
    <script type="text/javascript">
        jQuery('.datetimepicker').datetimepicker({ timepicker: false, format: 'd-m-Y', minDate: 0 });
		$("#FileUpload1").on("change",function(){
			$("div#file2").toggle($("#FileUpload1").val() != "");
			if(!$("div#file2").is("visible")) {
				$("#FileUpload2").val("")
			}
		});
		$("#FileUpload2").on("change",function(){
			$("div#file3").toggle($("#FileUpload2").val() != "");
			if(!$("div#file3").is("visible")) {
				$("#FileUpload3").val("")
			}
		});
		$("#FileUpload3").on("change",function(){
			$("div#file4").toggle($("#FileUpload3").val() != "");
			if(!$("div#file4").is("visible")) {
				$("#FileUpload4").val("")
			}
		});
		$("#FileUpload4").on("change",function(){
			$("div#file5").toggle($("#FileUpload4").val() != "");
			if(!$("div#file5").is("visible")) {
				$("#FileUpload5").val("")
			}
		});
		$("#FileUpload5").on("change",function(){
			$("div#file5").toggle($("#FileUpload5").val() != "");
		});
    </script>
</body>
</html>
