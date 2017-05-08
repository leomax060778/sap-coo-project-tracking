<%@ Page Language="VB" AutoEventWireup="false" CodeFile="sap_csv_ai_report.aspx.vb" Inherits="sap_csv_ai_report" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href='css/style.css' rel='stylesheet' type='text/css' />
    <script type='text/javascript' src='js/jquery-1.7.2.js'></script>
    <script type='text/javascript' src='js/jquery-latest.js'></script>
    <script type='text/javascript' src='js/jquery.tablesorter.js'></script>
    <script type='text/javascript'>jQuery(function () { $(document).ready(function () { $('#report').tablesorter(); }); });</script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Literal ID="report_table" runat="server"></asp:Literal>
        </div>
    </form>
</body>
</html>
