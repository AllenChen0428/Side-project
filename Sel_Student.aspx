<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Sel_Student.aspx.vb" Inherits="STDWEB_Sel_Student" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register Src="~/Inc/wuMaster.ascx" TagName="MySubmaster" TagPrefix="ucMySubmaster" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>學生個人課表</title>
    <script type="text/javascript">
        function init() {
            if (document.getElementById('<%=ThisYear.ClientID%>') != undefined) {
                document.getElementById('<%=ThisYear.ClientID%>').onchange = QueryForm;
                document.getElementById('<%=ThisTeam.ClientID%>').onchange = QueryForm;
            }

            if (document.getElementById('ThisYearTerm') != undefined) {
                document.getElementById('ThisYearTerm').onchange = QueryForm;
            }
        }

        function ConnectCos_Short(sYears, sTerm, sOpClass, sSerial, sCOSID, sTch_No) {
            //document.getElementById('Years').value = sYears;
            //document.getElementById('Term').value = sTerm;
            //document.getElementById('OpClass').value = sOpClass;
            //document.getElementById('Serial').value = sSerial;
            //document.getElementById('Cos_Short').value = sCos_Short;
            //form1.target = '_blank';
            //form1.method = 'post';
            //form1.action = "../TchWeb/Cur_SelTeaching.aspx";
            //form1.submit();
            //form1.target = '_self';
            //form1.action = 'Sel_Student.aspx';
            let form2 = document.createElement('form');
            form2.id = 'form2';
            form2.target = '_blank';
            form2.method = 'post';
            form2.action = '/TchWeb/Cur_Teaching_Print.aspx';

            let H_Years = document.createElement("input");
            let H_Term = document.createElement("input");
            let H_Op_Class = document.createElement("input");
            let H_Serial_No = document.createElement("input");
            let H_Cos_ID = document.createElement("input");
            let H_Tch_No = document.createElement("input");

            H_Years.setAttribute("type", "hidden");
            H_Years.setAttribute("name", "Years");
            H_Years.setAttribute("value", sYears);

            H_Term.setAttribute("type", "hidden");
            H_Term.setAttribute("name", "Term");
            H_Term.setAttribute("value", sTerm);

            H_Op_Class.setAttribute("type", "hidden");
            H_Op_Class.setAttribute("name", "Op_Class");
            H_Op_Class.setAttribute("value", sOpClass);

            H_Serial_No.setAttribute("type", "hidden");
            H_Serial_No.setAttribute("name", "Serial_No");
            H_Serial_No.setAttribute("value", sSerial);

            H_Cos_ID.setAttribute("type", "hidden");
            H_Cos_ID.setAttribute("name", "Cos_ID");
            H_Cos_ID.setAttribute("value", sCOSID);

            H_Tch_No.setAttribute("type", "hidden");
            H_Tch_No.setAttribute("name", "Tch_No");
            H_Tch_No.setAttribute("value", sTch_No);

            form2.appendChild(H_Years);
            form2.appendChild(H_Term);
            form2.appendChild(H_Op_Class);
            form2.appendChild(H_Serial_No);
            form2.appendChild(H_Cos_ID);
            form2.appendChild(H_Tch_No);

            document.body.appendChild(form2);

            form2.submit();

            document.body.removeChild(form2);
        }

        function SaveExcel() {
            document.getElementById('ToExcel').value = 'Y';
            document.getElementById('doQuery').value = 'Y';
            if (document.getElementById('language').value=="en") {
                let yearterm = document.getElementById('ThisYearTerm').value;
                document.getElementById('SubYear').value = String(yearterm).slice(0, 3);
                document.getElementById('SubTerm').value = String(yearterm).charAt(3);  
            } else {
               document.getElementById('SubYear').value = document.getElementById('ThisYear').value;
               document.getElementById('SubTerm').value = document.getElementById('ThisTeam').value; 
            }
            form1.target = '_blank';
            form1.action = 'Sel_Student.aspx';
            form1.submit();
        }

        function QueryForm() {
            document.getElementById('ToExcel').value = '';
            document.getElementById('doQuery').value = 'Y';
            form1.target = '_self';
            form1.action = 'Sel_Student.aspx';
            form1.submit();
        }

        function doChangeLanguage() {
            QueryForm();
        }
    </script>
    <style type="text/css">
    <!--
    @import url("../Inc/background.css");
    @import url("../Inc/stylec.css");
    @import url("../Inc/buttonStyle.css");  
    -->
	</style>

    <%--有開啟功能表時，按ENTER不反應--%>
    <script type="text/javascript" src="../Inc/PubJScript.js"></script>

</head>
<body onload="init();">
    <form id="form1" runat="server">
    <input type="hidden" id="doQuery" name="doQuery" />
    <input type="hidden" id="ToExcel" name="ToExcel" />
    <input type="hidden" id="SubYear" name="SubYear" />
    <input type="hidden" id="SubTerm" name="SubTerm" />
    <asp:HiddenField ID="Years" runat="server" />
    <asp:HiddenField ID="Term" runat="server" />
    <asp:HiddenField ID="OpClass" runat="server" />
    <asp:HiddenField ID="Serial" runat="server" />
    <asp:HiddenField ID="Cos_Short" runat="server" />
    <table width="98%" border="0" cellpadding="0">
        <tr>
            <td align="center">
                <ucMySubmaster:MySubmaster ID="thisSubmaster" runat="server" />
            </td>
        </tr>
    </table>
    <div class="content" id="TheWebMasterContent" align="center">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpdatePanel3" runat="server">
            <ContentTemplate>
                <table>
                    <tr>
                        <td>
                            <table class="title" align="center">
                                <tr>
                                    <td>
                                        <table id="table_cht" runat="server">
                                            <tr>
                                                <td>
                                                    <asp:DropDownList ID="ThisYear" runat="server" Width="60px"></asp:DropDownList>
                                                </td>
                                                <td>
                                                    <p style="margin: 3px;">學年度第</p>
                                                </td>
                                                <td>
                                                    <asp:DropDownList ID="ThisTeam" runat="server" Width="120px"></asp:DropDownList>
                                                </td>
                                                <td>
                                                    <p style="margin: 3px;">學期</p>
                                                </td>
                                            </tr>
                                        </table>
                                        <table id="table_eng" runat="server" visible="false">
                                            <tr>
                                                <td>
                                                    <asp:DropDownList ID="ThisYearTerm" runat="server" Width="280px"></asp:DropDownList>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="center"><asp:Label ID="lblSTD" runat="server" Text="Label"></asp:Label></td>
                                    <td>
                                        <p style="margin: 3px;">
                                            <input type="button" id="cmdFile" runat="server" value="匯出Excel檔" class="blue_L" onclick="SaveExcel();" />
                                        </p>
                                    </td>
                                    <td>
                                        <p id="lbLanguage" runat="server" style="margin: 3px;">語系</p>
                                    </td>
                                    <td>
                                        <p style="margin: 3px;">
                                            <select id="language" runat="server" style="width: 120px;" onchange="doChangeLanguage();">
                                                <option value="" selected="selected">繁體中文</option>
                                                <option value="en">英文</option>
                                            </select>
                                        </p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Literal ID="LitHtmlShow" runat="server"></asp:Literal>
                        </td>
                    </tr>
                </table>
            </ContentTemplate>
        </asp:UpdatePanel>
    </div>
    </form>
</body>
</html>



