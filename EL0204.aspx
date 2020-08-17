<%@ Page Language="VB" AutoEventWireup="false" CodeFile="EL0204.aspx.vb" Inherits="EL_EL0204" %>
<%@ Register Src="~/Inc/wuMaster.ascx" TagName="MySubmaster" TagPrefix="ucMySubmaster" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>課程教材維護</title>
    <style type="text/css">
        @import url("../Inc/stylec.css");
        @import url("../Inc/buttonStyle.css");
        @import url("../Inc/background.css");

        .Condition td {
            padding: 5px;
        }
    </style>
    <link type="text/css" rel="Stylesheet" href="../Inc/grid.css" />
    <link type="text/css" rel="stylesheet" href="/Inc/GridView_Style.css?Version=<%=Now().ToString("yyyyMMddHHmmss") %>" />
    <script>
        const Show_Count = 50;      //每頁顯示幾筆
        let Total_Count = 0;        //總比數
        let Now_Page = 1;           //現在頁數
        let Total_Page = 0;         //總頁數

        let Select_Progress_No = '';

        let Con_Years = '';
        let Con_Term = '';
        let Con_Class = '';
        let Con_Serial = '';

        let Tmp_Years = '';
        let Tmp_Term = '';
        let Tmp_Class = '';
        let Tmp_Serial = '';

        let Grd = {
            "Now_Row": 0,
            "New_Style": 'background-color:#FFFFBE;',
            "Old_Style": '',
            "New_Tr": document.createElement('tr'),
            "Old_Tr": document.createElement('tr')
        }
    </script>
    <script>
        window.onload = function () {
            ShowWaitDIv(true);

            Init_Control(function () {
                Query_List(function () {
                    setTimeout(function () {
                        ShowWaitDIv(false);
                    }, 500);
                });
                
            });
        };
    </script>
    <script>
        function Init_Control(callback) {

            document.getElementById('Client_Query').onclick = function () {
                Con_Years = document.getElementById('Con_Years').value;
                Con_Term = document.getElementById('Con_Term').value;
                Con_Class = document.getElementById('Con_Class').value;
                Con_Serial = document.getElementById('Con_Serial').value;

                ShowWaitDIv(true);
                Query_List(function () {
                    setTimeout(function () {
                        ShowWaitDIv(false);
                    }, 500);
                });
            };

            document.getElementById('Con_Years').onchange = function () {
                Re_Init_Control(1);
            };

            document.getElementById('Con_Term').onchange = function () {
                Re_Init_Control(1);
            };

            document.getElementById('Con_Class').onchange = function () {
                Re_Init_Control(2);
            };

            Init_Term_List({
                Control_ID: 'Con_Term',
                Show_All: false,
                All_Type: 1,
                Show_Code: false
            }, function () {
                Get_Now_YT(function (StrYear, StrTerm) {
                    Con_Years = StrYear;
                    Con_Term = StrTerm;
                    document.getElementById('Con_Years').value = Con_Years;
                    document.getElementById('Con_Term').value = Con_Term;

                    Init_Open_Class_List({
                        Control_ID: 'Con_Class',
                        Year: Con_Years,
                        Term: Con_Term,
                        Show_All: true,
                        All_Type: 1,
                        Show_Code: true
                    }, function () {
                        Con_Class = document.getElementById('Con_Class').value;

                        Init_Open_Cos_List({
                            Control_ID: 'Con_Serial',
                            Year: Con_Years,
                            Term: Con_Term,
                            Class: Con_Class, 
                            Show_All: true,
                            All_Type: 1,
                            Show_Code: true
                        }, function () {
                            Con_Serial = document.getElementById('Con_Serial').value;
                            callback();
                        });

                        
                    });

                    
                });
                
            });
            
        }

        function Re_Init_Control(Level) {
            Tmp_Years = document.getElementById('Con_Years').value;
            Tmp_Term = document.getElementById('Con_Term').value;
            Tmp_Class = document.getElementById('Con_Class').value;


            if (Level == 1) {
                Tmp_Class = '';
                Tmp_Serial = '';
            }

            if (Level == 2) {
                Tmp_Serial = '';
            }

            if (Level == 1) {

                Init_Open_Class_List({
                    Control_ID: 'Con_Class',
                    Year: Tmp_Years,
                    Term: Tmp_Term,
                    Show_All: true,
                    All_Type: 1,
                    Show_Code: true
                }, function () {
                    Init_Open_Cos_List({
                        Control_ID: 'Con_Serial',
                        Year: Tmp_Years,
                        Term: Tmp_Term,
                        Class: Tmp_Class,
                        Show_All: true,
                        All_Type: 1,
                        Show_Code: true
                    }, function () {

                    });

                });
            }


            if (Level == 2) {
                Init_Open_Cos_List({
                    Control_ID: 'Con_Serial',
                    Year: Tmp_Years,
                    Term: Tmp_Term,
                    Class: Tmp_Class,
                    Show_All: true,
                    All_Type: 1,
                    Show_Code: true
                }, function () {

                });
            }


        }
    </script>
    <script>
        function Query_List(callback) {
            let xhttp = new XMLHttpRequest();

            xhttp.onreadystatechange = function () {
                if (this.readyState == 4) {
                    if (this.status == 200) {
                        Show_List(this.responseText, callback);
                    } else {
                        callback();
                        return alert('伺服器錯誤');
                    }
                }
            };

            xhttp.open('Post', '/EL/HttpRequest/EL0204.ashx', true);

            let formdata = new FormData();
            formdata.append('Model', 'Query_List');

            formdata.append('Page', Now_Page);
            formdata.append('Show_Count', Show_Count);
            formdata.append('Years', Con_Years);
            formdata.append('Term', Con_Term);
            formdata.append('OP_Class', Con_Class);
            formdata.append('Serial', Con_Serial);

            xhttp.send(formdata);
        }

        function Show_List(JData, callback) {
            Check_Json_Data(JData, function () {
                let obj = JSON.parse(JData)[0];
                let Data = JSON.parse(JData)[0].Data;

                let GridView1 = document.getElementById('GridView1');
                let tbody = GridView1.getElementsByTagName('tbody')[0];

                while (tbody.hasChildNodes()) {
                    tbody.removeChild(tbody.firstChild);
                }

                let columnName = new Array('Years', 'Term', 'Class_Name', 'Serial', 'Cos_Name', 'Setting');

                if (Data == null) {
                    let tr = document.createElement('tr');
                    let td = document.createElement('td');

                    td.setAttribute('colspan', columnName.length);
                    td.setAttribute('style', 'height:35p; color:#FF0000; background-color:#FFFFFF; cursor:auto;');
                    td.innerText = '查無資料';
                    tr.appendChild(td);
                    tbody.appendChild(tr);
                } else {

                    for (i = 0; i < Data.length; i++) {
                        let tr = document.createElement('tr');

                        //tr.setAttribute('onclick', 'Select_Row(Grd, this, ' + parseInt((Now_Page - 1) * Show_Count + i + 1) + ', \'' + Data[i].Progress_No + '\');');

                        for (let x = 0; x < columnName.length; x++) {
                            let td = document.createElement('td');
                            let value = unescape(eval('Data[i].' + columnName[x] + ';'));

                            switch (columnName[x]) {
                                case 'Setting':
                                    td.innerHTML = '<input id="Material' + i + '" type="button" value="設定" class="blue_M" onclick="Setting_Material(\'' + Data[i].Serial + '\');" />';
                                    break;
                                default:
                                    td.innerText = value;
                                    break;
                            }


                            

                            tr.appendChild(td);
                        }

                        tbody.appendChild(tr);
                    }
                }

                Init_PageButton(obj.Total_Count, obj.Now_Page, obj.Total_Page);

                callback();
            });
        }

        //Init_PageButton('總筆數', '現在頁數', '總頁數');
        function Init_PageButton(Total_Count, Now_Page, Total_Page) {
            document.getElementById('Show_Count').innerText = Show_Count;
            document.getElementById('Total_Count').innerText = Total_Count;
            document.getElementById('Now_Page').innerText = Now_Page;
            document.getElementById('Total_Page').innerText = Total_Page;

            document.getElementById('Page_Down').removeAttribute('disabled');
            document.getElementById('Page_Down').onclick = Next_Query;
            document.getElementById('Page_Up').removeAttribute('disabled');
            document.getElementById('Page_Up').onclick = Up_Query;

            if (Now_Page == Total_Page) {
                document.getElementById('Page_Down').disabled = 'disabled'
                document.getElementById('Page_Down').onclick = '';
                document.getElementById('Page_Up').disabled = ''
                document.getElementById('Page_Up').onclick = Up_Query;
            }

            if (Now_Page == 1) {
                document.getElementById('Page_Down').disabled = '';
                document.getElementById('Page_Down').onclick = Next_Query;
                document.getElementById('Page_Up').disabled = 'disabled'
                document.getElementById('Page_Up').onclick = '';
            }

            if (Total_Page <= 1) {
                document.getElementById('Page_Down').disabled = 'disabled';
                document.getElementById('Page_Down').onclick = '';
                document.getElementById('Page_Up').disabled = 'disabled'
                document.getElementById('Page_Up').onclick = '';
            }
        }

        function Next_Query() {
            Now_Page = Now_Page + 1;
            ShowWaitDIv(true);
            Query_List(function () {
                setTimeout(function () {
                    ShowWaitDIv(false);
                }, 500);
            });
        }

        function Up_Query() {
            Now_Page = Now_Page - 1;
            ShowWaitDIv(true);
            Query_List(function () {
                setTimeout(function () {
                    ShowWaitDIv(false);
                }, 500);
            });
        }
    </script>
    <script>
        function Select_Row(GrdObject, New_Tr, Row, Progress_No) {
            Select_Row_Model(GrdObject, New_Tr, Row, function () {
                
                Select_Progress_No = Progress_No;

                //開放按鈕
                
                document.getElementById('Client_Update').removeAttribute('disabled');
                document.getElementById('Client_Update').className = 'blue_L';
                document.getElementById('Client_Update').onclick = Client_Update_Click;
                document.getElementById('Client_Delete').removeAttribute('disabled');
                document.getElementById('Client_Delete').className = 'blue_L';
                document.getElementById('Client_Delete').onclick = Client_Delete_Click;
                
                //查詢資料
                ShowWaitDIv(true);
                Query_Data(function () {
                    setTimeout(function () {
                        ShowWaitDIv(false);
                    }, 500);

                });
                

            }, function () {
                Clear_Data();
            });
        }

        function Clear_Data() {
            Select_Progress_No = '';

            document.getElementById('Progress_No').value = '';
            document.getElementById('Progress_Name').value = '';
            document.getElementById('Update_User').value = '';
            document.getElementById('Update_Time').value = '';

            
            document.getElementById('Client_Update').disabled = 'disabled';
            document.getElementById('Client_Update').className = 'black_L';
            document.getElementById('Client_Update').removeAttribute('onclick');
            document.getElementById('Client_Delete').disabled = 'disabled';
            document.getElementById('Client_Delete').className = 'black_L';
            document.getElementById('Client_Delete').removeAttribute('onclick');
        }
    </script>
    <script>
        function Query_Data(callback) {
            let xhttp = new XMLHttpRequest();

            xhttp.onreadystatechange = function () {
                if (this.readyState == 4) {
                    if (this.status == 200) {
                        Show_Data(this.responseText, callback);
                    } else {
                        callback();
                        return alert('伺服器錯誤');
                    }
                }
            };

            xhttp.open('Post', '/EL/HttpRequest/EL0202.ashx', true);

            let formdata = new FormData();
            formdata.append('Model', 'Query_Data');
            formdata.append('Years', Con_Years);
            formdata.append('Term', Con_Term);
            formdata.append('OP_Class', Con_Class);
            formdata.append('Serial', Con_Serial);
            formdata.append('Progress_No', Select_Progress_No);

            xhttp.send(formdata);
        }


        function Show_Data(JData, callback) {
            Check_Json_Data(JData, function () {
                let Data = JSON.parse(JData)[0].Data;

                if (Data == null) {
                    alert('查無資料');
                } else {
                    document.getElementById('Progress_No').value = unescape(Data[0].Progress_No);
                    document.getElementById('Progress_Name').value = unescape(Data[0].Progress_Name);
                    document.getElementById('Update_User').value = unescape(Data[0].Update_User);
                    document.getElementById('Update_Time').value = unescape(Data[0].Update_Time);
                }

                callback();
            });
        }
    </script>
    <script>
        function Insert_Data(success, failed) {
            let xhttp = new XMLHttpRequest();

            xhttp.onreadystatechange = function () {
                if (this.readyState == 4) {
                    if (this.status == 200) {
                        Show_Insert_Result(this.responseText, success, failed);
                    } else {
                        failed();
                        return alert('伺服器錯誤');
                    }
                }
            };

            xhttp.open('Post', '/EL/HttpRequest/EL0202.ashx', true);

            let formdata = new FormData();
            formdata.append('Model', 'Insert_Data');
            
            formdata.append('Years', Con_Years);
            formdata.append('Term', Con_Term);
            formdata.append('OP_Class', Con_Class);
            formdata.append('Serial', Con_Serial);
            formdata.append('Progress_No', document.getElementById('Progress_No').value);
            formdata.append('Progress_Name', document.getElementById('Progress_Name').value);

            xhttp.send(formdata);
        }


        function Show_Insert_Result(JData, success, failed) {
            Check_Json_Data(JData, function () {
                let Data = JSON.parse(JData)[0].Data;

                switch (Data) {
                    case true:
                        alert('新增成功');
                        success();
                        break;
                    case false:
                        alert('新增失敗');
                        failed();
                        break;
                    default:
                        alert(unescape(Data));
                        failed();
                        break;
                }
            });
        }
    </script>
    <script>
        function Client_Update_Click() {
            if (document.getElementById('Progress_Name').value == '') {
                return alert('請輸入進度名稱');
            }

            ShowWaitDIv(true);
            Update_Data(function () {
                Select_Row(Grd, null, 0, '');
                Clear_Data();
                Query_List(function () {
                    setTimeout(function () {
                        ShowWaitDIv(false);
                    }, 500);
                });
            }, function () {
                ShowWaitDIv(false);
            });
        }

        function Update_Data(success, failed) {
            let xhttp = new XMLHttpRequest();

            xhttp.onreadystatechange = function () {
                if (this.readyState == 4) {
                    if (this.status == 200) {
                        Show_Update_Result(this.responseText, success, failed);
                    } else {
                        failed();
                        return alert('伺服器錯誤');
                    }
                }
            };

            xhttp.open('Post', '/EL/HttpRequest/EL0202.ashx', true);
            
            let formdata = new FormData();
            formdata.append('Model', 'Update_Data');
            formdata.append('Years', Con_Years);
            formdata.append('Term', Con_Term);
            formdata.append('OP_Class', Con_Class);
            formdata.append('Serial', Con_Serial);
            formdata.append('Progress_No', Select_Progress_No);
            formdata.append('Progress_Name', document.getElementById('Progress_Name').value);

            xhttp.send(formdata);
        }

        function Show_Update_Result(JData, success, failed) {
            Check_Json_Data(JData, function () {
                let Data = JSON.parse(JData)[0].Data;

                switch (Data) {
                    case true:
                        alert('修改成功');
                        success();
                        break;
                    case false:
                        alert('修改失敗');
                        failed();
                        break;
                    default:
                        alert(unescape(Data));
                        failed();
                        break;
                }

            });
        }
    </script>
    <script>
        function Client_Delete_Click() {
            let r = confirm("確定是否要刪除?");
            if (r != true) {
                return alert('已取消刪除');
            }

            ShowWaitDIv(true);
            Delete_Data(function () {
                Now_Page = 1;
                Select_Row(Grd, null, 0, '');
                Clear_Data();
                Query_List(function () {
                    setTimeout(function () {
                        ShowWaitDIv(false);
                    }, 500);
                });
            }, function () {
                ShowWaitDIv(false);
            });
        }

        function Delete_Data(success, failed) {
            let xhttp = new XMLHttpRequest();

            xhttp.onreadystatechange = function () {
                if (this.readyState == 4) {
                    if (this.status == 200) {
                        Show_Delete_Result(this.responseText, success, failed);
                    } else {
                        failed();
                        return alert('伺服器錯誤');
                    }
                }
            };

            xhttp.open('Post', '/EL/HttpRequest/EL0202.ashx', true);
            
            let formdata = new FormData();
            formdata.append('Model', 'Delete_Data');
            formdata.append('Years', Con_Years);
            formdata.append('Term', Con_Term);
            formdata.append('OP_Class', Con_Class);
            formdata.append('Serial', Con_Serial);
            formdata.append('Progress_No', Select_Progress_No);

            xhttp.send(formdata);
        }

        function Show_Delete_Result(JData, success, failed) {
            Check_Json_Data(JData, function () {
                let Data = JSON.parse(JData)[0].Data;

                switch (Data) {
                    case true:
                        alert('刪除成功');
                        success();
                        break;
                    case false:
                        alert('刪除失敗');
                        failed();
                        break;
                    default:
                        alert(unescape(Data));
                        failed();
                        break;
                }

            });
        }
    </script>
    <script>
        function Setting_Material(Serial) {
            Show_SubPage('1000px', '800px', '10', '/EL/EL0204_Material.aspx?Years=' + Con_Years + '&Term=' + Con_Term + '&OP_Class=' + Con_Class + '&Serial=' + Serial, '設定教材');
        }
    </script>
    <script src="/Inc/PubJScript.js?Version=<%=Now().ToString("yyyyMMddHHmmss") %>"></script>
    <script src="/EL/Javascript/EL.js?Version=<%=Now().ToString("yyyyMMddHHmmss") %>"></script>
</head>
<body style="font-family:'Microsoft JhengHei'; font-size:13px;">
    <form id="form1" runat="server">
        <div id="Header" style="width:100%; min-width:1000px; height:90px; top:0px; left:0px; margin:0 auto; position:absolute;">
            <ucMySubmaster:MySubmaster ID="thisSubmaster" runat="server" />
        </div>
        <div id="TheWebMasterContent" style="width:100%; min-width:1024px; height:calc(100% - 100px); margin:0 auto; top:96px; border-style:none; background-color:#FFFFFF; position:absolute; overflow-y:scroll;">
            <div style="width:1000px; height:auto; min-height:5px; border-style:solid; border-color:#CCCCCC; margin:0 auto; background-image:url(/Images/bg_right.jpg); ">
                <br />
                <table border="0" style="width:100%; text-align:center; ">
                    <tr>
                        <td>
                            學年
                            <input id="Con_Years" type="text" maxlength="3" style="width:60px; text-align:center;" />
                            學期
                            <select id="Con_Term" style="width:80px;">
                            </select>
                        </td>
                        <td>
                            <input id="Client_Query" type="button" value="查詢"  class="orange_L" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            開課班級
                            <select id="Con_Class" style="width:150px;" >
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            科目
                            <select id="Con_Serial" style="width:150px;" >
                            </select>
                        </td>
                    </tr>
                </table>
                <br />
            </div>
            <br />
            <div style="width:1000px; max-height:400px; min-height:5px; border-style: solid; border-color: #CCCCCC; margin:0 auto; overflow-y:scroll; ">
                <table id="GridView1" border="0" class="GridView03">
                    <thead>
                        <tr>
                            <td style="width:80px;">
                                學年
                            </td>
                            <td style="width:60px;">
                                學期
                            </td>
                            <td style="width:250px;">
                                開課班級
                            </td>
                            <td style="width:60px;">
                                開課序號
                            </td>
                            <td style="">
                                科目
                            </td>
                            <td>
                                操作
                            </td>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td colspan="10" style="text-align:center; background-color:#FFFFFF; color:#FF0000; height:35px;">
                                <label style="">處理中...</label>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
            <div style="width:1000px; height:auto; min-height:5px; border-style:none; border-color: #CCCCCC; margin:0 auto; text-align:right;">
                <table style="width:100%;">
                    <tr>
                        <td style="text-align:left;">
                            
                        </td>
                        <td style="text-align:right;">
                            每頁
                            <span id="Show_Count">0</span> 
                            筆 共 
                            <span id="Total_Count">0</span> 
                            筆資料 
                            &nbsp;&nbsp;
                            目前在 
                            <span id="Now_Page">1</span> 
                            / 
                            <span id="Total_Page">0</span> 
                            頁
                            <input id="Page_Up" type="button" value="上一頁" onclick="Up_Query();" />
                            <input id="Page_Down" type="button" value="下一頁" onclick="Next_Query();" />
                        </td>
                    </tr>
                </table>
            </div>
        </div>

    </form>
</body>
</html>
