Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Text.RegularExpressions
Partial Class STDWEB_Sel_Student
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Cnn As System.Data.SqlClient.SqlConnection = Nothing '定義資料庫連接
        Dim objListItem As System.Web.UI.WebControls.ListItem = Nothing
        Dim strSQLCmd As String = "" '課表查詢指令
        Dim strTimeCmd As String = "" '現在學年期指令
        Dim strSessionCmd As String = "" '節次查詢指令
        Dim strNameCmd As String = ""
        Dim strLanguage As String = UTrim(Request.Params("language"))
        Dim strThisYearTerm As String = UTrim(Request.Form("ThisYearTerm"))

        Dim objRD As System.Data.SqlClient.SqlDataReader = Nothing '執行課表查詢指令
        Dim strError As String = ""

        Dim strYears As String = "" '紀錄當學年
        Dim strTerm As String = "" '紀錄當學期
        Dim strCurYears As String = "" '紀錄選課學年
        Dim strCurTerm As String = "" '紀錄選課學期
        Dim strStartYears As String = ""
        Dim strLastYears As String = ""
        Dim strThisYears As String = UTrim(Request.Form("ThisYear")) '紀錄當學年
        Dim strThisTerm As String = UTrim(Request.Form("ThisTeam")) '紀錄當學期
        Dim strSTD As String = "" '紀錄該名學生
        Dim strOther As String = ""
        Dim lngCredits As Single = 0.0  '學分數

        Dim http As String = "" '存放產生HTML語法變數

        '現在位置
        Me.thisSubmaster.setDbLocation("STDW301")

        Dim tmpCurTime As String = ""   '比對用節次
        Dim tmpClass As String = ""     '比對用開課班級
        Dim tmpCosID As String = ""     '比對用科目
        Dim strTchNo As String = ""     '教師代碼
        Dim strCurList As String = ""   '存放課表HTML語法變數
        Dim strToExcel As String = Request.Form("ToExcel")
        Dim strSubYear As String = Request.Form("SubYear")
        Dim strSubTerm As String = Request.Form("SubTerm")
        Dim strNewLine As String = IIf(strToExcel = "Y", "<br>", "<br>")
        Dim strURL As String = ""
        Dim CurArr(7, 15) As String     '課表使用陣列
        Dim I As Integer = 0
        Dim J As Integer = 0

        '陣列清空
        For I = 0 To 7
            For J = 0 To 15
                CurArr(I, J) = ""
            Next J
        Next I

        strTimeCmd = "select Now_Years, Now_Term,Cur_Years,Cur_Term from CUP_Time_Cntr"
        objRD = GetOfflineRS(Cnn, strTimeCmd, "現在學年期查詢", strError)
        If Not objRD Is Nothing Then
            While objRD.Read
                strYears = objRD("Now_Years")
                strTerm = objRD("Now_Term")
                strCurYears = objRD("Cur_Years")
                strCurTerm = objRD("Cur_Term")
            End While
            objRD.Close()
            objRD = Nothing
        Else
            Response.Write(strError)
            Response.End()
            Cnn.Close()
            Cnn.Dispose()
            Cnn = Nothing
            Exit Sub
        End If

        If strThisYears = "" Then
            strThisYears = strYears
            strThisTerm = strTerm
        End If

        If strThisYearTerm <> "" Then
            strThisYears = Mid(strThisYearTerm, 1, 3)
            strThisTerm = Mid(strThisYearTerm, 4, 1)
        End If

        '找學生名子
        strNameCmd = "select st.Std_Name," &
                     "  case isnull(st.Std_EngName,'') when '' then st.Std_Name else st.Std_EngName end as Std_Eng_Name," &
                     "  isnull(s5.Years,year(getdate())-1911) as Years " &
                     "from STD_Student01 st " &
                     "  left join (select min(Regi_Year) as Years " &
                     "             from STD_Student05 " &
                     "             where Std_No = '" & HttpContext.Current.Session("StdNo") & "' " &
                     "  ) as s5 on (1=1) " &
                     "where st.Std_No = '" & HttpContext.Current.Session("StdNo") & "' "
        objRD = GetOfflineRS(Cnn, strNameCmd, "學生名子查詢", strError)
        If Not objRD Is Nothing Then
            While objRD.Read
                strSTD = IIf(strLanguage = "", objRD("Std_Name"), objRD("Std_Eng_Name"))
                strStartYears = objRD("Years")
            End While
            objRD.Close()
            objRD = Nothing
        Else
            Response.Write(strError)
            Response.End()
            Cnn.Close()
            Cnn.Dispose()
            Cnn = Nothing
            Exit Sub
        End If

        If strLanguage = "" Then
            Me.table_cht.Visible = True
            Me.table_eng.Visible = False
            Me.cmdFile.Value = "匯出Excel"
            Me.lbLanguage.InnerText = "語系"

            Try
                If strCurYears > strYears Then
                    strLastYears = strCurYears
                Else
                    strLastYears = strYears
                End If

                ThisYear.Items.Clear()

                For I = CType(strStartYears, Integer) To CType(strLastYears, Integer)
                    objListItem = New System.Web.UI.WebControls.ListItem
                    With objListItem
                        .Value = I
                        .Text = I
                        .Selected = Microsoft.VisualBasic.IIf(I = CType(strThisYears, Integer), True, False)
                    End With
                    ThisYear.Items.Add(objListItem)
                    objListItem = Nothing
                Next
            Catch ex As Exception
                'Response.Write("strStartYears=" & strStartYears & ",strYears=" & strYears & "<br>")
            End Try

            Try


                Dim TermArray = New String() {"", "第一學期", "第二學期", "寒假", "暑期"}

                ThisTeam.Items.Clear()
                For I = 1 To 4
                    objListItem = New System.Web.UI.WebControls.ListItem
                    With objListItem
                        .Value = I
                        .Text = TermArray(I)
                        .Selected = Microsoft.VisualBasic.IIf(I = CType(strThisTerm, Integer), True, False)
                    End With
                    ThisTeam.Items.Add(objListItem)
                    objListItem = Nothing
                Next
            Catch ex As Exception
                'Response.Write("strThisTerm=" & strThisTerm & "<br>")
            End Try
        Else
            Me.table_cht.Visible = False
            Me.table_eng.Visible = True
            Me.cmdFile.Value = "Excel File"
            Me.lbLanguage.InnerText = "Language"

            strSQLCmd = "select Regi_Year,Regi_Term " &
                        "from STD_Student05 " &
                        "where Std_No='" & HttpContext.Current.Session("StdNo") & "' " &
                        "order by Regi_Year desc,Regi_Term desc"
            objRD = GetOfflineRS(Cnn, strSQLCmd, "註冊學年期", strError)

            If Not objRD Is Nothing Then
                Me.ThisYearTerm.Items.Clear()

                While objRD.Read
                    objListItem = New ListItem

                    With objListItem
                        .Text = "The " & objRD.Item(1) & " semester of the " & objRD.Item(0) & " academic year"
                        .Value = objRD.Item(0) & objRD.Item(1)
                        .Selected = IIf(objRD.Item(0) & objRD.Item(1) = strThisYears & strThisTerm, True, False)
                    End With

                    Me.ThisYearTerm.Items.Add(objListItem)
                End While
                objRD.Close()
                objRD = Nothing
            End If
        End If

        Call InitLanguage(Me.language, strLanguage)

        '節次判斷 依學生所屬日或夜 顯示對應的時間
        strSessionCmd = "select distinct cs.T_Section,sc.Years,sc.Term,ss1.Std_No,ss1.Std_Name,cs.DN_Mark," &
                        "   cs.Section_Name,cs.Start_Time,cs.End_Time," &
                        "   'Session '+convert(varchar,convert(int,cs.T_Section)) as Section_Name_Eng " &
                        "from CUP_Section as cs " &
                        "inner join RGP_Edu04 as e4 on cs.DN_Mark = e4.DN_Mark " &
                        "inner join STD_Student01 as ss1 on e4.Dept_No = ss1.Dept_No " &
                        "inner join SEL_Choose as sc on sc.Std_No = ss1.Std_No " &
                        "where ss1.Std_No = '" & HttpContext.Current.Session("StdNo") & "' " &
                        "and sc.Years = '" & strThisYears & "' " &
                        "and sc.Term = '" & strThisTerm & "' "

        objRD = GetOfflineRS(Cnn, strSessionCmd, "查節次時間", strError)

        If Not objRD Is Nothing Then
            While objRD.Read
                CurArr(0, objRD.Item("T_Section")) = IIf(strLanguage = "", objRD.Item("Section_Name"), objRD.Item("Section_Name_Eng")) & "<br>" & objRD.Item("Start_Time") & "<br>" & objRD.Item("End_Time")
            End While
            objRD.Close()
            objRD = Nothing
        Else
            Response.Write(strError)
            Response.End()
            Cnn.Close()
            Cnn.Dispose()
            Cnn = Nothing
            Exit Sub
        End If

        '修課學分數
        strSQLCmd = "select dbo.GetStdSelCredits('" & HttpContext.Current.Session("StdNo") & "','" & Right("0" & strThisYears, 3) & "','" & strThisTerm & "','" & strThisTerm & "','') as Cos_Credit"
        objRD = GetOfflineRS(Cnn, strSQLCmd, "查學生個人課表", strError)
        If Not objRD Is Nothing Then
            While objRD.Read
                lngCredits = objRD.Item("Cos_Credit")
            End While
            objRD.Close()
            objRD = Nothing
        End If

        '學生個人課表
        strSQLCmd = "execute SelCur_StdCosForm '" & Right("0" & strThisYears, 3) & "','" & strThisTerm & "','" & HttpContext.Current.Session("StdNo") & "'"

        'strSQLCmd = "Select distinct WeekTime = Substring(b.Cur_Time,1,1), a.Years, a.Term, a.OP_Class, a.Serial,Cur_Time = b.Cur_Time, " &
        '            "SectionTime = Convert(int,Substring(b.Cur_Time,2,2)),a.Cos_ID,g.Cos_Name, " &
        '            "TchName=dbo.GetCurTeacher(a.Years ,a.Term ,a.OP_Class ,a.Serial), " &
        '            "CrName=dbo.GetCurCR(a.Years ,a.Term ,a.OP_Class ,a.Serial), " &
        '            "Week_Mark = Case b.Week_Mark When '1' then '(單)' When '2' then '(雙)' else '' End " &
        '            "From SEL_Choose a Join CUR_ClassCur b On a.Years = b.Years and a.Term = b.Term " &
        '            "and a.OP_Class = b.OP_Class and a.Serial = b.Serial " &
        '            "Left Join CUR_TeacherCur c On b.Years = c.Years and b.Term = c.Term " &
        '            "and b.OP_Class = c.OP_Class and b.Serial = c.Serial and b.Cur_Time = c.Cur_Time and b.Week_Mark = c.Week_Mark " &
        '            "Left Join CUR_ClassroomCur d On b.Years = d.Years and b.Term = d.Term " &
        '            "and b.OP_Class = d.OP_Class and b.Serial = d.Serial and b.Cur_Time = d.Cur_Time and b.Week_Mark = d.Week_Mark " &
        '            "Left Join EMP_Main e On c.Tch_No = e.Code " &
        '            "Left Join CUP_Classroom f On d.Classroom = f.Classroom " &
        '            "Join CUP_Permanent g On a.Cos_ID = g.Cos_ID " &
        '            "Left Join CUR_Open1 o on a.Years = o.Years and a.Term = o.Term " &
        '            "and a.OP_Class = o.OP_Class and a.Serial = o.Serial " &
        '            "left join STD_Student01 as ss1 on ss1.Std_No = a.Std_No " &
        '            "Where a.Std_No = '" & HttpContext.Current.Session("StdNo") & "' and isnull(o.Stop_Mark,'')='' " &
        '            "and a.Years='" & Right("0" & strThisYears, 3) & "' " &
        '            "and a.Term='" & strThisTerm & "' " &
        '            "Order by b.Cur_Time"

        objRD = GetOfflineRS(Cnn, strSQLCmd, "查學生個人課表", strError)

        If Not objRD Is Nothing Then
            While objRD.Read
                If objRD.Item("OP_Class").ToString.Substring(0, 3) = "AOS" Then
                    If strOther <> "" Then
                        strOther = strOther & "、"
                    End If

                    strOther = strOther & objRD.Item("Cos_Name") & " (" & objRD.Item("Cos_Credit") & ")"
                Else
                    tmpCurTime = Trim(objRD.Item("Cur_Time"))

                    If Trim(objRD.Item("Cur_Time")) = tmpCurTime Then
                        strURL = IIf(strToExcel = "Y", objRD.Item("Cos_Name"), "<a href=" & Chr(34) & "javascript:ConnectCos_Short('" & objRD.Item("Years") & "'" & ",'" & objRD.Item("Term") & "'," & "'" & objRD.Item("Op_Class") & "'," & "'" & objRD.Item("Serial") & "'," & "'" & objRD.Item("Cos_ID") & "'," & "'" & objRD.Item("Tch_No") & "');" & Chr(34) & ">" & objRD.Item("Cos_Name") & "</a>")
                        CurArr(CInt(objRD.Item("WeekTime")), CInt(objRD.Item("SectionTime"))) &= strURL & "<br>" &
                                objRD.Item("TchName") + "<br>" &
                                objRD.Item("CrName")
                    End If
                End If

            End While
            objRD.Close()
            objRD = Nothing
        Else
            Response.Write(strError)
            Response.End()
            Cnn.Close()
            Cnn.Dispose()
            Cnn = Nothing
            Exit Sub
        End If

        strCurList = "<table id='bgBase' width=96% cellspacing=1 cellpadding=1 align='center' style='border:1 solid black;font-size:12px'>" & Chr(10)
        strCurList &= "<tr bgcolor='#dcdcdc'>" & Chr(10)
        strCurList &= "<td colspan='2' style='height:30px;text-align:center;font-size:13px;'>" & "本學期修課學分數：" & CType(lngCredits, String) & "</td>" & Chr(10)
        If strOther <> "" Then
            strCurList &= "<td colspan='6' style='height:30px;font-size:13px;'>" & "校際選課課程：" & strOther & "</td>" & Chr(10)
        Else
            strCurList &= "<td colspan='6' style='height:30px;'></td>" & Chr(10)
        End If
        strCurList &= "</tr>" & Chr(10)
        strCurList &= "<tr bgcolor='#dcdcdc'>" & Chr(10)
        strCurList &= "<td width='9%' align='center' height='36'>" & IIf(strLanguage = "", "節次", "Session") & "</td>" & Chr(10)
        strCurList &= "<td width='13%' align='center'>" & IIf(strLanguage = "", "一", "Monday") & "</td>" & Chr(10)
        strCurList &= "<td width='13%' align='center'>" & IIf(strLanguage = "", "二", "Tuesday") & "</td>" & Chr(10)
        strCurList &= "<td width='13%' align='center'>" & IIf(strLanguage = "", "三", "Wednesday") & "</td>" & Chr(10)
        strCurList &= "<td width='13%' align='center'>" & IIf(strLanguage = "", "四", "Thursday") & "</td>" & Chr(10)
        strCurList &= "<td width='13%' align='center'>" & IIf(strLanguage = "", "五", "Friday") & "</td>" & Chr(10)
        strCurList &= "<td width='13%' align='center'>" & IIf(strLanguage = "", "六", "Saturday") & "</td>" & Chr(10)
        strCurList &= "<td width='13%' align='center'>" & IIf(strLanguage = "", "日", "Sunday") & "</td>" & Chr(10)
        strCurList &= "</tr>" & Chr(10)

        For J = 1 To 15
            strCurList &= "<tr>" & Chr(10)
            For I = 0 To 7
                If I = 0 Then
                    strCurList &= "<td width='9%' align='center' bgcolor='#dcdcdc'>" & CurArr(I, J) & "</td>" & Chr(10)
                Else
                    strCurList &= "<td id='bgContent1' width='13%' align='center' bgcolor = '#FAF0E6'>" & CurArr(I, J) & "</td>" & Chr(10)
                End If
            Next I
            strCurList &= "</tr>" & Chr(10)
        Next J

        If strToExcel = "Y" Then
            strCurList = "<p align='center' style='margin:6px;'>" & IIf(strLanguage = "", strSubYear & "學年度 第" & strSubTerm & "學期 " & "學生" & strSTD & "&nbsp;個人課表", "The " & strSubTerm & " semester of the " & strSubYear & " academic year " & strSTD & "′Schedule") & "</p>" & strCurList
            Response.Clear()
            Response.AddHeader("content-disposition", "attachment;filename=PoolExport.xls")
            Response.ContentType = "application/vnd.xls"
            Response.Write(strCurList)
            Response.End()
        Else
            strCurList &= "</table>"

            lblSTD.Text = IIf(strLanguage = "", "學生" & strSTD & "&nbsp;個人課表", strSTD & "′Schedule")
            LitHtmlShow.Text = strCurList       '將課表HTML語法 以 LITERAL 顯示
        End If

        Cnn.Close()
        Cnn.Dispose()
        Cnn = Nothing
    End Sub
End Class
