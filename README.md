# excel-vba-note
時間對減，只顯示分或小時的格式
https://officeguide.cc/excel-add-or-subtract-time/

1.日期聯動

112/02/28	112/02/28	決算日
LEFT($B$3,7)&DAY(EOMONTH($B$13,-1))	112/01/31	上月底
IF(MID($B$1,5,2)-1<1,MID($B$1,1,3)-1&"/"&$B$10,MID($B$1,1,3)&"/"&$B$10)	112/01/01	上月初
MID($B$1,1,3)-1&"/"&MID($B$1,5,2)&"/"&DAY(EOMONTH($B$13,-12))	111/02/28	去年同期決算日
MID($B$1,1,3)&MID($B$1,4,3)	112/02	決算月
MID($B$1,1,3)-1&MID($B$1,4,3)	111/02	去年同月
MID($B$1,1,3)&MID($B$1,5,2)	11202	決算年月(無/)
IF(LEN($B$11)<5,"0"&$B$11,$B$11)	03/01	下月初(月日)
MID($B$1,5,5)	02/28	決算日(月日)
IF(LEN($B$12)<5,"0"&$B$12,$B$12)	01/01	上月初(月日)
IF(MID($B$1,5,2)+1>12,MID($B$1,5,2)+1-12,MID($B$1,5,2)+1)&"/01"	3/01	下月初(未補齊)
IF(MID($B$1,5,2)-1<1,MID($B$1,5,2)-1+12,MID($B$1,5,2)-1)&"/01"	1/01	上月初(未補齊)
TEXT(DATE(MID($B$1,1,3)+1911,MID($B$1,5,2),MID($B$1,8,2)),"yyyy/mm/dd")	2023/02/28	決算日(西元)
IF(DAY(TODAY())*1<11,YEAR(EOMONTH(TODAY(),-1))-1911&"/"&MONTH(EOMONTH(TODAY(),-1))&"/"&DAY(EOMONTH(TODAY(),-1)),YEAR(EOMONTH(TODAY(),0))-1911&"/"&MONTH(EOMONTH(TODAY(),0))&"/"&DAY(EOMONTH(TODAY(),0)))	112/2/28	決算日(當日判斷)
		
LEFT($B$26,3)&" ("&MID($B$26,5,2)&")"	111 (11)	"&$B$16&"
LEFT($B$20,3)&" ("&MID($B$20,5,2)&")"	111 (12)	"&$B$17&"
LEFT($B$2,3)&" ("&MID($B$2,5,2)&")"	112 (01)	"&$B$18&"
MID($B$1,1,3)&" ("&MID($B$1,5,2)&")"	112 (02)	"&$B$19&"
LEFT($B$21,7)&DAY(EOMONTH($B$13,-2))	111/12/31	上上月底
IF(MID($B$1,5,2)-2<1,MID($B$1,1,3)-1&"/"&$B$22,MID($B$1,1,3)&"/"&$B$22)	111/12/01	上上月初
IF(LEN($B$23)<5,"0"&$B$23,$B$23)	12/01	已補零
IF(MID($B$1,5,2)-2<1,MID($B$1,5,2)-2+12,MID($B$1,5,2)-2)&"/01"	12/01	未補零
LEFT($B$25,7)&DAY(EOMONTH($B$13,1))	112/03/31	下月底
IF(MID($B$1,5,2)+1>12,MID($B$1,1,3)+1&"/"&$B$8,MID($B$1,1,3)&"/"&$B$8)	112/03/01	下月初
LEFT($B$27,7)&DAY(EOMONTH($B$13,-3))	111/11/30	上上上月底
IF(MID($B$1,5,2)-3<1,MID($B$1,1,3)-1&"/"&$B$28,MID($B$1,1,3)&"/"&$B$28)	111/11/01	上上上月初
IF(LEN($B$29)<5,"0"&$B$29,$B$29)	11/01	已補零
IF(MID($B$1,5,2)-3<1,MID($B$1,5,2)-3+12,MID($B$1,5,2)-3)&"/01"	11/01	未補零
LEFT($B$20,3)&MID($B$20,5,2)	11112	"&$B$30&"
LEFT($B$2,3)&MID($B$2,5,2)	11201	"&$B$31&"
MID($B$1,1,3)&MID($B$1,5,2)	11202	"&$B$32&"
LEFT($B$20,3)	111	上上決算月(年份)
LEFT($B$2,3)	112	上個決算月(年份)
LEFT($B$1,3)	112	當前決算月(年份)
TEXT(DATE(MID($B$2,1,3)+1911,MID($B$2,5,2),MID($B$2,8,2)),"yyyy/mm/dd")	2023/01/31	上月底(西元)

2.匯出

Sub 匯出()

	'定義屬性
	Dim MYstr As String, i As Integer
    
	'定義Output File位置
	Open "D:\Practice\巨集\月結\03.dist_upr巨集\11202\01.複製新月工作表(dist_upr)_11202.txt" For Output As #1
	For i = 2 To 100
  	    MYstr = Worksheets("01.複製新月工作表(dist_upr)").Cells(i, 1)
	    Print #1, MYstr
	Next i
	Close #1

	'定義Output File位置
	Open "D:\Practice\巨集\月結\03.dist_upr巨集\11202\02.link改成新月份(dist_upr)_11202.txt" For Output As #2
	For i = 2 To 200
	    MYstr = Worksheets("02.link改成新月份(dist_upr)").Cells(i, 1)
	    Print #2, MYstr
	Next i
	Close #2

	End Sub


3.Call出子程序

	Sub 全部()

	    ''可以依序call出子程序如:Sub A()、Sub B()、...
	    Call link調整成新月份
	    Call 匯出
	
	End Sub
	
	Sub link調整成新月份()
	End sub
	
	Sub 匯出
	End sub

4.複製出新月工作表

	Sheets(Array("112 (01)", "112 (01)-美元", "112 (01)-歐元", "112 (01)-澳幣", _
	"112 (01)-南非幣", "112 (01)-人民幣")).Copy Before:=Sheets(1)

5.改正工作表名稱為新月份

跨年度
=IF(MID($B$1,5,2)<>"01","","    Sheets("""&$B$34&" (13)"").Name = """&$B$19&"""")

    Sheets("112 (01)-美元 (2)").Name = "112 (02)-美元"

6.link調整成新月份

    Sheets("112 (02)-美元").Activate
    Cells.Replace What:="syy_11201", Replacement:="syy_11202", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

="    Sheets("""&$B$19&"-美元"").Activate"
="    Cells.Replace What:=""syy_"&$B$31&""", Replacement:=""syy_"&$B$32&""", LookAt:=xlPart _"
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


7.複製到新工作簿

    Sheets(Array("未適格112 (02)", "dist_112 (02)-OIU", "dist_112 (02)", _
        "dist_112 (02)-歐元", "dist_112 (02)-南非幣", "dist_112 (02)-人民幣", "dist_112 (02)-澳幣", _
        "dist_112 (02)-美元", "112 (02)", "112 (02)-OIU", "112 (02)-歐元", "112 (02)-南非幣", _
        "112 (02)-人民幣", "112 (02)-澳幣", "112 (02)-美元", "112 (02)-業主權益下", "acc112 (02)", _
        "112 (02)-業主權益下(prt)")).Select
    
    Sheets(Array("未適格112 (02)", "dist_112 (02)-OIU", "dist_112 (02)", _
        "dist_112 (02)-歐元", "dist_112 (02)-南非幣", "dist_112 (02)-人民幣", "dist_112 (02)-澳幣", _
        "dist_112 (02)-美元", "112 (02)", "112 (02)-OIU", "112 (02)-歐元", "112 (02)-南非幣", _
        "112 (02)-人民幣", "112 (02)-澳幣", "112 (02)-美元", "112 (02)-業主權益下", "acc112 (02)", _
        "112 (02)-業主權益下(prt)")).Copy

8.貼死

    Sheets("112 (02)-業主權益下(prt)").UsedRange = Sheets("112 (02)-業主權益下(prt)").UsedRange.Value
    
    Sheets("acc112 (02)").UsedRange = Sheets("acc112 (02)").UsedRange.Value

9.另存新檔到指定資料夾並重新命名

        ChDir "D:\My Documents\rsv_st\Y112\總表"
    ActiveWorkbook.SaveAs Filename:= _
        "D:\My Documents\rsv_st\Y112\總表\rsv_acct11202(value).xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        
    Workbooks("rsv_acct112.xlsx").Sheets("未適格112 (02)").Activate


10.將all_txt逐一貼入工作表1()

    Dim str_txt() As String, line As Integer, i As Integer, txt As String
    line = 14
    Open "D:\My Documents\rsv_st\Y112\112data\rsvsql11201\all_txt\upr_arco_a.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, txt
        str_txt = Split(txt, "|")
        For i = 0 To UBound(str_txt)
            Sheets("工作表1").Cells(line, i + 1).Value = str_txt(i)
        Next i
        line = line + 1
    Loop
    Close #1
          
    line = 14
    Open "D:\My Documents\rsv_st\Y112\112data\rsvsql11201\all_txt\vupr_temp.txt" For Input As #2
    Do While Not EOF(2)
        Line Input #2, txt
        str_txt = Split(txt, "|")
        For i = 0 To UBound(str_txt)
            Sheets("工作表1").Cells(line, i + 12).Value = str_txt(i)
        Next i
        line = line + 1
    Loop
    Close #2

11.     資料列函數欄位補齊
	資料列數少於函數列數則刪除多餘欄位之函數
	資料列數多於函數列數則補齊缺少函數之欄位

    i = 14
    Sheets("工作表1").Select
    
    While ActiveSheet.Cells(i, 9) <> "" Or ActiveSheet.Cells(i, 7) <> ""
    
        If ActiveSheet.Cells(i, 7) = "" And ActiveSheet.Cells(i, 9) <> "" Then
        ActiveSheet.Range(Cells(i, 9), Cells(i, 10)).Select
        Selection.ClearContents
        ElseIf Cells(i, 7) <> "" And Cells(i, 9) = "" Then
        ActiveSheet.Range(Cells(i - 1, 9), Cells(i - 1, 10)).Select
        Selection.AutoFill Destination:=Range(Cells(i - 1, 9), Cells(i, 10)), Type:=xlFillDefault
        
        End If
        
        i = i + 1
        
    Wend

12.Sheet("DB")內容清除 巨集
'
    Sheets("DB").Select
    Range("Q1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

13.工作表1貼入DB 巨集

    Sheets("工作表1").Select
    Range("A14:E14").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Sheets("DB").Select
    Range("Q1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

14.白底 巨集
    Sheets("DB").Columns("P:AO").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


15.非1填黃底 巨集
   
    i = 1
    
    Sheets("DB").Select
    
    While ActiveSheet.Cells(i, 16) <> ""
    
        If ActiveSheet.Cells(i, 16) <> 1 Then
        ActiveSheet.Range(Cells(i, 16), Cells(i, 23)).Select
        
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
        End If
        i = i + 1
    Wend

1.以Excel開啟txt文件並匯入

開啟檔案ㄉ路徑
  Workbooks.OpenText Filename:="D:\My Documents\rsv_st\Y111\111data\rsvsql11109\all_txt\upr_arco_a_11109.txt"
資料剖析的那些打勾選項(連續空格設定為單1.其他都不勾選)
      , Origin:=950, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|",
設定分類選項 1~7為第幾類.後面那個1是設定顯示.隱藏.或格式
FieldInfo:=Array(Array(1, 1),Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1)),  TrailingMinusNumbers:=True


2.選擇性貼上值:
    ActiveSheet.PasteSpecial Format:="Unicode 文字", Link:=False, DisplayAsIcon _
        :=False, NoHTMLFormatting:=True

3.資料剖析(以|為分隔符號):
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1)), TrailingMinusNumbers:=True




4.將txt檔內容以|符號分割並逐直行貼入Excel中:
ps.改為line = x, Cells(line, i + y) 將可以指定從(x,y)欄位為貼入起始欄位
#1為給txt檔一個臨時編號


Sub file_txt()
    Dim str_txt() As String, line As Integer, i As Integer, txt As String
    line = 2
    Open "D:\My Documents\rsv_st\Y111\111data\rsvsql11109\all_txt\upr_arco_a_11109.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, txt
        str_txt = Split(txt, "|")
         For i = 0 To UBound(str_txt)
            Cells(line, i + 10).Value = str_txt(i)
         Next i
        line = line + 1
     Loop
    Close #1
    

 End Sub

5.Excel 大量字串取代 VBA
https://classic-blog.udn.com/mobile/WayCheng/2808290


6.程式執行完畢彈出執行時間

Sub Main()
    Dim start As Date

    start = now()
    '可以依序call出子程序如:Sub A()、Sub B()、...
    Call A
    Call B
    Call C
    Msgbox Format(now - start, "HH:mm:ss")

End Sub




































