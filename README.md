# Excel-Crawling
透過Excel爬下三民書局暢銷榜的書名、作者、價格....等,並自動整理成表格




Option Explicit
#If Win64 Then
  Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)                        '(3)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)                               '(4)
#End If


Private objIE As InternetExplorer

Sub sanmin()
    '建立IE物件
    Dim url, row   '宣告變數
    On Error Resume Next
    row = 1

        url = "https://www.sanmin.com.tw/promote/top/?id=WBYY&vs=grid&item=1101220&pi=1&vs=list"  '三民書局
    
    Set objIE = New InternetExplorer
    
        objIE.Visible = False
        objIE.Navigate2 (url)


'等待讀取完成

        While objIE.readyState <> READYSTATE_COMPLETE Or objIE.Busy = True
            DoEvents
            Sleep 100
        Wend
        Sleep 100

        Dim objDoc As HTMLElementCollection
        Set objDoc = objIE.document

        Dim BookName, Author, Price As IHTMLElement  '宣告變數指定型態
 

        For Each BookName In objDoc.getElementsByClassName("resultBooksInfor")
        '確認書名在網頁的位置
            Worksheets(1).Cells(row, 1) = row                                  '寫入工作表與儲存格
            Worksheets(1).Cells(row, 2) = BookName.innerText
            row = row + 1
        Next
                
        row = 1
        For Each Author In objDoc.getElementsByClassName("author")
        '確認書名在網頁的位置
            Worksheets(1).Cells(row, 3) = Author.innerText
            row = row + 1
        Next
        
        row = 1
        For Each Price In objDoc.getElementsByClassName("resultBooksLayout")
        '確認書名在網頁的位置
            Worksheets(1).Cells(row, 4) = Price.innerText
            row = row + 1
        Next
    
    Set objIE = Nothing
    MsgBox "Done ！"
End Sub






////也可以打成////



Option Explicit
#If Win64 Then
  Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)                        '(3)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)                               '(4)
#End If


Private objIE As InternetExplorer

Sub sanmin()
    '建立IE物件
    Dim url, row   '宣告變數
    On Error Resume Next
    row = 1

        url = "https://www.sanmin.com.tw/promote/top/?id=WBYY&vs=grid&item=1101220&pi=1&vs=list"  '三民書局
    
    Set objIE = New InternetExplorer
    
        objIE.Visible = False
        objIE.Navigate2 (url)


'等待讀取完成

        While objIE.readyState <> READYSTATE_COMPLETE Or objIE.Busy = True
            DoEvents
            Sleep 100
        Wend
        Sleep 100

        Dim objDoc As HTMLElementCollection
        Set objDoc = objIE.document

        Dim BookName As IHTMLElement  '宣告變數指定型態
 

        For Each BookName In objDoc.getElementsByClassName("resultBooksInfor")
        '確認書名在網頁的位置
            Worksheets(1).Cells(row, 1) = row                                  '寫入工作表與儲存格
            Worksheets(1).Cells(row, 2) = BookName.Children(0).innerText
            Worksheets(1).Cells(row, 3) = BookName.Children(1).Children(0).innerText
            Worksheets(1).Cells(row, 4) = BookName.Children(1).Children(1).innerText
            Worksheets(1).Cells(row, 5) = BookName.Children(2).innerText
            Worksheets(1).Cells(row, 6) = BookName.Children(4).Children(0).Children(0).innerText
            Worksheets(1).Cells(row, 7) = BookName.Children(4).Children(0).Children(1).innerText
            Worksheets(1).Cells(row, 8) = BookName.Children(4).Children(1).innerText
            row = row + 1

        Next
    
    Set objIE = Nothing
    MsgBox "Done ！"
End Sub


