# MyVBA
Sub test()
 
'本程序的作为是从wind上下载时间序列，将数据按年份、月份重新整理，并作图
'以观察数据的季节性特征
Dim arr As Object
Dim c, mydata As Range
Dim i, j, t As Integer
ScreenUpdating = False
 
Set d = CreateObject("scripting.dictionary") '创建字典对象，用于计算有多少年的数据
Set arr = Range("a4", [a65535].End(xlUp))
 
For Each c In arr

d(Year(c)) = ""
j = Year(c) - Year([a4])
i = Month(c)
t = d.Count
 
[e1].Offset(i, 0) = i & "月"
[e1].Offset(0, j + 1) = Year(c)
[e1].Offset(i, j + 1) = c.Offset(0, 1)

Next

Set mydata = ActiveSheet.[e1].Resize(13, t + 1) '整理后的数据
 
'计算最大值，最小值,用于设定坐标轴的最大值、最小值，以增强显示的美观度
Min = Application.WorksheetFunction.Min(mydata.Offset(1, 1))
Max = Application.WorksheetFunction.Max(mydata.Offset(1, 1))
 
'作图
Charts.Add.SetSourceData Source:=mydata, PlotBy:=xlColumns
'设定图表格式
With ActiveChart
.ChartType = xlLine
.Axes(xlValue).MinimumScale = Min * 0.94
.Axes(xlValue).MaximumScale = Max * 1.06
 
End With
Sheets(Sheets.Count).Activate
ScreenUpdating = True
 
End Sub


