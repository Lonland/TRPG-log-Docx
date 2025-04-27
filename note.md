# 需要工具
1. 安装了word的电脑
2. 海豹跑团log着色器[网站](log.weizaima.com)
3. 你的跑团log _如果直接由海豹导出会更好_
4. 愿意整理log的勤劳双手

> 以下内容包括少量**万人无我**剧透（用了万人无我log教学&懒得打码），担心被剧透请尽量避开看文字  
> 有疑问可以联系作者QQ:3065136818

# 单人团/非秘密团（单个链接）
### 一步到位版：选择下载海豹导出的**对话log**
### 复杂但有助于理解多人团版
1. 开启上方除首行缩进对齐&深色模式的全部开关-下载word
2. 选择偏好的名字格式: 左对齐or前有空格的左/中/右对齐(下面两个方法只有少部分不一样)
   > 需要竖线：右键-段落-制表符，设置你想要有竖线的长度竖线对齐
   1. 左对齐
      1. Ctrl+H 替换所有 **<名字>:** 为 **^t<名字>^t** (将名字替换为log里的所有人名)
         > 可以靠Ctrl+F 搜索 **>:** 是否存在，以确认是否全部替换完成  
         > 这一步两个格式是一样的
      2. Ctrl+H 替换所有 **^p** 为 **^l**, 再替换 **^l^t** 为 **^p**
         > 这一步是为了避免rp中的换行影响悬挂缩进  
         > 需要手动删除第一行最前面的制表符
      3. 全选文段，右键-段落-悬挂缩进-缩进量，缩进量需要大于最长的名字，可以多试几次找到最美观的
   2. 前有空格对齐
      1. Ctrl+H 替换所有 **<名字>:** 为 **^t<名字>^t** (将名字替换为log里的所有人名)
         > 可以靠Ctrl+F 搜索 **>:** 是否存在，以确认是否全部替换完成  
         > 这一步两个格式是一样的
      2. Ctrl+H 替换所有 **^p** 为 **^l**, 再替换 **^l^t** 为 **^p^t**
         > 这一步是为了避免rp中的换行影响悬挂缩进  
      3. 全选文段，右键-段落-制表符（需要设置两个制表符）
         1. 左对齐：设置你想缩进的量的长度，选左对齐（一般为2）+你想让rp缩进的长度，选左对齐（和下面悬挂缩进的量等同）
         2. 居中对齐：设置你想让名字居中对齐的量长度，选居中对齐（一般为最长名字长度一半）+你想让rp缩进的长度，选左对齐（和下面悬挂缩进的量等同）
         3. 右对齐：设置你想让名字右对齐的长度，选右对齐（一般为rp缩进量-2）+你想让rp缩进的长度，选左对齐（和下面悬挂缩进的量等同）
      4. 全选文段，右键-段落-悬挂缩进-缩进量
3. 调整不同人rp之间的间隔
   - 全选文段，右键-段落-段前段后间距均调为0.5行(根据自己喜好可以改编数值)

# 多人秘密团
> 需要工具新增excel  
> 标出**存档点**的部分是提醒你备份存档，避免有什么问题导致一切从头开始，熟练后可以不备份   
1. log染色界面
   1. 将所有log的pc颜色调整好（这个补救代码没写好，如果pc颜色没调好无法补救）
      > 没想好颜色可以把所有人颜色设置的不统一（请确保没设置错误），后期替换来进行修改
      > 但是如果pc之间颜色混了的话表格的根据名字替换颜色我没写代码。。。   
   2. 开关需关闭时间显示过滤和年月日不显示，其他打开，下载word文件
2. 将所有下载的word复制到同一个文件中，复制前记得标注一下这是那个群
   > 推荐使用2025HOx这种正常log中不会出现的标记，方便分段的时候靠ctrl+F搜索  
   > 前面有2025是为了方便第四步的时候把这里换行符留着
3. Ctrl+H 替换所有 **<名字>:** 为 **^t<名字>^t** (将名字替换为log里的所有人名)
   > 可以靠Ctrl+F 搜索 **>:** 是否存在，以确认是否全部替换完成
4. Ctrl+H 替换所有 **^p** 为 **{换行符}**, 再替换 **{换行符}2025** 为 **^p2025**
   > 区分rp内分段和不同rp间分段
5. ctrl+A复制全文，打开excel，在2B格复制，可以看见log已经分为三栏，时间/角色/RP
6. 在第一列填充你希望的小群排序顺序
   1. 推荐大群0，小群按照ho几就填几，多人小群就有几个人填几位
      > 防止有人不知道，excel只需要你填两个，然后选中格子下拉就可以自动填充同样的数字
   2. 给小群上底色，推荐是pc颜色亮度调为240（如果不喜欢底色，以后会补加框教程）   
**存档点**
7. 选择A-D四列，进行按照B列排序
8. 检索全文，出现比较明显的同时开多个小群的rp（主要表现形式是一整块都是有混乱底色），选择那一区块进行优先A列排序
9. 将排序完成的表格选择CD列复制回新的word文件中（这一步会比较卡）
10. 点击表格属性，在列那一栏进行宽度调整
11. 将{换行符}替换为^l
    > 或者^p，我个人喜欢^l因为如果去掉表格形式，^l也能快速匹配rp间不往前缩
12. 自行调整字体/行间距等进行美化
13. 如果需要插入标题/分割表格，在需要分割的下一行表格ctrl+shift+enter

# 神秘小巧思&小技巧
## 底页插入
1. 双击页眉处出现页眉编辑（也可以插入-页眉进入）  
2. 复制图片，调整图片格式为在文字下方，调整图片大小覆盖整个页面
## 目录
1. 用自带的标题一标题二插入标题   
> 一般我喜欢标题一设置段前分页
2. 在开头插入word自动目录  
3. 需要调整页码：   
   3.1 选中要设定为页码1的页数：布局-分隔符下一页    
   3.2 插入-页码-设置页码格式-页码编号-起始页码   
## 样式
设置样式用于替换，比如设置一个样式固定段落边框颜色-快速替换正文中的小窗部分
## 信息框
## 名字+下划线式log
设置名字样式：加上下划线，将所有名字替换成名字样式   
## 带头像log
先保存文件（方便之后图片插入位置不理想，调试代码时候不保存直接退出）
关闭自动保存！！！
第一张图片生成位置有问题很正常，代码运行结束手动调整一下
使用以下VB代码插入图片
```
Sub InsertImageBeforeEachNamedParagraph()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim para As Paragraph
    Dim shape As shape
    Dim nameMap As Object ' Dictionary
    Dim key As Variant
    Dim imgPath As String
    Dim paraText As String
    Dim paraIndex As Long: paraIndex = 0

    ' ? 创建名字与图片路径的映射表
    Set nameMap = CreateObject("Scripting.Dictionary")
    nameMap.Add "<名字1>", "D:\Users\...\1.jpg"
    nameMap.Add "<名字2>", "D:\Users\...\2.jpg"
    ' ?? 替换路径为你实际使用的图片路径
    startIndex = Selection.Paragraphs(1).Range.Start
    ' ? 遍历文档每一段
    For Each para In ActiveDocument.Paragraphs
    If para.Range.Start < startIndex Then GoTo ContinueLoop
        If Left(Trim(para.Range.Text), 1) <> "<" Then GoTo ContinueLoop
        paraIndex = paraIndex + 1
        ' ? 每处理一百段进行询问，如果不需要询问可以删除这个If，也可以看看效果
        If paraIndex Mod 100 = 0 Then
            If MsgBox("已处理 " & paraIndex & " 段，是否继续？", vbYesNo) = vbNo Then Exit Sub
        End If
        paraText = Trim(para.Range.Text)

        ' 遍历名字列表，逐个匹配
        For Each key In nameMap.Keys
            If paraText Like key & vbCr & "*" Then
                ' ? 插入对应图片
                Set shape = ActiveDocument.Shapes.AddPicture( _
                    FileName:=nameMap(key), _
                    LinkToFile:=False, _
                    SaveWithDocument:=True, _
                    Anchor:=para.Range)

                With shape
                    .LockAspectRatio = msoTrue
                    .Width = 40 ' 约等于 8 个汉字宽
                    
                    .WrapFormat.Type = wdWrapTight
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    .Left = InchesToPoints(0.5) ' 位于段落缩进区域

                    .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
                    .Top = 0 ' 补偿段前间距
                End With

                Exit For ' 匹配后退出内层循环
            End If
        Next key
ContinueLoop:
    Next para

    MsgBox "已为每个名字段落插入对应图片。", vbInformation
End Sub
```
## 名字染色错误补救（可以在手动输入npc名字后进行大面积npcrp颜色染色）
* 如果名字颜色没有混   
> 可以使用高级替换-格式-字体-颜色替换   
* rp出现混色（比如多人rp是统一颜色），通过以下代码补救   
用于格式是<名字> [tab] rp，如果不是该格式请自行微调代码   
```
Sub HighlightSpecificNameWithTab()
    Dim rng As Range
    Dim findText As String
    Dim highlightColor As Long

    ' **设定要查找的名字（包含尖括号）**
    findText = "<名字>^t"
    
    ' **设定高亮颜色（RGB 颜色值）**
    highlightColor = RGB(0, 0, 0) ' 例如：#FFFFFF 的 RGB(255, 255, 255)

    ' **移动到文档开头**
    Selection.HomeKey Unit:=wdStory

    ' **查找匹配的 "<名字>"**
    With Selection.Find
        .Text = findText
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True ' **区分大小写**
        .MatchWholeWord = False ' **允许部分匹配**
    End With

    ' **遍历全文，查找并染色**
    Do While Selection.Find.Execute
        ' **扩展范围至整行，包括“内容”**
        Set rng = Selection.Range
        rng.MoveEnd Unit:=wdParagraph, Count:=1 ' **扩展至换行符**
        
        ' **应用颜色**
        rng.Font.Color = highlightColor
    Loop

    MsgBox "所有 '" & findText & "' 相关行已染色！", vbInformation
End Sub

```

