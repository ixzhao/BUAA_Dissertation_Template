# BUAA_Dissertation_Template
北航硕博研究生毕业论文模板（Word）

点击下载  https://github.com/ixzhao/BUAA_Dissertation_Template/archive/refs/heads/main.zip


## 特点

- ✅参照了[《北京航空航天大学研究生撰写学位论文的规定》](http://graduate.buaa.edu.cn/info/1039/7831.htm)2021年11月的最新修订要求
- ✅专业/学术学位、硕士/博士，尽在一套模板
- ✅按顺序组织了论文结构，开箱即用
- ✅题注格式自动修改、公式自动编号，符合《规定》的格式要求
- ✅参考文献自动排版（需搭配NoteExpress使用）

模板基于Microsoft Office 2013 / 2019，如其他Office版本有兼容问题，请提[Issues](https://github.com/ixzhao/BUAA_Dissertation_Template/issues/new/choose)。

## 模板使用说明

1. 《规定》中理工类与社科文学类的唯一区别在于章节标题格式的不同。为了简化样式数量，区分了两个版本，社科文学类请使用带后缀的模板。

2. 请在选项-信任中心-信任中心设置-宏设置-启用宏中启用所有宏，如果你不明白[使用宏所带来的风险的信息](https://support.microsoft.com/zh-cn/office/%E5%90%AF%E7%94%A8%E6%88%96%E7%A6%81%E7%94%A8-office-%E6%96%87%E4%BB%B6%E4%B8%AD%E7%9A%84%E5%AE%8F-12b036fd-d140-4e74-b45e-16fed1a7e5c6)，不建议启用，你将需要手动创建“图”“表”这两个题注标签、手动修改每一个题注格式、手动编号排版公式。详细见[其他事项1](#macro)。
  
    运行宏的几种方式：
      - 通过视图--宏，或开发工具--宏，或Alt+F8，呼出宏对话框，双击指定宏或选中后点运行即可。
      - 将指定宏添加到功能区或快速访问工具栏。
      - 为指定宏添加自定义键盘快捷键。
  
3. 请在审阅-修订中显示所有批注，并将显示标注方式改为“在批注框中显示修订”。请阅读所有批注，按照批注进行更改；注意区分专业/学术学位硕/博士的中/英文封面和提名页，删除不需要的中/英文封面和提名页。更改完后，审阅-删除下拉按钮-删除文档中的所有批注。

4. 图/表题注格式自动修改：选中图/表，引用-插入题注或者运行宏InsertCaption，在弹出的题注对话框中直接输入即可，无需调整空格。

5. 公式自动编号：在需要插入公式的地方直接运行宏InsertEquation即可。

6. 参考文献自动排版，可以借助NoteExpress的样式文件，如果你平时用NoteExpress管理你的文献。但是NoteExpress自带的北航硕博论文样式已经很久不更新了，可以用本仓库中的“北航硕博论文.nes”替换。  
  
    在 C:\Users\用户名\AppData\Roaming\NoteExpress\Styles 路径下，直接替换原文件“北京航空航天大学博硕士论文.nes”（记得先备份再改名替换）；如果在授权IP范围以外，可以先将文件重命名为7个默认样式之一然后再替换。  

        NoteExpress的个人标准版是不允许在授权IP范围以外使用的，而且限制只能使用7个默认的样式，样式格式也不能编辑。所以推荐使用集团版，可以在学校图书馆首页-右侧快速链接-常用工具下载。或者到 http://www.inoteexpress.com/download_chs.htm#Downloads 找一个可以在授权IP范围外下载的高校集团版本，比如北大版 http://www.inoteexpress.com/support/cgi-bin/download_sch.cgi?code=BeiDa（该途径低调下载使用，说不定哪天页面就打不开了）。
  
7. 建议在开始-段落-显示编辑标记、在状态栏右键勾选“节”。
    
      ❗❗❗ **注意：** 
      在打印/输出PDF文件时，很有可能会在3个地方分别多出一页空白页：摘要前、第一章前、总结与展望前，个人建议将Word文件输出为PDF文件，然后使用Adobe Acrobat等工具删除多余的空白页。具体原因见[其他事项2](#blankpage)。


## 其他事项

1. <span><a name="macro"></a></span>模板是启用宏的Word文件，主要考虑3个方面：
   - 题注标签只能保存到本机上Office默认的Normal.dotm模板里（参考[这里](https://www.msofficeforums.com/word/15715-captions-self-defined.html#2)），所以只能通过宏，在每次打开文档时自动添加“图”和“表”这两个题注标签。如果不需要，可以进入宏工程后，在左侧窗口Project-ThisDocument双击，清空内容。  

   - 《规定》要求的题注格式为“图1空格空格XXX”，这与Word默认的题注格式不同，手动更改太麻烦，所以还是通过宏（借鉴的[这里](http://blog.sina.com.cn/s/blog_51817ae50102w8mz.html)），插入题注时自动删除标签与编号前的空格，自动在编号后添加2个空格。如果不需要，直接按Alt+F8在查看宏界面删除即可。  

   - Word对公式自动编号的支持不太好，格式也不符合要求，通过宏（借助表格排版）可以方便地实现编号排版。
  
2. <span><a name="blankpage"></a></span>由于使用了分节符、（页眉页脚）奇偶页不同等选项，导致Word会强制将每一节的第一页显示在纸的右页（正面），即新的一节总是从奇数页开始，如果上一节结束在奇数页，那么将多出一页偶数页，具体见[这个页面最下边的Important Warning](http://wordfaqs.ssbarnhill.com/BlankPage.htm#True_blank_pages)。  
虽然可以通过更改分节符的类型规避空白页，但是如果文档有新的修改后都可能改变奇偶页的布局，因此建议的方式是将Word文件输出为PDF后删除多余的空白页。[这里](https://answers.microsoft.com/en-us/msoffice/forum/all/word-print-preview-adds-extra-blank-pages/05710a98-838f-4b74-9b7c-e57b8c63eda3)给出的建议也是对导出的PDF文件进行修改。

  

  

















