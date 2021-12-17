# BUAA_Dissertation_Template
北航硕博研究生毕业论文模板（Word）


## 特点

- ✅参照了[《北京航空航天大学研究生撰写学位论文的规定》](http://graduate.buaa.edu.cn/info/1039/7831.htm)2021年11月的最新修订要求
  
- ✅专业/学术学位，硕士/博士，一套模板
  
- ✅按顺序组织了论文结构，开箱即用
  
- ✅自动修改题注格式，无需操心这里那里的空格
  
- ✅搭配NoteExpress使用，自动排版参考文献
  
- ……


## 模板使用方法

- 使用时请显示所有批注，按照批注进行更改即可，更改完后，审阅-删除下拉按钮-删除文档中的所有批注。

- 注意区分专业/学术学位硕/博士的中/英文封面和提名页，删除不需要的中/英文封面和提名页。

- Word公式编辑可以使用[这个工具](https://github.com/ixzhao/ixzhao.github.io)。
  
- 参考文献的自动排版，可以使用NoteExpress的样式文件“北航硕博论文.nes”，如果你平时用NoteExpress管理你的文献。
  用仓库里的nes文件直接替换C:\Users\你的用户名\AppData\Roaming\NoteExpress\Styles路径里的原文件（记得先备份）；如果在授权IP范围以外，可以替换7个默认样式之一的同名文件。
    ```
    NoteExpress的个人标准版是不允许在授权IP范围以外使用的，而且限制只能使用7个默认的样式，样式格式也不能编辑。所以推荐使用集团版，可以在学校图书馆下载。或者到 http://www.inoteexpress.com/download_chs.htm#Downloads 找一个可以下载的高校集团版本（该途径低调下载使用，说不定哪天页面就打不开了）。
    ```
  
- 建议启用宏，如果你不需要，可以视图-宏-查看宏-编辑，进入后删除，详细见[其他事项](##其他事项)。
  
- 题注插入：插入时，选中图/表，引用-插入题注，选择图/表标签以及题注位置；正文中引用时，引用-交叉引用，选择引用类型以及引用内容。
  
- 对于目录、图标清单以及域的自动更新，Ctrl+A全选，F9更新。
  
- 其他内容可参照《规定》填写。
  
  在线乞讨环节：
  如果本仓库大大地节省了你的时间，何尝不请我一杯小小的Coffee？

    <span><img src="https://raw.githubusercontent.com/ixzhao/BUAA_Dissertation_Template/main/image/alipay.jpg" width="40%"/>&nbsp;&nbsp;&nbsp;&nbsp;<img src="https://raw.githubusercontent.com/ixzhao/BUAA_Dissertation_Template/main/image/wechat.jpg" width="40%"/></span>



## 其他事项

- 模板是启用宏的Word文件，主要考虑2个方面：
  一是题注标签只能保存到本机上Office默认的Normal.dotm模板里（参考[这里](https://www.msofficeforums.com/word/15715-captions-self-defined.html#2)），所以只能通过宏（现查的VBA的编程规范，头发掉了一把😢），在每次打开文档时自动添加“图”和“表”这两个题注标签。如果不需要，可以进入宏工程后，在左侧窗口Project-ThisDocument双击，清空内容。
  二是《规定》要求的题注格式为“图1空格空格XXX”，这与Word默认的题注格式不同，手动更改太麻烦，所以还是通过宏（借鉴的[这里](http://blog.sina.com.cn/s/blog_51817ae50102w8mz.html)），插入题注时自动删除标签与编号前的空格，自动在编号后添加2个空格。如果不需要，直接在查看宏界面删除即可。
  
- 感谢 [heckBoxStudio/BUAAThesis](https://github.com/CheckBoxStudio/BUAAThesis)。
  
- 待补充。
  

















