# 项目介绍 #

都说创造来源于生活，最近就职后正赶上公司给客户（政府）试用一个项目，作为刚刚入职的菜鸟正兴趣勃勃研究项目代码解构，发现一个不合理地方，就是对POI操作真没一个像样工具类，而且政府项目对表格需求业务比较大，大部分代码逻辑都在处理解析表格去了，我想写一个能快速解析表格一个解析工具并对其对应的类进行封装，那后续岂不是只要专注于业务代码，完全符合MVC设计模式，当然具体网络上有没有更好用工具人，呸工具类呢，我暂时没有了解到，所以按照我的想法，准备对POI工具更深入封装，做Spring最擅长事情-----封装！

## 使用方法 ##

导入本项目jar包或直接复制项目核心代码（注意复制没有依赖）

本工具类提供DEMO测试类，请参考PoiToolTest.java

核心类ApachePOIUtils

静态方法readExcel

- file/inputStream 文件对象/输入流
- oClass 字节码对象
- indexSheet 工作表
- indexRowNum 第几行开始解析（使用注释失效）
- indexCell 第几列开始解析（使用注释失效）

返回static <T> List<T>

-------------
静态方法readExcel2String

- file/inputStream 文件对象/输入流
- indexSheet 工作表
- indexRowNum 第几行开始解析
- indexCell 第几列开始解析
- endCell 第几列结束解析(非必传)

返回List<String[]>

## 更新日记 ##

- 2020年4月11日 10:22:30

给静态方法readExcel2String添加一个参数，可指定解析结束的列数

移除对Hutool工具依赖,需要的工具自己实现

- 2020年4月1日 13:04:10

添加纯字符串解析，可以自定义自己解析规则

- 2020年3月31日 10:53:02

添加注解方式的解析模式，让解析更佳灵活，我准备让我这个工具类在正式项目上开始试用！

## 技术选型 ##

- Maven
- 反射原理
- 注解原理

# 其实我很看好我的想法，我也第一次当工具人，如果使用上有什么问题，或者bug出现，请反馈到这个邮箱谢谢，感谢你们一路支持。未完待续！ #
## mr.xinchen@icloud.com ##


