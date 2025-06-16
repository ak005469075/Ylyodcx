# 目的

为了**解决手动选择性粘贴**docx1文件的部分内容到docx2去，用python实现自动化处理

原理：docx的本质是xml文件的处理，可解压查看，更加规范，本项目用到python-docx库

平台：**windows 10**

# 介绍

默认适用于仅一个标题且行数为3行的情况

1.支持docx文件原标题、内容和末尾浮动图片(自动判断)的粘贴

2.输出名为docx文件原内容中的单号，找不到就输出为123.docx

# 结构

```bash
tips/
	utils.py #docx空行清除
core/
	Para_Handle.py #用于处理文本，标题和正文
	Pic_Handle.py #用于处理浮动图片
config/
	easy_settings.py #加载配置文件
	config.json #配置文件
muban.docx #以muban.docx为模板，替换后生成目标文件
main.py 启动
```

# 使用

输入原文件路径，输入需替换的序号

可以通过python项目，命令行中 python main.py启动

也可以下载release版本，muban.docx、config.json，exe缺一不可

命令行中 执行exe即可

# 注意事项

根据实际情况使用，需要调整config.json的参数值

config默认配置为

```bash
{
    "head_pos": 1,
    "head_pos_end": 3,
    "head_new_pos": 4
}
```

head_pos 为原docx中标题行的初始索引，(实际行-1)

head_pos_end 为原docx中标题的总行数

head_new_pos 为模板docx中标题行的初始索引，(实际行-1)



PS：空行也算行，所以原docx被做了无效空行段删除的处理

如效果图中，**输入docx** 中：

1.docx中，标题实际为1行，空行会被自动删除，标题索引为1，配置需调整为

```bash
{
    "head_pos": 1,
    "head_pos_end": 1,
    "head_new_pos": 4
}
```

2.docx中，**《你我》1**的行索引实则为1，配置默认

# 效果图

**模板docx：**

需替换位置xh、{{gh}}、{{zw}}

![](https://github.com/ak005469075/Ylyodcx/blob/master/example/muban.png)

**输入docx** 

1.docx(标题一行，不带图片)

![](https://github.com/ak005469075/Ylyodcx/blob/master/example/input1.png)

2.docx(标题三行，带图片)

![](https://github.com/ak005469075/Ylyodcx/blob/master/example/input2.png)

**输出docx**

![](https://github.com/ak005469075/Ylyodcx/blob/master/example/handle.png)

![](https://github.com/ak005469075/Ylyodcx/blob/master/example/output2.png)


