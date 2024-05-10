"此项目为我工作期间编写的电刺激按摩仪相关工具
# 处方生成工具
此工具为Python3编写，使用pandas和xlrd库，用于读取Excel文件，生成处方。
可以通过编写表格来生成处方对应的json文件和C代码。
## 安装使用环境
* 安装python3

* 安装python依赖包
```bash
pip install pandas xlrd
```

## 使用方法
### 1.编辑在线表格
* 低频
```
https://www.kdocs.cn/l/ca7to70qKJga
```
* 中频
```
https://www.kdocs.cn/l/ceMzKiam0qNy
```
* 打开链接里的在线表格模板
* 另存文件在自己的账号下
* 然后编辑表格内容
* 导出为.xls文件下载

### 2.运行脚本
#### 测试
将表格文件放到脚本同一目录下，运行脚本测试环境时候安装正确。
```bash
python xls2jsonMid.py testRecipeMid.xls W7-2F
```

#### 参数说明:

* 生成json
```bash
python3 xls2jsonLow.py <表格文件名.xls> <机器名>
```
```bash
python3 xls2jsonMid.py <表格文件名.xls> <机器名>
```
* 生成C
```bash
python3 xls2cLow.py <表格文件名.xls>
```
```bash
python3 xls2cMid.py <表格文件名.xls>
```
#### 范例
* 通过testRecipeLow.xls生成低频处方json文件，机器名为NK54
```bash
python xls2jsonLow.py testRecipeLow.xls NK54
```
* 通过testRecipeMid.xls生成中频C代码
```bash
python xls2cMid.py testRecipeMid.xls
```
## 注意
* 表格文件不能被加密，否则无法读取。
* 单个表格可以包含多个处方运行后生成“Mode_机器名_IDxxxx.json”文件，其中xxxx为处方ID。
* 生成的C代码文件为”outputMid.c"或"outputLow.c"，需要手动复制到对应的工程目录下。
* 填写的ParseType需要正确填写，否则会出现升级进去后无法切换到对应的模式的情况。
* 表格具体填写规则请参考表格模板。
