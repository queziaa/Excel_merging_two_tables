# 通过一列进行两表（Excel）合并

### 可以直接运行.py文件处理表格
### 需要安装 xlwt xlrd openpyxl
### 我的开发平台 python3 没有在测试过 python2
### 使用时直接将表格的完整路径作为运行参数

### 或使用Pyinstaller打包后使用 推荐打包后使用
### 打包后直接拖动表格到打包好EXE即可
### 不知道为什么有性能提升
### 打包命令添加-F参数打包为一个EXE文件
### Pyinstaller -F main.py

### 这里提供打包好的EXE
### 仅提供32位版本，在Win7X86下打包使用Pyinstaller，在Win7X86、Win7X64，Win10X64测试没有问题，且没有依赖问题[链接](https://github.com/queziaa/Excel_merging_two_tables/releases)

### 输入文件要求
### 1：输入文件需要拥有两个工作簿，名字任意
### 2：将以第一列的内容进行比对将两表合并
### 3：同一工作簿第一列不能出现相同内容
### 4：第一行视为表头直接复制至结果
### 5: 产生遗弃内容时，遗弃内容会在原工作簿同名工作簿下保存
