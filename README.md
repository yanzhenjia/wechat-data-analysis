# 公众号文章流量分析工具/wechat-data-analysis

这是一个用于整合公众号文章后台数据的本地工具。它会自动读取 `data/` 目录中的文章明细文件，完成去重、清洗、统计、评分，并同步输出一份 Excel 报表和一份 HTML 可视化报告。

<img width="1920" height="938" alt="image" src="https://github.com/user-attachments/assets/442ebcf3-3de6-4cdb-b941-5d3cb2d2a0e9" />

<img width="780" height="519" alt="image" src="https://github.com/user-attachments/assets/c1b668d5-3b44-43fe-95f3-6693376aa0bd" />


## 功能

- 自动扫描并读取公众号文章数据文件
- 支持 `.xls`、`.xlsx`、`.csv`、`.json`
- 自动识别重复文件并跳过
- 生成透明公式的 Excel 统计表
- 生成适合直接阅读的 HTML 可视化报告
- 提供桌面 GUI 界面，支持一键分析
- 支持打包成单文件 `exe`，可直接双击使用

## 输出内容

程序运行后会生成两份结果文件：

- `Excel`：文章统计和分析报表
- `HTML`：可视化分析页面

Excel 默认包含 5 个工作表：

- `汇总看板`
- `数据明细`
- `文章分析`
- `指标说明`
- `处理日志`

## 评分逻辑

当前报表使用以下口径：

- 阅读表现分 = `70% 阅读人数分位 + 30% 打开率分位`
- 互动表现分 = `35% 分享率分位 + 25% 点赞率分位 + 15% 在看率分位 + 10% 收藏率分位 + 15% 评论率分位`
- 留存表现分 = `65% 完读率分位 + 35% 平均停留时长分位`
- 转化表现分 = `100% 关注率分位`
- 综合指数 = `40% 阅读表现分 + 30% 互动表现分 + 20% 留存表现分 + 10% 转化表现分`

## 默认目录规则

程序默认优先使用软件所在目录下的：

- `data/` 作为输入目录
- `output/` 作为输出目录

如果软件所在目录不可写，则自动回退到桌面下的对应目录。

## 使用方式

### 方式一：直接运行 GUI

```bash
python main.py
```

打开后选择数据目录和输出目录，点击“开始分析”即可。

### 方式二：命令行运行

```bash
python gzh_analyzer.py --input data --output-dir output
```

可选参数：

- `--name`：自定义输出文件名前缀

示例：

```bash
python gzh_analyzer.py --input data --output-dir output --name 公众号文章流量分析
```

## 打包版 exe

项目已经支持打包为单文件 exe。生成后的文件位于：

```bash
dist/gzh_report_tool.exe
```

使用方法：

1. 将 `data/` 文件夹放在 exe 同级目录
2. 双击 `gzh_report_tool.exe`
3. 程序会自动读取输入数据，并把结果写入同级 `output/`

如果软件所在目录没有写权限，程序会自动改用桌面上的 `data/` 和 `output/`。

## 项目结构

- `main.py`：GUI 启动入口
- `gui.py`：桌面界面
- `gzh_analyzer.py`：数据解析、去重、统计、Excel/HTML 导出
- `data/`：输入数据目录
- `output/`：输出结果目录
- `requirements.txt`：依赖列表
- `构建程序.bat`：打包脚本

## 依赖安装

```bash
pip install -r requirements.txt
```

## 常见问题

### 1. 为什么有些文件没有被统计？

通常是因为文件是重复内容，程序会自动跳过重复文件，避免重复计入总表。

### 2. 为什么 Excel 里还有公式列？

为了保证统计过程可追溯，核心比率和分数都保留了公式，方便检查口径。

### 3. HTML 报告能做什么？

HTML 适合快速查看整体表现，包括：

- 总览指标
- 月度趋势
- 渠道结构
- TOP10 文章
- 潜力文章

## 重新打包

如果你修改了代码，可以重新执行打包命令：

```bash
pyinstaller --noconfirm --clean --onefile --windowed --name gzh_report_tool main.py
```

打包完成后，新的 exe 会生成在 `dist/` 目录下。

