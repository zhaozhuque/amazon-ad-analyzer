# 发布指南：打包EXE + 发布GitHub

---

## 方案一：打包成 EXE（推荐）

打包后生成一个独立的 .exe 文件，别人双击就能用，不需要装 Python。

### 步骤

**1. 打开命令提示符（CMD）**
- 按 `Win + R`，输入 `cmd`，回车

**2. 进入项目文件夹**
```
cd 你的文件夹路径
```
例如文件放在桌面的 amazon-ad-analyzer 文件夹：
```
cd C:\Users\你的用户名\Desktop\amazon-ad-analyzer
```

**3. 安装打包工具**
```
pip install pyinstaller
```

**4. 一键打包**
```
python build_exe.py
```

等待约1-2分钟，完成后在 `dist` 文件夹里找到 `AmazonAdAnalyzer.exe`。

**5. 测试**
- 双击 `AmazonAdAnalyzer.exe`
- 选择你的广告报告试一下
- 确认能正常生成诊断Excel

**6. 分发**
- 把 `AmazonAdAnalyzer.exe` 发给别人就行，单个文件，约30-50MB
- 对方不需要装任何东西，双击直接用

### 常见问题

Q: 打包报错 "No module named xxx"
A: 先安装缺少的模块：`pip install xxx`，再重新打包

Q: 杀毒软件误报
A: PyInstaller 打包的exe经常被误报，添加信任即可，这是正常现象

Q: exe文件太大
A: 正常的，因为它把 Python + pandas + numpy 都打包进去了

---

## 方案二：发布到 GitHub

### 前置准备

1. 注册 GitHub 账号：https://github.com/signup
2. 安装 Git：https://git-scm.com/downloads
   - 安装时一路默认即可

### 步骤

**1. 在 GitHub 上创建仓库**
- 打开 https://github.com/new
- Repository name 填：`amazon-ad-analyzer`
- Description 填：`亚马逊广告诊断工具 - 一键分析广告报告`
- 选 Public（公开）
- 不要勾选 "Add a README file"（我们已经有了）
- 点 Create repository

**2. 打开 CMD，进入项目文件夹**
```
cd C:\Users\你的用户名\Desktop\amazon-ad-analyzer
```

**3. 初始化并上传**
依次执行以下命令（复制粘贴即可）：
```
git init
git add .
git commit -m "first commit: amazon ad analyzer tool"
git branch -M main
git remote add origin https://github.com/你的用户名/amazon-ad-analyzer.git
git push -u origin main
```

第一次 push 会弹出 GitHub 登录窗口，登录即可。

**4. 上传打包好的 EXE（可选但推荐）**
- 打开你的仓库页面：`https://github.com/你的用户名/amazon-ad-analyzer`
- 点右侧 "Releases" -> "Create a new release"
- Tag version 填：`v1.0`
- Release title 填：`v1.0 首个发布版本`
- 描述里写：
  ```
  亚马逊广告诊断工具 v1.0

  功能：
  - 自动识别烧钱词、优质词、潜力词
  - 生成7个维度的Excel诊断报告
  - 支持中英文亚马逊后台报告
  - 完全离线运行

  使用：下载 AmazonAdAnalyzer.exe，双击运行即可。
  ```
- 把 `AmazonAdAnalyzer.exe` 拖到 "Attach binaries" 区域上传
- 点 "Publish release"

### 完成！

你的工具现在有了一个专业的 GitHub 页面，别人可以：
- 直接下载 EXE 使用
- 查看源代码
- 提 Issue 反馈问题
- Fork 修改

---

## 项目文件结构

```
amazon-ad-analyzer/
├── amazon_ad_tool.py      <- 主程序
├── build_exe.py           <- 打包脚本
├── start_tool.bat         <- Windows快速启动
├── requirements.txt       <- Python依赖
├── README.md              <- 项目说明（GitHub首页展示）
├── LICENSE                <- 开源协议
└── .gitignore             <- Git忽略规则
```
