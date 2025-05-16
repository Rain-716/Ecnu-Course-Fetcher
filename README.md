## 项目简介

本项目旨在自动化获取华东师范大学（ECNU）教务选课系统中的课程信息及选课人数统计，并将结果导出为 Excel 文件。通过 Selenium 驱动浏览器完成登录、验证码识别及页面抓取，结合 `ddddocr` 进行验证码识别，最终利用 pandas 与 openpyxl 整理并生成带有剩余名额统计的课程概览表。

## 功能特点

* 自动登录教务系统并识别验证码
* 定时抓取选课页面中的课程数据与人数统计
* 数据清洗与合并，计算剩余名额
* 导出格式化的 Excel 文件，自动适配列宽
* 可通过修改配置快速调整账户信息、抓取间隔及输出文件名

## 环境依赖

* Python 3.8 及以上
* Google Chrome 浏览器
* ChromeDriver（版本需与本机 Chrome 浏览器版本对应）
* 主要 Python 库：

  * `selenium`
  * `ddddocr`
  * `pandas`
  * `openpyxl`
  * `Pillow`

## 安装步骤

1. 克隆或下载本项目到本地：

   ```bash
   git clone <项目地址>
   cd <项目目录>
   ```

2. 创建并激活虚拟环境（可选）：

   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/macOS
   venv\\Scripts\\activate # Windows
   ```

3. 安装依赖：

   ```bash
   pip install -r requirements.txt
   ```

4. 下载并配置 ChromeDriver，将其路径添加到系统 `PATH` 或放置于项目根目录。

## 配置说明

在 `main.py`（或运行脚本）中，可通过修改以下常量来配置：

```python
USERNAME = '你的学号或账号'
PASSWORD = '你的密码'
LOGIN_URL = '登录页面 URL'
TARGET_URL = '选课数据接口 URL'
JS_FILE_COUNTS = '人数统计 JS 文件名'
JS_FILE_LESSONS = '课程数据 JS 文件名'
OUTPUT_EXCEL = '导出 Excel 文件名'
INTERVAL_SECONDS = 60  # 抓取循环间隔（秒）
```

## 使用方法

在项目根目录下，运行：

```bash
python main.py
```

程序会进入循环，每隔 `INTERVAL_SECONDS` 秒自动执行一次抓取任务，并将最新的课程概览与选课人数导出至 `OUTPUT_EXCEL`。

## 输出说明

* 输出文件：`lesson_overview_with_counts.xlsx`（可通过 `OUTPUT_EXCEL` 修改）
* 工作表名称：`课程概览`
* 包含字段：课程 ID、课程名称、授课教师、学分、当前选课人数、限选人数、剩余名额等
* 自动调整列宽，方便阅读

## 注意事项

1. 请确保 ChromeDriver 版本与本机 Chrome 浏览器版本匹配，否则可能无法正常启动浏览器。
2. 若验证码识别失败，可适当增大 Selenium 显示等待时间或检查网络环境。
3. 本脚本仅供学习和参考，请勿用于违规用途，请在学校政策允许范围内使用。

## 扩展与优化

* 可将抓取结果存入数据库，进行历史趋势分析
* 增加邮件或微信通知功能，实时推送剩余名额警报
* 支持多账号并发抓取

---