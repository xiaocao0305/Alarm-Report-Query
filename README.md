# Alarm Report Query

读取南盾报警主机 Alarm7 的历史报告库（Rpt），把报警记录按「用户账号 + 防区号」关联到用户数据库（Master）的防区配置，取出**防区位置**，以桌面窗口展示并支持导出 Excel。

**版本更新记录见 [UPDATE.md](UPDATE.md)。**

## 功能

- 按**日期范围**查询报警记录（SQL 层筛选，支持跨月自动定位月份库）
- 快捷按钮：当日报告 / 本周报告 / 本月报告
- 账号查询默认**完整精确匹配**，可选模糊查询
- 数据源选择：多套 Alarm7 由用户选择，记住上次选择，失效提示
- 查询在后台线程执行，窗口可操作，带进度条，支持取消
- 明确区分「没有记录」与「读取失败/缺月份」；失败时禁用导出
- 超 3000 条时界面与 Excel 注明「仅显示部分记录」
- 防区位置列名为「防区位置（当前配置）」，未匹配显示「未查到位置」，
  每次查询重新读取位置配置并记录读取时间
- 导出带样式的 Excel（表头、行颜色、冻结首行、自动筛选），
  所有字段按文本写入（防公式注入、保留前导 0），附「查询说明」页
- 修改 config.ini 后点「重新加载配置」立即生效，无需重启

## 技术栈

- Python 3.8（32 位，兼容 Win7+）
- wxPython（桌面界面）
- pyodbc（读取 Access .mdb，使用 Windows 自带驱动，免安装）
- openpyxl（导出 .xlsx）

## 目录结构

```
output/Alarm7ZoneTool/
  alarm7_zone_gui.py        # 桌面 GUI 版（当前主版本）
  config.ini.example        # 配置模板（复制为 config.ini 使用）
output/
  Alarm Zone Location Query Tool/   # 发布包（exe + 使用说明）
```

## 配置

程序不在代码里保存任何密码。首次运行会在程序同目录生成 `config.ini`（已被 .gitignore 排除），请在该文件中填写：

```ini
[paths]
alarm_dir =          # Alarm7 安装目录，留空自动探测
master_db =          # 用户数据库，留空用 alarm_dir\Master.mdb
master_pwd =         # 用户数据库密码（必填）
rpt_dir =            # 报告目录，留空用 alarm_dir\Rpt
rpt_pwd =            # 报告库密码（无密码留空）
```

密码只保存在本机 config.ini 中，不会进入代码仓库。

## 运行

```bash
pip install wxPython pyodbc openpyxl
python output/Alarm7ZoneTool/alarm7_zone_gui.py
```

## 打包为 exe

```bash
python -m PyInstaller --onefile --windowed --name "Alarm7防区位置查询工具" \
  --exclude-module numpy --exclude-module PIL \
  output/Alarm7ZoneTool/alarm7_zone_gui.py
```

## 版权

版权所有 2023-2030 IntelliVision Cloud。保留所有权利。
