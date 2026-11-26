# FineBI 分享链接 → Excel 定时采集

> **只想用、不想看原理？先看 [操作手册.md](操作手册.md)**：三步入上手、配置怎么填、出问题怎么查。
> 本文是技术说明：接口链路、载荷结构、版本兼容、踩过的坑——改动代码前读。

针对 `http://host:port/webroot/decision/link/xxxx` 这类 **FineBI 公共分享链接**，把仪表板里表格组件的数据抓下来落成 Excel。

- **一个页面上多个组件**：自动发现、自动重建载荷，**每个组件落成独立文件**，文件名 = `组件名_yyyyMMdd.csv`（默认 csv，可切 xlsx）；
- **在 Excel 里配置任务**，可以常驻按 cron 跑，也可以只做「跑一轮就退出」给外部调度器（ElasticJob / 计划任务）调用；
- **Windows 可直接部署**，详见 [DEPLOY-WINDOWS.md](DEPLOY-WINDOWS.md)。

运行环境：**Python 3.11**（零第三方依赖，`urllib` + `zipfile` 全标准库），已在 **3.11.9** 与 3.13.12 上实测；  
Java 8 只作为可选调度壳。

实测环境：FineBI 6.0.27 / FineReport 11.5.12；样例页面 2 个表格组件，各 **110000 行**（共 22 万行 / 12.8 MB）。  
同样一次全量：Python 3.13 约 **165 秒**，Python 3.11 约 **210 秒**（3.13 的解释器优化更快，功能一致）。  
6.1.x 的兼容性说明与升级验证方式见 [第八节](#八版本兼容finebi-60x--61x)。

---

## 一、数据交互链路（逆向结论）

分享链接没有走官方 OpenAPI，是前端渲染器自己发的内部请求，共四步：

| 步 | 请求                                                                                                             | 作用             | 关键产物                                                                           |
| - | -------------------------------------------------------------------------------------------------------------- | -------------- | ------------------------------------------------------------------------------ |
| 1 | `GET /webroot/decision/link/{code}`                                                                            | 打开分享页          | HTML 内联 `BI.sessionId`、分享令牌 `link`（JWT）、`reportId`、**`reportWidgets`（全部组件清单）** |
| 2 | `POST .../v5/design/report/share/report/{reportId}/widgets/detail`，body `{"ids":[realWidgetId,...]}`           | **批量**拿组件数据模型  | `data[realWidgetId]` = 该组件的模型                                                  |
| 3 | 由步 2 的模型重建取数载荷                                                                                                 | 每个组件一份 payload | —                                                                              |
| 4 | `POST .../v5/design/report/share/widget/data?engineType=2&reportId=..&entryType=2&link=<JWT>`，body 里 `page` 递增 | 逐组件取数翻页        | `data.header` / `data.items` / `data.hasNext` / `data.row`                     |

**请求头**必须带 `sessionId`（步骤 1 拿到的），**不需要 Cookie、不需要登录**。

步骤 2 的 `ids` 传的是 `realWidgetId`（不是 `widgetId`），一次调用可传多个，服务端按 id 返回字典；传了不存在的 id 会**静默忽略**，不报错。

### 多组件怎么走通

一个页面 N 个组件，就是 **1 次 detail + N 次取数**：

1. 从 HTML 的 `reportWidgets` 解析出 `[{widgetId, realWidgetId, title}]`；
2. 一次 `widgets/detail`，`ids` 传全部 `realWidgetId`；
3. 对每个组件的模型做统一重建（见下），得到 N 份 payload；
4. 逐个翻页取数，**一个组件一个 sheet**。

### 载荷重建的三个关键点（都是实测出来的）

| 字段                              | 怎么处理                                                                                                                                              | 不处理的后果                                      |
| ------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------- |
| `settings.tableStyle.totalRows` | **置 0（不限）**                                                                                                                                       | detail 里给的是 `1000`，静默只拿 1000 行。**这是最容易踩的坑** |
| `wId`                           | 用**外层** `widgetId`                                                                                                                                | detail 返回的 `wId` 是 `realWidgetId`，两者不同      |
| `tableName`                     | 字符串 → 单元素数组                                                                                                                                       | 类型不对                                        |
| `queryInfo`                     | 补 `reportId / reportWidgetId / widgetId / sessionId / title`                                                                                      | 上下文缺失                                       |
| 其余                              | 补 `explainConf / fields / filter / filterValues / drillOrder / resultFilter / uploadedImages / injection / legendFilter`，`realData=true`，`page=1` | —                                           |

> 上一版结论「必须用抓包模板、detail 重建只能取 1000 行」**是错的**。真正的原因是 `totalRows`，  
> 补上就完全等价：同一组件用重建载荷能拿满 110000 行，与抓包模板结果一致。  
> 顺带发现：抓包模板里 `totalRows` 写的是 `20000`，所以**之前用模板只取到了 2 万行，漏了 9 万行**。

### 趟过的坑

| 现象                                                | 原因                                                                                                                                                                                     |
| ------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `LoginInfoNotAvailableException: Token is empty!` | 取数路由少了 `/report/share`。分享页 moduleRouter 是 `design/report/share`，普通设计页才是 `design`                                                                                                       |
| 只拿到 1000 行 / `row` 从 110000 掉到 1000               | `settings.tableStyle.totalRows` 没置 0                                                                                                                                                   |
| 组件名取错                                             | `reportWidgets.title` 在未命名时是占位符「组件名称」或 `${..."__widget_name_id__"...}`，**不能**当组件名；真实名字在 `widgets/detail` 的 `name` 字段（如「全部明细」「客户号+名」）。取值顺序：`detail.name` → `reportWidgets.title` → 序号兜底 |
| 手机号导出成 `13,871,014,583`                           | BI 侧对该列配了千分位显示，`text` 带逗号、`value` 是 `1.3871014583E10` 科学计数。已在内置 `cell_text()` 里还原成 `13871014583`（只对「纯数字+千分位」生效，不会误伤住址里的逗号）                                                             |
| `page` 传成对象                                       | `page` 是整数                                                                                                                                                                             |
| 图表组件取不到行列                                         | 图表返回的是 `series` 结构，不是 `header/items`。本工具只导明细表组件，图表会被跳过并注明原因                                                                                                                            |

---

## 二、配置 Excel 的格式

`config.xlsx`，一行一个任务，第一行是表头。列按名字识别，**顺序随意**，中英文别名都认：

| 列名        | 必填    | 说明                                                                                        |
| --------- | ----- | ----------------------------------------------------------------------------------------- |
| 任务名       | 否     | 唯一标识，默认「任务N」；重名会自动加行号                                                                     |
| 启用        | 否     | `是/否`、`true/false`、`1/0`；**留空视为启用**                                                       |
| 分享链接      | **是** | 完整的 `/webroot/decision/link/xxxx` 地址                                                      |
| 组件        | 否     | **留空 = 页面上全部表格组件**；否则填组件名或 ID，多个用逗号分隔。名字支持模糊匹配                                            |
| 输出方式      | 否     | **留空或 `多文件` = 每个组件一个文件**（默认）；`单文件` = 合并成一个文件（csv 多组件时自动退回多文件）                              |
| 输出格式      | 否     | **`csv`（默认）** / `xlsx`。csv 无百万行上限（xlsx 单表上限 1,048,576 行）；xlsx 才有合并单元格/多级表头/长数字保精度           |
| payload文件 | 否     | **一般留空**（留空才走自动重建、支持多组件）。填了则退回「固定模板 + 单组件」旧模式                                             |
| 输出目录      | 否     | **可写绝对路径**（`D:\finebi\out`、`\\server\share\out`）；相对路径按**配置文件所在目录**解析，默认 `out`。详见下方「路径怎么写」 |
| 文件名       | 否     | 支持占位符；多文件模式留空默认 `{name}_{widget}_{yyyyMMdd}.csv` → `客户明细_全部明细_20261125.csv`；**扩展名由输出格式决定**             |
| cron表达式   | **是** | 见第三节                                                                                      |
| 备注        | 否     | 随便写                                                                                       |
| 表样式       | 否     | `自动`（默认）/ `明细表` / `分组表` / `交叉表`；自动识别不准时用它纠正，见第九节 |
| 分组列       | 否     | `合并`（默认）/ `填充` / `原样`；只影响分组表、交叉表的维度列呈现，见 9.5 |

文件名占位符：`{widget} {yyyy} {MM} {dd} {HH} {mm} {ss} {yyyyMMdd} {HHmmss} {yyyyMMdd_HHmmss} {date} {name}`

- `{widget}` = **组件真实名称**（**多文件模式必须包含**；漏写会自动补上，否则各组件会写到同一文件互相覆盖）；
  - 名字取自 `widgets/detail` 返回的 `name` 字段，也就是你在 BI 里看到的组件名；
  - 若该组件在 BI 里从没命名过（`reportWidgets.title` 是默认占位符 `${...__widget_name_id__...}`），  
    才退化成序号 `组件1`、`组件2`；
- 同一天重复跑会**覆盖**同名文件 —— 这是日报该有的幂等语义。
- **csv 说明**：分组列「合并」在 csv 下自动按「填充」处理（csv 无合并单元格）；身份证/客户号等长数字列
  用 Excel 双击打开会丢精度，请用「数据 → 自文本导入」查看，或该任务改 `xlsx` 输出。
  全局默认格式可用环境变量 `FINEBI_OUTPUT_FORMAT` 调整（`csv`/`xlsx`）。

### 路径怎么写（`输出目录` 列）

**可以写绝对路径**，推荐这么写——不受启动目录影响：

| 写法                   | 解析结果                 | 说明                            |
| -------------------- | -------------------- | ----------------------------- |
| `D:\finebi-data\out` | `D:\finebi-data\out` | Windows 绝对路径（**推荐**）          |
| `D:/finebi-data/out` | `D:\finebi-data\out` | 正斜杠也认，会自动归一化                  |
| `\\server\share\out` | `\\server\share\out` | UNC 网络路径                      |
| `/data/finebi/out`   | `/data/finebi/out`   | Linux / macOS 绝对路径            |
| `out`                | `<配置文件所在目录>\out`     | 相对路径                          |
| `..\data`            | `<配置文件所在目录>\..\data` | 支持 `..`                       |
| `~/Desktop/out`      | 家目录下                 | `~` 会展开                       |
| `%USERPROFILE%\out`  | 环境变量值                | `%VAR%` / `$VAR` 都会展开         |
| `"D:\out"`（带引号）      | `D:\out`             | Excel 里粘贴常带引号，会自动剥掉           |
| `D:out`（不带斜杠）        | **直接报错**             | 这是「D盘当前目录」，语义有歧义，程序拒绝执行并提示怎么写 |

几个实际要点：

- **目录不存在会自动创建**（含多级），不用先手工建文件夹；
- 相对路径的基准是**配置文件所在目录**，不是进程当前目录——所以你从任何地方启动、计划任务的「起始于」填错，落盘位置都不会变；
- 写 `D:\...` 时有权限要求：别往 `C:\Program Files` 这类需要管理员权限的目录写，  
  计划任务默认以普通用户跑，会报 `PermissionError`；建议用 `D:\` 或 `%USERPROFILE%\` 下自己的目录；
- 输出到网络盘（UNC）时先把网盘映射好并确认当前账户有写权限，否则会卡在写文件那一步。

> 改了页面上的组件（增删、改名）**不需要改配置**：留空 `组件` 列就自动跟着页面走。  
> 想确认组件叫什么，跑 `python3 collector.py --list-widgets` 或 `python3 finebi_export.py --list-widgets`。

---

## 三、cron 表达式

调度粒度是**分钟**。

| 写法                                                | 含义                             |
| ------------------------------------------------- | ------------------------------ |
| `30 8 * * *`                                      | 每天 8:30                        |
| `0 */2 * * *`                                     | 每 2 小时整点                       |
| `0 9 * * 1-5`                                     | 工作日 9:00                       |
| `0 15 9 ? * MON-FRI`                              | Quartz 写法也认（6 段，秒位必须是 0 或 `*`） |
| `0 0 1 JAN,JUL *`                                 | 1 月和 7 月 1 日零点（月份/星期支持英文缩写）    |
| `@hourly` `@daily` `@weekly` `@monthly` `@yearly` | 宏                              |

周：`0` 或 `7` = 周日，`1` = 周一。日与周同时被限定时取「或」（标准 cron 语义）。  
秒位写非 0 值会直接报错（例 `30 0 8 * * *`），不会静默降级成分钟级。

---

## 四、跑起来

> 下列命令以 macOS/Linux 的 `python3` 写法为例。**Windows 上把 `python3` 换成 `python`**  
> （多版本共存时用 `py -3.11`，或直接写 `C:\...\Python311\python.exe` 绝对路径）。  
> 代码零第三方依赖，Python **3.11** 即可，无需 `pip install`。

### 常驻调度（推荐：配置每小时自动重载）

```bash
cd finebi-export
python3 collector.py --config config.xlsx
```

行为：

- 每分钟对齐到整分后 2 秒判定一次 cron；
- **配置每小时强制重载一次**；`config.xlsx` 被改动（mtime 变化）时**立刻重载**，改完不用重启；
- 用 `任务名@年月日时分` 做幂等标记，存 `.collector_state.json`，**进程重启不会重复跑同一分钟**；
- 任务串行执行（FineBI 侧并发取数容易被限流）；
- 单个组件取数失败不影响其他组件；单个任务失败不影响进程存活；
- 日志：控制台 + `logs/collector.log`（按天切分，保留 30 天）。

### 交给外部调度器（推荐给你：手动 / ElasticJob / 计划任务）

不需要常驻进程，跑一轮就退出。**stdout 只输出 JSON，日志全部走 stderr**，方便调度平台解析：

```bash
python3 collector.py --once --json                 # 跑全部启用任务
python3 collector.py --once --json --job 客户明细   # 只跑指定任务（名字或片段）
```

退出码语义（调度器直接看这个）：

| 退出码 | 含义                     |
| --- | ---------------------- |
| `0` | 全部成功                   |
| `1` | 有任务失败                  |
| `2` | 配置或参数错误（配置读不到、任务名没匹配上） |

stdout 的 JSON 形如：

```json
{
 "ok": true, "total": 1, "failed": 0, "finishedAt": "2026-11-25 08:48:59",
 "results": [{
   "task": "客户明细", "ok": true, "rows": 220000, "seconds": 164.7,
   "files": [".../out/客户明细_组件1_20261125.csv", ".../out/客户明细_组件2_20261125.csv"]
 }]
}
```

Windows 上直接双击 `run_once.bat`，或 `run_once.bat --job 客户明细`。  
Java 8 侧接 ElasticJob 见 [DEPLOY-WINDOWS.md](DEPLOY-WINDOWS.md) 第三节。

### 其他模式

```bash
python check_env.py --net                    # 部署自检：环境/配置/目录/网络逐项检查
python3 collector.py --list                  # 只看解析结果 + 下次执行时间，不跑
python3 collector.py --list-widgets          # 看某个页面有哪些组件（改配置前先跑这个）
python3 collector.py --once --dry-run        # 只打印目标路径，不取数
python3 collector.py --once --max-pages 3    # 冒烟测试：每个组件只抓 3 页
python3 collector.py --verbose               # 逐页日志
```

### 挂成开机自启（launchd，macOS）

`~/Library/LaunchAgents/cn.bank.finebi.collector.plist`：

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>cn.bank.finebi.collector</string>
  <key>ProgramArguments</key><array>
    <string>/usr/bin/python3</string>
    <string>/绝对路径/finebi-export/collector.py</string>
    <string>--config</string>
    <string>/绝对路径/finebi-export/config.xlsx</string>
  </array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>StandardErrorPath</key><string>/绝对路径/finebi-export/logs/launchd.err</string>
</dict></plist>
```

```bash
launchctl load -w ~/Library/LaunchAgents/cn.bank.finebi.collector.plist
```

> 也可以不让进程常驻，用系统 cron 直接调 `--once`。但那样配置里的 cron 列就不起作用了，二选一。

---

## 五、单任务直跑（不需要配置时）

```bash
python3 finebi_export.py                          # 导出页面全部表格组件 -> 多 sheet
python3 finebi_export.py --list-widgets           # 只看组件清单
python3 finebi_export.py --widget 客户明细         # 只导指定组件
python3 finebi_export.py --split                  # 每组件一个文件
python3 finebi_export.py --max-pages 3            # 冒烟测试
python3 finebi_export.py --template payload_template.json   # 退回固定模板单组件模式
```

零第三方依赖：`urllib` 取数、`zipfile` 直接拼 xlsx（没装 openpyxl 也能跑）。

Excel 上的处理：

- **千分位还原**：BI 里配了千分位格式的列（手机号等）会导出成 `13,871,014,583`，已在取值时还原为 `13871014583`；
- **强制文本列**：列名含 `身份证/证件/客户号/账号/卡号/手机/电话/号码` 一律写文本 —— 否则 18 位身份证会变成 `1.10101E+17`，`0000000001` 掉成 `1`；
- **安全数值**：纯数字且无前导零、长度 < 15 的才写数值，其余保文本；
- 表头加粗+浅底、首行冻结、列宽自适应；
- 多 sheet 时组件名做 sheet 名，自动去 `[]:*?/\` 等非法字符、截断 31 字、重名加 `_2` 后缀。

---

## 六、Java 8 调度壳

**采集逻辑不要用 Java 重写**。`java8/` 下是一个 Java 8 兼容的调用壳，只做三件事：  
到点触发 → 用 `ProcessBuilder` 调 Python → 判断退出码。零第三方依赖，字节码目标已确认为 Java 8（52）。

```bash
cd java8 && mvn clean package
java -jar target/finebi-export-runner.jar \
  --python "C:\Python313\python.exe" \
  --work-dir "D:\apps\finebi-export" \
  --config "D:\apps\finebi-export\config.xlsx" \
  --job 客户明细 --max-pages 2
```

接入 ElasticJob 只需把 `FineBiExportJob.java.example` 改名启用，完整说明见  
[DEPLOY-WINDOWS.md](DEPLOY-WINDOWS.md) 第三节。

> `java/` 目录是上一版 **Java 17 + POI** 的独立实现（单任务、单组件），Java 8 项目用不到，  
> 你可以直接删掉；它的取数结论已全部并入 Python 版。



---

## 七、换一个仪表板 / 加一个组件怎么改

**都不用改代码**：

1. 把新分享链接填进 `config.xlsx` 的「分享链接」列；
2. `组件` 列留空（全导）或填想要的组件名；
3. 需要确认组件名就先跑 `--list-widgets`。

页面结构变了（组件增删改名）也一样，留空就自动跟随。

---

## 八、版本兼容（FineBI 6.0.x / 6.1.x）

**先说结论**：已在 **6.1.12** 上实测通过（明细表 / 分组表 / 交叉表各一个组件，共 24630 行，
一次跑通并写出 3 个文件）。6.1 仍属同一套 `v5` 报表渲染接口（命名空间 `report/share`），
但**加了三个新机制**，不处理就会失败或静默少数据，见 8.0。

### 8.0 6.1 相对 6.0 的四处变化（实测，必读）

| # | 变化 | 不处理的后果 | 现在的做法 |
|---|---|---|---|
| 1 | 取数 URL 必须带 `cid` | 报 `FineEmptyCIDException: CID is missing` | 自动生成并缓存，见 8.1 |
| 2 | 有**浏览器并发限制**（个人版许可 = 2） | 报 `BI concurrency has reached the limit of license` | 同机复用同一 cid，见 8.2 |
| 3 | 分组表/交叉表首次取数 `page=0`（不是 1） | 报 `cache Page_State expired`（看着像缓存坏了） | 按表型选起始页，见 8.3 |
| 4 | 数据响应按表型分化成三种结构 | 交叉表 0 行 0 列、分组表维度全空 | 三种结构分别解析，见 8.4 |

### 8.1 cid：6.1 新增的必填参数

取数请求的 URL 上多了三个参数：`showSectionError=true&cid=...&taskId=...`，其中 `cid` 是硬要求。

```
cid = <32位hex>#<13位毫秒时间戳>#<校验码>
校验码 = MD5(前 32 位 + "#" + 时间戳) 的末 8 位
```

这条规则的出处是 6.1.12 的 `com.finebi.dashboard.api.structure.session.BIBrowserConcurrencyManager#verifyCid`
（反编译得到），并用浏览器真实抓包验证过。脚本会自己算，不需要你管。

算出来的 cid 会**按站点缓存在 `.cid_cache`**，下次直接复用——这一点很重要，见下。

### 8.2 并发限制：一个 cid 占一个名额

6.1 引入浏览器并发控制，个人版许可上限是 **2**。占用是**按 cid 计**的：

- 同一台机器、同一个 cid 反复跑（哪怕每次都新建会话）→ 只占 **1 个**名额，实测连跑 3 次没问题；
- 如果同时用浏览器打开同一页面（浏览器有自己的 cid）→ 再占 1 个，两个加起来正好顶满，
  第三个请求就报并发超限。

所以：**脚本会长期占用 1 个并发名额**。生产环境上如果许可紧张，要么给这个任务留出名额，
要么让 BI 管理员调大并发数。别删 `.cid_cache`——删了会生成新 cid，历史上那个名额要等会话过期才释放。

### 8.3 page 的语义变了（最容易踩）

浏览器的真实请求里，交叉表和分组表首次取数用的是 `page=0`，只有明细表用 `page=1`。
我们按 6.0 的老逻辑从 `page=1` 开始，服务端当成「翻到下一页」，于是去找上一页的游标（`Page_State` 缓存），
找不到就抛 `cache Page_State expired`——报错文案很有误导性，实际跟缓存过期无关。

而且 6.1 的分组表/交叉表**不再按页码翻页**，改成「一次要多大给多大」：
由 `settings.tableStyle.rowCounts` / `colCounts` 控制。所以现在的做法是：

- 分组表/交叉表：`page=0` + 把 `rowCounts`/`colCounts` 放大到 `FINEBI_BIG_ROWS`（默认 200000），
  **一次取全**，实测 24426 行一次拿完，耗时 1~2 秒；
- 明细表：仍然逐页翻页（`page=1,2,3…`），但**每页行数被放大**——
  这个行数在载荷里，字段是 `settings.tableStyle.detailRowCounts`，服务端默认给 100。
  100 行/页时 11 万行要发 1100 次请求、跑 286 秒；放大到 5000 行/页只需 22 次、约 23 秒，
  同一份数据快 11 倍。默认值见 `FINEBI_DETAIL_PAGE_ROWS`（5000，设 0 则沿用服务端默认）。

  别再往上加太多：实测 `detailRowCounts=200000` 虽然一次就取全，但单次响应 260 MB，
  引擎算 19 秒，内存和带宽都不划算。几千行一页是性价比最好的区间。

数据没取全会给出明确提示，不用自己猜：

```
· 服务端标记还有更多数据（hasNext），本次没取全；调大环境变量 FINEBI_BIG_ROWS（当前 200000）后重试
```

### 8.4 三种响应结构（6.1 才有的分化）

| 表型 | 组件模型 `type` | 响应结构 | 落 Excel 的样子 |
|---|---|---|---|
| 明细表 | `4` | `header[]` + `items[[cell…]]`（平铺） | 单行表头 + 数据行 |
| 分组表 | `1` | `header[]` + `items[维度树]`，节点是 `{text, children[], values[]}` | 维度列纵向合并 + 指标列 |
| 交叉表 | `2` | `top{header,items}` + `left{header,items}`，是**视窗**不是全表 | 多级表头 + 行维度合并 |

要点：

- **分组表的根节点是整表的合计行**，`text` 为 `None`。它不能算作一个维度层，否则会多出一列空列，
  且所有行的首列都变成「合计」。接口不给这行标签，代码里补成「合计」并在日志里说明。
- **交叉表的值不在 `header/items` 里**，而在 `left` 的叶子节点上：叶子 `values` 按**列叶子顺序**排开，
  长度 = 列叶子数（例子里 6 个标签 × 2 个指标 = 12 个数）。`left.header` 的长度 = 行维度数 + 列叶子数，
  用它反推行维度数最稳。
- 交叉表是**视窗**结构，响应里带 `hasXNext` / `hasYNext`。如果为 true，说明只拿到可视区域，
  调大 `colCounts`/`rowCounts` 再取。
- 组件模型里的 `type` 字段是**最可靠的类型来源**，现在优先用它；取不到才退回 6.0 时代那套信号投票。

### 8.5 筛选条件（过滤组件传参）

**用法**：把参数直接拼在链接后面，参数名写**过滤组件的标题**。

```
http://host:37799/webroot/decision/link/nQcq?性别=女
http://host:37799/webroot/decision/link/nQcq?性别=女&姓名=汪瑾瑜
```

也可以在 `config.xlsx` 的「分享链接」列直接写带参数的链接，两种写法等价。多个值用重复参数：
`?性别=女&性别=男`（等值于「女 或 男」）。

**为什么不是「服务端认 URL 参数」**：实测把 sessionId/令牌/时间戳归一化后，
`link/nQcq` 和 `link/nQcq?性别=女` 返回的页面**完全一致**——服务端根本不看这个参数。
真正的处理在页面 JS 里：它读 `location.search`，把值赋给同名过滤组件，再重新发取数请求。
脚本走 HTTP 直连跳过了这层 JS，所以必须自己把参数翻成载荷里的筛选条件。

**载荷里长什么样**（这是 6.1 前端点筛选时发出的真实结构，两半缺一不可）：

```json
{
  "queryInfo": {
    "widgetIdValueMap": {
      "39424eb35b9a132c": {            // 过滤组件的完整 widgetBean，原样带过来
        "…": "…",
        "value": {"type": 1, "value": ["女"], "assist": []}   // 只改这一处
      }
    }
  },
  "filter": {
    "filterType": 34,
    "filterValue": [{"filterType": 56, "filterValue": {"widgetId": "39424eb35b9a132c"}}]
  }
}
```

只改 `value` 不写 `filter`，或只写 `filter` 不带 map，服务端都**静默返回全量**——不报错，
数据看起来完全正常，只是没筛。这是最难发现的一类问题，所以：

- 日志里会打出 `生效筛选：性别=女`，没有这行就说明没筛；
- 参数名写错时会告警 `参数 xxx 没有同名过滤组件，未生效（可用名：性别、姓名）`；
- 自检第 3 步会列出页面上所有过滤组件及其绑定字段。

**过滤组件从哪来**：`reportWidgets` 里带 `widgetBean` 的条目就是过滤组件，它**没有 `realWidgetId`**，
所以不能调 `widgets/detail`，也不需要——整份定义就在 `widgetBean` 里。这也意味着
6.1 的 `reportWidgets` 是**多层转义**的，解析时要一层层反转义再试 `json.loads`，
写死层数或直接用正则扫（要求条目带 `realWidgetId`）都会把过滤组件整批漏掉。

**已知限制**：

| 项 | 说明 |
|---|---|
| 值类型 | 只支持**文本类**过滤组件（值 = 字符串数组）。日期类、数值区间类的值是对象结构，按文本写会被服务端拒绝，见下条 |
| 日期/区间类组件 | **已明确不支持，会直接报错退出**，不会静默导全量。原因：`DateControlBean.value` 是 `DateValueWrapperBean`，里面套的是 `DateStaticValueBean{year,month,day,quarter,period,hour,minute,second}`；把 `["2026-11-23"]` 塞进去，Jackson 反序列化报 `Cannot deserialize DateStaticValueBean from Array value`，表现为**仪表板内所有组件逐个 400 失败**，很像接口坏了。要适配这类组件，先跑 `probe_filters.py` 把结构导出来 |
| 过滤组件之外 | 「联动」「钻取」「超链传参」属于另一些结构，不在本次范围内 |
| 参数名 | 只认组件标题或它绑定的字段名，不支持按组件 id 传 |

排查工具：
- `probe_filters.py` —— 打印每个过滤组件的值结构，判断能否按文本传、以及日期组件长什么样；
- `capture_payload.js`（需要 Node + Chrome）—— 把页面真实发出的取数请求体抓下来对比，
  版本升级后如果筛选突然失效，用它抓一份和 `build_payload()` 的输出 diff，比读代码快。

```bash
node capture_payload.js "http://localhost:37799/webroot/decision/link/nQcq?性别=女" cap.json 20000
```

### 哪些地方跟版本有关


| 环节 | 依赖的东西 | 6.0→6.1 风险 | 已做的兜底 |
|---|---|---|---|
| 分享页解析 | HTML 内联的 `BI.sessionId` / 令牌 `link` / `reportId` | 中 | 每一项都是 2~3 条候选规则，换写法不至于直接失败 |
| 组件清单 | `reportWidgets` 的 JSON | **最高**（最容易被改） | 括号配对截取 + `json.loads`；不规整时正则兜底，窗口从 600 放到 4000 |
| 组件模型 | `POST v5/.../report/share/report/{id}/widgets/detail` | 低 | 路径变化会明确报错，不会静默 |
| 明细表判定 | `settings.tableStyle` 是否存在 | 中 | 缺失时按「有字段、无图表特征」兜底，不会整页被误判成图表 |
| 行数上限 | `tableStyle.totalRows = 0`（0 = 不限） | 中 | 可用环境变量覆盖：`set FINEBI_TOTAL_ROWS=-1` |
| 翻页 | body 里的 `page` / `hasNext` / `items` | 低 | 自检会实取 2 页并比对首行，翻页失效当场暴露 |
| 明细表每页行数 | `tableStyle.detailRowCounts`（默认 100） | 低 | 主动放大到 5000；字段不被接受时只是每页少几行，`hasNext` 仍能翻到底 |
| 筛选条件 | 过滤组件的 `widgetBean` + `filter.filterType=56` 引用 | 中 | 只认这两种结构；不匹配时告警而不是静默导全量（见 8.5） |

最容易被忽略的是 **`reportWidgets` 的解析**：老实现用固定 600 字符窗口的正则，
6.1 只要在组件对象里多加几个字段，就会**静默少发现组件**（不报错，只是少导一个文件）。
现在改成按括号配对取出完整 JSON 再结构化解析，只有非标准 JSON 才回落到正则。

### 6.1 才有的四个坑（官方更新日志里能查到的）

| 变化 | 对本工具的影响 | 怎么办 |
|---|---|---|
| **6.1 公共链接支持「一次分享多个资源」**（主题内多个仪表板/组件/分析文档） | 页面里会出现多个 `reportId`，而工具**只处理第一个仪表板** | 自检第 2 步会 `[WARN]` 列出全部 id；要全导就给每个仪表板单独建公共链接、各配一行任务 |
| **6.1.9 新增「预览链接」** | 预览链接要求**登录平台**才能看，匿名跑不通 | 配任务必须用「公共链接」，不要用预览链接 |
| 公共链接可选**密码保护**与**有效期** | 带密码的链接取不到数据；到期后全线失败 | 给脚本单独开一个无密码、长期有效的公共链接，不要复用发给人的那条 |
| 表格默认列宽规则变更（`SystemOptimizationConfig.useNewColumnSizeRule`） | 只影响前端展示宽度，**不影响导出的数据** | 不用处理 |

> 上面第一条是最容易踩的：6.0 时代的链接天然只有一个报表，所以老链接升级后一般不受影响；
> 但只要你在 6.1 上**重新建**链接，就可能无意中建成了多资源链接。

### 升级前 / 升级后怎么验（各 1 条命令）

```bash
# 升级前：留一份基线指纹（version / 组件名 / 表头 / 总行数 / settings 字段）
python preflight.py --save before.json

# 升级后再跑，直接和基线对拍
python preflight.py --compare before.json
```

Windows 上双击 `preflight.bat`，或命令行 `preflight.bat --save before.json`。

排查结构差异时可以加 `--dump dump_dir`，会把分享页 HTML 和每个请求体/响应体原样存下来
（`00_share_page.html`、`01_req_detail.json`、`02_resp_data.json`…），失败时直接对着看。

自检走 5 步：站点与版本 → 分享链接 → 组件清单 → 组件模型 → 取数（2 页），
每步独立给 `OK` / `WARN` / `FAIL`，退出码 0 表示兼容、1 表示有阻塞。
对拍会打印一张「基线 / 当前」对照表，并标出 `settings` 字段的**新增**与**丢失**——
字段丢失通常才是真正需要改代码的信号。

### 真出问题了，按失败的步骤定位

| 失败步骤 | 大概率原因 | 处理 |
|---|---|---|
| 1 站点 | 地址/端口/防火墙 | 换到能访问 BI 的机器上跑 |
| 2 分享链接 | 链接失效、需登录、页面模板改了 | 浏览器重新打开链接确认；改 `bootstrap()` 里的候选规则 |
| 3 组件清单 | `reportWidgets` 结构变了 | 用浏览器 F12 看页面里的真实结构，按新结构补规则 |
| 4 组件模型 | detail 接口路径/参数变了 | 抓一次真实请求，比对路径与 body |
| 5 取数 | 载荷结构不被接受 / 翻页语义变了 | 看 `[FAIL]` 文案，配合 `--compare` 的字段差异改 `build_payload()` |

> 只升补丁（6.1.10 → 6.1.12）基本不用管；**跨小版本（6.0 → 6.1）务必先跑一次基线**，
> 这样万一行数或列名变了，是当场看见，而不是一个月后对数才发现。

---

## 九、表格类型与筛选（明细 / 分组 / 交叉）

BI 里三种表不是「换个样式」，**接口返回的结构就不一样**。这一节说清差异、落 Excel 的规则，
以及筛选条件怎么保证不丢。

> 下面 9.1 描述的是 6.0 的结构（也是 6.1 明细表的结构）；
> **6.1 的分组表和交叉表换了形态**（维度树 / 视窗），实测结构见 [8.4](#84-三种响应结构61-才有的分化)。

### 9.1 三类表在接口层的差异

| | 明细表 | 分组表 | 交叉表 |
|---|---|---|---|
| 组件模型信号 | `measures` 为空，字段全在 `dimensions` | `measures` 非空，维度在 `view['10000']` | 有列维度 `view['20000']`，或表头带 `children` |
| 表头 `header` | 平铺：`[{dId,text}…]` | 平铺（维度在前、指标在后） | 可能是**树**，只有叶子才是真实列 |
| 行 `items` | 每行长度 = 列数 | 合并区的行可能只返回首格 | 同样可能只返回部分格 |
| 若不处理会怎样 | 正常 | 维度列一片空白/重复，看不出分组 | 表头只剩一级、列错位 |
| 落 Excel | 单行表头 | 单行表头 + 维度列纵向合并 | **多级表头 + 合并单元格** |

### 9.2 类型判定与强制纠正

判定优先看**组件模型的 `type` 字段**（6.1 直接给：`4` 明细 / `2` 交叉 / `1` 分组），
取不到才按信号投票，依序：列维度 → 多级表头 → `measures` 非空 → 否则明细表。
判定结果和**依据**都会写进日志（`判定为 分组表（measures 非空（2 个指标））`），不会偷偷摸摸。

判错时两种纠正方式，都不用改代码：

- 配置列 `表样式` 填 `明细表` / `分组表` / `交叉表`（留空或 `自动` = 自动判定）；
- 命令行 `--kind detail|group|crosstab`。

### 9.3 多级表头怎么落 Excel

交叉表典型形态（「AUM」跨两个月）：

```
+--------+-------------------+
|        |       AUM         |   <- 第 1 行：跨 2 列
|  机构  +---------+---------+
|        |  1月    |  2月    |   <- 第 2 行：叶子列
+--------+---------+---------+
| 北京   |  10     |  20     |
```

落地规则：

- 叶子节点决定列数和列顺序；父节点占 `colspan`；
- **跨行规则**（这个最容易写反）：有子节点的只占自己那一层；叶子节点若不在最底层，要纵跨到最后一层；
- 合并写成 `mergeCells`，被覆盖的格子**不写值**（写了 Excel 打开会提示「修复文档」）；
- 冻结行数 = 表头实际行数，多级表头下方冻结，滚动时表头不跑。

### 9.4 行对齐：为什么不能按下标

交叉表/分组表里，合并区的第 2 行往往**只返回指标格，不返回维度格**：

```
行1: [北京][10][20]
行2:      [30][40]     <- 少一格，纯按下标会整体左移
```

所以对齐优先看每个 cell 自带的 `dId`：能对上列就放对应列；对不上才退化成按位置。
同一 `dId` 在一行里出现多次（交叉表的重复指标）时按顺序消费。
首行长度恰好等于列数时直接按位置（最快、也最稳，明细表走的就是这条路）。

### 9.5 分组列的三种处理（`分组列` 列 / `--group-cols`）

| 值 | 效果 | 适用 |
|---|---|---|
| `merge`（默认） | 只合并单元格，**数据一字不改** | 要跟 BI 页面长得一样、给人看 |
| `fill` | 同值向下填满，不合并 | 要拿去做透视/二次计算 |
| `raw` | 完全原样，接口给什么就是什么 | 要严格核对原始返回 |

合并只作用在**开头连续的维度列**（指标列绝不合并），且只合并「连续且非空且相同」的区间——
不会把两段不相邻的相同值硬连起来，也不会把空白区并成一块。

### 9.6 筛选条件：怎么保证不丢

页面上的筛选（组件内过滤、结果过滤器、日期参数）都在组件模型里。重建载荷时：

1. **整个模型深拷贝**，只改 `page` / `wId` / `queryInfo` / `tableStyle.totalRows` 四个地方，
   所以筛选条件、参数、联动配置**原样带回**，导出的就是页面看到的那个结果集；
2. 每个组件导出前把跟筛选/参数有关的配置**摘要打进日志**（如 `filterValues（2 项）、组件参数 1 个`），
   数量对不上就说明模型没带全；
3. **行数自证**：响应里的 `row` 是服务端按筛选条件算出的总行数，导出后拿它比对实际行数，
   少了就在日志里写 `行数不符：服务端报 110000 行，实际拿到 1000 行`——
   这正是之前 `totalRows` 被截断那类问题，以前是静默少数据，现在是当场能看见。

> 想让脚本**额外**加筛选（比如只导某个机构），目前建议在 BI 侧建两个组件/两条链接，
> 而不是在脚本里拼过滤条件：过滤器结构各版本差异大，硬拼容易取到错数据。

### 9.7 配置与命令

配置表新增两列（都不是必填，老配置不加也能跑）：

| 列名 | 取值 | 缺省 |
|---|---|---|
| `表样式` | 自动 / 明细表 / 分组表 / 交叉表 | 自动 |
| `分组列` | 合并 / 填充 / 原样 | 合并 |

命令行同名开关：`--kind`、`--group-cols`。

### 9.8 怎么验证与校准

```bash
python test_table_norm.py            # 三类表的纯逻辑自测（不需要 BI，52 项，含 6.1 的三种结构）
python probe_table.py --out probe/   # 探真实组件：表型、响应形态、表头层级、首行对齐结果
```

`probe_table.py` 会同时打印**响应形态**（平铺 / 分组树 / 交叉表视窗）和表型的判定依据，
6.1 上校准交叉表、分组表就靠它。

`probe_table.py` 会把每个组件的 `detail` / 第 1 页数据 / 表头树原样存到 `probe/`。
拿到真实的交叉表、分组表后，先跑它确认判定与列数，再跑正式导出。

### 9.9 已知限制

- **图表类组件不导**（`series` 结构，不是行列表），会被标为「跳过」并写原因；
- 交叉表的**合计/小计行**按接口原样落，不做二次计算，位置与 BI 页面一致；
- 一个公共链接只处理**第一个仪表板**（6.1 多资源链接见第八节）；
- 数字列按文本落盘（避免 18 位证件号被 Excel 截断精度），需要数值计算的在 Excel 里转一次；
- **6.1 分组表的合计行**：接口不给标签（`text` 为 `None`），代码补成「合计」。
  如果你不想要这一行，把配置列 `表样式` 留着自动、并在 Excel 里筛掉即可；
- **6.1 交叉表是视窗结构**：默认按 `FINEBI_BIG_ROWS=200000` 一次取全。
  行/列特别多的交叉表（比如上千列）可能仍会分多次，日志会给 `hasNext` 提示；
- **脚本常驻占用 1 个 BI 并发名额**（6.1 的并发按 cid 计）。许可紧张时请让管理员预留或调大，
  别删 `.cid_cache`（会让名额一直挂着到会话过期）。

---

## 十、合规提醒（重要）

样例组件含 **客户姓名、身份证号、手机号、住址、资产余额** 等个人金融信息，且是 **11 万条**：

- 落盘目录别放共享盘/桌面同步目录，建议限定权限（`chmod 700 out`）；
- 仓库 `out/` 里的验证产物**含真实客户记录，用完请删除**；
- 生产上更稳的做法是 FineBI 自带「定时调度 → 导出」把文件投递到受控位置，
  不让脚本把明文 PII 拉到个人电脑；
- 分享链接本身是**匿名公开**的（无登录、无口令、无水印），行内合规属风险项，建议收敛为组织内鉴权访问。

---

## 文件清单

```
finebi-export/
├── 操作手册.md              # ★ 使用者先看这个（配置/运行/排错/合规）
├── README.md                # 本文：接口链路、载荷结构、版本兼容（改代码前看）
├── DEPLOY-WINDOWS.md        # Windows 部署说明（Python 3.11 + Java 8）
├── config.xlsx              # 配置样例（改这个）
├── webadmin.bat             # 一键打开管理台（Windows 双击）
├── collector.py             # 调度入口：常驻 cron / --once / --json / --job
├── finebi_export.py         # 采集核心：发现组件 / 重建载荷 / 翻页 / 筛选 / 写文件
├── table_norm.py            # 三类表归一化：类型判定 / 多级表头 / 分组列 / 筛选摘要
├── xlsx_io.py               # xlsx 读写 + 文件名安全（Windows 规则）
├── cron_expr.py             # cron 解析与匹配（仅常驻模式用）
├── check_env.py             # 部署自检（新环境第一件事跑它）
├── preflight.py             # 版本兼容自检 + 升级前后对拍（见第八节）
├── probe_table.py           # 单组件结构探针（校准交叉表、分组表用）
├── probe_filters.py          # 过滤组件结构探针（日期/区间类能不能传值）
├── webadmin.py / admin.html   # 网页管理台：改配置、跑任务、看日志、下产出（python webadmin.py）
│                             #   内置 cron 自动调度：管理台开着时到点自动跑 config 里的定时任务
├── test_table_norm.py       # 三类表自测（纯逻辑，不需要 BI）
├── capture_payload.js       # 排错：headless Chrome 抓页面真实取数请求体（Node 18+，可选）
├── preflight.bat            # Windows 双击版自检
├── run_once.bat             # Windows 一键跑一轮
├── java8/                   # Java 8 调度壳（ProcessBuilder 调 Python，可选）
│   ├── pom.xml
│   └── src/main/java/finebi/{FineBiExportRunner,Main,FineBiExportJob.example}.java
├── payload_template.json    # 旧「固定模板模式」的抓包样本（正常模式不需要）
├── out/                     # 产出目录（可随时清空；内含真实客户数据，注意合规）
├── logs/                    # 运行日志（按天切分，保留 30 天）
├── .cid_cache               # 每个站点复用的 cid（★ 别删，见 8.2）
└── .collector_state.json    # 常驻模式的执行幂等记录（自动创建）
```
