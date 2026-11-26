# Windows 部署说明（Python 3.11 + Java 8）

采集本身是纯 Python、零第三方依赖；Java 8 只作为调度壳（ElasticJob / Quartz）把 Python 叫起来。
两条路互不影响，只用其中一条也行。

---

## 一、环境准备

| 项 | 要求 | 说明 |
|---|---|---|
| Python | **3.11**（3.7+ 语法都能跑，已在 3.11.9 / 3.13.12 实测；3.8 兼容性已按语法全量校验） | 安装时**务必勾选 `Add python.exe to PATH`** |
| 第三方包 | **不需要** | 取数用 `urllib`，写 xlsx/csv 用 `zipfile`/`csv`，全部标准库 |
| Java | 8（仅调度壳需要） | 不接 ElasticJob 就不用装 |
| 磁盘 | 预留 **≥ 5 GB** | 两个组件全量约 12.8 MB/次，按保留天数估算 |

**Windows 7 环境补充**（旧机器部署）：

- Win7 官方支持的最高 Python 是 **3.8.10**（3.9 起安装包不再支持 Win7），装这个版本即可跑通全套；
  `check_env.py` 会提示「低于实测版本 3.11」的 WARN——这是预期提醒，不是错误；
- 代码零第三方依赖，不存在"老 Python 装不到新库"的问题；
- **网页管理台（webadmin.bat）需要 Chromium 内核浏览器**：Win7 上用 Chrome 109 / Edge 109（各自在 Win7 的最终版本）打开；Win7 自带的 IE11 不支持页面里的 ES6 语法，会白屏。命令行采集本身不依赖任何浏览器；
- 3.8 已于 2024-10 停止官方维护，内网封闭环境无碍；若 BI 是 https，Win7 的旧根证书可能需要更新一次（默认 http 内网端口无此问题）。

**不要用 Microsoft Store 版 Python**：它的真实路径在 `...\WindowsApps\` 下，
计划任务和服务里常常找不到，报「找不到 python」。用官网安装包或绿色版解压版都行。

装完在 cmd 里确认：

```bat
python -V
where python
```

`where python` 应该指向 `C:\...\Python311\python.exe` 之类，而不是 `WindowsApps`。

> Python 官方从 3.11 起，Windows 安装包里已经自带 `py` 启动器。
> 装多个版本时用 `py -3.11 -V` 可以准确指定，别只依赖 `python`。

---

## 二、部署五步

### 1. 拷贝整个目录

把 `finebi-export` 整个目录拷到目标机器，例如 `D:\apps\finebi-export\`。

> 路径建议**纯英文**。中文路径本程序支持（内部全部 UTF-8），
> 但 Java 侧、计划任务、某些杀软对中文路径的处理更容易出岔子。

### 2. 自检（先跑这个，别急着导数据）

```bat
cd /d D:\apps\finebi-export
python check_env.py --net
```

逐项检查 Python 版本、标准库、程序文件、配置可读性、输出目录可写、FineBI 端口与分享链接可达。
全 `[ OK ]` 再往下走；有 `[FAIL]` 会告诉你具体原因和常见的修法。

### 3. 改配置

用 Excel 打开 `config.xlsx`，一行一个任务：

| 任务名 | 启用 | 分享链接 | 组件 | 输出方式 | 输出格式 | 输出目录 | 文件名 | cron表达式 |
|---|---|---|---|---|---|---|---|---|
| 客户明细 | 是 | `http://内网地址/webroot/decision/link/xxxx` | *(留空)* | `多文件` | *(留空=csv)* | `D:\data\finebi` | *(留空)* | `30 8 * * *` |

要点：

- **`组件` 留空** = 页面上全部表格组件，一个不落。页面以后增减组件，配置不用改。
- **`输出方式` 填 `多文件`**（或留空，默认就是多文件）= 每个组件一个文件。
- **`输出格式` 留空 = csv（默认）**。csv 没有百万行上限（xlsx 单表上限 1,048,576 行），
  数据源涨过百万也不会写挂；要合并单元格/Excel 直接打开不丢精度就填 `xlsx`。
- **`文件名` 里 `{widget}` 就是组件名**（取 BI 里看到的名字），留空走默认 `{name}_{widget}_{yyyyMMdd}`，
  产物形如 `客户明细_全部明细_20261125.csv`（任务名打头，多任务共用一个输出目录也不会混）；
  组件从没命名过才会退化成 `组件1_20261125.csv`。扩展名也可以写（会被输出格式替换），或干脆不写。
- **`输出目录` 建议写绝对路径**，如 `D:\data\finebi`：不受启动目录影响，计划任务/服务怎么启动都落同一处。
  相对路径按 **`config.xlsx` 所在目录** 解析（不是进程当前目录）。
  `~`、`%USERPROFILE%`、`\\server\share` 都支持，目录不存在会自动创建。
  完整的路径写法对照表见 README 第二节「路径怎么写」。

想确认页面上到底有几个组件、都叫什么：

```bat
python collector.py --list-widgets
```

### 4. 试跑一轮

```bat
python collector.py --list                 :: 先看任务解析和下次执行时间，不取数
python collector.py --once --max-pages 2   :: 冒烟：每组件只抓 2 页，几秒出结果
python collector.py --once                 :: 正式跑一轮全量
```

### 5. 全量跑

```bat
run_once.bat
```

或者直接 `python collector.py --once --json`。全量两个组件约 22 万行、**210 秒**、12.8 MB
（Python 3.11 实测；同样的数据在 3.13 上约 165 秒）。

---

## 三、三种调度方式，按需选一个

> **三者只能选一个**，否则会重复取数互相覆盖文件。

### 方式 A：纯手动（你现在的选择）

双击 `run_once.bat` 就行，或者命令行带参数：

```bat
run_once.bat --job 客户明细
```

同一天重复跑会**覆盖**同名文件（`客户明细_组件1_20261125.csv`），这正好是日报该有的行为。

### 方式 B：Windows 计划任务

```bat
schtasks /create /tn "FineBI导出" /tr "D:\apps\finebi-export\run_once.bat" ^
  /sc daily /st 08:30 /ru SYSTEM /rl HIGHEST
```

图形界面建任务时注意两点：

- **起始位置** 填不填都行：程序以 `config.xlsx` 所在目录为基准解析相对路径，不看当前目录。
  但**「程序或脚本」和「起始位置」里都用绝对路径**更稳，避免 PATH 里 `python` 指向别的版本；
- 勾选「不管用户是否登录都要运行」，否则锁屏/注销后不跑。

### 方式 C：ElasticJob（Java 8）

Java 侧不重复实现取数逻辑，只负责「到点调 Python + 判断成败」。

**编译调度壳**：

```bat
cd /d D:\apps\finebi-export\java8
mvn clean package
```

产物 `target\finebi-export-runner.jar`，先单独验证一下：

```bat
java -jar target\finebi-export-runner.jar ^
  --python C:\Python311\python.exe ^
  --work-dir D:\apps\finebi-export ^
  --config D:\apps\finebi-export\config.xlsx ^
  --job 客户明细
```

**接进 ElasticJob**：把 `src\main\java\finebi\FineBiExportJob.java.example` 改名为 `.java`，
按文件头部注释加上 ElasticJob 依赖（3.x 用 `elasticjob-lite-spring-boot-starter:3.0.4`），
Job 配置里：

```yaml
elasticjob:
  jobs:
    finebiCustomerDaily:
      elasticJobClass: finebi.FineBiExportJob
      cron: 0 30 8 * * ?          # 这里用 Quartz 6 段写法
      shardingTotalCount: 1
      jobParameter: 客户明细       # 一个分片对应一个配置里的任务名
```

几个要点：

- **调用时必须带 `--once`**（`FineBiExportRunner` 内部已经带了），绝不要再启动 Python 侧的常驻调度器，否则同一时刻会有两套在跑；
- 用了 ElasticJob 之后，`config.xlsx` 里的 `cron表达式` 列只当备注看，真正的时间由 ElasticJob 的 cron 控制；
- 失败时 `FineBiExportRunner` 抛 `IllegalStateException`，ElasticJob 会按你的 failover / 重试策略处理；
- `FineBiExportJob` 用 `System.getProperty("finebi.python")` 等读路径，也可以在代码里写死常量。

---

## 四、Windows 专有问题（都已处理，但你需要知道）

| 现象 | 原因 | 现在怎么处理 |
|---|---|---|
| 日志中文乱码 / `UnicodeEncodeError` | Windows 控制台默认 GBK | 已处理：cmd 窗口下走 Windows Unicode API 天然不乱码；管道/重定向只降级罕见字符不报错。**三个 `.bat` 必须保持 ANSI/GBK 编码**（cmd 按 GBK 解析 bat，存成 UTF-8 会整段乱码报错） |
| 子进程把 Java 卡死 | stdout/stderr 管道写满 | `FineBiExportRunner` 用**两个线程**并发读，不会死锁 |
| 报「文件被占用，无法写入」 | 文件正在 Excel 里打开 | 已翻译成明确中文提示；关掉 Excel 重跑即可 |
| 组件名带 `/\:*?"<>|` 导致写文件失败 | Windows 文件名规则 | 自动替换成 `_`，并处理 `CON/PRN/AUX/NUL/COM1-9/LPT1-9` 这类保留名 |
| 两个组件名只差大小写 | Windows 不区分大小写，会互相覆盖 | 文件名去重时会按不区分大小写判断，第二个自动加 `_2` |
| 路径过长报错 | 260 字符限制 | 文件名主干已限制在 80 字符内；目录别嵌太深 |
| 首次运行很慢/被杀软拦 | 杀软扫描大量临时写入 | 把输出目录加进白名单 |
| 计划任务跑出来目录不对 | 起始位置没设 | 起始位置设为程序目录，或配置里写绝对路径 |

---

## 五、目录结构

```
finebi-export/
├── 操作手册.md             使用者手册（配置/运行/排错，先看这个）
├── README.md               技术说明（接口链路、版本兼容）
├── DEPLOY-WINDOWS.md       本文
├── check_env.py            部署自检（Windows 上第一个跑它）
├── preflight.py            版本兼容自检 + 升级前后对拍
├── preflight.bat           Windows 双击版自检
├── run_once.bat            Windows 一键跑一轮
├── collector.py            调度入口（--once / --json / --job）
├── table_norm.py           三类表归一化（明细/分组/交叉）
├── probe_table.py          单组件结构探针
├── probe_filters.py        过滤组件结构探针（日期类筛选排查）
├── webadmin.py / admin.html   网页管理台（python webadmin.py，浏览器里操作；内置 cron 自动调度，
│                              开着管理台到点自动跑，与 collector.py 常驻模式二选一）
├── webadmin.bat            一键打开管理台（双击，自动开浏览器）
├── test_table_norm.py      三类表自测（纯逻辑）
├── capture_payload.js      排错用：抓页面真实请求体（需 Node 18+，可选）
├── config.xlsx             配置（改这个）
├── finebi_export.py        采集核心（发现组件 / 重建载荷 / 翻页 / 写文件）
├── cron_expr.py            cron 解析（仅常驻模式用）
├── xlsx_io.py              xlsx 读写 + Windows 文件名安全
├── java8/                  Java 8 调度壳
│   ├── pom.xml
│   └── src/main/java/finebi/
│       ├── FineBiExportRunner.java   ProcessBuilder 调用 + 并发读管道
│       ├── Main.java                 命令行入口
│       └── FineBiExportJob.java.example  ElasticJob 示例（改名启用）
├── out/                    产出目录
└── logs/                   运行日志（按天切分，保留 30 天）
```

---

## 六、验证过的清单

**Python 3.11.9 与 3.13.12 两个版本都实测通过**（同一套代码，未用任何 3.11 以上才有的 API）；
逻辑与 Windows 一致，差异只在路径和编码，已按 Windows 规则处理：

- 页面 2 个组件 → 各自独立成文件 `客户明细_组件1_20261125.csv`（11 列 × 11 万行）、`客户明细_组件2_20261125.csv`（2 列 × 11 万行）
- 全量 22 万行 / **209.6 秒** / 12.8 MB，退出码 0（Python 3.11.9）
- 数据校验：手机号 11 位纯数字 110000/110000，列对齐异常 0，客户号 0000000001→0000110000 连续
- `--json` 输出可被调度器解析，日志与结构化结果分走 stderr / stdout
- 文件名安全：保留名、禁用字符、结尾空格与点、大小写冲突、超长，全部通过单测
- 自检脚本 `check_env.py --net` 全项 OK

**Windows 真机还没跑过**（我这边是 macOS）。到 Windows 上第一件事跑 `python check_env.py --net`，
如果有 `[FAIL]` 把输出贴给我，我按实际情况修。

---

## 七、FineBI 升级时怎么做（6.1.10 → 6.1.12 或 6.0 → 6.1）

补丁升级（6.1.10 → 6.1.12）接口不会动，按下面走一遍只是求个心安；
**跨小版本（6.0 → 6.1）必须做**，因为内部 JSON 结构可能变，而变了之后往往不报错、只是少数据。

升级前（趁还没升，先留证据）：

```bat
preflight.bat --save before.json
```

升级后：

```bat
preflight.bat --compare before.json
```

看三处就够：

1. **结论行**：`兼容` 直接收工；`不兼容` 看下面的阻塞项。
2. **对照表**：版本 / 组件数 / 组件名 / 列数 / 表头 / 总行数 / 每页行数，任一列出现 `<-- 变化` 都要确认是否符合预期。
3. **settings 字段**：`[WARN] 新增` 一般无碍；`[FAIL] 丢失` 才是需要改代码的信号（尤其涉及 `tableStyle`）。

行数疑似被截断（总行数正好 1000 / 10000 这类整数）时，先用环境变量试，不用改代码：

```bat
set FINEBI_TOTAL_ROWS=-1
run_once.bat --max-pages 3
```

平时也可以把 `preflight.py` 挂到每月一次的计划任务里，只做监控不做导出——
它只发 5~6 个请求，比全量导出轻得多，能在结构悄悄变化时提前发现。

如果你在 6.1 上**重建**过链接，注意两点：别建成「一次分享多个资源」的链接（只会导第一个仪表板，
自检第 2 步会给 `[WARN]`），也别开密码保护（脚本取不到数）。用 6.0 时代的老链接一般没这问题。

### 7.1 已经升到 6.1 了，Windows 上要多看三件事

6.1 和 6.0 在取数链路上有四处实质变化（cid 必填、并发限制、page 语义、响应结构分化），
代码已经全部适配，在 6.1.12 上实测跑通。Windows 上只需要注意三点：

| 事项 | 怎么确认 | 备注 |
|---|---|---|
| **并发名额** | 日志里若出现 `BI concurrency has reached the limit of license (limit = 2)` | 脚本常驻占 1 个名额（按 cid 计）。别删 `.cid_cache`；许可紧张就找 BI 管理员调大 |
| **分组/交叉表取全没** | 日志出现 `服务端标记还有更多数据（hasNext）` 就是没取全 | 加大：`set FINEBI_BIG_ROWS=1000000` 再跑 |
| **报 `cache Page_State expired`** | 说明 `page` 起始值用错了 | 现在已按表型自动处理；若自己改过代码，记住分组/交叉表首次是 `page=0` |
| **明细表导得很慢** | 日志里几十上百次「第 N 页 +100 行」 | 每页行数在载荷的 `tableStyle.detailRowCounts`，默认 100。已自动放大到 5000（11 万行 286s → 23s）；想再调：`set FINEBI_DETAIL_PAGE_ROWS=10000`，设 0 恢复服务端默认 |
| **筛选没生效，导的还是全量** | 日志里没有「生效筛选：…」这一行 | 链接后拼 `?过滤组件名=值`，名字写页面上那个过滤组件的**标题**（不是数据字段名）。写错会告警并列出可用名：`参数 xxx 没有同名过滤组件，未生效（可用名：性别、姓名）`。详见 README 8.5 |

跑之前建议先探一次结构，确认表型判对（不会写文件、不落数据）：

```bat
python probe_table.py --base http://<BI地址>:37799 --code <链接码>
```

输出里每个组件会给出「表型 + 响应形态 + 列数/行数 + 首行对齐结果」，
分组表应显示「分组树」、交叉表应显示「交叉表视窗」。

---

## 八、三类表（明细 / 分组 / 交叉）在 Windows 上要注意什么

不需要额外装任何东西，但有三点值得知道：

1. **第一次接新页面，先跑探针**：`python probe_table.py --out probe/`
   会打印每个组件的表类型、表头层级、首行对齐结果，并把原始报文存到 `probe\`。
   判定不对就用配置列 `表样式` 或 `--kind` 强制指定。
2. **分组表的维度列默认合并单元格**（`merge`）。如果下游要拿 Excel 做透视/二次计算，
   把 `分组列` 列改成 `填充`；要严格核对原始返回就用 `原样`。
3. **行数要核**：日志里每个组件会打 `取数完成：N 行`，若同时出现
   `行数不符：服务端报 X 行，实际拿到 Y 行`，说明被截断，先查 `totalRows`（见第七节）。

纯逻辑自测随时可跑，不需要连 BI：

```bat
python test_table_norm.py
```
