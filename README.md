# SAP 数据自动化

一套定时运行的脚本：**从 SAP 拉取报表 → 本地处理 → 写入 Google Sheets**。

由 `run_all.ps1` 依次执行 7 个独立脚本，全程约 9–10 分钟。

其中 `00`/`01`/`02` 每次会处理**上一个日历月 + 当前月**（见「跨月补偿」），所以比其余脚本慢一倍。

---

## ⚠️ 运行期间的重要注意事项

### ✅ SAP 可以放心开着

**脚本不会关掉你的 SAP 会话。**

- `run_all.ps1` 只在检测到 SAP **没在运行时**才启动一个；本来就开着的话，全程不动它，收尾也不关
- 各脚本在**自己的新会话**里干活（约 3 秒开一个），用完只关掉自己那个窗口
- 所以你在自己电脑上开着 SAP 做事，自动化不会打断你

> 这是 2026-10-07 改的。**在此之前，每个脚本都会 `taskkill` 掉机器上所有 SAP 会话**（包括你正在用的登录会话），那个行为已经没有了。

### ⚠️ 但 Excel 仍然要注意

| 脚本的行为 | 后果 |
|---|---|
| 用 `excel.Workbooks(1)` 取"第一个打开的工作簿" | 如果你开着别的 Excel 文件，脚本可能拿到的**不是它自己导出的那个** |
| 对取到的工作簿执行 `wb.SaveAs(自动化输出路径)` | 那个文件会被**改名并移动到** `O:\My Drive\071 - SAP 数据\...` |
| 紧接着执行 `wb.Close(SaveChanges=False)` | **未保存的修改直接丢弃，不弹提示** |

**结论：跑 `run_all.ps1` 之前，请先保存并关闭所有 Excel 工作簿。SAP 不用管。**

如果这台机器同时还要给人用，考虑把定时任务安排在无人时段（主要是避开 Excel 那一项）。

---

## 运行方式

### 全量运行（7 个项目）

```powershell
cd "C:\Users\kelland zhao\Projects\SAP_Google_AutoRun"
.\run_all.ps1
```

执行顺序：`00 → 01 → 02 → 03 → 04 → 06 → 07`（**不含 `05` 和 `08`**，见下方「项目清单」）。

结果同时输出到控制台和 `run_all.log`。开头会打印本次运行的**代码版本**，便于事后核对：

```
========================================
  SAP 自动运行 - 共 7 个项目
  代码版本: 09e73e7
  开始: 2026-10-05 13:35:31
========================================
```

### 单独运行某个项目

```powershell
cd "C:\Users\kelland zhao\Projects\SAP_Google_AutoRun"
python "03 - Safety_Stock_ZSE16\main.py" ; "退出码 = $LASTEXITCODE"
```

> VS Code 的「运行 Python 文件」按钮**不会**打印退出码，要判断成败必须在终端里手动输入上面的命令。

### 定时运行

用 Windows 任务计划程序触发 `run_all.ps1`。两点要求：

1. 必须配置为**「只在用户登录时运行」**——SAP GUI Scripting 和 Excel COM 都需要交互式桌面，在 session 0（无桌面）里会失败。
2. 任务的「起始于」目录、`run_all.ps1` 里的 `$Root`、你 `git pull` 的目录，**三者必须是同一个**，否则跑的是另一份代码。

### 跨月补偿

`00`、`01`、`02` 三个 KPI 脚本每次运行都会按时间顺序处理：**上一个日历月 + 当前月**。

这样即使月初当天是假期、任务未运行，或当天 SAP/Excel/网络失败，下一次运行也会重新处理上个月，并把结果写回对应的 `MasterData` 行。月级处理失败会以退出码 `1` 结束，避免把失败结果写入 Google Sheets。

这项改动只防止今后再次漏更新；已经漏掉的历史月份（例如 `202608`）需要单独做一次 SAP 回填，确认回填结果后再刷新 DOMO。

---

## 退出码

| 退出码 | 含义 |
|---|---|
| `0` | 成功。**包括"本月无数据、跳过上传"这种情况**——那是正常结果，不是故障 |
| `1` | 失败。SAP 连不上、导出失败、Google Sheets 写入失败等 |

`run_all.ps1` 依据退出码统计成功/失败，所以**这个数字是可信的**。若某个脚本报失败，它的完整输出（含 traceback）会打印在控制台上。

历史的坑：2026-10-05 之前，所有脚本出错时都只打印错误然后正常退出（退出码 0），导致 `run_all` 永远显示"7 个全部成功"。该问题已修复。

---

## 项目清单

### 由 `run_all.ps1` 执行

| 目录 | SAP 事务码 | 产出 | 写入的工作表 |
|---|---|---|---|
| `00 - Maintenance_Plan_Adherence_Sync` | IW39 | 维护计划遵守率 + 未执行工单明细 | `MasterData`、`Maintenance_Plan_Adherence_Gap` |
| `01 - Maintenance_Effectiveness_Sync` | IW47 | 维护有效性 | `MasterData` |
| `02 - Critical_A_ &_H_equipment_with_Maintenance_Plan_Sync` | IH08 → IP18 | A 级关键设备中有维护计划的比例 + 无计划设备清单 | `MasterData`、`无保养计划A类设备` |
| `03 - Safety_Stock_ZSE16` | ZSE16 | 安全库存全量 | `安全库存数据` |
| `04 - Total_Workorder` | IW39 | 全年工单总表 | `Total_Workorder` |
| `06 - IM_Equipment_Number` | IH08 | 设备编号（含 Tag 列） | `Equipment_Number_EAM` |
| `07 - Inventory_MB52` | MB52 | 库存 | `MasterData` |

`00`、`01`、`02` 写入**同一张** Google 表格的 `MasterData` 工作表，各自占不同的列区间。

### `00` 的附加产出：`Maintenance_Plan_Adherence_Gap`

`00` 还会把**未执行**（【系统状态】不含 CNF，且状态非空）的工单明细写进同一张 spreadsheet 的 `Maintenance_Plan_Adherence_Gap` 工作表。12 列：前 10 列取自 IW39 导出（参考日期、通知、订单、功能位置、设备、描述、成本中心、系统状态、ABC 标识、订单类型），后 2 列 `写入时间`、`月份` 由脚本生成。列的位置按**表头名**去找，不写死列号。

- **数据每次都刷新，写入时间锁在该月首次写入的时刻** —— 所以每天跑，这一列也不会被刷成新时间
- 刷新只动正在处理的月份：其余月份的历史行原样保留（靠最后一列 `月份` 区分）
- 某月没有未执行工单时，该月的行会被清空（连同上一次记下的写入时间）；之后若又出现，会记一个新的时间
- 写入用 `RAW` 方式：所有格子按文本原样落库，带前导零的订单号/设备号不会被 Sheets 当数字吃掉
- 导出文件里找不到上面那 10 个表头时：**只打警告并跳过 Gap 写入**，`MasterData` 照常更新、退出码 0（Gap 会停更，直到你注意到那行警告）。而写 Gap 表本身失败（网络/权限/表被删）则退出码 1，下次重跑

> 逻辑（找列、日期归一化、按月拆行）放在仓库根的 `gap_utils.py`，不 import `win32com`，所以它的单元测试在任何机器上都能跑：`python -m unittest discover -s tests`。

`03` 和 `07` 写入**另一张**表格：`03` 用 `安全库存数据` 工作表，`07` 用 `MasterData` 工作表。

### 不在 `run_all.ps1` 中（当前挂起）

这两个项目**因 SAP 系统层面的问题暂时不可用**，已挂起。不是脚本本身的问题，恢复后需单独手动运行。

| 目录 | SAP 事务码 | 说明 |
|---|---|---|
| `05 - Stock_Turnover` | MC.7 | 库存周转率 → 写入 `Summary` |
| `08 - WorkOrderAutoAcquisition` | IW39 | 工单 → 写入 `Database`。目录内另有一个 SAP 录制的 `.vbs` |

---

## 环境要求

### 运行机（Windows）

- **Windows** + 已安装并可用的 **SAP GUI**（需启用 GUI Scripting）
- **Microsoft Excel**（脚本通过 COM 驱动它导出/另存）
- **Python 3.11**，路径：`C:\Users\kelland zhao\scoop\apps\python311\current\python.exe`
- **`O:` 盘已映射**到 Google Drive（输出目录都在 `O:\My Drive\071 - SAP 数据\` 下）

### Python 依赖

```powershell
pip install pywin32 openpyxl gspread google-auth requests urllib3 pandas numpy
```

| 包 | 用途 |
|---|---|
| `pywin32` | `win32com.client` 驱动 SAP GUI 和 Excel |
| `openpyxl` / `pandas` / `numpy` | 读写和清洗导出的 Excel / 文本 |
| `gspread` + `google-auth` | 写入 Google Sheets |

> 项目里**没有 `requirements.txt`**，依赖需手工安装。

### 凭据

**Google 服务账号**：把 `pyreadsp-*.json` 放在**仓库根目录**（与 `run_all.ps1` 同级）。该文件已被 `.gitignore` 排除，不会进仓库，换机器时需要手工放一份。

同时，该服务账号的邮箱必须被加为目标 Google 表格的**编辑者**，否则会报"找不到工作簿"。

**SAP 凭据**：脚本里写的是字面量 `'-user=USERNAME', '-pw=PASSWORD'`，**这是有意的占位符，不是配置遗漏**。公司电脑上 SAP 靠 SSO 登录，这两个参数不生效。不要试图"补全"它们。

（这段占位符统一定义在仓库根的 `sap_common.py` 里，各脚本通过 `start_sap()` 调用。）

---

## 常见故障

| 现象 | 原因 / 处理 |
|---|---|
| `TimeoutError: SAP 会话在 60 秒内未就绪` | SAP 没能启动或没能登录。检查 `sap_common.py` 里的 `SAP_SYSTEM` / `SAP_CLIENT`（应为 `LAP` / `321`） |
| `❌ Google Sheets 错误: 找不到工作簿` | 服务账号邮箱没被加为那个表格的编辑者，或 `pyreadsp-*.json` 不在仓库根目录 |
| `⚠️ 警告: ...没有数据，跳过上传（不算失败）` | **正常**。本月无数据，脚本正常结束、退出码 0 |
| `⚠️ 警告: 没有找到打开的 Excel 工作簿` | SAP 导出没成功弹出 Excel。看前一步的导出日志 |
| 卡在 `正在等待 Excel 打开...` 很久 | Excel 有弹窗挡着，或开着一堆工作簿 |
| `机器上没有正在运行的 SAP 进程，跳过。` | **正常**。SAP 本来就没开着，无需关闭 |

---

## 已知限制

1. **Excel 仍是独占的**——见开头的警告。这是当前设计的前提，不是 bug。
2. **仍有大量固定秒数等待**——9 个 `main.py` 里还有 **58 处** `time.sleep()`，单轮合计约 **118 秒**。SAP 启动相关的那批已改成轮询，**其余散布在 SAP 操作步骤之间，尚未处理**——机器慢的时候可能等不够。
3. **连接 Google Sheets 时关闭了 TLS 校验**（`verify = False`），为兼容公司代理。属于可收紧项。

> **历史（都已解决，留作背景）**
>
> - **2026-10-06** 公共代码重复——`close_SAP` / `start_sap` / `get_resource_path` 曾在 9 个文件里各有一份，现抽出到 `sap_common.py`。
> - **2026-10-07** 误杀 SAP 会话——原先每个脚本都会 `taskkill` 掉机器上所有 SAP，现已改为「`run_all.ps1` 启动一次 + 各脚本开自己的新会话」。
> - **2026-10-07** 退出码——原先脚本出错只打印、仍以 0 退出，`run_all` 永远显示"全部成功"。现在失败会以退出码 1 结束。
> - **2026-10-07** 跨月补偿——`00`/`01`/`02` 现在每次都补跑上一个日历月，月初漏跑不再导致上月数据长期不更新。

---

## 目录里还有什么

| 文件 | 说明 |
|---|---|
| `run_all.ps1` | 编排器：依次运行 7 个项目并统计成败 |
| **`sap_common.py`** | **各脚本共用的公共模块**，提供 `get_resource_path` / `start_sap` / `wait_for_sap_session` / `close_sap`。SAP 的系统、客户端、语言等配置也在这里 |
| `run_all.log` | 运行日志（含代码版本、每步耗时、每个脚本的退出码） |
| `*/batch_update.py` | **一次性历史回填脚本**（`00`/`01`/`02` 各一份），用硬编码的月份列表补跑历史数据，日常运行不会执行。**尚未迁到 `sap_common`** |
| `08 - .../*.vbs`、`*/*.txt` | SAP GUI 录制脚本及其导出产物，历史遗留 |
| `*/.gitignore` | 各项目自己的忽略规则 |

各脚本开头这两行的作用，是让它们**既能被 `run_all.ps1` 调用、也能单独运行**：

```python
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from sap_common import get_resource_path, start_sap, wait_for_sap_session, close_sap
```

> ⚠️ `get_resource_path()` 的基准目录是**仓库根**，不是项目目录。用它取密钥时写 `get_resource_path('pyreadsp-xxx.json')`，**不要写 `'../pyreadsp-xxx.json'`**。

---

## 维护

代码通过 GitHub 同步：本地改 → `git push` → 运行机 `git pull`。

提交信息格式：`VYYYY-MM-DD.XX_简短描述`（同一天内 `.XX` 递增）。
