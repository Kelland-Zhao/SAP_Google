# SAP 数据自动化

一套定时运行的脚本：**从 SAP 拉取报表 → 本地处理 → 写入 Google Sheets**。

由 `run_all.ps1` 依次执行 7 个独立脚本，全程约 9 分钟。

---

## ⚠️ 运行期间的重要注意事项

**这套脚本会独占这台电脑的 SAP GUI 和 Excel。运行期间不要用它们做别的工作。**

| 脚本的行为 | 后果 |
|---|---|
| 每个脚本开始和结束时各执行一次 `taskkill /im saplogon.exe /t /f` | **强制关闭所有 SAP GUI 进程**，包括你自己正在用的登录会话 |
| 用 `excel.Workbooks(1)` 取"第一个打开的工作簿" | 如果你开着别的 Excel 文件，脚本可能拿到的**不是它自己导出的那个** |
| 对取到的工作簿执行 `wb.SaveAs(自动化输出路径)` | 那个文件会被**改名并移动到** `O:\My Drive\071 - SAP 数据\...` |
| 紧接着执行 `wb.Close(SaveChanges=False)` | **未保存的修改直接丢弃，不弹提示** |

**结论：跑 `run_all.ps1` 之前，先关掉所有 Excel 工作簿，并退出 SAP。**

如果这台机器同时还要给人用，考虑把定时任务安排在无人时段，或挪到专用机 / 专用 Windows 账号（独立账号有自己独立的 SAP 会话）。

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
| `00 - Maintenance_Plan_Adherence_Sync` | IW39 | 维护计划遵守率 | `MasterData` |
| `01 - Maintenance_Effectiveness_Sync` | IW47 | 维护有效性 | `MasterData` |
| `02 - Critical_A_ &_H_equipment_with_Maintenance_Plan_Sync` | IH08 → IP18 | A 级关键设备中有维护计划的比例 + 无计划设备清单 | `MasterData`、`无保养计划A类设备` |
| `03 - Safety_Stock_ZSE16` | ZSE16 | 安全库存全量 | `安全库存数据` |
| `04 - Total_Workorder` | IW39 | 全年工单总表 | `Total_Workorder` |
| `06 - IM_Equipment_Number` | IH08 | 设备编号（含 Tag 列） | `Equipment_Number_EAM` |
| `07 - Inventory_MB52` | MB52 | 库存 | `MasterData` |

`00`、`01`、`02` 写入**同一张** Google 表格的 `MasterData` 工作表，各自占不同的列区间。

`03` 和 `07` 写入**另一张**表格：`03` 用 `安全库存数据` 工作表，`07` 用 `MasterData` 工作表。

### 不在 `run_all.ps1` 中，需手动运行

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

> ⚠️ `05 - Stock_Turnover/main.py` 顶部有句注释写「请在使用前替换 USERNAME 和 PASSWORD」，**那句话已经过时**，忽略它。

---

## 常见故障

| 现象 | 原因 / 处理 |
|---|---|
| `错误: 无法连接到 SAP GUI Scripting Engine 或找不到活动会话` | SAP 没起来，或系统名/客户端号不对。检查 `sap_auto_logo()` 里的 `-system=` / `-client=` |
| `❌ Google Sheets 错误: 找不到工作簿` | 服务账号邮箱没被加为那个表格的编辑者，或 `pyreadsp-*.json` 不在仓库根目录 |
| `⚠️ 警告: ...没有数据，跳过上传（不算失败）` | **正常**。本月无数据，脚本正常结束、退出码 0 |
| `⚠️ 警告: 没有找到打开的 Excel 工作簿` | SAP 导出没成功弹出 Excel。看前一步的导出日志 |
| 卡在 `正在等待 Excel 打开...` 很久 | Excel 有弹窗挡着，或开着一堆工作簿 |
| `taskkill` 输出 `ERROR: The process "saplogon.exe" not found.` | **正常噪音**。机器上本来就没有 SAP 在跑，紧接着那句"关闭成功"是脚本的固定文案 |

---

## 已知限制

1. **独占 SAP 和 Excel**——见开头的警告。这是当前设计的前提，不是 bug。
2. **固定秒数等待**——脚本用 `time.sleep()` 等待 SAP 和 Excel 就绪（9 个 `main.py` 里共 91 处）。机器慢的时候可能等不够，快的时候白等。
3. **每个脚本都重启一次 SAP**——7 个项目各做一次"关掉 SAP → 重新启动"，每轮约 2.5 分钟花在这上面。
4. **公共代码重复**——`close_SAP()`、`sap_auto_logo()`、`get_resource_path()` 等函数在 12 个文件里各有一份，改一处要改十二处。
5. **连接 Google Sheets 时关闭了 TLS 校验**（`verify = False`），为兼容公司代理。属于可收紧项。

---

## 目录里还有什么

| 文件 | 说明 |
|---|---|
| `run_all.ps1` | 编排器：依次运行 7 个项目并统计成败 |
| `run_all.log` | 运行日志（含代码版本、每步耗时、每个脚本的退出码） |
| `*/batch_update.py` | **一次性历史回填脚本**（`00`/`01`/`02` 各一份），用硬编码的月份列表补跑历史数据，日常运行不会执行 |
| `08 - .../*.vbs`、`*/*.txt` | SAP GUI 录制脚本及其导出产物，历史遗留 |
| `*/.gitignore` | 各项目自己的忽略规则 |

---

## 维护

代码通过 GitHub 同步：本地改 → `git push` → 运行机 `git pull`。

提交信息格式：`VYYYY-MM-DD.XX_简短描述`（同一天内 `.XX` 递增）。
