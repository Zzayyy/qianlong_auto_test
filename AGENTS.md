# 项目协作说明

## 项目定位

这是一个面向 Windows 期权交易客户端的 GUI 自动化项目，覆盖查询、下单、撤单、组合申报和交易系统设置检查，并通过客户端档案兼容不同交易软件。

## 运行与验证

- 开发环境安装：`python -m pip install -r requirements.txt`
- 启动桌面工具：`python GUI自动化工具2/main.py`
- 执行查询驱动：`python 行情交易/查询/run_query.py`
- 运行测试（优先用系统 `python`，无需虚拟环境）：`python -m unittest discover -s tests -v`
  - 若使用虚拟环境，则把 `python` 换成对应解释器（如 `venv\Scripts\python.exe` 或 `.venv\Scripts\python.exe`）；本机当前为系统 Python 3.12（`python` 在 PATH 中），无 `.venv`。
- 真实 GUI 自动化依赖 Windows 和已登录的交易客户端；仅在控件访问/UIPI 证据明确时考虑提升整个 GUI 进程，输出目录 ACL 或路径错误不能用管理员权限代替诊断。单元测试不等同于真实客户端验证。
- **不要自动写测试、不要主动跑测试**：除非用户明确要求写测试或跑测试，否则不得擅自新建/修改测试文件，也不得执行 `python -m unittest` 或任何测试脚本。用户说到写测试/跑测试时再写/再跑。

## 技术栈

- Python、Tkinter、pywinauto、pywin32/uiautomation
- pandas、openpyxl、python-calamine、xlrd
- RapidOCR、ONNX Runtime、mss
- PyInstaller 用于桌面程序打包

## 目录与约定

- `GUI自动化工具2/`：桌面 GUI、任务调度和用户配置。
- `core/`：窗口定位、客户端档案、导出、比对和共用自动化逻辑。
- `行情交易/`：包含 `查询/`、`下单/`、`撤单/`、`组合申报/` 等行情交易业务脚本与模板。
- `超级策略/`：牛市认购、牛市认沽、熊市认购、熊市认沽、卖出跨式、卖宽跨式的一键开仓入口与组合申报；共享流程位于 `core/super_strategy.py`，自绘菜单 OCR 位于 `core/tactics_panel.py`。
- `交易系统设置/`：可从任意主界面打开的设置检查脚本。
- `交易系统设置/` 各面板的预期标准存于 `交易系统设置/标准/<client_id>/<panel>.json`；`标准/default/` 仅作未知客户端的通用兜底（如 `default/超价设置.json`）。7 客户端均已建标准目录（`qianlong/`、`guotai_haitong/`、`zhongtai/`、`huabao/`、`guangfa/`、`dongwu/`、`dongzheng/`；钱龙和国泰为原始预设，华宝/广发/东吴/东证从实机抓取，中泰从钱龙基线复制并按需补抓）；联储（`lianchu/`）目录已建但**暂时为空**——该客户端当前版本的交易系统设置组件不通用，暂不放标准，比对走 `default/` 兜底，待后续版本组件对齐后再补抓；粤开（`yuekai/`）目录同样未创建，比对走 `default/` 兜底。
- `交易系统设置/抓取自定义标准.py` 是排除批量执行的配置工具：只读采集客户端界面，写入 `标准/<client_id>/<panel>.json` 并在覆盖前生成 `.bak`（不删除）。`--panel` 可限定只采某面板；`--super-price` 额外采集超价参数弹窗（多页 SysListView32 滚动 + PrintWindow OCR）。
- 标准自定义边界：**一键炒单设置**仅国泰海通独有（钱龙、guotai_haitong、zhongtai、huabao、guangfa、dongwu 其余客户端都暂无该面板，标准 unsupported 默认不抓）。其快捷键表格属用户可自由自定义内容（增删改行/单元格），默认全量抓取或 `抓取自定义标准.py --client guotai_haitong --panel 一键炒单设置` 即以其当前表格为基线，也可直接手改 JSON；**超价设置**目前是各客户端预设默认标准（钱龙 9 行 / 国泰 24 行，买=1 卖=-1），尚不支持用户自由自定义——普通抓取不触碰 `超价设置.json`（它不在 PANEL_MODULES，是期权设置里的子弹窗），只有显式 `--super-price` 才以客户端当前弹窗覆盖（仍属预设，非用户自由表格）。
- `tests/`：不依赖真实客户端或带条件跳过的自动化测试。
- `clients.json` 是客户端差异的权威配置；查询项由 `queries.json` 驱动。
- **报价方式（下单面板下拉）按客户端配置**：`clients.json` 的 `quote_types`（可选 `quote_type_map`），由 `core/quote_type.py` 统一解析，三个下单脚本据此取白名单/索引。未配置的客户端沿用内置钱龙基准 9 项（行为与改动前一致）。**中泰证券期权宝XTP 后三项不同**：钱龙系为 `市价转限/市价FAK/市价F0K`，XTP 为 `全额成交或撤销/对手方最优转限价/市价剩余转撤销`（前 6 项相同）。该控件是「自绘可输入下拉」：`GetWindowText`/`WM_GETTEXT` 跨进程恒为空，**只有从主窗口出发的 UIA `get_value()` 能读到真实文本**（`connect(handle=edit)` 或 `connect(process=pid)` 取不到 ValuePattern），故校验优先精确文本、读不到才退回 `WM_GETTEXTLENGTH` 长度比较；**只比长度会误判**——`全额成交或撤销` 与 `市价剩余转撤销` 同为 7 字/14 字节。
- 新增脚本后同步更新 `GUI自动化工具2/config.py` 中的脚本清单；可复用逻辑优先放入 `core/`。
- 新增超级策略还必须同步入口脚本、`core/super_strategy.py` 目标白名单、`core/tactics_panel.py` 的 `SUPER_TARGETS`/`FORMAL_TARGET_CELLS` 以及对应测试；目标跨页时要显式增加滚动页定位，不能假设仍在顶部。
- 运行产物写入配置的输出目录；`交易系统设置_测试结果/`、日志、虚拟环境和构建目录不得提交。
- UI 组件树、截图和调试转储可能包含账户及本机窗口标题，新增或更新前必须脱敏，未脱敏内容不得提交。
- 文件与文本默认使用 UTF-8；保留现有中文业务命名，避免无关批量改名。

## 工作记忆

- AI 跨会话记忆入口：`.codebuddy/memory/MEMORY.md`（项目长期记忆）和 `.codebuddy/memory/YYYY-MM-DD.md`（每日事件，append-only）。
- 同步备份仅作为产品改名前的历史归档（`.workbuddy/`），不再刷新；详见 `.gitignore` 注释。
- **`.codebuddy/` 与 `.workbuddy/` 均为本机工作区私有目录（已被 `.gitignore` 排除），不会随仓库分发**。需要让其他协作者/新会话看到的信息，请写回 `README.md`、`AGENTS.md` 或 `clients.json`，不要把 `.codebuddy` 当共享真相源。

## 当前状态与下一步

- 当前分支包含多客户端支持（qianlong、guotai_haitong、zhongtai、huabao、guangfa、dongwu、dongzheng、lianchu、yuekai、zhongtai_xtp 十客户端；华宝 2026-08-04 新增，广发 2026-08-05 新增，东吴 2026-08-10 新增，东证期货 2026-08-12 新增，联储 2026-08-19 新增，粤开 2026-09-04 新增，中泰XTP 2026-09-17 新增，树拓扑/下单区/菜单项与钱龙不同，细节见 `clients.json` 各客户端 `description` 与 `native_tree_profile`）、统一查询驱动、组合申报、交易系统设置自动检查和状态映射。联储当前版本的交易系统设置组件不通用，`标准/lianchu/` 暂为空，比对走 `default/` 兜底。粤开 68 节点，四个下单脚本（期权下单/三键/四键/快速）与全选撤单的控件 auto_id 已与钱龙对齐，`标准/yuekai/` 未创建（走 `default/` 兜底）。
- **中泰证券期权宝XTP（`zhongtai_xtp`）**：与既有 `zhongtai`（`中泰证券期权宝(仿真行情)`，72/73 节点）是**不同客户端**，XTP 树仅 **26 节点 = 12 根 + 查询5 + 组合申报容器6 + 资金调拨3**。无 键盘下单/快速下单/普通下单/二键下单/单向下单，故只有**三个下单脚本**（期权下单/三键/四键）+ 全选撤单；查询容器仅 5 子，**无 通知查询/结算单/协议行权**；`资金调拨` 未接入（用户明确『资金调拨不用管』）。**关键点：根下同时存在叶子「组合申报」[4] 与容器「组合申报」[10]**——`core/native_tree.py::select_tree_path` 的 `need_child` 逻辑会跳过无子节点的叶子、优先命中容器，故 `\组合申报\组合申报` 正确落到 [10,0]（与其它客户端语义一致），叶子 [4] 不接入任何脚本；位置指纹同样指向 [10]，两条定位路径结论一致。`window_key` 与 `zhongtai` 相同（均为「中泰证券期权宝」），用户确认两者不会同时打开，故不做更精确区分；GUI 任务始终按 `GUI_CLIENT_ID` 精确路由，仅『脱离 GUI 直跑』的标题推断会因多命中而退化为纯文本定位（不会误选）。查询导出：资金持仓→数据输出弹窗（`export_type:null` 覆盖）、期权合约→数据输出弹窗（base 本无 `export_type`，无需覆盖）。三个下单脚本控件 auto_id 已实机核对与钱龙一致（四键页 18070=买平 / 18072=卖平）。**报价方式下拉项不同**：XTP 后三项为 `全额成交或撤销/对手方最优转限价/市价剩余转撤销`（钱龙系为 `市价转限/市价FAK/市价F0K`），已写入 `clients.json` 的 `quote_types`，由 `core/quote_type.py` 解析；Excel 模板按 XTP 新名字填写即可。`标准/zhongtai_xtp/` 未创建，走 `default/` 兜底。
- 粤开通知查询第 9 项为独有菜单 **风险通知查询**（其它 8 客户端无此菜单，已写入各自 `unsupported`）；其通知查询前 7 项与广发完全重叠，因此 `tests/test_client_profiles.py` 的跨客户端过滤断言需把 `yuekai` 与 `guangfa`/`lianchu` 一并列入跳过名单。
- 超级策略菜单采用 `PrintWindow` 小区域 RapidOCR：固定菜单格先做 recognition-only 精确校验，失败后才回退区域检测；固定坐标不得绕过文字校验直接点击。
- “加入标的/一键开仓”首次共用一次操作栏检测，后续缓存相对位置并逐按钮 OCR 复核；布局指纹或文字不符时必须回退完整检测，不能直接使用旧缓存。
- 启用“加入标的”时，客户端先返回一键开仓结果、再延迟返回加入标的结果；成功笔数允许不同。成功、失败、停市、未登录和未知弹窗都要用 Win32 逐个关闭，完整弹窗链静默后才返回或在运行日志报错。
- 交易系统设置的单独运行、任务中心和报告中心一键运行共用批次目录；全部任务结束或停止时生成一次总 TXT、Excel 和批次汇总 JSON。
- 报告中心直接接收本次完成批次的完整路径；启动、切换路径或手动刷新时只解析并显示最近 20 个有效批次，历史目录不会自动清理。
- “删除选中批次”会先校验批次目录，再将整个目录移入 Windows 回收站。
- GUI 脚本清单中“价格提醒设置”显示编号为 7，但文件仍名为 `6_价格提醒设置.py`；调整文件编号时必须同步配置引用。
- 执行导出或报告任务前要核对 `GUI自动化工具2/config.json` 的 `output_dirs` 是否属于当前账户且可写；该文件仍含机器相关配置，可移植默认路径尚未落地。
- 涉及点击、下单、撤单或修改交易设置的改动，除单元测试外还需在指定客户端进行受控人工验证。
- 超级策略最新 OCR 缓存和延迟双结果框已通过自动化回归；更新后的真实客户端端到端成功/失败链复验仍是待办，不得仅凭单元测试宣称 live verified。
