# 期权交易 GUI 自动化工具

面向 Windows 期权交易客户端的 GUI 自动化项目，覆盖**查询、下单、撤单、组合申报、超级策略一键开仓、交易系统设置自动检查**，并通过客户端档案兼容不同交易软件（钱龙、国泰海通、中泰、华宝）。

## 功能特性

- **行情交易**
  - 查询：资金持仓、策略持仓、当日委托/成交、历史委托/成交、期权合约、结算单等，由 `queries.json` 数据表统一驱动，支持按客户端覆盖导出参数。
  - 下单：期权下单 / 快速下单 / 三键 / 四键，Excel 模板驱动批量自动化下单（单 sheet，按「菜单」列区分）。
  - 撤单：全选撤单自动化。
  - 组合申报：组合申报 / 拆分申报全自动，Win32 控件原语实现。
- **超级策略**：牛市认购 / 牛市认沽 / 熊市认购 / 熊市认沽 / 卖出跨式 / 卖宽跨式一键开仓；组合申报支持「市场 × 策略」复选框多选，按持仓派生合约批量依次组合。
- **交易系统设置**：委托设置、期权设置、自动拆单设置、自动追单设置、快捷设置、价格提醒设置、一键炒单设置共 7 个面板的自动化检查；比对标准外置为 `标准/<client_id>/<panel>.json`，可用 `抓取自定义标准.py` 只读采集当前客户端标准（覆盖前自动备份 `.bak`）。
- **桌面 GUI 工具**（Tkinter）：任务中心、报告中心、定时调度、历史记录、参数配置，支持脚本拖拽运行；PyInstaller 打包后 exe 自带 Python 运行环境（`--_run_script` 脚本运行器模式）。

## 环境要求

- Windows（真实 GUI 自动化依赖 Windows 与已登录的交易客户端）
- Python 3.12
- 依赖安装：`python -m pip install -r requirements.txt`

## 快速开始

```powershell
# 1. 安装依赖
python -m pip install -r requirements.txt

# 2. 启动桌面工具
python GUI自动化工具2/main.py
```

> 部分客户端（`clients.json` 中 `requires_elevation: true`）需要以管理员身份运行，GUI 启动时会自动检测并以 UAC 重新拉起。

### 测试

```powershell
python -m unittest discover -s tests -v
```

测试不依赖真实客户端，可在无客户端环境运行。

## 目录结构

```
├── GUI自动化工具2/          桌面 GUI 与任务调度
│   ├── main.py            启动入口（含 PyInstaller 脚本运行器模式）
│   ├── config.py          应用环境、用户配置、脚本清单（SCRIPTS_CONFIG）
│   ├── config.json        用户配置（输出目录、客户端、导出格式等）
│   ├── gui/               主窗口、任务中心、报告中心、调度视图、回收站等
│   └── engine/            任务构建与执行、定时调度（task.py / runner.py / scheduler.py）
├── core/                  共用自动化逻辑
│   ├── window.py          窗口查找、激活、面板切换、倒计时
│   ├── native_tree.py     客户端功能树定位（文本定位 + 位置指纹兜底）
│   ├── clients.py         客户端档案解析、菜单路径归一化、别名
│   ├── runner.py          查询-输出-导出 通用执行入口
│   ├── export_dialog.py / save_as_dialog.py   导出 / 另存为对话框处理
│   ├── super_strategy.py  超级策略六类一键开仓 + 组合申报批量
│   ├── tactics_panel.py   自绘菜单 OCR（PrintWindow + RapidOCR）
│   ├── combination_order.py  组合/拆分申报 Win32 控件原语
│   ├── settings_standard.py / settings_window.py  交易系统设置标准读写与弹窗
│   ├── one_click_settings.py  一键炒单 OCR 解析与比对纯函数
│   └── workspace.py       工作区 / 输出目录管理
├── 行情交易/
│   ├── 查询/              run_query.py 通用查询驱动 + queries.json
│   ├── 下单/              期权下单等 Excel 驱动批量下单、模板
│   ├── 撤单/              全选撤单自动化
│   └── 组合申报/          组合/拆分申报全自动脚本
├── 超级策略/              六类一键开仓脚本 + 组合申报
├── 交易系统设置/           7 个面板脚本 + 标准/<client>/ JSON + 抓取自定义标准.py
├── tests/                自动化单元测试（不依赖真实客户端）
├── clients.json          客户端差异权威配置（窗口关键字、功能树指纹、菜单映射、不支持项）
└── requirements.txt
```

## 多客户端支持

- 现有客户端：`qianlong`（钱龙模拟期权宝，默认）、`guotai_haitong`（国泰海通证券期权宝）、`zhongtai`（中泰证券期权宝）、`huabao`（华宝证券期权宝，2026-08-04 新增）。
- 客户端差异统一配置在 `clients.json`：
  - `window_key`：主窗口关键字；
  - `native_tree_profile`：功能树节点数、位置指纹（Win11 文本被屏蔽时的兜底定位）；
  - `menu_map` / `menu_aliases`：菜单路径映射与别名归一化；
  - `script_redirects`：界面相同但脚本模型不同的重定向（如中泰快速下单改跑期权下单脚本）；
  - `unsupported`：该客户端不支持、需过滤的菜单项。
- 运行时以环境变量 `GUI_CLIENT_ID` 区分客户端，`core.window` / `core.runner` / 各业务脚本据此自动解析窗口与菜单路径。
- 交易系统设置标准按 `标准/<client_id>/` → `标准/default/` → 代码内嵌兜底 三级优先级加载。

## 关键机制

- **查询驱动**：`行情交易/查询/run_query.py` 由 `GUI_QUERY_KEY`（queries.json 的 key）指明查询项，`queries.json` 中 `client_overrides` 按客户端覆盖导出参数（值为 `null` 时删除基础参数），脚本本身不关心客户端。
- **菜单定位**：优先文本定位；Win11 屏蔽文本时回退位置指纹（校验节点数后按位置列表导航），位置不符 fail-closed。
- **超级策略 OCR**：`PrintWindow` 小区域 RapidOCR，固定菜单格先做 recognition-only 精确校验，失败才回退区域检测；固定坐标不得绕过文字校验直接点击。
- **任务执行**：任务中心/报告中心一键运行共用批次目录，任务全部结束或停止时生成一次总 TXT、Excel 和批次汇总 JSON。

## 打包

使用 PyInstaller（已在 `requirements.txt` 中），打包后 exe 自带 Python 运行环境与依赖，可通过 `main.exe --_run_script <脚本路径>` 在无 Python 环境下执行子脚本。

## 注意事项

- 涉及点击、下单、撤单或修改交易设置的改动，除单元测试外还需在指定客户端进行受控人工验证；单元测试不等同于真实客户端验证。
- 输出目录、日志、虚拟环境和构建目录不得提交；UI 组件树、截图和调试转储可能包含账户及本机窗口信息，提交前必须脱敏。
- 执行导出或报告任务前核对 `GUI自动化工具2/config.json` 的 `output_dirs` 是否属于当前账户且可写。
- 文件与文本默认 UTF-8；保留现有中文业务命名。
