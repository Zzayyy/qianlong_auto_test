# 期权交易客户端 GUI 自动化

面向 Windows 期权交易客户端的 GUI 自动化项目，覆盖**查询、下单、撤单、组合申报、超级策略一键开仓、交易系统设置自动检查**，并通过客户端档案兼容不同交易软件（钱龙、国泰海通、中泰、华宝、广发、东吴）。

## 快速开始

```powershell
# 安装依赖（系统 Python 3.12）
python -m pip install -r requirements.txt

# 启动桌面工具
python GUI自动化工具2/main.py
```

详细约定见 [`AGENTS.md`](./AGENTS.md)。

## 技术栈

- Python 3.12、Tkinter（桌面 GUI）
- pywinauto / pywin32 / uiautomation
- pandas、openpyxl、python-calamine、xlrd
- RapidOCR + ONNX Runtime、mss
- PyInstaller

## 目录结构

| 目录 | 用途 |
|---|---|
| `GUI自动化工具2/` | 桌面 GUI、任务调度、用户配置 |
| `core/` | 窗口定位、客户端档案、导出、比对、共用自动化 |
| `行情交易/` | 查询、下单、撤单、组合申报 等业务脚本与模板 |
| `超级策略/` | 牛市认购/认沽、熊市认购/认沽、卖出跨式/宽跨式 一键开仓 |
| `交易系统设置/` | 设置自动检查脚本 + `标准/<client_id>/` 期望值 |
| `tests/` | 不依赖真实客户端或带条件跳过的自动化测试 |

## 多客户端支持

通过 `clients.json` 驱动 `core/clients.py::normalize_menu_path_for_client` 与 `core/window.py::switch_panel` 完成菜单别名/重定向。客户端档案的关键差异：

| 客户端 | 菜单数量 | 备注 |
|---|---|---|
| 钱龙 | 70 | 基线 |
| 国泰海通 | 70 | `一键炒单设置` 独有面板 |
| 中泰 | 72/73 | 双层组合申报；|
| 华宝 | 52 | 拓扑不同（24 根 + 组合 8 + 查询 16 + 通知 4） |
| 广发 | 57 | 通知查询有 6 项专属； |
| 东吴 | 48 | 三键/四键下单嵌在「其他下单方式」容器； |

客户端差异的完整位置指纹与 unsupported 清单见 `clients.json` 各客户端条目；新增客户端或大改客户端时需同步 `clients.json`、`queries.json` 与 `tests/test_client_profiles.py`。

## 交易系统设置 —— 标准目录

期望值外置于 `交易系统设置/标准/<client_id>/<panel>.json`，由 `core/settings_standard.py::load_standard` 加载；优先级：

1. `标准/<client>/<panel>.json`
2. `标准/default/<panel>.json`（兜底）
3. 代码内嵌默认值

**抓取工具**：`交易系统设置/抓取自定义标准.py --client <id> --panel <panel>`（覆盖前自动生成 `.bak`）。

**客户端覆盖**：
- 钱龙、guotai_haitong：预设完整
- 中泰、huabao、guangfa、dongwu：已抓取至 `标准/`
- 默认目录 `标准/default/`：7 面板兜底

**自定义边界**：一键炒单快捷键表格属用户可自定义；超价设置是各客户端预设默认，普通抓取不触碰，限定 `--super-price` 才覆盖。

**面板分布**：
- 钱龙/华宝/广发/东吴：7 面板 + 超价
- 国泰海通：7 面板 + 超价（且有独立一键炒单）
- 中泰：6 面板 + 超价（无一键炒单，因为客户端确实没有该面板）

## 超级策略

- 菜单扫描：`core/tactics_panel.py::PrintWindow` 小区域 RapidOCR，先固定菜单格精确识别，失败再回退区域检测；**固定坐标不得绕过文字校验直接点击**
- 一键开仓/加入标的首次共用一次操作栏检测，结果写入缓存；布局指纹或文字不符必须回退完整检测
- 启用「加入标的」时客户端先返回一键开仓结果、再延迟返回加入标的结果；成功笔数允许不同
- 成功/失败/停市/未登录/未知弹窗都用 Win32 逐个关闭，整条弹窗链静默后才返回
- 流程共享在 `core/super_strategy.py`；目标白名单与 `FORMAL_TARGET_CELLS` 在该文件统一维护

## 已知限制与待实机清单

- 真实 GUI 自动化依赖 Windows + 已登录的交易客户端；本仓库不保证其他平台
- 涉及点击、下单、撤单、修改设置的改动**单元测试外还需实机验证**
- 超级策略最新 OCR 缓存与延迟双结果框已通过自动化回归，真实客户端端到端成功/失败链复验仍待办
- `GUI自动化工具2/config.json` 含机器相关路径，可移植默认路径尚未落地；执行导出或报告前核对 `output_dirs`
- 客户端：新版本或客户端大改后，相关 `标准/*.json` 可能需要重新抓取
