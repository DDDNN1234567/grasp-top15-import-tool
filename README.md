# 管家婆辉煌 TOP15 双向直导工具

千业管理通（商贸通 Bas 库）↔ 管家婆辉煌 TOP15（标准账套，SQL Server）**双向 100% 直导**工具，覆盖基本资料、**16 方案完整预设价格（含客户自定义门店价）**、客户价格追踪、**当前结存库存**、条码、辅助单位（真实换算率）。

主数据导入走管家婆**官方存储过程通道**（`p_hh_BasicInsertB_Commodity` 单位 / `p_hh_BasicInsertK_Commodity` 仓库 / `p_hh_BascInsertP_Commodity` 商品 / `p_hh_BasicInsertE_Commodity` 职员 / `p_hh_ModInitGoodsStocks_SNMng_Package` 库存），由官方过程自动维护树编号、拼音、层级、子级计数与增量标记；价格、条码、辅助单位批量直写。

## 功能（三个通道 + 维护工具）

| 通道 | 数据 | 方式 |
|---|---|---|
| Excel 模板 | 6-sheet：单位/商品/仓库/职员/库存/价格追踪 | 官方存储过程逐行 |
| 正向直连 | 千业 Bas 库 → 管家婆：四表 + 16 方案价格 + 条码 + 辅助单位 + 价格追踪 + 当前结存 | 级联数字化 → 官方过程 + 明细直写 |
| **反向导出** | 管家婆 → 千业 Bas 库：八表回写，**主键 BaseNo 100% 还原** | 源BaseNo 溯源 + 显式列名批量插入 |
| 数据安全 | 导入/导出前自动备份，一键回滚 | 事务保证原子性 |
| 差异对比 | 源-目标核心表行数 + 商品 BaseNo 集合差异（增量发现，只读） | — |
| 缓存 | 清理客户端 localData，重登即见 | — |

## 核心能力

1. **16 方案完整价格**：自动补齐自定义价格类型（0104-0118，IsSystem=0），客户自定义门店价、现金价、预设价全部导入不漏；主表零售价（SalePriceBasic）优先于 PPM 槽8。
2. **真实换算率**：千业 `Percent*` = 辅助单位换算率（非毛利率），双向 URate 一致（个→箱 48/24/108 等）。
3. **当前结存口径**：库存取 `StorageGoodsInfo` 当前结存，不读历史期初表；负库存符号自动对齐（数量与金额同号）。
4. **主键 100% 还原**：反向按备注“源BaseNo:N”溯源，保留历史删除缺口，不压缩编号。
5. **通用设计**：不写死具体账套，CCC/CCC2 及任意含自定义价格列的账套均可处理。

## 快速开始

1. 环境：Windows + 管家婆辉煌 TOP15 + SQL Server 实例 + ODBC Driver 17 for SQL Server + Python 3.9+（pyodbc、xlrd、tkinter）。
2. 图形界面：`python gpj_top15_gui.py`（或打包版 exe）。四个页签：Excel 通道 / 千业→管家婆正向直连 / 管家婆→千业反向导出 / 维护工具。
3. 各页签填源/目标连接 → 测试连接 → 勾选模块 → 一键直导。耗时任务独立线程、独立连接、全局 busy 锁。
4. 重开管家婆查看“基本信息”。

## 命令行

```bat
:: 正向（千业 → 管家婆）
python gpj_top15_srcimport.py CCC unit,warehouse,goods,employee,stockini,price,barcode,assistunit,pricefollow

:: 反向（管家婆 → 千业 Bas 库）
python gpj_top15_srcexport.py qysrc_test
```

源库密码通过环境变量 `GPJ_SRC_PWD` 传入，代码不含任何明文凭据。

## 文件清单

| 文件 | 说明 |
|---|---|
| `gpj_top15_core.py` | 引擎：官方过程包装 / 备份回滚 / 父编号解析 / Excel 编排 |
| `gpj_top15_gui.py` | tkinter 四页签图形界面（v2.0） |
| `gpj_top15_srcimport.py` | 正向引擎：级联数字化 + 16 方案 + 真实换算率 + 当前结存 |
| `gpj_top15_srcexport.py` | 反向引擎：源BaseNo 溯源 + 八表回写 |
| `管家婆辉煌TOP15数据库词典.md` | 管家婆表逐列字典、官方过程签名与错误码 |
| `千业管理通-管家婆辉煌对应数据词典.md` | 双向字段映射与核验结论（v2.0） |
| `管家婆TOP15一键导入工具_使用说明.md` | 用户手册 |

## 打包为 exe

```bat
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "管家婆TOP15双向直导工具" ^
  --hidden-import tkinter --hidden-import _tkinter --hidden-import pyodbc --collect-submodules pyodbc ^
  --add-data "<python>\Lib\tcl8.6;lib\tcl8.6" --add-data "<python>\Lib\tk8.6;lib\tk8.6" ^
  gpj_top15_gui.py
```

## 验证结论（2026-11-02 双向端到端实跑，全绿）

**正向（CCC 634MB → HH15new）**：

| 模块 | 目标行数 | 结果 |
|---|---|---|
| 单位/仓库/职员/商品 | 883 / 9 / 3 / 24729 | 失败 0 |
| 预设价格（19 类型） | **136875** | 孤儿 0 |
| 条码 | 22992 | 失败 0 |
| 辅助单位 | 24765 | 真实 URate |
| 客户价格追踪 | **1601528**（400382×4） | — |
| 当前结存 | **6204** | 失败 0 |

读取链路 xw_GetBasicData：B=884/P=24730/K=16/E=4；非数字/长度非法/deleted/updatetag/拼音/悬空/孤儿**全部 0**。

**反向（HH15new → qysrc_test）权威核对 12 项全绿**：四主表+条码行数一致；商品 BaseNo 集合 100% 还原（含缺口）、全名一致；PPM 16 方案逐键逐值一致（88113）；价格跟踪 400382；结存数量逐值一致、金额 2 位精度一致（负库存 25013 符号修正单独记录）。

**系统级边界（非缺陷）**：管家婆只存非零价、金额统一 2 位小数（源 4 位四舍五入，367 条）、辅助单位价格由基本价×换算率派生、条码首尾空格清洗。

## 免责声明

本工具面向自有账套数据迁移场景，操作数据库前请务必备份；使用官方存储过程可最大限度保证与软件原生行为一致。
