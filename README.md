# 管家婆辉煌 TOP15 一键导入工具

将 **基本资料（单位/商品/仓库/职员）、预设价格、客户价格追踪、期初库存** 一键导入管家婆辉煌 TOP15 账套（SQL Server）的工具。

所有主数据导入均走管家婆**官方存储过程通道**（`p_hh_BasicInsertB_Commodity` 单位 / `p_hh_BasicInsertK_Commodity` 仓库 / `p_hh_BascInsertP_Commodity` 商品 / `p_hh_BasicInsertE_Commodity` 职员 / `p_hh_ModInitGoodsStocks_SNMng_Package` 期初库存），由官方过程自动维护树编号、拼音、层级、子级计数与增量标记（updatetag / xw_BaseUpdateTag），软件内"基本信息"立即可见。价格明细、条码、辅助单位按技能经验批量直写（UnitID 数字串、PRTypeId 标准字典）。

## 功能

| 通道 | 数据 | 方式 |
|---|---|---|
| Excel 模板 | 单位/商品/仓库/职员/期初库存/客户价格追踪 6-sheet | 官方存储过程逐行 |
| 源库直连 | 千业商贸通 Bas* 库（mdf/ldf 附加后）→ 四表 + 价格 + 条码 + 辅助单位 + 价格追踪 + 期初库存 | 级联数字化 → 官方过程 + 明细直写 |
| 数据安全 | 导入前自动备份 9 张表（`_impbak_<时间戳>`），一键回滚 | 事务保证原子性 |
| 缓存 | 导入完成自动清理客户端 localData，重登即见 | — |

## 快速开始

1. 环境：Windows + 管家婆辉煌 TOP15（`F:\GRASPⅡTOP15\`）+ SQL Server 实例 + ODBC Driver 17 for SQL Server + Python 3.9+（pyodbc、xlrd、tkinter）。
2. 图形界面：`python gpj_top15_gui.py`（也可用打包版 exe）。
3. 填服务器实例 / 账套库 / 账号密码 → 连接自检 → 选择 Excel 模板（或勾选"源库直连"）→ 勾选模块 → 一键导入。
4. 导入完成自动备份可回滚；重开管家婆查看"基本信息"。

Excel 模板规格见 `管家婆辉煌TOP15数据库词典.md` 第 6 节与 `管家婆TOP15一键导入模板.xlsx`（示例行自动跳过）。

## 源库直连（千业商贸通 → 管家婆）

1. 将千业软件数据文件 `CCC.mdf / CCC.ldf`（或 千业测试 / 测试）附加到本机 SQL Server（2008 Express 可附加 634MB 生产库）。
2. 命令行：`python gpj_top15_srcimport.py qysrc_ccc unit,warehouse,goods,employee,stockini,price,barcode,assistunit,pricefollow`
3. 引擎自动完成：Bas 四表级联数字化（保留层级、纯数字 5 位倍数树编号、原编号入备注）→ 官方过程导入 → 明细直写（预设价格 PRTypeId 0001/0002-0004/0005/0007/0103、条码 UnitID='0'、辅助单位、客户价格追踪 0005-0008 ×4 展开）→ updatetag 归位 → 备份可回滚。

> 源库密码通过环境变量 `GPJ_SRC_PWD` 传入，仓库代码不含任何明文凭据。

## 文件清单

| 文件 | 说明 |
|---|---|
| `gpj_top15_core.py` | 引擎：官方过程包装 / 备份回滚 / 父编号解析 / 导入编排 |
| `gpj_top15_gui.py` | tkinter 图形界面 |
| `gpj_top15_srcimport.py` | 源库直连引擎（Bas 级联数字化 + 明细直写） |
| `管家婆辉煌TOP15数据库词典.md` | 9 表逐列字典、13 种价格类型、6 个官方过程签名与错误码 |
| `千业管理通-管家婆辉煌对应数据词典.md` | 千业 Bas* 表 ↔ 管家婆 9 表**双向**字段映射与核验结论 |
| `管家婆TOP15一键导入工具_使用说明.md` | 用户手册 |
| `管家婆TOP15一键导入模板.xlsx` | 官方规格 6-sheet 模板 |

## 打包为 exe

```bat
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "管家婆TOP15一键导入工具" ^
  --hidden-import tkinter --hidden-import _tkinter --hidden-import pyodbc --collect-submodules pyodbc ^
  --add-data "<python>\Lib\tcl8.6;lib\tcl8.6" --add-data "<python>\Lib\tk8.6;lib\tk8.6" ^
  gpj_top15_gui.py
```

## 验证结论（2026-11-02 实库，CCC 生产库 634MB → HH15new 全量）

- **Excel 通道**：6-sheet 测试工作簿 12/12 行成功，0 失败；父类自定义编号自动解析为树编号；xw_GetBasicData B/P/K/E 显示链路验证通过。
- **源库直连（千业 CCC 生产库全量导入）**：

| 模块 | 源（qysrc_ccc） | 目标（HH15new） | 结果 |
|---|---|---|---|
| 单位 | BasUnitTypeInner 883 | btype | 882 成功 / 0 失败 |
| 仓库 | BasStockInner 9 | Stock | 8 成功 / 0 失败 |
| 职员 | BasPersonInner 3 | employee | 2 成功 / 0 失败 |
| 商品 | BasPtypeInner 24729 | ptype | **24728 成功 / 0 失败** |
| 预设价格 | 价格列 + BasPTypePriceMx | xw_P_PtypePrice | **72591 行** |
| 条码 | BasPTypeBarMx 22992 | xw_PtypeBarCode | 22992 行 |
| 辅助单位 | BasicUnit + AssistUnit | xw_PtypeUnit | 24765 行 |
| 客户价格追踪 | price 800743（去重 400382 组） | xw_P_PriceFollow | **1601528 行**（×4 展开） |
| 期初库存 | StorageIniGoods（Qty≠0 5150） | IniGoodsStocks | **5150 行** |

- **数据完整性全绿**：编号非数字 0 / 长度非法 0 / deleted 0 / updatetag<10000 0 / 拼音空 0；父子悬空 0；**明细孤儿引用全部 0**（价格/条码/辅助单位/价格跟踪/期初库存）。
- 关键修复：官方过程把传入编号存 `pusercode`、实际树编号另行分配 → 明细直写经 **pusercode 反查实际树编号**后引用 100% 匹配；条码预清洗防重复拒插；(PTypeId,PRTypeId) 全局去重防主键冲突；单位拼音 fn_GetPyCode 补齐。
- 导入前自动备份（`_impbak_20261102174546`）可一键回滚；客户端缓存已清，重登管家婆即可核对。

## 免责声明

本工具面向自有账套数据迁移场景，操作数据库前请务必备份；使用官方存储过程导入可最大限度保证与软件原生行为一致。
