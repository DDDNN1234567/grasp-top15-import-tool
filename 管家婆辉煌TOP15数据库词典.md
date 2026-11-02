# 管家婆辉煌 TOP15 数据库词典与表结构（HH15new 实库实测）

> 实测环境：`DESKTOP-8P5G8VQ\SQLSERVER`（SQL Server 2016 SP2 Express，13.0.5026.0）· 账套库 `HH15new` · 安装目录 `F:\GRASPⅡTOP15\`
> 账套状态：标准 V13/TOP15 账套（基本信息七表齐全），导入前为空账套（各表仅根节点）
> 本文数据来源：**实库 INFORMATION_SCHEMA 逐列导出 + 官方存储过程定义 + 官方 excel导入配置 XML + 公开资料对照**

---

## 0. 账套形态与表总览

| 软件功能 | TOP15 主表 | identity 列 | 根节点 | 实测行数(导入前→演练后) |
|---|---|---|---|---|
| 单位信息 | `btype` | brec | 00000「全部单位」 | 1 → 4 |
| 商品信息 | `ptype` | prec | 00000「全部商品」 | 1 → 4 |
| 仓库信息 | **`Stock`（注意：非旧版 C1Type）** | krec | 00000「全部仓库」 | 1 → 2 |
| 职员信息 | `employee` | erec | 00000「管理员」 | 1 → 2 |
| 期初库存 | `IniGoodsStocks` | GOODSORDERID | — | 0 → 2 |
| 单位价格跟踪 | `xw_P_PriceFollow` | 无 | — | 0 → 8 |
| 商品预设价格 | `xw_P_PtypePrice` | 无 | — | 0 → 10 |
| 商品辅助单位 | `xw_PtypeUnit` | 无 | 每商品基本单位 1 行 | 0 → 3 |
| 商品条码 | `xw_PtypeBarCode` | 无 | PTypeId/UnitID/BarCode | 0 → 2 |
| 增量标记 | `xw_BaseUpdateTag` | 无 | 每模块 1 行 | 15 行 |

**TOP15 关键差异（相对旧版 V13.32/商贸通）**：仓库主表为 `Stock`（`xw_GetBasicData 'K'` 分支直接读 `from stock k`），旧版文档中的 `C1Type` 在本库仅存根节点「全部」且不被仓库链路读取。导入接口一律走官方存储过程，无需手工维护编号/拼音/子级计数/updatetag。

---

## 1. 基本信息表列字典（实库逐列导出）

### 1.1 btype（单位信息，52 列）

| 列名 | 类型 | 说明 |
|---|---|---|
| btypeid | varchar(25) NOT NULL | 树编号（软件自动生成，5 位倍数级联，纯数字） |
| parid | varchar(50) | 父树编号（根=0） |
| leveal | smallint | 层级（0=根，1=一级，2=二级…） |
| soncount / bsonnum | int | 曾有儿子数 / 现有儿子数（客户端树展开依赖） |
| busercode | varchar(26) | **单位编号（用户自定义编号，Excel「单位编号」列）** |
| bfullname | varchar(200) | 单位全名（Excel「单位全名」列） |
| bname | varchar(38) | 单位简名 |
| Area | varchar(1000) | 单位地址 |
| TelAndAddress | varchar(66) | 电话与地址（旧字段） |
| PostCode / Person / TaxNumber | varchar(50/120/50) | 邮编 / 联系人 / 税号 |
| BankAndAcount / BankAccounts | varchar(50) | 开户行（旧）/ 银行账号 |
| bcomment | varchar(1000) | 备注 |
| ARTotal/APTotal/ARTotal00/APTotal00 | numeric(22,10) | 应收/应付发生额、期初应收/应付 |
| YRTotal/YPTotal/YRTotal00/YPTotal00 | numeric(22,10) | 预收/预付发生额、期初预收/预付 |
| R_WARNUP / P_WARNUP | numeric(22,10) | 应收款上限 / 应付款上限 |
| deleted | bit | 是否删除（导入必须 0） |
| bnamepy | varchar(66) | 全名拼音（官方过程自动生成） |
| brec | int identity | 记录号 |
| ParRec | int | 父记录号 |
| AreaTypeID | varchar(25) | 地区 |
| Fax / EMail / linker / linkerTel | varchar(50/100/50/60) | 传真 / 邮件 / 联系人 / 联系人电话 |
| SKQX / skqxType / skqxMonthDay | int | 结算期限(天) / 结算期限类型 / 月结日 |
| ROWINDEX | int | 同级排序 |
| ChangeDay / ChangeRate | int / numeric | 换货期限(天) / 换货比例(%) |
| isStop | int | 是否停用 |
| updatetag | int | 增量标记（官方过程自动维护） |
| DefInput / GatherBtypeId | varchar(25/50) | 默认经手人 / 结算单位 |
| UpFlag / province / city / region | int / varchar(50)*3 | 更新标志 / 省市区 |
| moPhone | varchar(60) | 手机 |

### 1.2 ptype（商品信息，57 列）

| 列名 | 类型 | 说明 |
|---|---|---|
| ptypeid | varchar(50) NOT NULL | 树编号（自动生成） |
| ParId | varchar(50) | 父树编号 |
| leveal | smallint | 层级 |
| psonnum / soncount | int | 现有儿子数 / 曾有儿子数 |
| CanModify | char(1) | 可修改标志 |
| pusercode | varchar(50) | **商品编号（用户自定义，Excel「商品编号」列）** |
| pfullname | varchar(256) | 商品全名 |
| pname | varchar(30) | 简名 |
| Standard / Type / Area | varchar(120/400/400) | 规格 / 型号 / 产地 |
| UsefulLifeMonth / UsefulLifeDay | smallint | 保质期(月)/(天) |
| pcomment | varchar(250) | 备注 |
| deleted | bit | 是否删除 |
| costmode | int | 成本算法（0=加权平均/4=先进先出等） |
| pnamepy | varchar(66) | 拼音（官方自动） |
| warnup / warndown | numeric(22,10) | 库存上限 / 下限 |
| prec | int identity | 记录号 |
| ParRec / RowIndex | int | 父记录号 / 同级排序 |
| isStop | int | 停用 |
| LastBuyPrice | numeric(22,10) | 最近进价 |
| updatetag | int | 增量标记 |
| Custom1-4 | varchar(25) | 自定义字段 |
| SideUnit | varchar(10) | 辅助计量单位串 |
| pgManCode / pgholunit / pgholunitrate | int/varchar(10)/numeric | 包装管理 / 整箱单位 / 整箱换算 |
| SNManCode | int | 序列号管理 |
| RepPtypeid | varchar(50) | 替代品 |
| Ptypetype | int | 商品属性（0=库存商品，非 0 不可入期初库存） |
| punitname | varchar(255) | 单位名称 |
| pptypeid | varchar(50) | 品牌 |
| KWManCode | int | 口味/批次管理（=1 时不可入期初库存） |
| isWeight | int | 计价方式 |
| UpFlag / MenGuid / Alid / Qqdid | int/varchar(50)/varchar(200)*2 | 更新标志 / 菜单 / 阿里 / 钉钉 |
| DefBtypeID | varchar(50) | 默认供应商 |
| PoolDiscount | numeric(22,10) | 联营扣率(%) |
| AllowModifyPrice | int | 前台议价 |
| DefaultTax | numeric(22,10) | 默认税率 |
| sourcewayzz/wg/ww (+Day) | int | 是否自制/外购/委外 及周期 |
| PJobManCode | int | 启用批次 |
| updateFlag / ossfilename | int / varchar(256) | 更新标志 / 图片 |

### 1.3 Stock（仓库信息，20 列）— TOP15 主表

| 列名 | 类型 | 说明 |
|---|---|---|
| ktypeid | varchar(25) NOT NULL | 树编号 |
| parid | varchar(50) | 父树编号 |
| leveal | int | 层级 |
| ksonnum / soncount | int | 现有儿子数 / 曾有儿子数 |
| kfullname | varchar(100) | 仓库全名 |
| kname | varchar(30) | 仓库简称 |
| kusercode | varchar(26) | **仓库编号（自定义，Excel「仓库编号」列）** |
| kcomment | varchar(256) | 备注 |
| deleted | bit | 删除 |
| knamepy | varchar(66) | 拼音 |
| krec | int identity | 记录号 |
| ParRec / ROWINDEX | int | 父记录号 / 排序 |
| isStop / updatetag | int | 停用 / 增量标记 |
| Address / Person / Tel | varchar(256/100/50) | 地址 / 负责人 / 电话 |
| SaveMenid / UpFlag / usesl / isMrpControl | int | 保存菜单 / 标志 / 使用货位 / 物料控制 |

### 1.4 employee（职员信息，27 列）

| 列名 | 类型 | 说明 |
|---|---|---|
| etypeid | varchar(25) NOT NULL | 树编号 |
| Parid / leveal | varchar(50)/smallint | 父树编号 / 层级 |
| soncount / esonnum | int | 曾有儿子数 / 现有儿子数 |
| efullname / ename | varchar(66)/varchar(30) | 职员全名 / 简名 |
| eusercode | varchar(26) | **职员编号（自定义）** |
| Department / Tel / Address | varchar(25/20/66) | 部门 / 电话 / 地址 |
| ecomment | varchar(256) | 备注 |
| deleted | bit | 删除 |
| enamepy | varchar(66) | 拼音 |
| erec | int identity | 记录号 |
| ParRec / Pay / Incumbency | int/numeric(22,10)/int | 父记录 / 工资 / 在职状态 |
| Birthday | varchar(10) | 生日 |
| EMail / ROWINDEX / TopTotal | varchar(50)/int/numeric | 邮件 / 排序 / 每单优惠限额 |
| isStop / updatetag / duty | int/int/varchar(100) | 停用 / 增量 / 职务 |
| LowLimitDiscount | numeric(22,10) | 最低折扣下限 |
| SaveMenid | UpFlag | int | 保存菜单 / 更新标志 |

### 1.5 C1Type（旧版仓库表，14 列）— 本库仅存根节点

`C1typeid/PARID/LEVEAL/C1sonnum/SONCOUNT/C1usercode/C1fullname/C1namepy/C1comment/DELETED/C1rec/PARREC/ROWINDEX/updatetag`。TOP15 仓库链路已切换至 `Stock`，C1Type 保留根节点「全部」但不参与仓库读取。

---

## 2. 业务明细表列字典

### 2.1 xw_P_PtypePrice（商品预设价格，6 列）

| 列名 | 类型 | 说明 |
|---|---|---|
| PTypeId | varchar(50) | 商品树编号 |
| UnitID | varchar(50) | 单位序号（基本单位=0，辅1=1，辅2=2） |
| PRTypeId | varchar(50) | 价格类型（见 §3 价格类型字典） |
| Price | numeric(22,10) | 价格 |
| UpDateDt | datetime | 更新日期 |
| KeyId | varchar(50) | GUID |

### 2.2 xw_P_PriceFollow（单位价格跟踪，10 列）

| 列名 | 类型 | 说明 |
|---|---|---|
| PTypeId / BTypeId | varchar(50) | 商品树编号 / 单位树编号 |
| UnitID | varchar(20) | 计量单位序号（0=基本单位） |
| Price | numeric(22,10) | 价格或折扣 |
| KTypeId | varchar(50) | 仓库（导入为''） |
| BillDt | datetime | 单据日期（导入为 1900-01-01） |
| PRTypeId | varchar(50) | 0005 最近进价 / 0006 最近进货折扣 / 0007 最近售价 / 0008 最近销售折扣 |
| PPType | int | 单据类型标记（导入=6） |
| KeyId / JGTypeId / MDTypeId | varchar(50) | GUID / 价格组 / 门店 |

### 2.3 IniGoodsStocks（期初库存，13 列）

| 列名 | 类型 | 说明 |
|---|---|---|
| PtypeId | varchar(50) | 商品树编号 |
| KtypeId | varchar(25) | 仓库树编号 |
| JobNumber / OutFactoryDate | varchar(50)/varchar(10) | 批号 / 生产日期 |
| Qty / Price / Total | numeric(22,10) | 数量 / 单价(=金额÷数量，官方自动) / 金额 |
| GoodsOrder / GOODSORDERID | int / identity | 行序 / 记录号 |
| pgholInqty | int | 整箱数量 |
| GoodsBatchID | varchar(50) | 批次号 |
| BtypeID | varchar(50) | 往来单位 |
| BuyDate / UsefulEndDate | varchar(10) | 进货日期 / 有效期至 |

### 2.4 xw_PtypeUnit（辅助单位，6 列）

`PTypeId / Unit1(单位名) / URate(换算率) / IsBase(1=基本) / Comment / Ordid(0=基本,1=辅1,2=辅2) / keyid(GUID)`

### 2.5 xw_PtypeBarCode（条码，6 列）

`PTypeId / UnitID(数字串'0'/'1'/'2'，禁止单位名) / BarCode / Ordid / keyid / kwtypeid`

---

## 3. 价格类型字典（xw_P_PriceType 实库 13 种）

| PRTypeId | 名称 | 备注 |
|---|---|---|
| 0001 | 零售价 | 预设 |
| 0002 | 预设售价1 | 预设 |
| 0003 | 预设售价2 | 预设 |
| 0004 | 预设售价3 | 预设 |
| 0005 | 最近进价 | 跟踪 |
| 0006 | 最近进货折扣 | 跟踪 |
| 0007 | 最近售价 | 跟踪 |
| 0008 | 最近销售折扣 | 跟踪 |
| 0013 | 配送价 | 预设 |
| 0101 | 最近折前进价 | 跟踪 |
| 0102 | 最近折后进价 | 预设（商品导入 @BuyDiscountPrice） |
| 0103 | 最低售价 | 预设 |
| 0014 | 库存均价 | 计算值 |

---

## 4. 官方导入存储过程接口（授权导入通道）

> 来源：`F:\GRASPⅡTOP15\excel导入配置\配置文件(XMl_XLS)\*.xml`（官方 Excel 导入控件 xwImport.ocx 的映射配置）+ HH15new `sys.procedures` 实读。**直接调用官方过程 = 编号/拼音/层级/子级计数/updatetag 全部由官方维护，无需直写七表。**

### 4.1 调用规范速查

| 模块 | 存储过程 | 关键参数 | 输出 |
|---|---|---|---|
| 单位 | `p_hh_BasicInsertB_Commodity` | @Parid=**父树编号**(00000/00001…)、@FullName=单位编号(busercode)、@UserCode=单位全名、@Name=简名、@prtypeid=适用价格(中文：零售价/预设售价1/2/3/配送价/最低售价) | @RltTypeID=新树编号、@errorValue |
| 仓库 | `p_hh_BasicInsertK_Commodity` | @parid=父树编号、@fullname=仓库编号、@usercode=仓库全名、@name=简称 | @rlttypeid、@errorValue |
| 商品 | `p_hh_BascInsertP_Commodity` | @Parid=父树编号、@FullName=商品编号、@UserCode=商品全名、@Name=简名、@unit=基本单位、@UCode=条码、@SalePrice=零售价、@PrePrice1-3=预设售价1-3、@SendPrice=配送价、@LowSalePrice=最低售价、@BuyDiscountPrice=最近折后进价、@unit1/2+@unitRate1/2+@UCode1/2=辅助单位、@defaulttax=税率、@allowprice=前台议价、@pjobmancode=启用批次 | @RltTypeID、@errorValue |
| 职员 | `p_hh_BasicInsertE_Commodity` | @parid=父树编号、@fullname=职员编号、@usercode=职员全名、@duty=职务、@LowLimitDiscount=折扣下限、@TopTotal=优惠限额 | @rlttypeid、@errorValue |
| 期初库存 | `p_hh_ModInitGoodsStocks_SNMng_Package` | @szptypeid=**商品编号(pusercode)**、@szPUserCode=商品全名、@szktypeid=**仓库编号(kusercode)**、@szKUserCode=仓库全名、@dQty=数量、@dTotal=金额、@JobNumber=批号、@CreateDate=生产日期、@UserEndDate=有效期至 | @OrderID、@errorValue |
| 客户价格追踪 | `p_hh_ImportPtypePriceGZ` | @szbtypeid=**单位编号(busercode)**、@szbfullname=单位全名、@szptypeid=**商品编号(pusercode)**、@szpfullname=商品全名、@punit=计量单位、@buyPrice=最近进价、@salePrice=最近售价、@buyDiscount=进货折扣(1-10)、@saleDisCount=销售折扣(1-10) | @OrderID、@errorValue |

**要点**：
- 基本资料四过程 @Parid 传**树编号**（根=00000）；期初库存与价格跟踪传**自定义编号**（内部自动转树编号）。
- 期初库存要求商品 `Ptypetype=0`（库存商品）且 `KWManCode<>1`（非口味管理），`Qty<>0`；开账后（sysdata iniover=1）不可导期初应收应付。
- 价格跟踪同一 单位×商品×计量单位 存在时**先删后插**（幂等）；折扣缺省 1.0、价格缺省 0。
- 商品价格仅当数值>0 时写入 xw_P_PtypePrice；辅助单位必须有基本单位。
- 返回码 <0 失败，@errorValue 携带中文原因；单位过程 0 成功。

### 4.2 单位过程错误码

| 返回码 | 含义 |
|---|---|
| -1 / -2 | 不能同时输入期初应收+应付 / 预收+预付 |
| -3 | 父单位已停用 |
| -4 | 单位全名为空 |
| -5 | 分层数达上限 |
| -10/-11 | 期初应收应付为负 / 期初预收预付为负 |
| -13 | 系统已开账不能录入期初 |
| -14 | 期初金额过大 |
| -15 | 换货比例不在 0-100 |
| -17 | 没有找到父类 |
| -18 | 已存在相同信息 |
| -19 | 空编号已存在 |
| -20 | 单位编号为空 |

---

## 5. 官方 Excel 模板规格（食品.xls 六 sheet，与存储过程参数一一对应）

| Sheet | 列（表头） |
|---|---|
| 商品 | 父类编号、商品编号、商品全名、简名、型号、规格、产地、成本算法、商品属性、联营扣率(%)、默认供应商(编号)、品牌(编号)、计价方式、有效期(天)、备注、基本单位、条码、零售价、预设售价1、预设售价2、预设售价3、配送价、最近折后进价、最低售价、辅助单位1、辅1换算关系、辅1条码、辅1零售价、辅1预设售价1-3、辅助单位2、辅2换算关系、辅2条码、辅2零售价、辅2预设售价1-3、口味管理、默认税率、前台议价、启用批次、是否自制、自制周期、是否外购、外购周期（46 列） |
| 仓库 | 父类编号、仓库编号、仓库全名、仓库简称、备注、仓库地址、仓库负责人、负责人电话（8 列） |
| 单位 | 父类编号、单位编号、单位全名、单位简名、单位地址、单位电话、手机一、手机二、传真、邮编、联系人一、联系人二、开户银行、银行账号、税号、期初应收款、期初应付款、期初预收款、期初预付款、应收款上限、应付款上限、结算期限(天)、备注、换货期限(天)、换货比例(%)、电子邮件、默认经手人(编号)、地区(编号)、结算单位(编号)、适用价格（30 列） |
| 职员 | 父类编号、职员编号、职员全名、生日、职员电话、职员职务、职员地址、最低折扣下限、每单优惠限额、备注（10 列） |
| 库存信息 | 商品编号、商品全名、仓库编号、仓库全名、数量、金额、批号、生产日期、有效期至（9 列） |
| 单位价格跟踪 | 单位编号、单位全名、商品编号、商品全名、计量单位、最近进价、最近售价、最近进货折扣、最近销售折扣（9 列） |

---

## 6. 读取链路与增量机制

- 客户端基本信息读取：`EXEC dbo.xw_GetBasicData @cMode, @nUpdateTag, @EtypeID, @beginRec, @endRec`
  - @cMode：B=单位 / P=商品 / K=仓库(stock) / E=职员 / D=部门 / A=地区
  - **@EtypeID 必须传 `''`**（传 NULL 返回 0 行假象）；@nUpdateTag=0 全量，>0 增量。
  - 过程可能多结果集，取数需遍历 nextset。
- 增量标记：主表 updatetag + `xw_BaseUpdateTag` 每模块一行，官方导入过程**自动维护**（实测 BTYPE 1→4、PTYPE 2→5、STOCK/EMPLOYEE→2，与新增行 updatetag 同步），客户端登录即增量拉取。
- 客户端缓存：`F:\GRASPⅡTOP15\Data\localData\<账套>-00000-*`，改库后清缓存重登更稳妥。

---

## 7. 公开资料表名对照（联网核实）

| 公开资料表名 | 对应实库/功能 | 来源 |
|---|---|---|
| employee | 职员信息表 ✅ | CSDN《管家婆数据库表.doc》 |
| Ptype | 库存商品信息表 ✅ | 同上 |
| btype | 往来单位 ✅ | 同上 |
| Stock | 仓库信息表 ✅（TOP15 主表） | 同上 |
| IniGoodsStocks | 期初商品库存分布表 ✅ | 原创力《SQL数据库的管家婆出错分析》 |
| GoodsStocks | 商品库存分布表（本库存在） | 同上 |
| Price | 客户跟踪价格表（旧版，TOP15 为 xw_P_PriceFollow） | CSDN《管家婆数据库表.doc》 |
| Dlyndx/Dlybuy/BakDlyOrder | 单据索引/进货单明细等业务单表（本库存在） | 豆丁《管家婆辉煌版数据表结构》 |

> 公开资料多为辉煌 9.x/II 旧版词典，与新库列名存在差异（如旧版 `Price` 表 → TOP15 `xw_P_PriceFollow`；旧版字段 `TelAndAddress` → TOP15 拆分 Address/Tel）。**以本文实库词典为准**。

---

## 8. 一键导入工具设计基线（v1.0 方案）

- 通道：官方 6 存储过程逐行调用（autocommit=True 连接，防静默回滚）
- 模板：内置生成官方规格 6-sheet Excel（.xlsx），支持 .xls/.xlsx 读取（xlrd/openpyxl）
- 导入顺序：单位 → 仓库 → 商品 → 职员 → 期初库存 → 客户价格追踪（依赖前置存在）
- 交互：连接配置 → 选 Excel → 预览/列校验 → 一键导入（实时进度+失败行定位+汇总）→ 清缓存/验证
- 额外：导入前 9 表备份（SELECT INTO 可回滚，排除 identity 列+显式事务）、父类编号自定义编号自动解析为树编号、模板示例行自动跳过

---

## 9. 商品过程错误码与补充发现（v1.0 实跑核对）

### 9.1 p_hh_BascInsertP_Commodity 关键错误码（源码核对）

| 返回码 | 含义 |
|---|---|
| -3 | 成本算法不在合法集（默认置移动加权平均） |
| -4/-5 | 父类不存在/未停用校验失败 |
| -14/-17 | 统一成本算法不一致 / 父类属性不一致（error15：商品的属性与父类不是同一属性） |
| -18/-19 | 父类商品存在替换品不能分类 |
| -20/-21/-22 | 编号为空/重复/条码类冲突（error20：模板中商品条码在数据库中已经被使用） |
| -26 | 条码为空 |
| -27 | 商品条码与组合套件商品条码重复 |
| -28/-29 | 辅助单位1/2 录入信息有误（@unitN<>'' 但换算关系为空或非数字） |
| -30 | 基本单位为空，不能有辅助单位（@unit='' 且 @unit1/@unit2 非空） |
| -31/-32 | 父类为口味管理商品不能分类 / 父类为总部商品专用类不允许导入 |

> 排错要点：error30 常见根因是 Excel 中“基本单位”列与“辅助单位”列错位（列数不足时右移），或基本单位列留空却填了辅助单位。

### 9.2 辅助函数与单位字典

- `dbo.fn_CharToBit(@szVarchar, @CutLength)`：仅做**字节截断**（中文按 2 字节计），不查字典不转码。商品过程用它截断单位名/条码/全名等字段长度。
- 计量单位字典：`xw_UnitGroupM`（单位组主表）+ `xw_UnitGroupD`（单位组明细）；基本单位按“单位名文本”传入商品过程（如“件”），无需预查字典。

### 9.3 v1.0 端到端实跑结论（2026-11-02 HH15new）

- 构造含父子层级的 6-sheet 测试工作簿：单位（分类+2 客户）、仓库 1、商品（分类+2 商品含零售/预设 1-3/配送/折后进价/最低/辅助单位/条码）、职员 1、期初库存 2、价格跟踪 2。
- 一键导入 **12/12 行全绿**，0 失败；btype 4 / ptype 4 / Stock 2 / employee 2 / xw_P_PtypePrice 15 / xw_P_PriceFollow 8 / IniGoodsStocks 2 / xw_PtypeUnit 4 / xw_PtypeBarCode 3。
- 树编号级联正确（00006=分类、0000600001=客户A）；父类编号填自定义编号（FL01）自动解析为树编号；官方过程自动维护编号/拼音/层级/子级计数/updatetag/xw_BaseUpdateTag。
- xw_GetBasicData B=5/P=5/K=9/E=3 读取链路正常。
- 备份→回滚验证通过（rollback 须排除 identity 列 + BEGIN TRAN 原子性）。

---

## 10. 双向直导（v2.0，千业 Bas* ↔ HH15new）

### 10.1 源库与附加

- 千业软件数据目录 `E:\QySoft\gtSystemTYV86PRO\WorkData\`：CCC.mdf / ccc2.mdf（各 634MB）等。
- mdf/ldf 对普通进程拒绝读（透明加密），**SQL Server Express 附加成功**后经 SQL 通道读取。
- CCC 与 CCC2 实测为同一份库的两个副本（214 表/列定义/校验和/文件大小全一致）；工具按通用 16 方案设计，不写死账套。

### 10.2 引用基线

- Bas 四表 BaseNo int 主键，引用全部可 JOIN BaseNo（100%）；千业库无标准 7 表，属 Bas 旧版商贸通。

### 10.3 16 方案完整价格（v2.0 关键升级）

- `ensure_custom_pricetypes` 自动补齐自定义价格类型（0104-0118，IsSystem=0/IsVisible=1），`F_GetPriceInfo` 动态读取——**客户自定义门店价、现金价、预设价全部导入，不再跳过**。
- 千业 IndexID 1-16 → 管家婆 PRTypeId：1-7/9-16 → 0104-0112/0111-0118，8 → 0001（系统零售价）；主表 RecentPrice/RecentPrice1 → 0005/0007。
- 零售价 0001 双来源：主表 SalePriceBasic 优先于 PPM 槽8；(PTypeId,PRTypeId) 全局去重，DELETE 全表幂等。

### 10.4 真实换算率（Percent 语义修正）

- `PercentBasic/One/Two/Three` 是**辅助单位换算率（UnitRate），不是毛利率**（UpdateDB.sql 行 52447-52449、53330；官方过程 sp_p_hh_BascInsertP 行 404/434）。
- 正向写 xw_PtypeUnit：基本行 IsBase=1/URate=1，辅助行 IsBase=0/URate=Percent 真实值（12-108，如箱=108）。

### 10.5 当前结存与负库存

- 库存取 `StorageGoodsInfo` **当前结存**（非历史期初 StorageIniGoods），写 IniGoodsStocks。
- 负库存符号自动对齐：数量与金额异号时以数量符号为准、按 |单价| 重算金额（商品 25013：+219→-219）。

### 10.6 反向导出与主键还原

- 管家婆 → 千业 Bas 八表回写；`rebase_map` 优先解析备注“源BaseNo:N”，**BaseNo 集合 100% 还原（含历史删除缺口）**；INSERT 显式列名，普通 executemany。

### 10.7 实库验证（2026-11-02 双向端到端，全绿）

| 方向 | 结果 |
|---|---|
| 正向 CCC→HH15new | btype 883/ptype 24729/Stock 9/employee 3；xw_P_PtypePrice **136875**；xw_P_PriceFollow **1601528**；条码 22992；辅助单位 24765；当前结存 **6204**，失败 0 |
| 反向 HH15new→qysrc_test | 四主表+条码行数一致；BaseNo 集合 100% 还原、全名一致；PPM 88113 逐值一致；price 400382；结存数量逐值一致、金额 2 位精度一致；权威核对 **12 项全绿** |

系统级边界（非缺陷）：管家婆只存非零价、金额统一 2 位小数（源 4 位四舍五入，367 条）、辅助单位价格由基本价×换算率派生、条码首尾空格清洗。发布 GitHub（DDDNN1234567/grasp-top15-import-tool）。
