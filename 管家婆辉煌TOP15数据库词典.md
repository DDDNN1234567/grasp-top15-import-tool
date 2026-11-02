# 管家婆辉煌 TOP15 数据库词典

> 版本：v1.0（2026-11-02 实库核验版，账套 HH15new / DESKTOP-8P5G8VQ\SQLSERVER）
> 依据：管家婆辉煌 TOP15（`F:\GRASPⅡTOP15\`）账套库逐列实测 + 官方存储过程导出源码

---

## 一、核心表清单（9 表）

| 表 | 用途 | 说明 |
|---|---|---|
| btype | 单位（往来单位/供应商/客户） | 根 00000"全部单位"；busercode=自定义编号 |
| ptype | 商品 | 根 00000"全部商品"；pusercode=自定义编号 |
| Stock | 仓库 | 根 00000"全部仓库"；kusercode=自定义编号 |
| employee | 职员 | 根 00000"管理员"；eusercode=自定义编号 |
| IniGoodsStocks | 期初库存 | GOODSORDERID identity；商品×仓库 |
| xw_P_PriceFollow | 单位价格跟踪 | PK=(PTypeId,UnitID,PRTypeId,BTypeId,KTypeId) |
| xw_P_PtypePrice | 商品预设价格 | PK=(PTypeId,UnitID,PRTypeId)；13 种标准类型 |
| xw_PtypeUnit | 商品辅助单位 | 每商品基本单位 1 行 + 辅单位 |
| xw_PtypeBarCode | 商品条码 | UnitID 必须数字串'0'/'1'/'2' |

> 注：管家婆 2016 Express（HH15new）仓库主表名为 `Stock`（旧版 V13.32 为 `C1Type`），导入/对导以实测列名为准。

---

## 二、btype（单位，51 列）

| 列 | 类型 | 说明 |
|---|---|---|
| btypeid | varchar | **树编号**（5 位倍数纯数字，父子=前缀；客户端 flPtypeid=LEFT(编号,LEN-5)） |
| parid | varchar | 父编号 |
| leveal | smallint | 层级（0=根，1/2/3…） |
| soncount / bsonnum | int | 子级数（客户端树展开标记，必须真实值） |
| busercode | varchar | **自定义编号**（Excel 模板/源库直连填此列） |
| bfullname / bname | varchar | 全名 / 简名 |
| Area / TelAndAddress | varchar | 地区 / 电话地址 |
| PostCode / Fax / EMail | varchar | 邮编 / 传真 / 邮箱 |
| Person / linker / linkerTel / moPhone | varchar | 联系人一 / 联系人二 / 联系人电话 / 手机 |
| TaxNumber / BankAndAcount / BankAccounts | varchar | 税号 / 开户银行 / 银行账号 |
| bcomment | varchar | 备注（源库导入写"来源:Bas;原编号:xxx"） |
| ARTotal / APTotal / ARTotal00 / APTotal00 | numeric | 应收/应付（当前/期初） |
| YRTotal / YPTotal / YRTotal00 / YPTotal00 | numeric | 预收/预付（当前/期初） |
| ARTotal00Year / APTotal00Year | numeric | 期初应收/应付（年度） |
| R_WARNUP / P_WARNUP | numeric | 应收/应付上限 |
| SKQX / ChangeDay / ChangeRate | int/numeric | 结算期限(天) / 换货期限(天) / 换货比例(%) |
| DefInput | varchar | 适用价格 |
| GatherBtypeId | varchar | 结算单位 |
| deleted | smallint | 必须 0（新版读取链 fn_* 过滤 deleted=1） |
| bnamepy | varchar | 拼音（源空则归位阶段用 fn_GetPyCode(bfullname) 补齐） |
| updatetag | int | 增量标记：>0 且推荐统一 10000 |
| brec / ParRec / ROWINDEX | int | 自增记录号 / 父记录号 / 行序 |
| AreaTypeID / skqxType / skqxMonthDay / province / city / region | int/varchar | 地区类型 / 结算期限类型 / 省市县 |
| isStop / UpFlag | int | 停用 / 更新标志 |

---

## 三、ptype（商品，55 列）

| 列 | 类型 | 说明 |
|---|---|---|
| ptypeid | varchar | 树编号（父子=前缀） |
| ParId | varchar | 父编号 |
| leveal / psonnum / soncount | smallint/int | 层级 / 子级数 |
| pusercode | varchar | 自定义编号 |
| pfullname / pname | varchar | 全名 / 简名 |
| Standard / Type / Area | varchar | 规格 / 型号 / 产地 |
| UsefulLifeMonth / UsefulLifeDay | int | 有效期（月/天） |
| pcomment | varchar | 备注 |
| costmode | varchar | 成本算法 |
| pnamepy | varchar | 拼音（官方过程自动生成） |
| warnup / warndown | numeric | 库存上下限 |
| prec / ParRec / RowIndex | int | 自增记录号 / 父记录号 / 行序 |
| LastBuyPrice | numeric | 最近进价 |
| SideUnit / punitname / pgholunit / pgholunitrate | varchar/numeric | 辅助单位 / 基本单位名 / 整箱单位 / 整箱换算 |
| DefaultTax | numeric | 默认税率 |
| Ptypetype | varchar | 商品属性 |
| PoolDiscount | numeric | 联营扣率(%) |
| CanModify / AllowModifyPrice / isWeight / isStop / UpFlag | int | 可否修改/前台议价/称重/停用/更新 |
| Custom1-4 / MenGuid / Alid / Qqdid | varchar | 自定义字段 |
| DefBtypeID | varchar | 默认供应商 |
| RepPtypeid / pptypeid / pgManCode / SNManCode / KWManCode / PJobManCode | varchar | 替代商品/商品类型/经手人/序列号/口味/工序 |
| sourcewayzz / sourcewaywg / sourcewayww | varchar | 自制/外购/委外 |
| deleted | smallint | 必须 0 |
| updatetag | int | 增量标记：10000 |
| updateFlag / ossfilename | int/varchar | 更新标志 / 图片 |

---

## 四、Stock（仓库，23 列）

| 列 | 类型 | 说明 |
|---|---|---|
| ktypeid | varchar | 树编号 |
| parid / leveal / ksonnum / soncount | varchar/smallint/int | 父编号 / 层级 / 子级数 |
| kfullname / kname | varchar | 全名 / 简称 |
| kusercode | varchar | 自定义编号 |
| kcomment | varchar | 备注 |
| Address / Person / Tel | varchar | 仓库地址 / 负责人 / 电话 |
| knamepy | varchar | 拼音 |
| deleted | smallint | 必须 0 |
| krec / ParRec / ROWINDEX | int | 记录号 |
| updatetag | int | 10000 |
| SaveMenid / UpFlag / usesl / isMrpControl | int | 保管员 / 更新 / 是否数量 / MRP |

---

## 五、employee（职员，28 列）

| 列 | 类型 | 说明 |
|---|---|---|
| etypeid | varchar | 树编号 |
| Parid / leveal / soncount / esonnum | varchar/smallint/int | 父编号 / 层级 / 子级数 |
| efullname / ename | varchar | 全名 / 简名 |
| eusercode | varchar | 自定义编号 |
| Department / Tel / Address | varchar | 部门 / 电话 / 地址 |
| ecomment | varchar | 备注 |
| Pay / Incumbency / duty | varchar/numeric | 工资 / 职务 / 职务说明 |
| Birthday / EMail | varchar | 生日 / 邮箱 |
| TopTotal | numeric | 每单优惠限额 |
| LowLimitDiscount | numeric | 最低折扣下限 |
| deleted | smallint | 必须 0 |
| enamepy | varchar | 拼音 |
| erec / ParRec / ROWINDEX | int | 记录号 |
| updatetag | int | 10000 |
| SaveMenid / UpFlag | int | 更新标志 |

---

## 六、Excel 导入模板列（官方规格 6-sheet）

### 6.1 单位 sheet（30 列）

`父类编号 / 单位编号 / 单位全名 / 单位简名 / 单位地址 / 单位电话 / 手机一 / 手机二 / 传真 / 邮编 / 联系人一 / 联系人二 / 开户银行 / 银行账号 / 税号 / 期初应收款 / 期初应付款 / 期初预收款 / 期初预付款 / 应收款上限 / 应付款上限 / 结算期限(天) / 备注 / 换货期限(天) / 换货比例(%) / 电子邮件 / 默认经手人(编号) / 地区(编号) / 结算单位(编号) / 适用价格`

### 6.2 仓库 sheet（8 列）

`父类编号 / 仓库编号 / 仓库全名 / 仓库简称 / 备注 / 仓库地址 / 仓库负责人 / 负责人电话`

### 6.3 商品 sheet（46 列）

`父类编号 / 商品编号 / 商品全名 / 简名 / 型号 / 规格 / 产地 / 成本算法 / 商品属性 / 联营扣率(%) / 默认供应商(编号) / 品牌(编号) / 计价方式 / 有效期(天) / 备注 / 基本单位 / 条码 / 零售价 / 预设售价1 / 预设售价2 / 预设售价3 / 配送价 / 最近折后进价 / 最低售价 / 辅助单位1 / 辅1换算关系 / 辅1条码 / 辅1零售价 / 辅1预设售价1 / 辅1预设售价2 / 辅1预设售价3 / 辅助单位2 / 辅2换算关系 / 辅2条码 / 辅2零售价 / 辅2预设售价1 / 辅2预设售价2 / 辅2预设售价3 / 口味管理 / 默认税率 / 前台议价 / 启用批次 / 是否自制 / 自制周期 / 是否外购 / 外购周期`

### 6.4 职员 sheet（10 列）

`父类编号 / 职员编号 / 职员全名 / 生日 / 职员电话 / 职员职务 / 职员地址 / 最低折扣下限 / 每单优惠限额 / 备注`

### 6.5 库存信息 sheet（9 列）

`商品编号 / 商品全名 / 仓库编号 / 仓库全名 / 数量 / 金额 / 批号 / 生产日期 / 有效期至`

### 6.6 单位价格跟踪 sheet（9 列）

`单位编号 / 单位全名 / 商品编号 / 商品全名 / 计量单位 / 最近进价 / 最近售价 / 最近进货折扣 / 最近销售折扣`

---

## 七、商品价格类型（xw_P_PtypePrice.PRTypeId）

| PRTypeId | 名称 | 导入来源 |
|---|---|---|
| 0001 | 零售价 | BasPtypeInner.SalePriceBasic / BasPTypePriceMx IndexID=8 |
| 0002 | 预设售价1 | BasPtypeInner.SalePriceOne / BasPTypePriceMx IndexID=1 |
| 0003 | 预设售价2 | BasPtypeInner.SalePriceTwo / BasPTypePriceMx IndexID=2 |
| 0004 | 预设售价3 | BasPtypeInner.SalePriceThree / BasPTypePriceMx IndexID=3 |
| 0005 | 最近进价（折后） | BasPtypeInner.RecentPrice |
| 0007 | 最近售价（折前） | BasPtypeInner.RecentPrice1 |
| 0013 | 配送价 | Excel 模板配送价 |
| 0101 | 折前价格 | — |
| 0102 | 折后价格 | — |
| 0103 | 最低售价 | BasPtypeInner.MinPrice |

> 单位价格跟踪 xw_P_PriceFollow.PRTypeId：0005 最近进价 / 0006 最近进货折扣 / 0007 最近售价 / 0008 最近销售折扣（PPType=6，KTypeId=''，BillDt=1900-01-01）。

---

## 八、显示与增量规则（导入后必须满足）

1. **编号**：纯数字、长度≥5 且为 5 的倍数（00001…/0000100001/…）；含字母 → 列表行不渲染；<5 位 → flPtypeid 计算异常弹 "K开发提示：flPtypeid"。
2. **级联编号**：有层级数据必须父=子去末 5 位（树构建=编号前缀父子，不是 ParId）。
3. **子级计数**：soncount/psonnum/bsonnum/k sonnum/esonnum 必须真实（非根=COUNT(编号 LIKE 父编号+5 下划线 且 LEN=父+5)；根 00000=leveal=1 且 ParId='00000'）。
4. **deleted=0**：读取链（fn_PtypePriceall / p_hh_GetBSGoodsTree）过滤 deleted=1。
5. **updatetag=10000 + xw_BaseUpdateTag=0**：客户端按增量请求，数据表 updatetag 必须大值且增量表清零。
6. **条码 UnitID 数字串**：'0'/'1'/'2'，传单位名文本 → fn_GetBarCodeStr 245 崩溃，商品列表整块不显示。
7. **清客户端缓存**：改库后必须清 `F:\GRASPⅡTOP15\Data\localData\<账套>-00000-*` 并重登（旧缓存显示旧树）。
8. **验证链路**：`EXEC xw_GetBasicData @cMode,0,'',0,1000000`（@EtypeID 必须传 ''，@nUpdateTag<=0 全量），返回非根行数=导入数即显示 OK。

---

## 九、官方导入存储过程（6 个）

### 9.1 签名与返回

| 过程 | 用途 | 返回 |
|---|---|---|
| p_hh_BasicInsertB_Commodity | 单位 | @rc/@rlttypeid/@errorValue |
| p_hh_BasicInsertK_Commodity | 仓库 | @rc/@rlttypeid/@errorValue |
| p_hh_BascInsertP_Commodity | 商品 | @rc/@rlttypeid/@errorValue |
| p_hh_BasicInsertE_Commodity | 职员 | @rc/@rlttypeid/@errorValue |
| p_hh_ModInitGoodsStocks_SNMng_Package | 期初库存 | @rc/@OrderID/@errorValue |
| p_hh_ImportPtypePriceGZ | 单位价格跟踪 | @rc/@OrderID/@errorValue |

> 错误码：-13 已开账不能录期初；-17 没有找到父类；-18 已存在相同信息；-20 编号为空/条码已占用；-28/-29 辅助单位信息有误；-30 基本单位为空不能有辅助单位。

### 9.2 关键行为（实库验证）

1. **传入编号 → *usercode**：官方过程把调用方传入编号写入 pusercode/busercode/kusercode/eusercode（自定义编号），实际树编号 ptypeid/btypeid/ktypeid/etypeid 按过程规则另分配（实测 pusercode=00001 ↔ ptypeid=00012）。
2. **明细引用必须反查**：价格/条码/辅助单位/价格跟踪直写前，用 `resolve_real_ids` 按 *usercode 反查实际树编号，否则引用孤儿（全量首轮 89% 孤儿 → 修复后 0）。
3. 过程自动维护：拼音（商品）、层级、子级计数、updatetag、xw_BaseUpdateTag。
4. 期初库存依赖商品/仓库已存在；价格跟踪依赖单位/商品已存在。

---

## 十、源库直连（千业商贸通 → 管家婆）

### 10.1 千业源库

`E:\QySoft\gtSystemTYV86PRO\WorkData\*.mdf/*.ldf` 被千业软件透明加密、普通进程读被拒（PermissionError 13）；SQL Server 2008 Express（DESKTOP-8P5G8VQ\SQLEXPRESS）可附加：qysrc_ccc（CCC.mdf 634MB 生产库）/ qysrc_qys / qysrc_test。三库均为 Bas 旧版商贸通（22 张 Bas* 表，无标准 7 表）。

### 10.2 映射基线

Bas 表引用基线 = **BaseNo（int 主键）**：price.ptypeid、StorageIniGoods.PtypeId、BasPTypeBarMx.PTypeID、BasPTypePriceMx.PTypeID 100% 可 JOIN BasPtypeInner.BaseNo。

### 10.3 级联数字化

BaseNo → 新树编号：BaseLevel=1 → 00001 起 5 位行序；BaseLevel>1 → 父编号+5 位行序；根(BaseLevel=0) → 目标保留自身根，跳过。原编号写备注 `来源:Bas;原编号:<BaseCode>;<Memo>`。

### 10.4 价格映射

| 源 | → PRTypeId |
|---|---|
| BasPtypeInner.SalePriceBasic | 0001 |
| SalePriceOne/Two/Three | 0002/0003/0004 |
| RecentPrice / RecentPrice1 | 0005 / 0007 |
| MinPrice | 0103 |
| BasPTypePriceMx IndexID 1/2/3（预设售价1-3） | 0002/0003/0004 |
| BasPTypePriceMx IndexID 8（零售价） | 0001 |
| BasPTypePriceMx IndexID 4-7/9-16（门店/现金价等 12 种） | 无标准类型，跳过 |

去重：(PTypeId,PRTypeId) 全局唯一；UnitID='0'。

### 10.5 价格追踪 / 期初库存 / 条码 / 辅助单位

- price（800743 行，Units=-1 旧/1 新）→ GROUP BY (ptypeid,btypeid) MAX 聚合去重（400382 组）→ 每组 ×4 行（0005 进价/0006 进货折扣/0007 售价/0008 销售折扣），KTypeId=''、BillDt=1900-01-01、PPType=6、KeyId=NEWID()、UnitID='0'。
- StorageIniGoods（Qty≠0 5150 行）→ 官方过程 p_hh_ModInitGoodsStocks_SNMng_Package。
- BasPTypeBarMx（22992 行）→ xw_PtypeBarCode：UnitID='0'、Ordid=IndexID、keyid=NEWID；条码预清洗（同码仅首个保留，其余置空）。
- BasPtypeInner.BasicUnit/AssistUnitOne..Three → xw_PtypeUnit：URate=1 兜底、Ordid=0/1/2/3、IsBase=1/0。

### 10.6 实库验证（2026-11-02 CCC → HH15new）

- 试跑 120 商品全链路全绿（xw_GetBasicData 读取链路正常）→ 回滚。
- 全量：单位 882 / 仓库 8 / 职员 2 / 商品 24728（重复条码预清洗后 0 失败）/ 条码 22992 / 辅助单位 24765 / **预设价格 72591（(PTypeId,PRTypeId) 去重）** / **客户价格追踪 1601528（400382 组 ×4）** / 期初库存 5150。
- 引用完整性：明细直写经 **pusercode 反查实际树编号**（官方过程把传入编号存 *usercode、ptypeid 另行分配）后，价格/条码/辅助单位/价格跟踪孤儿引用全部为 0；父子悬空 0；编号非数字/长度非法/deleted/updatetag 全达标；单位拼音 fn_GetPyCode 补齐。
- 导入前 9 表备份可回滚；验证后发布 GitHub（DDDNN1234567/grasp-top15-import-tool）。
