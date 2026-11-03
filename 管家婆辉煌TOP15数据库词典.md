# 管家婆辉煌 TOP15 数据库词典（目标库 HH15new）

> 实测环境：SQL Server 2016 Express（13.0.5026.0），实例 `DESKTOP-8P5G8VQ\SQLSERVER`，库 HH15new，共 758 张表。
> 本词典记录导入相关的核心表、官方存储过程与读取链路，供工具开发与故障排查使用。

## 一、核心表（9 张）

| 表 | 功能 | 编号列 | 根节点 | identity |
|---|---|---|---|---|
| btype | 单位信息 | btypeid | 00000 全部单位 | brec |
| ptype | 商品信息 | ptypeid | 00000 全部商品 | prec |
| Stock | 仓库信息 | ktypeid | 00000 全部 | C1rec |
| employee | 职员信息 | etypeid | 00000 管理员 | erec |
| IniGoodsStocks | 期初库存 | — | — | GOODSORDERID |
| xw_P_PriceFollow | 单位价格跟踪 | — | — | 无 |
| xw_P_PtypePrice | 商品价格 | — | — | 无 |
| xw_PtypeUnit | 商品辅助单位 | — | — | 无 |
| xw_PtypeBarCode | 商品条码 | — | — | 无 |

## 二、主表关键列

### 1. btype（单位）
- 编号层级：btypeid（纯数字，5 的倍数）、leveal（层级，根=0）、ParId（父编号）；
- 名称：bfullname（全名）、bname（简称）、bnamepy（拼音）；
- 期初：Artotal00/Aptotal00（应收/应付）、YRTotal00/YPTotal00（预收/预付）；
- 限额：ArUp/ApUp、skqx（结算期限）；
- 显示控制：deleted（必须 0）、updatetag（>0）、soncount/bsonnum（直接子级数）；
- 其他：baddress、btel、btaxno、bbankacct、bcomment、busercode（用户编号）。

### 2. ptype（商品）
- 编号层级：ptypeid、leveal、ParId；
- 名称：pfullname、pname、pnamepy；
- 属性：ptype、pstandard、parea、pmodel、ptypetype、ppooldiscount、pLifeDay、pdefaulttax；
- 显示控制：deleted、updatetag、soncount/psonnum；
- 其他：pcomment、pusercode。

### 3. Stock（仓库）
- ktypeid、leveal、ParId、kfullname、kname、kaddress、kperson、ktel、kcomment、kusercode、soncount。

### 4. employee（职员）
- etypeid、leveal、ParId、efullname、ename、ebirthday、etel、eduty、eaddress、esex、ecomment、eusercode。

## 三、明细表关键列

### 1. xw_P_PtypePrice（商品价格）
- 主键：(PTypeId, UnitID, PRTypeId)；
- 列：PTypeId、UnitID（'0' 数字串）、PRTypeId（价格类型）、Price、UpDateDt、KeyId；
- 价格类型见 xw_P_PriceType。

### 2. xw_PtypeUnit（辅助单位）
- 列：PTypeId、Unit1（单位名）、URate（换算率，1 辅助单位=URate 基本单位）、IsBase、Comment、Ordid、keyid；
- Ordid：0=基本单位，1/2/3=辅助单位。

### 3. xw_PtypeBarCode（条码）
- 列：PTypeId、UnitID（必须数字串 '0'）、BarCode、Ordid、keyid、kwtypeid。

### 4. xw_P_PriceFollow（单位价格跟踪）
- 列：PTypeId、BTypeId、UnitID、Price、KTypeId、BillDt、PRTypeId、PPType、KeyId、JGTypeId、MDTypeId；
- PRTypeId：0005 进价 / 0006 进货折扣 / 0007 售价 / 0008 销售折扣；PPType=6。

### 5. IniGoodsStocks（期初库存，13 列）
- PtypeId、KtypeId、JobNumber、OutFactoryDate、Qty、Price、Total、GoodsOrder、GOODSORDERID（identity）、
  pgholInqty、GoodsBatchID、BtypeID、BuyDate、UsefulEndDate；无触发器。

## 四、价格类型 xw_P_PriceType

- 共 28 行；系统价：0001 零售、0002-0004 预设售价、0005 最近进价、0007 最近售价、0103 最低售价；
- 自定义价：0104-0118（IsSystem=0），F_GetPriceInfo 以 IsSystem=0 动态读取；
- 列：PRTypeId、PRName、PRDisName、PRType、IsVisible、IsSystem、IsEdit、IsSet、Comment、URec、ParRec、USonnum、updatetag、Ordid、IsBaseVis。

## 五、业务表（v3.0 新增备份）

### 1. Commission / IniCommission（委托/受托，各 11 列）
- PTYPEID、BTYPEID、TYPE、QTY、TOTAL、PRODATE、BLOCKNO、COSTTOTAL、tax、taxtotal、goodsorderId；
- TYPE：1 受托（计应付 AP）、4 委托发出（计应收 AR）、5 委托退回（计应收）。

### 2. lendborrow / lendborrow00（借进/借出，7 列）
- btypeid、ptypeid、qty、total、ktypeid、price、goodsorderId。

### 3. 虚拟仓库
- 9999999999 借出仓、9999999998 借入仓；实体仓库为 00012-00019（8 个）+根 00000，无虚拟仓库行。

## 六、官方导入存储过程

| 存储过程 | 用途 | 关键输出 |
|---|---|---|
| p_hh_BasicInsertB_Commodity | 单位 | @RltTypeID（实际树编号）、@errorValue |
| p_hh_BasicInsertK_Commodity | 仓库 | @RltTypeID、@errorValue |
| p_hh_BascInsertP_Commodity | 商品 | @RltTypeID、@errorValue |
| p_hh_BasicInsertE_Commodity | 职员 | @RltTypeID、@errorValue |
| p_hh_ModInitGoodsStocks_SNMng_Package | 期初库存 | @OrderID、@errorValue |
| p_hh_ImportPtypePriceGZ | 价格跟踪 | @OrderID、@errorValue |

要点：传入的“单位编号”被当作 @UserCode（存 busercode），实际树编号 btypeid 由过程按系统规则自动分配；
父编号须经 ParentResolver 动态映射。

## 七、读取链路（软件显示）

### 1. xw_GetBasicData
- 签名：`xw_GetBasicData @cMode, @nUpdateTag, @EtypeID, @beginRec, @endRec`；
- @cMode：B 单位 / P 商品 / K 仓库 / E 职员；@EtypeID 必须传空串；@nUpdateTag<=0 全量；
- 能返回非根行数即证明基本信息显示链路 OK。

### 2. 树与过滤
- 树编号前缀父子：父编号=子编号去掉末 5 位（flPtypeid=LEFT(编号,LEN-5)）；
- fn_PtypePriceall 内部 WHERE deleted=0；p_hh_GetBSGoodsTree 走该函数；
- 子级计数控制树能否展开。

## 八、开账标志 sysdata

- 四列：SubNo/SubName/SubValue/SubComment；
- iniover=1 已开账（inioverDate 空），iniover=0 未开账；
- 已开账时官方过程拒绝录入期初应收应付/期初库存，需反开账；业务单据表全部 0 行时反开账安全。

## 九、客户端缓存

- 路径：`F:\GRASPⅡTOP15\Data\localData\<账套>-00000-*`；
- 改库后清缓存并重新登录，避免界面显示旧数据。
