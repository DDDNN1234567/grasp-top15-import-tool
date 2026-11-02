# -*- coding: utf-8 -*-
"""
管家婆辉煌 TOP15 一键导入工具 - 反向导出引擎（v1.2 新增）
管家婆 HH15new（标准 7 表） -> 千业商贸通 Bas 库（Bas* 表）
  树编号 -> BaseNo 重排（父先子后，按树编号排序）
  四主表 / 明细 / 16 价格方案 / 条码 / 辅助单位(含真实换算率) / 期初库存 / 客户价格跟踪 回写
依赖：gpj_top15_core、gpj_top15_srcimport（复用 PPM_FULL_MAP 双向映射）
"""
import os, sys, io, re, datetime
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except Exception:
        pass
import gpj_top15_core as core
import gpj_top15_srcimport as imp

SRC_SERVER = r'DESKTOP-8P5G8VQ\SQLSERVER'    # 管家婆实例
DST_SERVER = r'DESKTOP-8P5G8VQ\SQLEXPRESS'   # 千业实例
SRC_DB = 'HH15new'
DST_DB_DEFAULT = 'qysrc_test'
USER = 'sa'
DST_PWD = os.environ.get('GPJ_SRC_PWD', '')

# 管家婆 PRTypeId -> 千业 PPM IndexID（PPM_FULL_MAP 的逆映射；系统零售价 0001->8）
PR_TO_INDEX = {pr: ix for ix, (pr, _name) in imp.PPM_FULL_MAP.items()}
# 管家婆最近价/最低价 -> BasPtypeInner 列（不进 PPM 明细）
PR_TO_INNER = {'0005': 'RecentPrice', '0007': 'RecentPrice1', '0103': 'MinPrice'}

# 反向库存目标：当前结存表（与正向 STOCK_SOURCE 口径一致；特殊场景可改 StorageIniGoods）
DST_STOCK_TABLE = 'StorageGoodsInfo'
EXPORT_TABLES = ['BasUnitTypeInner', 'BasPtypeInner', 'BasStockInner', 'BasPersonInner',
                 'BasPTypeBarMx', 'BasPTypePriceMx', 'price', DST_STOCK_TABLE]


# ---------------- 工具 ----------------
def _s(v, n=None):
    """文本清洗：None->''，可选截断到 n 字节对应长度（按字符数保守截断）"""
    s = '' if v is None else str(v).strip()
    if n and len(s) > n:
        s = s[:n]
    return s


def _n(v):
    """数值清洗：None/空->0"""
    try:
        return float(v) if v is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


def _bit(v):
    try:
        return 1 if int(v) else 0
    except (TypeError, ValueError):
        return 0


def _py(cur, name, existing):
    """拼音：优先用源值，空则调用管家婆 dbo.fn_GetPyCode 现算（仓库/职员源拼音常为空）"""
    p = _s(existing, 66)
    if p:
        return p
    try:
        cur.execute("SET NOCOUNT ON; SELECT dbo.fn_GetPyCode(?)", name)
        r = cur.fetchone()
        return _s(r[0] if r else '', 66)
    except Exception:
        return ''


# ---------------- 编号重排 ----------------
def rebase_map(src_cur, table, idc, cmc):
    """管家婆树编号 -> 千业 BaseNo
    优先解析备注中的“源BaseNo:N”（正向溯源，实现主键 100% 还原，含历史缺口）；
    无标记的行（管家婆原生新增）连续编号，自动跳过已占用值。按树编号排序保证父先子后。"""
    src_cur.execute("SET NOCOUNT ON; SELECT %s, %s FROM dbo.%s ORDER BY %s" % (idc, cmc, table, idc))
    rows = src_cur.fetchall()
    id2bn, used = {}, set()
    for iid, comment in rows:                      # 第一遍：显式溯源 BaseNo
        m = re.search(r'源BaseNo:(\d+)', comment or '')
        if m:
            bn = int(m.group(1))
            id2bn[iid.strip()] = bn
            used.add(bn)
    nxt = 1
    for iid, comment in rows:                      # 第二遍：无标记连续编号
        key = iid.strip()
        if key in id2bn:
            continue
        while nxt in used:
            nxt += 1
        id2bn[key] = nxt
        used.add(nxt)
        nxt += 1
    return id2bn


def _parent_no(id2bn, parid):
    """父树编号 -> 父 BaseNo；根/缺失 -> 0"""
    if not parid:
        return 0
    return id2bn.get(parid.strip(), 0)


# ---------------- 辅助单位信息 ----------------
def load_unit_info(src_cur):
    """xw_PtypeUnit -> {PTypeId: {ordid: (unit, rate, isbase)}}"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, Unit1, URate, IsBase, Ordid FROM dbo.xw_PtypeUnit")
    u = {}
    for p, un, rate, isbase, ordid in src_cur.fetchall():
        try:
            u.setdefault(p.strip(), {})[int(ordid)] = (un, _n(rate), _bit(isbase))
        except (TypeError, ValueError):
            continue
    return u


# ---------------- 四主表导出 ----------------
def export_btype(src_cur, id2bn):
    """btype -> BasUnitTypeInner（49 列）"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.btype ORDER BY btypeid")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        iid = r['btypeid'].strip()
        data.append((
            id2bn[iid], _parent_no(id2bn, r['parid']), int(r['leveal']),          # BaseNo/ParentNo/BaseLevel
            _s('', 100), int(r['bsonnum'] or 0),                                  # BasePath/SonNum
            _s(r['busercode'], 30), _s(r['bname'], 60), _s(r['bfullname'], 60),  # BaseCode/Name/FullName
            0,                                                                   # AreaId
            _s(r['TelAndAddress'], 100), _s('', 30), _s(r['PostCode'], 20),     # Address/UnitPhone/PostCode
            _s(r['linker'], 20), _s(r['linkerTel'], 30), _s(r['bnamepy'], 30), # LinkMan/LinkPhone/PyNo
            _s(r['TaxNumber'], 20), _s(r['BankAndAcount'], 200), _s(r['bcomment'], 200),  # TaxNum/Bank/Memo
            _s(r['EMail'], 50), _s('', 20), _s('', 50),                          # Email/QQ/Msn
            _n(r['ARTotal']), _n(r['APTotal']), 0, 0,                            # AR/AP（管家婆无应收/应付上限列）
            _n(r['ARTotal00']), _n(r['APTotal00']), int(r['SKQX'] or 0), 0,                 # IniAR/IniAP/SkLimit/DeftPrice
            _n(r['R_WARNUP']), _n(r['P_WARNUP']), int(r['ChangeDay'] or 0),                 # 警告/换货日
            _n(r['ChangeRate']), _s(r['Fax'], 20), _bit(r['deleted']), int(r['ROWINDEX'] or 0),  # ChgRate/Fax/Del/RowIndex
            0, 0, 0, 0, 0, 0,                                                  # SendGoods* / PersonID / VestID
            _n(r['YRTotal00']), _n(r['YPTotal00']), _n(r['YRTotal']), _n(r['YPTotal']),     # Prepay
            _bit(r['isStop']), 0, _s('', 100),                                  # Stoped/Discount/SpaceDate
        ))
    return data


def export_stock(src_cur, id2bn):
    """Stock -> BasStockInner（15 列）"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.Stock ORDER BY ktypeid")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        iid = r['ktypeid'].strip()
        data.append((
            id2bn[iid], _parent_no(id2bn, r['parid']), int(r['leveal']),
            int(r['ksonnum'] or 0), _s('', 100),                               # SonNum/BasePath
            _s(r['kfullname'], 66), _s(r['kname'], 30), _s(r['kusercode'], 26),
            _s(r['kcomment'], 256), _bit(r['deleted']), _py(src_cur, r['kfullname'], r['knamepy']),
            int(r['ROWINDEX'] or 0), 0, _bit(r['isStop']), _s('', 10),         # StType/Stoped/Onlintype
        ))
    return data


def export_employee(src_cur, id2bn):
    """employee -> BasPersonInner（31 列）"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.employee ORDER BY etypeid")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        iid = r['etypeid'].strip()
        data.append((
            id2bn[iid], _parent_no(id2bn, r['Parid']), int(r['leveal']),
            _s('', 100), int(r['esonnum'] or 0), int(r['soncount'] or 0),       # BasePath/SonNum/SonCount
            _s(r['eusercode'], 26), _s(r['ename'], 30), _s(r['efullname'], 66),
            _py(src_cur, r['efullname'], r['enamepy']), _s('', 6), 0,            # PyNo(空则现算)/Sex/DeptID
            _s(r['Tel'], 20), _s('', 20), _s(r['Address'], 66),                 # LinkTel/HandTel/Address
            _s(r['ecomment'], 256), _bit(r['deleted']), _s(r['EMail'], 100),
            None, _s('', 18), _n(r['Pay']), _s(r['duty'], 20),                  # InJobDate/IDCard/Commission/Incumbency
            _s(r['Birthday'], 10), _n(r['TopTotal']), 0, _s('', 10),            # EveryTotal/EducateLevel
            int(r['ROWINDEX'] or 0), _s('', 20), 0, _bit(r['isStop']), _s('', 50),  # KTypeID/ChangeKType/Stoped/APPType
        ))
    return data


def export_ptype(src_cur, id2bn, unit_info, inner_extra):
    """ptype -> BasPtypeInner（49 列）；单位/换算率来自 xw_PtypeUnit，最近价/最低价来自 inner_extra"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.ptype ORDER BY ptypeid")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        iid = r['ptypeid'].strip()
        u = unit_info.get(iid, {})
        bn = id2bn[iid]
        # 单位与换算率
        basic = u.get(0, ('', 1, 1))
        a1 = u.get(1, ('', 1, 0)); a2 = u.get(2, ('', 1, 0)); a3 = u.get(3, ('', 1, 0))
        extra = inner_extra.get(bn, {})
        data.append((
            bn, _parent_no(id2bn, r['ParId']), int(r['leveal']),
            int(r['psonnum'] or 0), int(r['soncount'] or 0), _s('', 100),       # SonNum/SonCount/BasePath
            _s(r['pusercode'], 200), _s(r['pname'], 100), _s(r['pfullname'], 100),
            _s(r['Standard'], 50), _s('', 50), _s(r['Area'], 50),              # Standards/Model/ProArea
            _s(basic[0], 10), _s(a1[0], 10), _s(a2[0], 10), _s(a3[0], 10),     # BasicUnit/Assist*
            # 价格列（商品主表；默认 0，由 PPM/明细补充）
            _n(extra.get('SalePriceBasic')), 0, 0, 0,
            _n(basic[1]), _n(a1[1]), _n(a2[1]), _n(a3[1]),                      # Percent* 换算率
            int(r['UsefulLifeDay'] or 0),
            _n(extra.get('RecentPrice')), _n(extra.get('RecentPrice1')),        # RecentPrice/RecentPrice1
            int(r['costmode'] or 0), _n(r['warnup']), _n(r['warndown']),
            _bit(r['deleted']), _s(r['pnamepy'], 50), int(r['RowIndex'] or 0),
            0, 0,                                                                # isSerial/PChage
            _s(r['pcomment'], 3000), 0, _n(r['PoolDiscount']), 0,              # Memo/DeductType/DeductValue/MemberPrice
            _n(extra.get('MinPrice')), 0, _bit(r['isStop']), _s('', 50),       # MinPrice/Preprice1/Stoped/PtypePrice
            '0', '0', _s('', 10), 0, 0, int(r['Ptypetype'] or 0),              # BUnit/SUnit/OnLin*/gyid/szGoodsType
        ))
    return data


# ---------------- 明细导出 ----------------
def _unit_name_for(unit_info, iid, ordid):
    u = unit_info.get(iid, {})
    cell = u.get(ordid)
    return cell[0] if cell else ''


def load_inner_extra(src_cur, id2bn_p):
    """xw_P_PtypePrice 中最近价/最低价(0005/0007/0103) -> {商品BaseNo: {列: 值}}
    以及零售价 0001 -> SalePriceBasic（写主表）"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, PRTypeId, Price FROM dbo.xw_P_PtypePrice")
    out = {}
    for p, pr, price in src_cur.fetchall():
        bn = id2bn_p.get(p.strip())
        if bn is None:
            continue
        d = out.setdefault(bn, {})
        if pr in PR_TO_INNER:
            d[PR_TO_INNER[pr]] = _n(price)
        elif pr == '0001':
            d['SalePriceBasic'] = _n(price)
    return out


def export_barcodes(src_cur, id2bn_p, unit_info):
    """xw_PtypeBarCode -> BasPTypeBarMx（6 列）；每商品首个 IsDefault=1"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, UnitID, BarCode, Ordid FROM dbo.xw_PtypeBarCode ORDER BY PTypeId")
    seen = set()
    data = []
    for p, uid, bc, ordid in src_cur.fetchall():
        iid = p.strip()
        bn = id2bn_p.get(iid)
        if bn is None or not bc:
            continue
        try:
            oi = int(ordid)
        except (TypeError, ValueError):
            oi = 0
        isdef = 1 if bn not in seen else 0
        seen.add(bn)
        sub = _unit_name_for(unit_info, iid, oi)
        data.append((bn, _s(sub, 20), oi, _s('', 20), _s(bc, 50), isdef))
    return data


def export_ppm(src_cur, id2bn_p, unit_info):
    """xw_P_PtypePrice -> BasPTypePriceMx（6 列，16 方案）；最近价/折扣价不进 PPM"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, UnitID, PRTypeId, Price FROM dbo.xw_P_PtypePrice")
    data = []
    for p, uid, pr, price in src_cur.fetchall():
        iid = p.strip()
        bn = id2bn_p.get(iid)
        ix = PR_TO_INDEX.get(pr)
        if bn is None or ix is None:
            continue
        try:
            ui = int(uid)
        except (TypeError, ValueError):
            ui = 0
        sub = _unit_name_for(unit_info, iid, ui)
        name = imp.PPM_FULL_MAP[ix][1]
        data.append((bn, _s(sub, 10), ix, _n(price), _s(name, 50), _s('', 20)))
    return data


def export_pricefollow(src_cur, id2bn_p, id2bn_b):
    """xw_P_PriceFollow(PPType=6) -> price：按 (商品,客户) 把 4 行还原成 1 条"""
    src_cur.execute("""SET NOCOUNT ON;
SELECT PTypeId, BTypeId, PRTypeId, Price FROM dbo.xw_P_PriceFollow WHERE PPType=6""")
    agg = {}
    for p, b, pr, price in src_cur.fetchall():
        pbn = id2bn_p.get(p.strip()); bbn = id2bn_b.get(b.strip())
        if pbn is None or bbn is None:
            continue
        agg.setdefault((pbn, bbn), {})[pr] = _n(price)
    data = []
    for (pbn, bbn), d in agg.items():
        data.append((
            pbn, bbn,
            d.get('0005', 0), d.get('0007', 0), d.get('0006', 0), d.get('0008', 0),  # cost/sale/buydisc/saledisc
            _s('', 10), _s('', 10), 0, 1,                                          # BuyDate/SaleDate/IsPrint/Units
        ))
    return data


def export_stockini(src_cur, id2bn_p, id2bn_k):
    """IniGoodsStocks -> 当前结存表 StorageGoodsInfo（排除 identity GOODSORDERID，无 BasePath，10 列）"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.IniGoodsStocks ORDER BY GoodsOrder")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        pbn = id2bn_p.get(r['PtypeId'].strip()); kbn = id2bn_k.get(r['KtypeId'].strip())
        if pbn is None or kbn is None:
            continue
        data.append((
            pbn, kbn, _s(r['JobNumber'], 20), _s(r['OutFactoryDate'], 13),
            _n(r['Qty']), _n(r['Price']), _n(r['Total']), int(r['GoodsOrder'] or 0),
            _s('', 100), _s('', 100),                                                # PPath/KPath
        ))
    return data


# ---------------- 目标库清空/写入 ----------------
def _bulk_insert(dst_cur, table, data, columns, batch=2000):
    if not data:
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.%s" % table)
        return 0
    ph = ','.join('?' * len(columns))
    sql = "SET NOCOUNT ON; INSERT INTO dbo.%s (%s) VALUES (%s)" % (table, ','.join(columns), ph)
    # 注：不启用 fast_executemany——跨版本千业库列顺序/类型推断差异下，
    # 快速批量绑定偶发 varchar->numeric 8114；普通 executemany 逐行推断更稳
    dst_cur.fast_executemany = False
    n = 0
    for i in range(0, len(data), batch):
        dst_cur.executemany(sql, data[i:i + batch])
        n += len(data[i:i + batch])
    return n


def backup_dst(dst_cur):
    stamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    pre = '_expbak_' + stamp
    for t in EXPORT_TABLES:
        dst_cur.execute("SET NOCOUNT ON; SELECT * INTO %s%s FROM dbo.%s" % (t, pre, t))
    return pre


# ---------------- 编排 ----------------
def export_to_db(src_cur, dst_cur, logcb):
    """管家婆 -> 千业 一键反向导出：备份目标 -> 重排编号 -> 清空 -> 四主表 -> 明细"""
    pre = backup_dst(dst_cur)
    logcb('INFO', '备份', 0, '已备份目标 8 表（前缀 %s）' % pre)
    stats = {}
    # 1. 编号重排
    id2bn_b = rebase_map(src_cur, 'btype', 'btypeid', 'bcomment')
    id2bn_p = rebase_map(src_cur, 'ptype', 'ptypeid', 'pcomment')
    id2bn_k = rebase_map(src_cur, 'Stock', 'ktypeid', 'kcomment')
    id2bn_e = rebase_map(src_cur, 'employee', 'etypeid', 'ecomment')
    logcb('INFO', '映射', 0, '编号重排：单位 %d / 商品 %d / 仓库 %d / 职员 %d' %
          (len(id2bn_b), len(id2bn_p), len(id2bn_k), len(id2bn_e)))
    # 2. 辅助单位 + 主表额外价格
    unit_info = load_unit_info(src_cur)
    inner_extra = load_inner_extra(src_cur, id2bn_p)
    # 3. 清空目标 8 表（先关外键无关，Bas 表无强外键）
    for t in EXPORT_TABLES:
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.%s" % t)
    # 4. 四主表
    cols_b = ['BaseNo', 'ParentNo', 'BaseLevel', 'BasePath', 'SonNum', 'BaseCode', 'Name', 'FullName',
              'AreaId', 'Address', 'UnitPhone', 'PostCode', 'LinkMan', 'LinkPhone', 'PyNo', 'TaxNum',
              'BankAndAcount', 'Memo', 'Email', 'QQNo', 'MsnNo', 'ARTotal', 'APTotal', 'ARLimit', 'APLimit',
              'IniARTotal', 'IniAPTotal', 'SkLimit', 'DeftPrice', 'ARWaringUp', 'PRWaringUp', 'ChgGoodDay',
              'ChgGoodScale', 'FaxNum', 'Deleted', 'RowIndex', 'SendGoods00', 'SendGoods', 'SendGoodSJ00',
              'SendGoodSJ', 'PersonID', 'VestID', 'IniPrepayARTotal', 'IniPrepayAPTotal', 'CurPrepayARTotal',
              'CurPrepayAPTotal', 'Stoped', 'Discount', 'SpaceDate']
    stats['BasUnitTypeInner'] = _bulk_insert(dst_cur, 'BasUnitTypeInner', export_btype(src_cur, id2bn_b), cols_b)

    cols_k = ['BaseNo', 'ParentNo', 'BaseLevel', 'SonNum', 'BasePath', 'FullName', 'Name', 'BaseCode',
              'Memo', 'deleted', 'PyNo', 'ROWINDEX', 'StType', 'Stoped', 'Onlintype']
    stats['BasStockInner'] = _bulk_insert(dst_cur, 'BasStockInner', export_stock(src_cur, id2bn_k), cols_k)

    cols_e = ['BaseNo', 'ParentNo', 'BaseLevel', 'BasePath', 'SonNum', 'SonCount', 'BaseCode', 'Name',
              'FullName', 'PyNo', 'Sex', 'DeptID', 'LinkTel', 'HandTel', 'Address', 'Memo', 'Deleted',
              'EMail', 'InJobDate', 'IDCard', 'CommissionPercent', 'Incumbency', 'Birthday', 'TopTotal',
              'EveryTotal', 'EducateLevel', 'RowIndex', 'KTypeID', 'ChangeKType', 'Stoped', 'APPType']
    stats['BasPersonInner'] = _bulk_insert(dst_cur, 'BasPersonInner', export_employee(src_cur, id2bn_e), cols_e)

    cols_p = ['BaseNo', 'ParentNo', 'BaseLevel', 'SonNum', 'SonCount', 'BasePath', 'BaseCode', 'Name',
              'FullName', 'Standards', 'Model', 'ProArea', 'BasicUnit', 'AssistUnitOne', 'AssistUnitTwo',
              'AssistUnitThree', 'SalePriceBasic', 'SalePriceOne', 'SalePriceTwo', 'SalePriceThree',
              'PercentBasic', 'PercentOne', 'PercentTwo', 'PercentThree', 'UseFullLifeDay', 'RecentPrice',
              'RecentPrice1', 'CostType', 'WaringUp', 'WaringDown', 'Deleted', 'PyNo', 'RowIndex', 'isSerial',
              'PChage', 'Memo', 'DeductType', 'DeductValue', 'MemberPrice', 'MinPrice', 'Preprice1',
              'Stoped', 'PtypePrice', 'BUnit', 'SUnit', 'OnLinType', 'OnlinPrice', 'gyid', 'szGoodsType']
    stats['BasPtypeInner'] = _bulk_insert(dst_cur, 'BasPtypeInner',
                                          export_ptype(src_cur, id2bn_p, unit_info, inner_extra), cols_p)
    # 5. 明细
    stats['BasPTypeBarMx'] = _bulk_insert(dst_cur, 'BasPTypeBarMx',
                                          export_barcodes(src_cur, id2bn_p, unit_info),
                                          ['PTypeID', 'SubUnit', 'IndexID', 'BarName', 'BarCode', 'IsDefault'])
    stats['BasPTypePriceMx'] = _bulk_insert(dst_cur, 'BasPTypePriceMx',
                                            export_ppm(src_cur, id2bn_p, unit_info),
                                            ['PTypeID', 'SubUnit', 'IndexID', 'AdvPrice', 'DisPlayName', 'PyNo'])
    stats['price'] = _bulk_insert(dst_cur, 'price',
                                  export_pricefollow(src_cur, id2bn_p, id2bn_b),
                                  ['ptypeid', 'btypeid', 'costprice', 'saleprice', 'buydiscount',
                                   'salediscount', 'BuyDate', 'SaleDate', 'IsPrint', 'Units'])
    stats[DST_STOCK_TABLE] = _bulk_insert(dst_cur, DST_STOCK_TABLE,
                                          export_stockini(src_cur, id2bn_p, id2bn_k),
                                          ['PtypeId', 'KtypeId', 'JobNumber', 'OutFactoryDate', 'Qty',
                                           'Price', 'Total', 'GoodsOrder', 'PPath', 'KPath'])
    return pre, stats


def main(argv):
    dst_db = argv[1] if len(argv) > 1 else DST_DB_DEFAULT
    logcb = lambda tag, m, ln, msg: print('[%s][%s] %s' % (tag, m, msg))
    src = core.connect(SRC_SERVER, SRC_DB, USER, DST_PWD)
    dst = core.connect(DST_SERVER, dst_db, USER, DST_PWD)
    print('源 %s.%s / 目标 %s.%s' % (SRC_SERVER, SRC_DB, DST_SERVER, dst_db))
    pre, stats = export_to_db(src.cursor(), dst.cursor(), logcb)
    for k, v in stats.items():
        print('%-20s %d' % (k, v))
    src.close(); dst.close()
    print('完成（目标备份前缀 %s）。' % pre)


if __name__ == '__main__':
    main(sys.argv)
