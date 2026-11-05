# -*- coding: utf-8 -*-
"""
管家婆辉煌 TOP15 一键导入工具 - 反向导出引擎 v3.0
管家婆 HH15new（标准表） -> 千业商贸通 Bas 库（Bas* 表）
v3.0 修复（实机验收问题）：
  1) 四主表 BasePath 按"父路径-自己BaseNo"正确生成（根='1'），修复 fn_GetBasePath 路径塌缩
  2) StorageGoodsInfo 的 PPath/KPath = 商品/仓库 BasePath + '-'，修复期初库存界面为空
  3) 辅助单位空槽位 Percent=0（不再误写1），修复辅助单位描述错误
  4) 职员 Name 空时回退全名、Sex 默认男，修复职员信息为空
  5) 新增委托受托(Commission/IniCommission)、借进借出(lendborrow) 四张业务表映射
  6) 价格方案动态解析（不再硬编码），自定义价格按名称匹配、新槽位自动分配
依赖：gpj_top15_core、gpj_top15_srcimport（PPM_FULL_MAP 作为已知名称字典兜底）
"""
import os, sys, io, re, datetime
from collections import Counter
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

# 管家婆系统价 -> 千业 BasPtypeInner 主表列（不进 PPM 明细）
PR_TO_INNER = {'0005': 'RecentPrice', '0007': 'RecentPrice1', '0103': 'MinPrice'}
# 管家婆零售价固定锚点 -> 千业 IndexID 8（千业 8 号槽位即零售价）
RETAIL_PR, RETAIL_INDEX = '0001', 8

# 反向库存目标：当前结存表（与正向 STOCK_SOURCE 口径一致；特殊场景可改 StorageIniGoods）
DST_STOCK_TABLE = 'StorageGoodsInfo'
# 备份/清空范围：原 8 表 + 委托受托/借进借出 4 张业务表
EXPORT_TABLES = ['BasUnitTypeInner', 'BasPtypeInner', 'BasStockInner', 'BasPersonInner',
                 'BasPTypeBarMx', 'BasPTypePriceMx', 'price', DST_STOCK_TABLE,
                 'StorageCommission', 'StorageInitCommiss',
                 'BillQcLendAndBorrow', 'BillLendOrBorrow']


# ---------------- 基础清洗 ----------------
def _s(v, n=None):
    """文本清洗：None->''，可选按字符数截断"""
    s = '' if v is None else str(v).strip()
    if n and len(s) > n:
        s = s[:n]
    return s


def _n(v):
    """数值清洗：None/空/非法 -> 0"""
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
    """拼音：优先源值，空则调 fn_GetPyCode 现算"""
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
    优先解析备注"源BaseNo:N"（正向溯源，100%还原含历史缺口）；无标记行连续编号、自动跳过占用值。"""
    src_cur.execute("SET NOCOUNT ON; SELECT %s, %s FROM dbo.%s ORDER BY %s" % (idc, cmc, table, idc))
    rows = src_cur.fetchall()
    id2bn, used = {}, set()
    for iid, comment in rows:                       # 第一遍：显式溯源
        m = re.search(r'源BaseNo:(\d+)', comment or '')
        if m:
            bn = int(m.group(1))
            id2bn[iid.strip()] = bn
            used.add(bn)
    nxt = 1
    for iid, comment in rows:                       # 第二遍：无标记连续编号
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


def build_basepath(id2bn, parid_of, level_of):
    """通用 BasePath 生成：根(level0)='1'，子级=父BasePath+'-'+自己BaseNo。
    id2bn: 管家婆编号->BaseNo；parid_of/level_of: 编号->父编号/层级。返回 {BaseNo: BasePath}"""
    bn_par, bn_lv = {}, {}
    for iid, bn in id2bn.items():
        parid = (parid_of.get(iid) or '').strip()
        bn_par[bn] = id2bn.get(parid, 0)
        bn_lv[bn] = int(level_of.get(iid) or 0)
    path = {}
    for lv in sorted(set(bn_lv.values())):
        for bn in [b for b in bn_lv if bn_lv[b] == lv]:
            pbn = bn_par[bn]
            if lv == 0 or pbn == 0:
                path[bn] = '1'
            else:
                path[bn] = path.get(pbn, '1') + '-' + str(bn)
    return path


# ---------------- 价格方案动态解析 ----------------
class PriceSchemeResolver:
    """管家婆 PRTypeId <-> 千业 IndexID 动态双向解析（不依赖固定账套）。
    规则：零售价固定 8 号槽；最近价/最低价走主表；自定义价格按名称匹配，
    匹配不上的自动分配 1-16 空槽（避开零售价）。"""

    def __init__(self, src_cur, dst_cur):
        # 1. 管家婆价格类型名称
        src_cur.execute("SET NOCOUNT ON; SELECT PRTypeId, PRName FROM dbo.xw_P_PriceType")
        self.gp_name = {pr: _s(name) for pr, name in src_cur.fetchall()}
        # 2. 实际有数据的 PRTypeId（只处理这些，避免空系统价占槽）
        src_cur.execute("SET NOCOUNT ON; SELECT DISTINCT PRTypeId FROM dbo.xw_P_PtypePrice")
        active_pr = [r[0] for r in src_cur.fetchall()]
        # 3. 千业现有 IndexID -> 名称（目标库清空前读，取众数）
        dst_cur.execute("SET NOCOUNT ON; SELECT IndexID, DisPlayName FROM dbo.BasPTypePriceMx")
        bag = {}
        for ix, nm in dst_cur.fetchall():
            bag.setdefault(int(ix), []).append(_s(nm))
        self.qy_name = {ix: Counter(names).most_common(1)[0][0] for ix, names in bag.items()}
        # PPM_FULL_MAP 已知名称兜底
        for ix, (_pr, name) in imp.PPM_FULL_MAP.items():
            self.qy_name.setdefault(ix, name)
        # 名称 -> 槽位反查
        name2ix = {nm: ix for ix, nm in self.qy_name.items()}
        # 4. 构建双向映射
        self.pr2ix, self.ix2pr = {}, {}
        if RETAIL_PR in active_pr:
            self._bind(RETAIL_PR, RETAIL_INDEX)
        for pr in sorted(active_pr):                      # 按编号确定性排序
            if pr == RETAIL_PR or pr in PR_TO_INNER:
                continue
            nm = self.gp_name.get(pr, '')
            ix = name2ix.get(nm)
            if ix is not None and ix not in self.ix2pr:  # 名称强匹配
                self._bind(pr, ix)
            else:                                        # 分配空槽
                for cand in range(1, 64):
                    if cand != RETAIL_INDEX and cand not in self.ix2pr:
                        ix = cand
                        self._bind(pr, ix)
                        self.qy_name[ix] = nm
                        name2ix[nm] = ix
                        break

    def _bind(self, pr, ix):
        self.pr2ix[pr] = ix
        self.ix2pr[ix] = pr

    def index_of(self, pr):
        return self.pr2ix.get(pr)

    def name_of_index(self, ix):
        return self.qy_name.get(ix, '')


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
def export_btype(src_cur, id2bn, path):
    """btype -> BasUnitTypeInner（49 列）"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.btype ORDER BY btypeid")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        iid = r['btypeid'].strip()
        bn = id2bn[iid]
        data.append((
            bn, _parent_no(id2bn, r['parid']), int(r['leveal']),
            path[bn], int(r['bsonnum'] or 0),                                    # BasePath/SonNum
            _s(r['busercode'], 30), _s(r['bname'], 60), _s(r['bfullname'], 60),
            0,
            _s(r['TelAndAddress'], 100), _s('', 30), _s(r['PostCode'], 20),
            _s(r['linker'], 20), _s(r['linkerTel'], 30), _s(r['bnamepy'], 30),
            _s(r['TaxNumber'], 20), _s(r['BankAndAcount'], 200), _s(r['bcomment'], 200),
            _s(r['EMail'], 50), _s('', 20), _s('', 50),
            _n(r['ARTotal']), _n(r['APTotal']), 0, 0,
            _n(r['ARTotal00']), _n(r['APTotal00']), int(r['SKQX'] or 0), 0,
            _n(r['R_WARNUP']), _n(r['P_WARNUP']), int(r['ChangeDay'] or 0),
            _n(r['ChangeRate']), _s(r['Fax'], 20), _bit(r['deleted']), int(r['ROWINDEX'] or 0),
            0, 0, 0, 0, 0, 0,
            _n(r['YRTotal00']), _n(r['YPTotal00']), _n(r['YRTotal']), _n(r['YPTotal']),
            _bit(r['isStop']), 0, _s('', 100),
        ))
    return data


def export_stock(src_cur, id2bn, path):
    """Stock -> BasStockInner（15 列）"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.Stock ORDER BY ktypeid")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        iid = r['ktypeid'].strip()
        bn = id2bn[iid]
        data.append((
            bn, _parent_no(id2bn, r['parid']), int(r['leveal']),
            int(r['ksonnum'] or 0), path[bn],                                    # SonNum/BasePath
            _s(r['kfullname'], 66), _s(r['kname'], 30), _s(r['kusercode'], 26),
            _s(r['kcomment'], 256), _bit(r['deleted']), _py(src_cur, r['kfullname'], r['knamepy']),
            int(r['ROWINDEX'] or 0), 0, _bit(r['isStop']), _s('', 10),
        ))
    return data


def export_employee(src_cur, id2bn, path):
    """employee -> BasPersonInner（31 列）；Name 空回退全名、Sex 默认男"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.employee ORDER BY etypeid")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        iid = r['etypeid'].strip()
        bn = id2bn[iid]
        data.append((
            bn, _parent_no(id2bn, r['Parid']), int(r['leveal']),
            path[bn], int(r['esonnum'] or 0), int(r['soncount'] or 0),
            _s(r['eusercode'], 26),
            _s(r['ename'] or r['efullname'], 30),                                # Name 回退全名
            _s(r['efullname'], 66),
            _py(src_cur, r['efullname'], r['enamepy']),
            '男', 0,                                                            # Sex 默认男 / DeptID
            _s(r['Tel'], 20), _s('', 20), _s(r['Address'], 66),
            _s(r['ecomment'], 256), _bit(r['deleted']), _s(r['EMail'], 100),
            None, _s('', 18), _n(r['Pay']), _s(r['duty'], 20),
            _s(r['Birthday'], 10), _n(r['TopTotal']), 0, _s('', 10),
            int(r['ROWINDEX'] or 0), _s('', 20), 0, _bit(r['isStop']), _s('', 50),
        ))
    return data


def export_ptype(src_cur, id2bn, path, unit_info, inner_extra):
    """ptype -> BasPtypeInner（49 列）；空辅助槽位 Percent=0"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.ptype ORDER BY ptypeid")
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        iid = r['ptypeid'].strip()
        u = unit_info.get(iid, {})
        bn = id2bn[iid]
        basic = u.get(0, ('', 1, 1))
        # 空槽位默认 rate=0（修复辅助单位描述错误）
        a1 = u.get(1, ('', 0, 0)); a2 = u.get(2, ('', 0, 0)); a3 = u.get(3, ('', 0, 0))
        extra = inner_extra.get(bn, {})
        data.append((
            bn, _parent_no(id2bn, r['ParId']), int(r['leveal']),
            int(r['psonnum'] or 0), int(r['soncount'] or 0), path[bn],
            _s(r['pusercode'], 200), _s(r['pname'], 100), _s(r['pfullname'], 100),
            _s(r['Standard'], 50), _s('', 50), _s(r['Area'], 50),
            _s(basic[0], 10), _s(a1[0], 10), _s(a2[0], 10), _s(a3[0], 10),
            _n(extra.get('SalePriceBasic')), 0, 0, 0,
            _n(basic[1]), _n(a1[1]), _n(a2[1]), _n(a3[1]),                       # Percent*（空槽=0）
            int(r['UsefulLifeDay'] or 0),
            _n(extra.get('RecentPrice')), _n(extra.get('RecentPrice1')),
            int(r['costmode'] or 0), _n(r['warnup']), _n(r['warndown']),
            _bit(r['deleted']), _s(r['pnamepy'], 50), int(r['RowIndex'] or 0),
            0, 0,
            _s(r['pcomment'], 3000), 0, _n(r['PoolDiscount']), 0,
            _n(extra.get('MinPrice')), 0, _bit(r['isStop']), _s('', 50),
            '0', '0', _s('', 10), 0, 0, int(r['Ptypetype'] or 0),
        ))
    return data


# ---------------- 明细导出 ----------------
def _unit_name_for(unit_info, iid, ordid):
    cell = unit_info.get(iid, {}).get(ordid)
    return cell[0] if cell else ''


def load_inner_extra(src_cur, id2bn_p):
    """xw_P_PtypePrice 中最近价/最低价(0005/0007/0103) -> {商品BaseNo:{列:值}}；零售价 0001 -> SalePriceBasic"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, PRTypeId, Price FROM dbo.xw_P_PtypePrice")
    out = {}
    for p, pr, price in src_cur.fetchall():
        bn = id2bn_p.get(p.strip())
        if bn is None:
            continue
        d = out.setdefault(bn, {})
        if pr in PR_TO_INNER:
            d[PR_TO_INNER[pr]] = _n(price)
        elif pr == RETAIL_PR:
            d['SalePriceBasic'] = _n(price)
    return out


def export_barcodes(src_cur, id2bn_p, unit_info):
    """xw_PtypeBarCode -> BasPTypeBarMx（6 列）；每商品首个 IsDefault=1"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, UnitID, BarCode, Ordid FROM dbo.xw_PtypeBarCode ORDER BY PTypeId")
    seen, data = set(), []
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


def export_ppm(src_cur, id2bn_p, unit_info, resolver):
    """xw_P_PtypePrice -> BasPTypePriceMx（6 列）；按动态解析映射，最近价/折扣价不进 PPM"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, UnitID, PRTypeId, Price FROM dbo.xw_P_PtypePrice")
    data = []
    for p, uid, pr, price in src_cur.fetchall():
        iid = p.strip()
        bn = id2bn_p.get(iid)
        ix = resolver.index_of(pr)
        if bn is None or ix is None:
            continue
        try:
            ui = int(uid)
        except (TypeError, ValueError):
            ui = 0
        sub = _unit_name_for(unit_info, iid, ui)
        data.append((bn, _s(sub, 10), ix, _n(price),
                     _s(resolver.name_of_index(ix), 50), _s('', 20)))
    return data


def export_pricefollow(src_cur, id2bn_p, id2bn_b):
    """xw_P_PriceFollow(PPType=6) -> price：按(商品,客户)把4行还原1条"""
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
            d.get('0005', 0), d.get('0007', 0), d.get('0006', 0), d.get('0008', 0),
            _s('', 10), _s('', 10), 0, 1,
        ))
    return data


def export_stockini(src_cur, id2bn_p, id2bn_k, path_p, path_k):
    """IniGoodsStocks -> 当前结存表 StorageGoodsInfo（10 列，identity GOODSORDERID 排除）；PPath/KPath 补齐"""
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
            path_p[pbn] + '-', path_k[kbn] + '-',                                  # PPath/KPath
        ))
    return data


# ---------------- 委托受托 / 借进借出（v3.0 新增） ----------------
def export_commission(src_cur, id2bn_p, id2bn_b, path_p, src_tbl, dst_kind):
    """管家婆 Commission/IniCommission -> 千业 StorageCommission/StorageInitCommiss（11 列）
    字段语义双方一致，仅编号引用需转换；BasePath=商品路径+'-'。"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.%s" % src_tbl)
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        pbn = id2bn_p.get(r['PTYPEID'].strip()); bbn = id2bn_b.get(r['BTYPEID'].strip())
        if pbn is None or bbn is None:
            continue
        data.append((
            pbn, bbn, int(r['TYPE']),                                             # PTYPEID/BTYPEID/TYPE
            _n(r['QTY']), _n(r['TOTAL']), _s(r['PRODATE'], 10),                  # Quantity/ToTal/Prodate
            _n(r['COSTTOTAL']), _n(r['tax']), _n(r['taxtotal']),                 # CostToTal/TaxPercent/TaxTotal
            path_p[pbn] + '-', _s(r['BLOCKNO'], 50),                             # BasePath/BlockNo
        ))
    return data


def export_lendborrow(src_cur, id2bn_p, id2bn_b, id2bn_k, path_p, src_tbl):
    """管家婆 lendborrow/lendborrow00 -> 千业 BillQcLendAndBorrow/BillLendOrBorrow（7 列）
    虚拟仓库(99999开头)透传，普通仓库转 BaseNo；BasePath=商品路径+'-'。"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM dbo.%s" % src_tbl)
    cols = [d[0] for d in src_cur.description]
    data = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        pbn = id2bn_p.get(r['ptypeid'].strip()); bbn = id2bn_b.get(r['btypeid'].strip())
        if pbn is None or bbn is None:
            continue
        kid = r['ktypeid'].strip()
        if kid.startswith('99999'):                                              # 虚拟借出/借入仓
            ktype = kid
        else:
            kbn = id2bn_k.get(kid)
            ktype = str(kbn) if kbn is not None else kid
        data.append((
            bbn, pbn, _n(r['qty']), _n(r['price']),                              # UnitId/PTypeId/Quantity/Price
            _s(ktype, 50), _n(r['total']), path_p[pbn] + '-',                    # KTypeId/total/BasePath
        ))
    return data


# ---------------- 目标库写入 ----------------
def _bulk_insert(dst_cur, table, data, columns, batch=2000):
    if not data:
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.%s" % table)
        return 0
    ph = ','.join('?' * len(columns))
    sql = "SET NOCOUNT ON; INSERT INTO dbo.%s (%s) VALUES (%s)" % (table, ','.join(columns), ph)
    dst_cur.fast_executemany = False                                            # 跨版本逐行推断更稳
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
def _read_par_level(src_cur, table, idc):
    """读 {编号: 父编号}、{编号: 层级}"""
    src_cur.execute("SET NOCOUNT ON; SELECT %s, parid, leveal FROM dbo.%s" % (idc, table))
    parid_of, level_of = {}, {}
    for iid, parid, lv in src_cur.fetchall():
        k = iid.strip()
        parid_of[k] = parid or ''
        level_of[k] = lv or 0
    return parid_of, level_of


def export_to_db(src_cur, dst_cur, logcb):
    """管家婆 -> 千业 一键反向导出 v3.0"""
    pre = backup_dst(dst_cur)
    logcb('INFO', '备份', 0, '已备份目标 12 表（前缀 %s）' % pre)
    stats = {}
    # 1. 编号重排
    id2bn_b = rebase_map(src_cur, 'btype', 'btypeid', 'bcomment')
    id2bn_p = rebase_map(src_cur, 'ptype', 'ptypeid', 'pcomment')
    id2bn_k = rebase_map(src_cur, 'Stock', 'ktypeid', 'kcomment')
    id2bn_e = rebase_map(src_cur, 'employee', 'etypeid', 'ecomment')
    logcb('INFO', '映射', 0, '编号重排：单位 %d / 商品 %d / 仓库 %d / 职员 %d' %
          (len(id2bn_b), len(id2bn_p), len(id2bn_k), len(id2bn_e)))
    # 2. 各表 BasePath（通用生成）
    par_b, lv_b = _read_par_level(src_cur, 'btype', 'btypeid')
    par_p, lv_p = _read_par_level(src_cur, 'ptype', 'ptypeid')
    par_k, lv_k = _read_par_level(src_cur, 'Stock', 'ktypeid')
    par_e, lv_e = _read_par_level(src_cur, 'employee', 'etypeid')
    path_b = build_basepath(id2bn_b, par_b, lv_b)
    path_p = build_basepath(id2bn_p, par_p, lv_p)
    path_k = build_basepath(id2bn_k, par_k, lv_k)
    path_e = build_basepath(id2bn_e, par_e, lv_e)
    # 3. 价格方案解析（须在目标清空写之前，构造器内读目标旧名称）
    resolver = PriceSchemeResolver(src_cur, dst_cur)
    logcb('INFO', '价格', 0, '价格方案动态映射：%d 个 PRTypeId -> %d 个 IndexID 槽位' %
          (len(resolver.pr2ix), len(resolver.ix2pr)))
    # 4. 辅助单位 + 主表额外价格
    unit_info = load_unit_info(src_cur)
    inner_extra = load_inner_extra(src_cur, id2bn_p)
    # 5. 清空目标 12 表
    for t in EXPORT_TABLES:
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.%s" % t)
    # 6. 四主表
    cols_b = ['BaseNo', 'ParentNo', 'BaseLevel', 'BasePath', 'SonNum', 'BaseCode', 'Name', 'FullName',
              'AreaId', 'Address', 'UnitPhone', 'PostCode', 'LinkMan', 'LinkPhone', 'PyNo', 'TaxNum',
              'BankAndAcount', 'Memo', 'Email', 'QQNo', 'MsnNo', 'ARTotal', 'APTotal', 'ARLimit', 'APLimit',
              'IniARTotal', 'IniAPTotal', 'SkLimit', 'DeftPrice', 'ARWaringUp', 'PRWaringUp', 'ChgGoodDay',
              'ChgGoodScale', 'FaxNum', 'Deleted', 'RowIndex', 'SendGoods00', 'SendGoods', 'SendGoodSJ00',
              'SendGoodSJ', 'PersonID', 'VestID', 'IniPrepayARTotal', 'IniPrepayAPTotal', 'CurPrepayARTotal',
              'CurPrepayAPTotal', 'Stoped', 'Discount', 'SpaceDate']
    stats['BasUnitTypeInner'] = _bulk_insert(dst_cur, 'BasUnitTypeInner',
                                             export_btype(src_cur, id2bn_b, path_b), cols_b)

    cols_k = ['BaseNo', 'ParentNo', 'BaseLevel', 'SonNum', 'BasePath', 'FullName', 'Name', 'BaseCode',
              'Memo', 'deleted', 'PyNo', 'ROWINDEX', 'StType', 'Stoped', 'Onlintype']
    stats['BasStockInner'] = _bulk_insert(dst_cur, 'BasStockInner',
                                          export_stock(src_cur, id2bn_k, path_k), cols_k)

    cols_e = ['BaseNo', 'ParentNo', 'BaseLevel', 'BasePath', 'SonNum', 'SonCount', 'BaseCode', 'Name',
              'FullName', 'PyNo', 'Sex', 'DeptID', 'LinkTel', 'HandTel', 'Address', 'Memo', 'Deleted',
              'EMail', 'InJobDate', 'IDCard', 'CommissionPercent', 'Incumbency', 'Birthday', 'TopTotal',
              'EveryTotal', 'EducateLevel', 'RowIndex', 'KTypeID', 'ChangeKType', 'Stoped', 'APPType']
    stats['BasPersonInner'] = _bulk_insert(dst_cur, 'BasPersonInner',
                                           export_employee(src_cur, id2bn_e, path_e), cols_e)

    cols_p = ['BaseNo', 'ParentNo', 'BaseLevel', 'SonNum', 'SonCount', 'BasePath', 'BaseCode', 'Name',
              'FullName', 'Standards', 'Model', 'ProArea', 'BasicUnit', 'AssistUnitOne', 'AssistUnitTwo',
              'AssistUnitThree', 'SalePriceBasic', 'SalePriceOne', 'SalePriceTwo', 'SalePriceThree',
              'PercentBasic', 'PercentOne', 'PercentTwo', 'PercentThree', 'UseFullLifeDay', 'RecentPrice',
              'RecentPrice1', 'CostType', 'WaringUp', 'WaringDown', 'Deleted', 'PyNo', 'RowIndex', 'isSerial',
              'PChage', 'Memo', 'DeductType', 'DeductValue', 'MemberPrice', 'MinPrice', 'Preprice1',
              'Stoped', 'PtypePrice', 'BUnit', 'SUnit', 'OnLinType', 'OnlinPrice', 'gyid', 'szGoodsType']
    stats['BasPtypeInner'] = _bulk_insert(dst_cur, 'BasPtypeInner',
                                          export_ptype(src_cur, id2bn_p, path_p, unit_info, inner_extra),
                                          cols_p)
    # 7. 明细
    stats['BasPTypeBarMx'] = _bulk_insert(dst_cur, 'BasPTypeBarMx',
                                          export_barcodes(src_cur, id2bn_p, unit_info),
                                          ['PTypeID', 'SubUnit', 'IndexID', 'BarName', 'BarCode', 'IsDefault'])
    stats['BasPTypePriceMx'] = _bulk_insert(dst_cur, 'BasPTypePriceMx',
                                            export_ppm(src_cur, id2bn_p, unit_info, resolver),
                                            ['PTypeID', 'SubUnit', 'IndexID', 'AdvPrice',
                                             'DisPlayName', 'PyNo'])
    stats['price'] = _bulk_insert(dst_cur, 'price',
                                  export_pricefollow(src_cur, id2bn_p, id2bn_b),
                                  ['ptypeid', 'btypeid', 'costprice', 'saleprice', 'buydiscount',
                                   'salediscount', 'BuyDate', 'SaleDate', 'IsPrint', 'Units'])
    stats[DST_STOCK_TABLE] = _bulk_insert(dst_cur, DST_STOCK_TABLE,
                                          export_stockini(src_cur, id2bn_p, id2bn_k, path_p, path_k),
                                          ['PtypeId', 'KtypeId', 'JobNumber', 'OutFactoryDate', 'Qty',
                                           'Price', 'Total', 'GoodsOrder', 'PPath', 'KPath'])
    # 8. 委托受托（当前/期初）
    stats['StorageCommission'] = _bulk_insert(
        dst_cur, 'StorageCommission',
        export_commission(src_cur, id2bn_p, id2bn_b, path_p, 'Commission', 'cur'),
        ['PTYPEID', 'BTYPEID', 'TYPE', 'Quantity', 'ToTal', 'Prodate',
         'CostToTal', 'TaxPercent', 'TaxTotal', 'BasePath', 'BlockNo'])
    stats['StorageInitCommiss'] = _bulk_insert(
        dst_cur, 'StorageInitCommiss',
        export_commission(src_cur, id2bn_p, id2bn_b, path_p, 'IniCommission', 'ini'),
        ['PTYPEID', 'BTYPEID', 'TYPE', 'Quantity', 'ToTal', 'Prodate',
         'CostToTal', 'TaxPercent', 'BasePath', 'TaxTotal', 'BlockNo'])
    # 9. 借进借出（当前/期初）
    stats['BillQcLendAndBorrow'] = _bulk_insert(
        dst_cur, 'BillQcLendAndBorrow',
        export_lendborrow(src_cur, id2bn_p, id2bn_b, id2bn_k, path_p, 'lendborrow00'),
        ['UnitId', 'PTypeId', 'Quantity', 'Price', 'KTypeId', 'total', 'BasePath'])
    stats['BillLendOrBorrow'] = _bulk_insert(
        dst_cur, 'BillLendOrBorrow',
        export_lendborrow(src_cur, id2bn_p, id2bn_b, id2bn_k, path_p, 'lendborrow'),
        ['UnitId', 'PTypeId', 'Quantity', 'Price', 'KTypeId', 'total', 'BasePath'])
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
