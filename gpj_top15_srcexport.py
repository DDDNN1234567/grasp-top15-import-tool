# -*- coding: utf-8 -*-
"""
管家婆辉煌 TOP15 -> 千业商贸通 Bas 库 反向导出引擎（v1.1 新增）
把管家婆标准账套（ptype/btype/Stock/employee 等）按千业 Bas 结构重建并回写：
  Bas四主表（单位/商品/仓库/职员）：重排 BaseNo/ParentNo/BaseLevel，重建 BasePath 层级路径
  条码/辅助单位/预设价格（16 槽位）：编号映射后批量直写
  期初库存（当前结存）：写 StorageGoodsInfo
  委托/受托、借进/借出：同步对应业务表
关键修复（v3.0，针对“SQL 行数对但千业软件界面不显示”）：
  · 通用 build_basepath 按真实层级重建 BasePath（不依赖 BaseNo 大小）
  · 辅助单位空槽默认换算率写 0（不再误写 1）
  · 价格方案动态解析 PRTypeId↔IndexID（系统价 + 自定义价名称匹配/空槽分配）
  · 职员名称/性别回退、仓库职员补全、库存路径补全
"""
import os, sys, io, datetime
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except Exception:
        pass
import gpj_top15_core as core

# 管家婆 -> 千业（命令行 main 测试用；GUI 不依赖这些常量）
SRC_SERVER = r'DESKTOP-8P5G8VQ\SQLSERVER'   # 管家婆 TOP15 实例
DST_SERVER = r'DESKTOP-8P5G8VQ\SQLEXPRESS'   # 千业后端实例
SRC_DB = 'HH15new'
DST_DB = 'qysrc_test'
SRC_USER = 'sa'
SRC_PWD = os.environ.get('GPJ_SRC_PWD', '')

# 管家婆 PRTypeId -> 千业 IndexID（系统价固定锚点；自定义价由 PriceSchemeResolver 动态解析）
SYSTEM_PR_TO_INDEX = {'0005': None, '0007': None, '0103': None, '0001': 8}  # 零售固定 Index8

# 千业 Bas 价格槽位名称（用于与管家婆自定义价格名称匹配）
INDEX_NAMES = {
    1: '魏庙够多多', 2: '蟠桃够多多', 3: '乔家湖', 4: '金客隆', 5: '开元华联',
    6: '悦民坊', 7: '苏堤集市', 8: '零售价', 9: '双沟时代茂超市', 10: '现金价',
    11: '预设售价', 12: '柳新佳佳购', 13: '鹿庄新一佳', 14: '丰县百亿超市',
    15: '百亿招商场店', 16: '花语城直播基地',
}

# 虚拟仓库（借进/借出）：两软件编号一致
LENDOUT_STOCK = 9999999999
BORROWIN_STOCK = 9999999998


# ---------------- 编号映射（管家婆树编号 -> 千业 BaseNo） ----------------
def build_baseno_map(src_cur, table, idcol, levelcol):
    """管家婆树编号按层级重排为连续 BaseNo（leveal1 从 1 起，子节点按父分组连续）。
    返回 dict(树编号 -> BaseNo)，根(level<=0)映射 0。
    使用真实父子关系（树编号前缀），不依赖编号大小。"""
    src_cur.execute("SET NOCOUNT ON; SELECT %s, %s, ParId FROM dbo.%s WHERE %s>0 ORDER BY %s" %
                    (idcol, levelcol, table, levelcol, idcol))
    recs = [(i.strip(), int(lv), (p or '').strip()) for i, lv, p in src_cur.fetchall()]
    # 按树编号前缀确定父子（父编号 = 子编号去掉末 5 位）；level1 父=根
    id_set = {i for i, lv, p in recs}
    children = {}
    for i, lv, p in recs:
        if lv == 1 or p not in id_set:
            children.setdefault('__root__', []).append(i)
        else:
            children.setdefault(p, []).append(i)
    newno = {}
    counter = [0]

    def assign(parent_key, parent_baseno):
        kids = sorted(children.get(parent_key, []))
        for i in kids:
            counter[0] += 1
            bn = counter[0]
            newno[i] = bn
            assign(i, bn)
    assign('__root__', 0)
    return newno


def build_basepath(dst_cur, table, bn_map, idcol, namecol):
    """通用 BasePath 重建：按真实层级自顶向下生成千业路径（BasePath = 父 BasePath + '-' + 自己 BaseNo）。
    根 level0 路径='1'；level1 = '1-<BaseNo>'；level>1 = 父路径 + '-' + BaseNo。
    不依赖 BaseNo 大小，正确处理父编号大于子编号的情况。返回 dict(树编号 -> BasePath)。"""
    dst_cur.execute("SET NOCOUNT ON; SELECT %s, %s, ParId FROM dbo.%s WHERE %s>0" %
                    (idcol, namecol, table, 'leveal'))
    recs = {i.strip(): (int(lv), (p or '').strip()) for i, lv, p in dst_cur.fetchall()}
    paths = {}
    # 根路径固定 '1'
    def path_of(treeid):
        if treeid in paths:
            return paths[treeid]
        if treeid not in recs:
            return '1'
        lv, par = recs[treeid]
        bn = bn_map.get(treeid)
        if bn is None:
            return '1'
        if lv == 1 or par not in recs:
            paths[treeid] = '1-%d' % bn
        else:
            paths[treeid] = path_of(par) + '-%d' % bn
        return paths[treeid]
    for treeid in recs:
        path_of(treeid)
    return paths


# ---------------- 价格方案解析 ----------------
class PriceSchemeResolver:
    """动态解析管家婆 PRTypeId -> 千业 IndexID。
    系统价：0001 零售固定 Index8；0005/0007/0103 走主表语义（不占 PPM 槽位）。
    自定义价（0104+）：先按名称与千业槽位名称强匹配，未命中再分配空槽。"""

    def __init__(self, src_cur):
        self.pr_to_index = {'0001': 8}
        self.used = {8}
        # 读取管家婆价格类型名称
        src_cur.execute("SET NOCOUNT ON; SELECT PRTypeId, PRName FROM dbo.xw_P_PriceType WHERE IsSystem=0 ORDER BY PRTypeId")
        custom = [(p.strip(), (n or '').strip()) for p, n in src_cur.fetchall()]
        name_to_index = {v: k for k, v in INDEX_NAMES.items()}
        free_slots = [i for i in range(1, 17) if i not in self.used]
        for pr, name in custom:
            ix = None
            if name in name_to_index and name_to_index[name] not in self.used:
                ix = name_to_index[name]
            elif free_slots:
                ix = free_slots.pop(0)
            if ix is not None:
                self.pr_to_index[pr] = ix
                self.used.add(ix)

    def index(self, prtypeid):
        return self.pr_to_index.get((prtypeid or '').strip())


# ---------------- 主表导出 ----------------
def export_unit(src_cur, dst_cur, bn_map, paths):
    """btype -> BasUnitTypeInner（含 BasePath/层级/编号）"""
    src_cur.execute("SET NOCOUNT ON; SELECT btypeid, leveal, bfullname, bname, bcomment, bnamepy,"
                    " baddress, btel, btaxno, bbankacct FROM dbo.btype WHERE leveal>0 ORDER BY btypeid")
    data = []
    for tid, lv, full, name, comment, py, addr, tel, tax, bank in src_cur.fetchall():
        tid = tid.strip()
        bn = bn_map.get(tid)
        if bn is None:
            continue
        # 父 BaseNo
        pbn = 0
        if lv > 1 and len(tid) > 5:
            pbn = bn_map.get(tid[:-5], 0)
        data.append((bn, pbn, lv, str(bn), full or name or '', name or '',
                     py or '', paths.get(tid, '1'), 0, 0, comment or '', addr or '',
                     tel or '', '', '', tax or '', bank or '', 0, 0, 0, 0, 0, 0, 0, '', '', ''))
    dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.BasUnitTypeInner WHERE BaseNo>0")
    dst_cur.fast_executemany = True
    dst_cur.executemany("""
SET NOCOUNT ON; INSERT INTO dbo.BasUnitTypeInner
(BaseNo, ParentNo, BaseLevel, BaseCode, FullName, Name, PyNo, BasePath, Deleted, Stoped,
 Memo, Address, UnitPhone, LinkPhone, PostCode, TaxNum, BankAndAcount, IniARTotal, IniAPTotal,
 IniPrepayARTotal, IniPrepayAPTotal, ARLimit, APLimit, SkLimit, LinkMan, FaxNum, Email)
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", data)
    return len(data)


def export_stock(src_cur, dst_cur, bn_map, paths):
    """Stock -> BasStockInner（仓库；含根保留）"""
    src_cur.execute("SET NOCOUNT ON; SELECT ktypeid, leveal, kfullname, kname, kcomment, kaddress, kperson, ktel"
                    " FROM dbo.Stock WHERE leveal>0 ORDER BY ktypeid")
    data = []
    for tid, lv, full, name, comment, addr, person, tel in src_cur.fetchall():
        tid = tid.strip()
        bn = bn_map.get(tid)
        if bn is None:
            continue
        pbn = 0
        if lv > 1 and len(tid) > 5:
            pbn = bn_map.get(tid[:-5], 0)
        data.append((bn, pbn, lv, str(bn), full or name or '', name or '',
                     paths.get(tid, '1'), 0, 0, comment or '', addr or '', person or '', tel or ''))
    dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.BasStockInner WHERE BaseNo>0")
    dst_cur.fast_executemany = True
    dst_cur.executemany("""
SET NOCOUNT ON; INSERT INTO dbo.BasStockInner
(BaseNo, ParentNo, BaseLevel, BaseCode, FullName, Name, BasePath, Deleted, Stoped, Memo, Address, Person, Tel)
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", data)
    return len(data)


def export_employee(src_cur, dst_cur, bn_map, paths):
    """employee -> BasPersonInner（职员；名称回退、性别默认男）"""
    src_cur.execute("SET NOCOUNT ON; SELECT etypeid, leveal, efullname, ename, ecomment, etel, eduty,"
                    " eaddress, ebirthday, esex FROM dbo.employee WHERE leveal>0 ORDER BY etypeid")
    data = []
    for tid, lv, full, name, comment, tel, duty, addr, bday, sex in src_cur.fetchall():
        tid = tid.strip()
        bn = bn_map.get(tid)
        if bn is None:
            continue
        pbn = 0
        if lv > 1 and len(tid) > 5:
            pbn = bn_map.get(tid[:-5], 0)
        full = full or name or ''
        name = name or full
        sx = (sex or '男') or '男'
        data.append((bn, pbn, lv, str(bn), full, name, paths.get(tid, '1'),
                     0, 0, comment or '', tel or '', duty or '', addr or '', bday or '', sx))
    dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.BasPersonInner WHERE BaseNo>0")
    dst_cur.fast_executemany = True
    dst_cur.executemany("""
SET NOCOUNT ON; INSERT INTO dbo.BasPersonInner
(BaseNo, ParentNo, BaseLevel, BaseCode, FullName, Name, BasePath, Deleted, Stoped, Memo, LinkTel, Incumbency, Address, Birthday, Sex)
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", data)
    return len(data)


def export_goods(src_cur, dst_cur, bn_map, paths):
    """ptype -> BasPtypeInner（商品主表；价格/单位/条码走明细）"""
    src_cur.execute("SET NOCOUNT ON; SELECT ptypeid, leveal, pfullname, pname, pcomment, ptype, pstandard,"
                    " parea, pmodel, ptypetype, ppooldiscount, pLifeDay, pdefaulttax FROM dbo.ptype WHERE leveal>0 ORDER BY ptypeid")
    data = []
    for tid, lv, full, name, comment, typ, std, area, model, ptype, pool, lifeday, tax in src_cur.fetchall():
        tid = tid.strip()
        bn = bn_map.get(tid)
        if bn is None:
            continue
        pbn = 0
        if lv > 1 and len(tid) > 5:
            pbn = bn_map.get(tid[:-5], 0)
        data.append((bn, pbn, lv, str(bn), full or name or '', name or '', '',
                     paths.get(tid, '1'), 0, 0, comment or '', model or '', std or '',
                     area or '', typ or '', ptype or '', pool or 0, lifeday or '', tax or '',
                     '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''))
    dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.BasPtypeInner WHERE BaseNo>0")
    dst_cur.fast_executemany = True
    dst_cur.executemany("""
SET NOCOUNT ON; INSERT INTO dbo.BasPtypeInner
(BaseNo, ParentNo, BaseLevel, BaseCode, FullName, Name, PyNo, BasePath, Deleted, Stoped,
 Memo, Model, Standards, ProArea, Type, PtypeType, DeductValue, UseFullLifeDay, TaxRate,
 BasicUnit, AssistUnitOne, AssistUnitTwo, AssistUnitThree,
 PercentBasic, PercentOne, PercentTwo, PercentThree,
 SalePriceBasic, SalePriceOne, SalePriceTwo, SalePriceThree,
 RecentPrice, RecentPrice1, MinPrice, BarCode, SideUnit, pgManCode, snManCode, kwManCode,
 Pgholunit, Pgholunitrate, defaulttax, allowprice, pjobmancode,
 sourcewayzz, sourcewayzzDay, sourcewaywg, sourcewaywgDay, sourcewayww, sourcewaywwDay)
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", data)
    return len(data)


# ---------------- 明细导出 ----------------
def export_barcodes(src_cur, dst_cur, bn_map_goods):
    """xw_PtypeBarCode -> BasPTypeBarMx（UnitID 数字串；PTypeID 转 BaseNo）"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, BarCode, Ordid FROM dbo.xw_PtypeBarCode"
                    " WHERE BarCode IS NOT NULL AND LTRIM(RTRIM(BarCode))<>''")
    data = []
    for pid, bc, ordid in src_cur.fetchall():
        bn = bn_map_goods.get((pid or '').strip())
        if bn is None:
            continue
        data.append((bn, bc.strip(), ordid or 0))
    dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.BasPTypeBarMx WHERE 1=1")
    dst_cur.fast_executemany = True
    dst_cur.executemany(
        "SET NOCOUNT ON; INSERT INTO dbo.BasPTypeBarMx (PTypeID, BarCode, IndexID) VALUES (?,?,?)", data)
    return len(data)


def export_units(src_cur, dst_cur, bn_map_goods):
    """xw_PtypeUnit -> BasPtypeInner 单位列（基本/辅助单位 + 换算率）。
    按商品聚合：Ordid0=基本单位，1/2/3=辅助单位；空槽位换算率写 0（不写 1）。"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, Unit1, URate, Ordid FROM dbo.xw_PtypeUnit ORDER BY PTypeId, Ordid")
    per = {}
    for pid, unit, rate, ordid in src_cur.fetchall():
        per.setdefault((pid or '').strip(), {})[int(ordid or 0)] = ((unit or '').strip(), float(rate or 0))
    upd = []
    for pid, d in per.items():
        bn = bn_map_goods.get(pid)
        if bn is None:
            continue
        bu = d.get(0, ('', 0))
        a1 = d.get(1, ('', 0)); a2 = d.get(2, ('', 0)); a3 = d.get(3, ('', 0))
        upd.append((bu[0], a1[0], a2[0], a3[0],
                    bu[1] or 1, a1[1], a2[1], a3[1], bn))
    dst_cur.fast_executemany = True
    dst_cur.executemany("""
SET NOCOUNT ON; UPDATE dbo.BasPtypeInner SET
 BasicUnit=?, AssistUnitOne=?, AssistUnitTwo=?, AssistUnitThree=?,
 PercentBasic=?, PercentOne=?, PercentTwo=?, PercentThree=?
 WHERE BaseNo=?""", upd)
    return len(upd)


def export_prices(src_cur, dst_cur, bn_map_goods, resolver):
    """xw_P_PtypePrice -> BasPTypePriceMx（16 槽位）+ BasPtypeInner 系统价列。
    自定义/零售按 resolver 映射 IndexID；0005/0007/0103 回写主表 RecentPrice/RecentPrice1/MinPrice。"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeId, PRTypeId, Price FROM dbo.xw_P_PtypePrice")
    ppm = []
    inner = {}  # bn -> [recent, recent1, min]
    for pid, pr, price in src_cur.fetchall():
        bn = bn_map_goods.get((pid or '').strip())
        if bn is None:
            continue
        pr = (pr or '').strip()
        v = float(price or 0)
        if pr in ('0005', '0007', '0103'):
            a = inner.setdefault(bn, [0, 0, 0])
            if pr == '0005':
                a[0] = v
            elif pr == '0007':
                a[1] = v
            else:
                a[2] = v
        else:
            ix = resolver.index(pr)
            if ix is not None:
                ppm.append((bn, ix, v))
    dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.BasPTypePriceMx WHERE 1=1")
    dst_cur.fast_executemany = True
    dst_cur.executemany(
        "SET NOCOUNT ON; INSERT INTO dbo.BasPTypePriceMx (PTypeID, IndexID, AdvPrice) VALUES (?,?,?)", ppm)
    if inner:
        dst_cur.executemany(
            "SET NOCOUNT ON; UPDATE dbo.BasPtypeInner SET RecentPrice=?, RecentPrice1=?, MinPrice=? WHERE BaseNo=?",
            [(a[0], a[1], a[2], bn) for bn, a in inner.items()])
    return len(ppm), len(inner)


def export_storage_goods(src_cur, dst_cur, bn_map_goods, bn_map_stock, gpaths, kpaths):
    """IniGoodsStocks -> StorageGoodsInfo（当前结存；补 PPath/KPath，数量<>0）。"""
    src_cur.execute("SET NOCOUNT ON; SELECT PtypeId, KtypeId, Qty, Price, Total, JobNumber, OutFactoryDate"
                    " FROM dbo.IniGoodsStocks WHERE Qty<>0")
    data = []
    for pid, kid, qty, price, total, job, odate in src_cur.fetchall():
        pbn = bn_map_goods.get((pid or '').strip())
        kbn = bn_map_stock.get((kid or '').strip())
        if pbn is None or kbn is None:
            continue
        data.append((pbn, kbn, float(qty), float(price or 0), float(total or 0),
                     job or '', odate or '',
                     gpaths.get((pid or '').strip(), '1'), kpaths.get((kid or '').strip(), '1')))
    dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.StorageGoodsInfo WHERE 1=1")
    dst_cur.fast_executemany = True
    dst_cur.executemany("""
SET NOCOUNT ON; INSERT INTO dbo.StorageGoodsInfo
(PTypeId, KTypeId, Qty, Price, Total, JobNumber, OutFactoryDate, PPath, KPath)
VALUES (?,?,?,?,?,?,?,?,?)""", data)
    return len(data)


# ---------------- 业务表导出（委托/受托、借进/借出） ----------------
def export_commission(src_cur, dst_cur, bn_map_goods, bn_map_unit, bn_map_stock, gpaths):
    """管家婆 Commission/IniCommission -> 千业 StorageCommission/StorageInitCommiss（11 列）。
    TYPE 语义：1 受托 / 4 委托发出 / 5 委托退回；BasePath 取商品路径+'-'。"""
    total_n = 0
    for sp_t, dst_t in [('Commission', 'StorageCommission'), ('IniCommission', 'StorageInitCommiss')]:
        src_cur.execute("SET NOCOUNT ON; SELECT PTYPEID, BTYPEID, TYPE, QTY, TOTAL, PRODATE, BLOCKNO,"
                        " COSTTOTAL, tax, taxtotal FROM dbo.%s" % sp_t)
        rows = []
        for pid, bid, typ, qty, tot, prodate, block, cost, tax, taxtot in src_cur.fetchall():
            pbn = bn_map_goods.get((pid or '').strip())
            ubn = bn_map_unit.get((bid or '').strip())
            if pbn is None or ubn is None:
                continue
            rows.append((pbn, ubn, int(typ or 1), float(qty or 0), float(tot or 0),
                         str(prodate or ''), float(cost or 0), float(tax or 0), float(taxtot or 0),
                         gpaths.get((pid or '').strip(), '1') + '-', block or ''))
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.%s WHERE 1=1" % dst_t)
        dst_cur.fast_executemany = True
        dst_cur.executemany("""
SET NOCOUNT ON; INSERT INTO dbo.%s
(PTYPEID, BTYPEID, TYPE, Quantity, ToTal, Prodate, BlockNo, CostToTal, TaxPercent, TaxTotal, BasePath)
VALUES (?,?,?,?,?,?,?,?,?,?,?)""" % dst_t, rows)
        total_n += len(rows)
    return total_n


def export_lendborrow(src_cur, dst_cur, bn_map_goods, bn_map_unit, bn_map_stock, gpaths):
    """管家婆 lendborrow/lendborrow00 -> 千业 BillQcLendAndBorrow/BillLendOrBorrow（7 列）。
    ktypeid 转仓库 BaseNo；BasePath 取商品路径。"""
    total_n = 0
    for sp_t, dst_t in [('lendborrow', 'BillQcLendAndBorrow'), ('lendborrow00', 'BillLendOrBorrow')]:
        src_cur.execute("SET NOCOUNT ON; SELECT btypeid, ptypeid, qty, total, ktypeid, price FROM dbo.%s" % sp_t)
        rows = []
        for bid, pid, qty, tot, kid, price in src_cur.fetchall():
            pbn = bn_map_goods.get((pid or '').strip())
            ubn = bn_map_unit.get((bid or '').strip())
            kbn = bn_map_stock.get((kid or '').strip(), 0)
            if pbn is None or ubn is None:
                continue
            rows.append((ubn, pbn, float(qty), float(price or 0), kbn, float(tot or 0),
                         gpaths.get((pid or '').strip(), '1')))
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.%s WHERE 1=1" % dst_t)
        dst_cur.fast_executemany = True
        dst_cur.executemany("""
SET NOCOUNT ON; INSERT INTO dbo.%s
(UnitId, PTypeId, Quantity, Price, KTypeId, total, BasePath)
VALUES (?,?,?,?,?,?,?)""" % dst_t, rows)
        total_n += len(rows)
    return total_n


# ---------------- 编排 ----------------
def _backup_target(dst_cur):
    stamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    pre = '_expbak_' + stamp
    for t in EXPORT_TABLES:
        dst_cur.execute("SET NOCOUNT ON; SELECT * INTO %s%s FROM dbo.%s" % (t, pre, t))
    return pre


EXPORT_TABLES = ['BasUnitTypeInner', 'BasStockInner', 'BasPersonInner', 'BasPtypeInner',
                 'BasPTypeBarMx', 'BasPTypePriceMx', 'StorageGoodsInfo',
                 'StorageCommission', 'StorageInitCommiss', 'BillQcLendAndBorrow', 'BillLendOrBorrow', 'price']


def export_to_db(src_cur, dst_cur, logcb):
    """管家婆 -> 千业 一键反向导出。返回 (备份前缀, 统计dict)。"""
    pre = _backup_target(dst_cur)
    logcb('INFO', '备份', 0, '目标千业库已备份 %d 张表（前缀 %s）' % (len(EXPORT_TABLES), pre))

    # 1. 编号映射
    bn_unit = build_baseno_map(src_cur, 'btype', 'btypeid', 'leveal')
    bn_stock = build_baseno_map(src_cur, 'Stock', 'ktypeid', 'leveal')
    bn_emp = build_baseno_map(src_cur, 'employee', 'etypeid', 'leveal')
    bn_goods = build_baseno_map(src_cur, 'ptype', 'ptypeid', 'leveal')

    # 2. BasePath 路径（在源库结构上按真实层级重建）
    up = build_basepath(src_cur, 'btype', bn_unit, 'btypeid', 'bfullname')
    kp = build_basepath(src_cur, 'Stock', bn_stock, 'ktypeid', 'kfullname')
    ep = build_basepath(src_cur, 'employee', bn_emp, 'etypeid', 'efullname')
    gp = build_basepath(src_cur, 'ptype', bn_goods, 'ptypeid', 'pfullname')

    stats = {}
    # 3. 主表
    stats['BasUnitTypeInner'] = export_unit(src_cur, dst_cur, bn_unit, up)
    stats['BasStockInner'] = export_stock(src_cur, dst_cur, bn_stock, kp)
    stats['BasPersonInner'] = export_employee(src_cur, dst_cur, bn_emp, ep)
    stats['BasPtypeInner'] = export_goods(src_cur, dst_cur, bn_goods, gp)

    # 4. 价格方案解析（依赖主表价格类型）
    resolver = PriceSchemeResolver(src_cur)

    # 5. 明细
    stats['BasPTypeBarMx'] = export_barcodes(src_cur, dst_cur, bn_goods)
    stats['辅助单位'] = export_units(src_cur, dst_cur, bn_goods)
    n_ppm, n_inner = export_prices(src_cur, dst_cur, bn_goods, resolver)
    stats['BasPTypePriceMx'] = n_ppm
    stats['StorageGoodsInfo'] = export_storage_goods(src_cur, dst_cur, bn_goods, bn_stock, gp, kp)

    # 6. 业务表（委托/受托、借进/借出）
    stats['委托受托'] = export_commission(src_cur, dst_cur, bn_goods, bn_unit, bn_stock, gp)
    stats['借进借出'] = export_lendborrow(src_cur, dst_cur, bn_goods, bn_unit, bn_stock, gp)

    logcb('INFO', '导出', 0, '反向导出完成（备份前缀 %s，可回滚）' % pre)
    return pre, stats


def main(argv):
    logcb = lambda tag, m, ln, msg: print('[%s][%s] %s' % (tag, m, msg))
    src = core.connect(SRC_SERVER, SRC_DB, SRC_USER, SRC_PWD)
    dst = core.connect(DST_SERVER, DST_DB, SRC_USER, SRC_PWD)
    print('源 %s.%s -> 目标 %s.%s' % (SRC_SERVER, SRC_DB, DST_SERVER, DST_DB))
    pre, stats = export_to_db(src.cursor(), dst.cursor(), logcb)
    for k, v in stats.items():
        print('  %s = %d' % (k, v))
    src.close(); dst.close()
    print('完成。')


if __name__ == '__main__':
    main(sys.argv)
