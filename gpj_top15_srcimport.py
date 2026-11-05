# -*- coding: utf-8 -*-
"""
管家婆辉煌 TOP15 一键导入工具 - 源库直连引擎（v1.1 新增）
读取已附加的千业商贸通源库（qysrc_ccc / qysrc_qys / qysrc_test 的 Bas* 表）：
  Bas四表(单位/商品/仓库/职员) → 级联数字化 → 官方存储过程逐行导入 HH15new
  预设价格/条码/辅助单位/单位价格跟踪 → 批量直写（技能 §4B 已验证语义）
  期初库存 → 官方存储过程（p_hh_ModInitGoodsStocks_SNMng_Package）
依赖：gpj_top15_core（复用 connect/call_*/backup/rollback/clear_localdata/import_sheet）
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

SRC_SERVER = r'DESKTOP-8P5G8VQ\SQLEXPRESS'   # 千业软件后端实例（可读 E:\ 源库）
DST_SERVER = r'DESKTOP-8P5G8VQ\SQLSERVER'     # 管家婆 TOP15 实例
DST_DB = 'HH15new'
SRC_USER = 'sa'
SRC_PWD = os.environ.get('GPJ_SRC_PWD', '')  # 发布版：从环境变量读取，不在源码明文保存
SRC_DBS = {'CCC': 'qysrc_ccc', '千业测试': 'qysrc_qys', '测试': 'qysrc_test'}

# 库存数量来源（关键业务口径）：
#   StorageGoodsInfo = 当前结存（经全部进销单据后的实际库存，默认，导入新账套作为期初）
#   StorageIniGoods  = 历史期初（仅在需要原样复制建账期初时使用）
STOCK_SOURCE = 'StorageGoodsInfo'

# BasPtypeInner 价格列 -> 目标 PRTypeId（仅保留生产库实际有数据的列；SalePriceOne/Two/Three、MinPrice 实测全空）
INNER_COL_PR = [('SalePriceBasic', '0001'), ('RecentPrice', '0005'), ('RecentPrice1', '0007')]

# 千业 BasPTypePriceMx 16 槽位 -> 管家婆 PRTypeId（按名称语义；系统价 0001 + 自定义门店价 0104-0118）
# 依据：管家婆 xw_P_PriceType 已手工建 0104-0112（IsSystem=0），F_GetPriceInfo 以 IsSystem=0 动态读取
PPM_FULL_MAP = {
    1: ('0104', '魏庙够多多'), 2: ('0105', '蟠桃够多多'), 3: ('0106', '乔家湖'),
    4: ('0107', '金客隆'), 5: ('0108', '开元华联'), 6: ('0109', '悦民坊'),
    7: ('0110', '苏堤集市'), 8: ('0001', '零售价'), 9: ('0111', '双沟时代茂超市'),
    10: ('0112', '现金价'), 11: ('0113', '预设售价'), 12: ('0114', '柳新佳佳购'),
    13: ('0115', '鹿庄新一佳'), 14: ('0116', '丰县百亿超市'),
    15: ('0117', '百亿招商场店'), 16: ('0118', '花语城直播基地'),
}


def ensure_custom_pricetypes(dst_cur):
    """确保 xw_P_PriceType 含 16 方案所需全部价格类型：系统价跳过，自定义缺则插入
    （IsSystem=0/IsVisible=1/IsEdit=1/IsSet=1，updatetag=10000），F_GetPriceInfo 自动动态取价"""
    dst_cur.execute("SET NOCOUNT ON; SELECT PRTypeId FROM dbo.xw_P_PriceType")
    have = {r[0] for r in dst_cur.fetchall()}
    n = 0
    for ix, (pr, name) in PPM_FULL_MAP.items():
        if pr == '0001' or pr in have:
            continue
        dst_cur.execute("""SET NOCOUNT ON;
INSERT INTO dbo.xw_P_PriceType
 (PRTypeId,PRName,PRDisName,PRType,IsVisible,IsSystem,IsEdit,IsSet,Comment,URec,ParRec,USonnum,updatetag,Ordid,IsBaseVis)
 VALUES (?,?,?,0,1,0,1,1,'',0,0,0,10000,'',1)""", pr, name, name)
        n += 1
    return n


# ---------------- 级联数字化 ----------------
def digi_map(src_cur, db, table):
    """Bas 表级联数字化：BaseNo -> 新树编号（leveal1=00001起，leveal>1=父编号+5位行序）
    返回 dict(BaseNo -> newid)；根(level<=0)映射 '00000'（目标库保留自身根）。
    v3.0：基于真实父子关系 BFS 遍历，正确处理"父 BaseNo 大于子 BaseNo"（如 257 父=327），
    不再依赖 BaseNo 大小顺序；同一父下直接子节点按 BaseNo 升序分配行序。"""
    from collections import deque
    src_cur.execute("SET NOCOUNT ON; SELECT BaseNo, ParentNo, BaseLevel FROM %s.dbo.%s" % (db, table))
    pn_of, lv_of, child_map = {}, {}, {}
    for bn, pn, lv in src_cur.fetchall():
        bn = int(bn)
        pn_of[bn] = int(pn) if pn else 0
        lv_of[bn] = int(lv) if lv else 0
    roots = [bn for bn in pn_of if lv_of[bn] <= 0]
    newid = {bn: '00000' for bn in roots}
    for bn in pn_of:                                          # 按真实父节点分组
        if lv_of[bn] <= 0:
            continue
        pn = pn_of[bn] if pn_of[bn] in pn_of else (roots[0] if roots else 0)
        child_map.setdefault(pn, []).append(bn)
    queue = deque(roots)
    while queue:                                              # 自根向下 BFS
        parent = queue.popleft()
        for i, kid in enumerate(sorted(child_map.get(parent, [])), 1):
            if lv_of[parent] <= 0:                            # level1 不拼根前缀，纯5位
                newid[kid] = '%05d' % i
            else:                                             # level>1 父编号+5位行序
                newid[kid] = newid[parent] + '%05d' % i
            queue.append(kid)
    for bn in pn_of:                                          # 兜底：悬空/成环节点挂根
        if bn not in newid:
            newid[bn] = '00000'
    return newid


# ---------------- Bas -> 官方过程行（与 Excel 导入同构） ----------------
def build_unit_rows(src_cur, db, mp):
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM %s.dbo.BasUnitTypeInner ORDER BY BaseLevel, BaseNo" % db)
    cols = [d[0] for d in src_cur.description]
    rows = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        if r['BaseLevel'] is None or int(r['BaseLevel']) <= 0:
            continue
        comment = '来源:Bas;原编号:%s;%s' % (r.get('BaseCode') or '', r.get('Memo') or '')
        rows.append({
            '父类编号': mp.get(int(r['BaseNo']), '00000')[:-5] if int(r['BaseNo']) in mp and len(mp[int(r['BaseNo'])]) > 5 else '00000',
            '单位编号': mp.get(int(r['BaseNo']), ''),
            '单位全名': r.get('FullName') or r.get('Name') or '',
            '单位简名': r.get('Name') or '',
            '单位地址': r.get('Address') or '', '单位电话': r.get('UnitPhone') or '',
            '手机一': r.get('LinkPhone') or '', '手机二': '', '传真': r.get('FaxNum') or '',
            '邮编': r.get('PostCode') or '', '联系人一': r.get('LinkMan') or '', '联系人二': '',
            '开户银行': r.get('BankAndAcount') or '', '银行账号': '', '税号': r.get('TaxNum') or '',
            '期初应收款': r.get('IniARTotal') or 0, '期初应付款': r.get('IniAPTotal') or 0,
            '期初预收款': r.get('IniPrepayARTotal') or 0, '期初预付款': r.get('IniPrepayAPTotal') or 0,
            '应收款上限': r.get('ARLimit') or 0, '应付款上限': r.get('APLimit') or 0,
            '结算期限(天)': r.get('SkLimit') or '', '备注': comment,
            '换货期限(天)': r.get('ChgGoodDay') or '', '换货比例(%)': r.get('ChgGoodScale') or '',
            '电子邮件': r.get('Email') or '', '默认经手人(编号)': '', '地区(编号)': '',
            '结算单位(编号)': '', '适用价格': '',
        })
    return rows


def build_warehouse_rows(src_cur, db, mp):
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM %s.dbo.BasStockInner ORDER BY BaseLevel, BaseNo" % db)
    cols = [d[0] for d in src_cur.description]
    rows = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        if r['BaseLevel'] is None or int(r['BaseLevel']) <= 0:
            continue
        comment = '来源:Bas;原编号:%s;%s' % (r.get('BaseCode') or '', r.get('Memo') or '')
        rows.append({
            '父类编号': '00000',
            '仓库编号': mp.get(int(r['BaseNo']), ''),
            '仓库全名': r.get('FullName') or r.get('Name') or '',
            '仓库简称': r.get('Name') or '', '备注': comment,
            '仓库地址': '', '仓库负责人': '', '负责人电话': '',
        })
    return rows


def build_employee_rows(src_cur, db, mp):
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM %s.dbo.BasPersonInner ORDER BY BaseLevel, BaseNo" % db)
    cols = [d[0] for d in src_cur.description]
    rows = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        if r['BaseLevel'] is None or int(r['BaseLevel']) <= 0:
            continue
        comment = '来源:Bas;原编号:%s;%s' % (r.get('BaseCode') or '', r.get('Memo') or '')
        rows.append({
            '父类编号': '00000',
            '职员编号': mp.get(int(r['BaseNo']), ''),
            '职员全名': r.get('FullName') or r.get('Name') or '',
            '生日': r.get('Birthday') or '', '职员电话': r.get('LinkTel') or '',
            '职员职务': r.get('Incumbency') or '', '职员地址': r.get('Address') or '',
            '最低折扣下限': '', '每单优惠限额': r.get('EveryTotal') or '', '备注': comment,
        })
    return rows


def build_goods_rows(src_cur, db, mp, bar_map):
    """商品行：只传基本字段（价格/辅助单位走直写，避免官方过程写价格与直写冲突）"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM %s.dbo.BasPtypeInner ORDER BY BaseLevel, BaseNo" % db)
    cols = [d[0] for d in src_cur.description]
    rows = []
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        if r['BaseLevel'] is None or int(r['BaseLevel']) <= 0:
            continue
        bn = int(r['BaseNo'])
        newid = mp.get(bn, '')
        parent = '00000'
        if len(newid) > 5:
            parent = newid[:-5]
        comment = '来源:Bas;原编号:%s;%s' % (r.get('BaseCode') or '', (r.get('Memo') or '')[:200])
        barcode = bar_map.get(bn, '')
        rows.append({
            '父类编号': parent, '商品编号': newid, '商品全名': r.get('FullName') or r.get('Name') or '',
            '简名': r.get('Name') or '', '型号': r.get('Model') or '', '规格': r.get('Standards') or '',
            '产地': r.get('ProArea') or '', '成本算法': '', '商品属性': '',
            '联营扣率(%)': r.get('DeductValue') or '', '默认供应商(编号)': '', '品牌(编号)': '',
            '计价方式': '', '有效期(天)': r.get('UseFullLifeDay') or '', '备注': comment,
            '基本单位': r.get('BasicUnit') or '件', '条码': barcode,
            '零售价': '', '预设售价1': '', '预设售价2': '', '预设售价3': '',
            '配送价': '', '最近折后进价': '', '最低售价': '',
            '辅助单位1': '', '辅1换算关系': '', '辅1条码': '', '辅1零售价': '',
            '辅1预设售价1': '', '辅1预设售价2': '', '辅1预设售价3': '',
            '辅助单位2': '', '辅2换算关系': '', '辅2条码': '', '辅2零售价': '',
            '辅2预设售价1': '', '辅2预设售价2': '', '辅2预设售价3': '',
            '口味管理': '', '默认税率': '', '前台议价': '', '启用批次': '', '是否自制': '',
            '自制周期': '', '是否外购': '', '外购周期': '',
        })
    # 条码预清洗：源库同一条码被多个商品引用时，首个保留、其余置空（管家婆不允许条码重复占用）
    seen_bc = set()
    for r in rows:
        bc = (r.get('条码') or '').strip()
        if bc:
            if bc in seen_bc:
                r['条码'] = ''
            else:
                seen_bc.add(bc)
    return rows


def build_stockini_rows(src_cur, db, mp_p, mp_k, name_p, name_k):
    """库存结存：默认取 StorageGoodsInfo 当前结存（非历史期初；Qty<>0）-> 官方过程行
    两表结构一致；STOCK_SOURCE 可切换为 StorageIniGoods 特殊场景"""
    src_cur.execute("SET NOCOUNT ON; SELECT * FROM %s.dbo.%s WHERE Qty<>0 ORDER BY PtypeId, KtypeId"
                    % (db, STOCK_SOURCE))
    cols = [d[0] for d in src_cur.description]
    rows = []
    sign_fixed = 0
    for rec in src_cur.fetchall():
        r = dict(zip(cols, rec))
        p = mp_p.get(int(r['PtypeId']), '')
        k = mp_k.get(int(r['KtypeId']), '')
        if not p or not k:
            continue
        qty = float(r.get('Qty') or 0)
        total = float(r.get('Total') or 0)
        price = float(r.get('Price') or 0)
        # 符号对齐：负库存 + 移动加权可能产生"数量与金额异号/负单价"，管家婆官方过程要求二者同号
        # 以数量符号为准，按 |单价| 重算金额（边界数据，导入日志与文档会显式记录）
        if qty and (qty > 0) != (total > 0):
            base = abs(price) if price else (abs(total) / abs(qty) if total else 0.0)
            total = qty * base
            sign_fixed += 1
        rows.append({
            '商品编号': p, '商品全名': name_p.get(int(r['PtypeId']), ''),
            '仓库编号': k, '仓库全名': name_k.get(int(r['KtypeId']), ''),
            '数量': qty, '金额': total,
            '批号': r.get('JobNumber') or '', '生产日期': r.get('OutFactoryDate') or '',
            '有效期至': '',
        })
    if sign_fixed:
        print('  [符号修正] 负库存异号成本对齐 %d 条（以数量符号重算金额，请核对）' % sign_fixed)
    return rows


# ---------------- 明细直写 ----------------
def resolve_real_ids(dst_cur):
    """反查官方过程实际分配的树编号：pusercode(传入编号) -> ptypeid/btypeid/ktypeid/etypeid
    官方过程将传入编号存为 *usercode，实际树编号按过程规则分配（实测 pusercode=00001 ↔ ptypeid=00012）"""
    real = {}
    for t, idc, codec in [('ptype', 'ptypeid', 'pusercode'), ('btype', 'btypeid', 'busercode'),
                          ('Stock', 'ktypeid', 'kusercode'), ('employee', 'etypeid', 'eusercode')]:
        m = {}
        dst_cur.execute("SET NOCOUNT ON; SELECT %s, %s FROM dbo.%s WHERE leveal>0 AND %s IS NOT NULL" % (idc, codec, t, codec))
        for iid, cid in dst_cur.fetchall():
            m[cid.strip()] = iid.strip()
        real[t] = m
    return real
def _delete_scope(dst_cur, table, id_col, src_ids, incremental):
    """全量模式清空整表；增量模式只删源商品对应行（临时表限定），保留目标独有数据。"""
    if not incremental:
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.%s WHERE 1=1" % table)
        return
    dst_cur.execute("IF OBJECT_ID('tempdb..#srcids') IS NOT NULL DROP TABLE #srcids;"
                    " CREATE TABLE #srcids (id varchar(50))")
    dst_cur.fast_executemany = False
    dst_cur.executemany("INSERT INTO #srcids VALUES (?)", [(x,) for x in set(src_ids)])
    dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.%s WHERE [%s] IN (SELECT id FROM #srcids)" % (table, id_col))
    dst_cur.execute("DROP TABLE #srcids")


def write_barcodes(src_cur, dst_cur, db, mp, real=None, incremental=False):
    """BasPTypeBarMx -> xw_PtypeBarCode（UnitID='0' 数字串，Ordid=IndexID，keyid=NEWID）
    real: {映射编号->实际树编号} 反查表（缺省直用映射编号）"""
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeID, BarCode, IndexID FROM %s.dbo.BasPTypeBarMx WHERE BarCode IS NOT NULL AND LTRIM(RTRIM(BarCode))<>''" % db)
    rp = (real or {}).get('ptype', {})
    data = []
    for pt, bc, ix in src_cur.fetchall():
        p = mp.get(int(pt), '')
        if not p:
            continue
        p = rp.get(p, p)
        data.append((p, '0', bc.strip(), int(ix or 0)))
    _delete_scope(dst_cur, 'xw_PtypeBarCode', 'PTypeId', [d[0] for d in data], incremental)
    n = 0
    dst_cur.fast_executemany = True
    for i in range(0, len(data), 2000):
        chunk = data[i:i + 2000]
        dst_cur.executemany(
            "SET NOCOUNT ON; INSERT INTO dbo.xw_PtypeBarCode (PTypeId, UnitID, BarCode, Ordid, keyid, kwtypeid) VALUES (?,?,?,?,NEWID(),'')",
            chunk)
        n += len(chunk)
    return n


def write_units(src_cur, dst_cur, db, mp, real=None, incremental=False):
    """BasPtypeInner 基本/辅助单位 -> xw_PtypeUnit（Ordid=0/1/2/3）
    URate 取千业真实换算率列 PercentBasic/PercentOne/PercentTwo/PercentThree
    （UpdateDB.sql 权威佐证：Percent*=UnitRate*，1 辅单位=URate 基本单位；0/NULL 回退 1）"""
    src_cur.execute("SET NOCOUNT ON; SELECT BaseNo, BasicUnit, AssistUnitOne, AssistUnitTwo, AssistUnitThree,"
                    " PercentBasic, PercentOne, PercentTwo, PercentThree"
                    " FROM %s.dbo.BasPtypeInner WHERE BaseLevel>0" % db)
    rp = (real or {}).get('ptype', {})
    data = []

    def _rate(v):
        try:
            x = float(v)
            return x if x and x > 0 else 1.0
        except (TypeError, ValueError):
            return 1.0

    for bn, bu, a1, a2, a3, pb, r1, r2, r3 in src_cur.fetchall():
        p = mp.get(int(bn), '')
        if not p:
            continue
        p = rp.get(p, p)
        bu = (bu or '').strip()
        a1 = (a1 or '').strip(); a2 = (a2 or '').strip(); a3 = (a3 or '').strip()
        if bu:
            data.append((p, bu, _rate(pb), 1, 0))
        if a1 and a1 != bu:
            data.append((p, a1, _rate(r1), 0, 1))
        if a2 and a2 != bu and a2 != a1:
            data.append((p, a2, _rate(r2), 0, 2))
        if a3 and a3 != bu and a3 != a1 and a3 != a2:
            data.append((p, a3, _rate(r3), 0, 3))
    _delete_scope(dst_cur, 'xw_PtypeUnit', 'PTypeId', [d[0] for d in data], incremental)
    n = 0
    dst_cur.fast_executemany = True
    for i in range(0, len(data), 2000):
        chunk = data[i:i + 2000]
        dst_cur.executemany(
            "SET NOCOUNT ON; INSERT INTO dbo.xw_PtypeUnit (PTypeId, Unit1, URate, IsBase, Comment, Ordid, keyid) VALUES (?,?,?,?,'',?,NEWID())",
            chunk)
        n += len(chunk)
    return n


def write_prices(src_cur, dst_cur, db, mp, real=None, incremental=False):
    """预设价格完整 16 方案直写：
    1) 自动补齐缺失的自定义价格类型（IsSystem=0，F_GetPriceInfo 动态读取）
    2) BasPtypeInner：SalePriceBasic->0001 零售价 / RecentPrice->0005 最近进价 / RecentPrice1->0007 最近售价
    3) BasPTypePriceMx 16 槽位 -> PPM_FULL_MAP（门店价/零售价，按名称语义）
    UnitID='0'；引用用官方过程实际树编号（real 反查）；(PTypeId,PRTypeId) 去重；DELETE 全表幂等"""
    nc = ensure_custom_pricetypes(dst_cur)
    rp = (real or {}).get('ptype', {})
    seen = set()
    data = []
    # 1) BasPtypeInner 价格列
    src_cur.execute("SET NOCOUNT ON; SELECT BaseNo, SalePriceBasic, RecentPrice, RecentPrice1"
                    " FROM %s.dbo.BasPtypeInner WHERE BaseLevel>0" % db)
    for bn, spb, rcp, rp1 in src_cur.fetchall():
        p = rp.get(mp.get(int(bn), ''), '')
        if not p:
            continue
        for val, pr in [(spb, '0001'), (rcp, '0005'), (rp1, '0007')]:
            v = float(val or 0)
            if v != 0 and (p, pr) not in seen:
                seen.add((p, pr))
                data.append((p, '0', pr, v))
    # 2) BasPTypePriceMx 16 槽位
    src_cur.execute("SET NOCOUNT ON; SELECT PTypeID, IndexID, AdvPrice FROM %s.dbo.BasPTypePriceMx WHERE AdvPrice<>0" % db)
    for pt, ix, adv in src_cur.fetchall():
        p = rp.get(mp.get(int(pt), ''), '')
        m = PPM_FULL_MAP.get(int(ix or 0))
        if not p or not m:
            continue
        pr = m[0]
        if (p, pr) in seen:
            continue
        seen.add((p, pr))
        data.append((p, '0', pr, float(adv)))
    # 全量清空重写 / 增量只删源商品（幂等可重跑）
    _delete_scope(dst_cur, 'xw_P_PtypePrice', 'PTypeId', [d[0] for d in data], incremental)
    n = 0
    dst_cur.fast_executemany = True
    for i in range(0, len(data), 5000):
        chunk = data[i:i + 5000]
        dst_cur.executemany(
            "SET NOCOUNT ON; INSERT INTO dbo.xw_P_PtypePrice (PTypeId, UnitID, PRTypeId, Price, UpDateDt, KeyId) VALUES (?,?,?,?,GETDATE(),NEWID())",
            chunk)
        n += len(chunk)
    return n, nc


def write_pricefollow(src_cur, dst_cur, db, mp_p, mp_b, real=None, incremental=False):
    """price 表(千业客户价格) -> xw_P_PriceFollow：每条记录展开 4 行
    (0005 最近进价 / 0006 最近进货折扣 / 0007 最近售价 / 0008 最近销售折扣)，PPType=6
    按 (ptypeid,btypeid) 去重（源库 Units=-1/1 重复记录）；引用用官方过程实际树编号"""
    src_cur.execute("""
SET NOCOUNT ON;
SELECT ptypeid, btypeid,
       MAX(costprice) AS cp, MAX(saleprice) AS sp,
       MAX(buydiscount) AS bd, MAX(salediscount) AS sd
FROM %s.dbo.price GROUP BY ptypeid, btypeid""" % db)
    rp = (real or {}).get('ptype', {})
    rb = (real or {}).get('btype', {})
    data = []
    for pt, bt, cp, sp, bd, sd in src_cur.fetchall():
        p = rp.get(mp_p.get(int(pt), ''), '')
        b = rb.get(mp_b.get(int(bt), ''), '')
        if not p or not b:
            continue
        if cp or sp or (bd and float(bd) != 0) or (sd and float(sd) != 0):
            data.append((p, b, '0', cp or 0, '0005'))
            data.append((p, b, '0', float(bd or 0), '0006'))
            data.append((p, b, '0', sp or 0, '0007'))
            data.append((p, b, '0', float(sd or 0), '0008'))
    if incremental:
        dst_cur.execute("IF OBJECT_ID('tempdb..#srcids') IS NOT NULL DROP TABLE #srcids;"
                        " CREATE TABLE #srcids (id varchar(50))")
        dst_cur.fast_executemany = False
        dst_cur.executemany("INSERT INTO #srcids VALUES (?)", [(x,) for x in set(d[0] for d in data)])
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.xw_P_PriceFollow WHERE PPType=6 AND PTypeId IN (SELECT id FROM #srcids)")
        dst_cur.execute("DROP TABLE #srcids")
    else:
        dst_cur.execute("SET NOCOUNT ON; DELETE FROM dbo.xw_P_PriceFollow WHERE PPType=6")
    n = 0
    dst_cur.fast_executemany = True
    for i in range(0, len(data), 10000):
        chunk = data[i:i + 10000]
        dst_cur.executemany(
            "SET NOCOUNT ON; INSERT INTO dbo.xw_P_PriceFollow (PTypeId, BTypeId, UnitID, Price, KTypeId, BillDt, PRTypeId, PPType, KeyId, JGTypeId, MDTypeId) VALUES (?,?,?,?,'','1900-01-01',?,6,NEWID(),'','')",
            chunk)
        n += len(chunk)
    return n


# ---------------- 主表字段动态补全（官方过程未覆盖字段） ----------------
# 规格：(目标表, 源Bas表, real键, 目标id列, [(源列, 目标列, 类型 num/bit), ...])
# 仅纳入语义明确的字段；源有目标无的字段走备注追加（COMMENT_SPECS）。
FILL_SPECS = [
    ('btype', 'BasUnitTypeInner', 'btype', 'btypeid', [
        ('ARTotal', 'ARTotal', 'num'),              # 当前应收（区别于期初 ARTotal00）
        ('APTotal', 'APTotal', 'num'),              # 当前应付
        ('CurPrepayARTotal', 'YRTotal', 'num'),     # 当前预收
        ('CurPrepayAPTotal', 'YPTotal', 'num'),     # 当前预付
        ('ARWaringUp', 'R_WARNUP', 'num'),
        ('PRWaringUp', 'P_WARNUP', 'num'),
        ('Stoped', 'isStop', 'bit'),
    ]),
    ('ptype', 'BasPtypeInner', 'ptype', 'ptypeid', [
        ('RecentPrice', 'LastBuyPrice', 'num'),     # 最近进价（主表冗余列）
        ('WaringUp', 'warnup', 'num'),
        ('WaringDown', 'warndown', 'num'),
        ('Stoped', 'isStop', 'bit'),
    ]),
    ('Stock', 'BasStockInner', 'Stock', 'ktypeid', [
        ('Stoped', 'isStop', 'bit'),
    ]),
    ('employee', 'BasPersonInner', 'employee', 'etypeid', [
        ('Stoped', 'isStop', 'bit'),
    ]),
]

# 源有目标无、追加到备注的字段：(目标表, 源表, real键, id列, 备注列, [(源列, 中文标签)])
COMMENT_SPECS = [
    ('employee', 'BasPersonInner', 'employee', 'etypeid', 'ecomment',
     [('Sex', '性别'), ('IDCard', '身份证'), ('EducateLevel', '学历')]),
    ('btype', 'BasUnitTypeInner', 'btype', 'btypeid', 'bcomment',
     [('QQNo', 'QQ'), ('MsnNo', 'MSN')]),
]


def _conv(v, t):
    if t == 'bit':
        return 1 if v in (True, 1, '1') else 0
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def fill_main_fields(src_cur, dst_cur, db, mp, real, spec, logcb=None):
    """单个目标表的字段补全：读源→映射实际树编号→临时表→一条 JOIN UPDATE 批量更新。返回更新行数。"""
    target, src_table, real_key, id_col, cols = spec
    real_map = (real or {}).get(real_key, {})
    src_cols = ', '.join('[%s]' % s for s, _, _ in cols)
    src_cur.execute("SET NOCOUNT ON; SELECT BaseNo, %s FROM %s.dbo.%s" % (src_cols, db, src_table))
    records = []
    for row in src_cur.fetchall():
        code = mp.get(int(row[0]))
        if not code:
            continue
        tid = real_map.get(code.strip(), code.strip())
        records.append((tid,) + tuple(_conv(row[i], t) for i, (_, _, t) in enumerate(cols, start=1)))
    if not records:
        return 0
    tgt_cols = [tg for _, tg, _ in cols]
    coldefs = ', '.join('[%s] %s' % (tg, ('int' if t == 'bit' else 'numeric(18,4)'))
                        for _, tg, t in cols)
    dst_cur.execute("IF OBJECT_ID('tempdb..#fill') IS NOT NULL DROP TABLE #fill;"
                    " CREATE TABLE #fill (id varchar(50) NOT NULL, %s)" % coldefs)
    qmarks = ', '.join(['?'] * (len(cols) + 1))
    # 不用 fast_executemany：它对临时表 numeric 列的批量绑定在混合类型下会误报 8114；
    # 普通 executemany 实测 883 行仅 0.06s，24k 行也在亚秒级，足够快。
    dst_cur.fast_executemany = False
    dst_cur.executemany("INSERT INTO #fill VALUES (%s)" % qmarks, records)
    setclause = ', '.join('t.[%s]=f.[%s]' % (tg, tg) for tg in tgt_cols)
    dst_cur.execute("UPDATE t SET %s FROM dbo.%s t JOIN #fill f ON t.[%s]=f.id"
                    % (setclause, target, id_col))
    dst_cur.execute("DROP TABLE #fill")
    # UPDATE...FROM JOIN 在 autocommit 下 rowcount 常返回 -1，无参考价值；返回候选行数
    return len(records)


def append_comment_fields(src_cur, dst_cur, db, mp, real, spec, logcb=None):
    """把源有目标无的字段以“标签:值”追加到备注（幂等：已含标签不重复）。返回追加行数。"""
    target, src_table, real_key, id_col, comment_col, fields = spec
    real_map = (real or {}).get(real_key, {})
    src_cols = ', '.join('[%s]' % s for s, _ in fields)
    src_cur.execute("SET NOCOUNT ON; SELECT BaseNo, %s FROM %s.dbo.%s" % (src_cols, db, src_table))
    n = 0
    for row in src_cur.fetchall():
        code = mp.get(int(row[0]))
        if not code:
            continue
        tid = real_map.get(code.strip(), code.strip())
        pieces = []
        for i, (s, label) in enumerate(fields, start=1):
            v = row[i]
            if v is not None and str(v).strip() not in ('', '0'):
                pieces.append('%s:%s' % (label, str(v).strip()))
        if not pieces:
            continue
        dst_cur.execute("SET NOCOUNT ON; SELECT [%s] FROM dbo.%s WHERE [%s]=?"
                        % (comment_col, target, id_col), tid)
        rr = dst_cur.fetchone()
        if not rr:
            continue
        cur_comment = rr[0] or ''
        add = '; '.join(p for p in pieces if p.split(':')[0] not in cur_comment)
        if not add:
            continue
        new_comment = (cur_comment + '; ' + add) if cur_comment else add
        dst_cur.execute("SET NOCOUNT ON; UPDATE dbo.%s SET [%s]=? WHERE [%s]=?"
                        % (target, comment_col, id_col), new_comment, tid)
        n += 1
    return n


# ---------------- 增量同步（基于 usercode/BaseNo 集合差异） ----------------
# 模块 -> (目标表, usercode列, 行内编号key, 可更新字段[(行key, 目标列)])
INCREMENTAL_META = {
    'unit': ('btype', 'busercode', '单位编号', [
        ('单位全名', 'bfullname'), ('单位简名', 'bname'), ('单位电话', 'TelAndAddress'),
        ('手机一', 'moPhone'), ('传真', 'Fax'), ('联系人一', 'Person'), ('备注', 'bcomment')]),
    'goods': ('ptype', 'pusercode', '商品编号', [
        ('商品全名', 'pfullname'), ('简名', 'pname'), ('规格', 'Standard'), ('备注', 'pcomment')]),
    'warehouse': ('Stock', 'kusercode', '仓库编号', [
        ('仓库全名', 'kfullname'), ('仓库简称', 'kname'), ('备注', 'kcomment')]),
    'employee': ('employee', 'eusercode', '职员编号', [
        ('职员全名', 'efullname'), ('备注', 'ecomment')]),
}


def existing_usercodes(dst_cur, table, uc_col):
    """目标库当前已有的 usercode 集合（=已导入的源 BaseNo 传入编号）"""
    dst_cur.execute("SET NOCOUNT ON; SELECT [%s] FROM dbo.%s WHERE [%s] IS NOT NULL AND [%s]<>''"
                    % (uc_col, table, uc_col, uc_col))
    return set(r[0] for r in dst_cur.fetchall())


def split_incremental(rows, id_key, existing):
    """按 usercode 差异分流：新增（目标没有）/ 更新（目标已有）。目标独有行不在此处理，自动保留。"""
    new_rows, upd_rows = [], []
    for r in rows:
        (upd_rows if r.get(id_key) in existing else new_rows).append(r)
    return new_rows, upd_rows


def update_existing_rows(dst_cur, module, upd_rows):
    """共有行：只更新名称/联系/备注等可变字段，不动树编号与层级。返回更新行数。"""
    table, uc_col, id_key, fields = INCREMENTAL_META[module]
    setclause = ', '.join('[%s]=?' % tc for _, tc in fields)
    sql = "SET NOCOUNT ON; UPDATE dbo.%s SET %s, updatetag=10000 WHERE [%s]=?" % (table, setclause, uc_col)
    n = 0
    for r in upd_rows:
        vals = [r.get(k) or '' for k, _ in fields]
        dst_cur.execute(sql, vals + [r[id_key]])
        n += 1
    return n


# ---------------- 编排 ----------------
def import_from_db(src_cur, dst_cur, db, modules, logcb, incremental=False):
    """源库直连一键导入：备份 → 四表官方过程 → 明细直写 → 字段补全 → 期初库存 → 归位"""
    # 1. 备份
    pre = core.backup_tables(dst_cur)
    logcb('INFO', '备份', 0, '已备份 %d 张表（含业务表，前缀 %s）' % (len(core.BAK_TABLES), pre))
    # 1b. 已开账账套：官方过程拒绝录入“期初应收应付/期初库存”。在所有官方过程之前统一反开账，
    #     全部完成（含异常）后在 finally 恢复，保证单位/库存等期初数据均可录入。
    dst_cur.execute("SET NOCOUNT ON; SELECT SubValue FROM dbo.sysdata WHERE SubName='iniover'")
    _r = dst_cur.fetchone()
    _was_open = _r[0] if _r else '0'
    if _was_open == '1':
        dst_cur.execute("SET NOCOUNT ON; UPDATE dbo.sysdata SET SubValue='0' WHERE SubName='iniover'")
        logcb('INFO', '开账', 0, '检测到账套已开账，导入前统一反开账（已确认无业务单据，安全），完成后自动恢复')
    try:
        # 2. 级联数字化
        logcb('INFO', '映射', 0, '级联数字化 Bas 编号...')
        mp_u = digi_map(src_cur, db, 'BasUnitTypeInner')
        mp_p = digi_map(src_cur, db, 'BasPtypeInner')
        mp_k = digi_map(src_cur, db, 'BasStockInner')
        mp_e = digi_map(src_cur, db, 'BasPersonInner')
        # 默认条码映射（IsDefault=1 优先）
        src_cur.execute("SET NOCOUNT ON; SELECT PTypeID, BarCode, IndexID FROM %s.dbo.BasPTypeBarMx WHERE BarCode IS NOT NULL AND LTRIM(RTRIM(BarCode))<>'' ORDER BY IndexID" % db)
        bar_map = {}
        for pt, bc, ix in src_cur.fetchall():
            bar_map.setdefault(int(pt), bc.strip())
        # 3. 四表官方过程（复用 core.import_sheet 通道与日志）
        resolver = core.ParentResolver(dst_cur)
        def _prep(module, rows_all):
            if not incremental:
                return rows_all
            tbl, ucc, idk, _ = INCREMENTAL_META[module]
            exist = existing_usercodes(dst_cur, tbl, ucc)
            new_rows, upd_rows = split_incremental(rows_all, idk, exist)
            nu = update_existing_rows(dst_cur, module, upd_rows)
            logcb('INFO', module, 0, '增量分流：更新 %d 行、待新增 %d 行（目标独有数据保留不删）' % (nu, len(new_rows)))
            return new_rows

        if 'unit' in modules:
            rows = _prep('unit', build_unit_rows(src_cur, db, mp_u))
            logcb('INFO', 'unit', 0, '── 开始导入单位（新增 %d 行）──' % len(rows))
            ok, fail, _ = core.import_sheet(dst_cur, 'unit', rows, logcb, resolver)
            logcb('INFO', 'unit', 0, '── 单位完成：成功 %d，失败 %d ──' % (ok, fail))
        if 'warehouse' in modules:
            rows = _prep('warehouse', build_warehouse_rows(src_cur, db, mp_k))
            logcb('INFO', 'warehouse', 0, '── 开始导入仓库（新增 %d 行）──' % len(rows))
            ok, fail, _ = core.import_sheet(dst_cur, 'warehouse', rows, logcb, resolver)
            logcb('INFO', 'warehouse', 0, '── 仓库完成：成功 %d，失败 %d ──' % (ok, fail))
        if 'employee' in modules:
            rows = _prep('employee', build_employee_rows(src_cur, db, mp_e))
            logcb('INFO', 'employee', 0, '── 开始导入职员（新增 %d 行）──' % len(rows))
            ok, fail, _ = core.import_sheet(dst_cur, 'employee', rows, logcb, resolver)
            logcb('INFO', 'employee', 0, '── 职员完成：成功 %d，失败 %d ──' % (ok, fail))
        if 'goods' in modules:
            rows = _prep('goods', build_goods_rows(src_cur, db, mp_p, bar_map))
            logcb('INFO', 'goods', 0, '── 开始导入商品（新增 %d 行，价格/辅助单位/条码走直写）──' % len(rows))
            ok, fail, _ = core.import_sheet(dst_cur, 'goods', rows, logcb, resolver)
            logcb('INFO', 'goods', 0, '── 商品完成：成功 %d，失败 %d ──' % (ok, fail))
        # 4. 明细直写（引用官方过程实际树编号：pusercode 反查）
        real = resolve_real_ids(dst_cur)
        if 'barcode' in modules:
            n = write_barcodes(src_cur, dst_cur, db, mp_p, real, incremental)
            logcb('INFO', 'barcode', 0, '条码直写完成：%d 行' % n)
        if 'assistunit' in modules:
            n = write_units(src_cur, dst_cur, db, mp_p, real, incremental)
            logcb('INFO', 'assistunit', 0, '辅助单位直写完成：%d 行' % n)
        if 'price' in modules:
            n, nc = write_prices(src_cur, dst_cur, db, mp_p, real, incremental)
            logcb('INFO', 'price', 0, '预设价格直写完成：%d 行（自动新建自定义价格类型 %d 个，共 16 方案）' % (n, nc))
        if 'pricefollow' in modules:
            n = write_pricefollow(src_cur, dst_cur, db, mp_p, mp_u, real, incremental)
            logcb('INFO', 'pricefollow', 0, '单位价格跟踪直写完成：%d 行' % n)
        # 4b. 主表字段动态补全（官方过程未覆盖：当前应收应付/最近进价/停用/预警等）
        mp_by_real = {'btype': mp_u, 'ptype': mp_p, 'Stock': mp_k, 'employee': mp_e}
        for spec in FILL_SPECS:
            n = fill_main_fields(src_cur, dst_cur, db, mp_by_real[spec[2]], real, spec, logcb)
            if n:
                logcb('INFO', '补全', 0, '%s 动态补全 %d 行（%s）' %
                      (spec[0], n, '、'.join(c[1] for c in spec[4])))
        for spec in COMMENT_SPECS:
            n = append_comment_fields(src_cur, dst_cur, db, mp_by_real[spec[2]], real, spec, logcb)
            if n:
                logcb('INFO', '补全', 0, '%s 备注追加 %d 行' % (spec[0], n))
        # 5. 期初库存（官方过程，依赖商品/仓库已导入）
        if 'stockini' in modules:
            name_p = {}
            name_k = {}
            src_cur.execute("SET NOCOUNT ON; SELECT BaseNo, FullName FROM %s.dbo.BasPtypeInner" % db)
            for bn, fn in src_cur.fetchall():
                name_p[int(bn)] = fn or ''
            src_cur.execute("SET NOCOUNT ON; SELECT BaseNo, FullName FROM %s.dbo.BasStockInner" % db)
            for bn, fn in src_cur.fetchall():
                name_k[int(bn)] = fn or ''
            rows = build_stockini_rows(src_cur, db, mp_p, mp_k, name_p, name_k)
            # 反开账已在流程开头统一处理（iniover=0），此处直接走官方库存过程
            logcb('INFO', 'stockini', 0, '── 开始导入期初库存（%d 行，取当前结存 StorageGoodsInfo）──' % len(rows))
            ok, fail, _ = core.import_sheet(dst_cur, 'stockini', rows, logcb)
            logcb('INFO', 'stockini', 0, '── 期初库存完成：成功 %d，失败 %d ──' % (ok, fail))
        # 6. 归位：主表 updatetag=10000 + xw_BaseUpdateTag=0（官方过程已自动维护，兜底执行）
        dst_cur.execute("SET NOCOUNT ON; UPDATE dbo.btype SET updatetag=10000 WHERE updatetag<10000;"
                        "UPDATE dbo.ptype SET updatetag=10000 WHERE updatetag<10000;"
                        "UPDATE dbo.Stock SET updatetag=10000 WHERE updatetag<10000;"
                        "UPDATE dbo.employee SET updatetag=10000 WHERE updatetag<10000;"
                        "UPDATE dbo.xw_BaseUpdateTag SET UpdateTag=0")
        # 6b. 单位拼音补齐（健壮化：函数缺失/异常绝不阻断主流程，源 PyNo 兜底）
        try:
            dst_cur.execute("SET NOCOUNT ON; SELECT OBJECT_ID('dbo.fn_GetPyCode')")
            if dst_cur.fetchone()[0] is not None:
                dst_cur.execute("SET NOCOUNT ON; UPDATE dbo.btype SET bnamepy=dbo.fn_GetPyCode(bfullname)"
                                " WHERE leveal>0 AND (bnamepy IS NULL OR bnamepy='')")
                logcb('INFO', '归位', 0, '单位拼音补齐完成（数据库函数）')
            else:
                raise RuntimeError('目标库缺少 fn_GetPyCode')
        except Exception as pe:
            try:
                src_cur.execute("SET NOCOUNT ON; SELECT BaseNo, PyNo FROM %s.dbo.BasUnitTypeInner" % db)
                rp = real.get('btype', {})
                recs = []
                for bn, pyno in src_cur.fetchall():
                    code = mp_u.get(int(bn))
                    if code and pyno and str(pyno).strip():
                        recs.append((str(pyno).strip(), rp.get(code.strip(), code.strip())))
                dst_cur.fast_executemany = True
                dst_cur.executemany("SET NOCOUNT ON; UPDATE dbo.btype SET bnamepy=? WHERE btypeid=? "
                                    "AND (bnamepy IS NULL OR bnamepy='')", recs)
                logcb('WARN', '归位', 0, '数据库拼音函数不可用，已用源 PyNo 兜底 %d 行（原因：%s）' % (len(recs), pe))
            except Exception as pe2:
                logcb('WARN', '归位', 0, '拼音补齐未完成（不影响核心数据，可在管家婆内手工生成）：%s' % pe2)
        logcb('INFO', '归位', 0, 'updatetag 归位完成（备份前缀 %s，可在管家婆内核对后回滚）' % pre)
        return pre
    except Exception as e:
        logcb('FAIL', '导入', 0, '导入中止：%s（备份前缀 %s 可回滚）' % (e, pre))
        raise
    finally:
        # 无论成功或异常，恢复账套原开账状态
        if _was_open == '1':
            dst_cur.execute("SET NOCOUNT ON; UPDATE dbo.sysdata SET SubValue='1' WHERE SubName='iniover'")
            logcb('INFO', '开账', 0, '已恢复开账状态（iniover=1）')


MODULES = ['unit', 'warehouse', 'goods', 'employee', 'stockini', 'price', 'barcode', 'assistunit', 'pricefollow']
MODULE_CN = {'unit': '单位', 'warehouse': '仓库', 'goods': '商品', 'employee': '职员',
             'stockini': '期初库存', 'price': '预设价格', 'barcode': '条码',
             'assistunit': '辅助单位', 'pricefollow': '客户价格追踪'}


def main(argv):
    db = argv[1] if len(argv) > 1 else 'qysrc_ccc'
    mods = argv[2].split(',') if len(argv) > 2 else MODULES
    logcb = lambda tag, m, ln, msg: print('[%s][%s] %s' % (tag, m, msg))
    src = core.connect(SRC_SERVER, db, SRC_USER, SRC_PWD)
    dst = core.connect(DST_SERVER, DST_DB, SRC_USER, SRC_PWD)
    print('源库 %s / 目标 %s.%s' % (db, DST_SERVER, DST_DB))
    core.check_account(dst.cursor())
    import_from_db(src.cursor(), dst.cursor(), db, mods, logcb)
    src.close(); dst.close()
    print('完成。')


if __name__ == '__main__':
    main(sys.argv)
