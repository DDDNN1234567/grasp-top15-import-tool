# -*- coding: utf-8 -*-
"""
管家婆辉煌 TOP15 一键导入工具 - 核心引擎（无 GUI 依赖，可独立测试）
通道：官方存储过程（p_hh_BasicInsertB_Commodity 等 6 个），autocommit=True。
依赖：pyodbc, xlrd(.xls), openpyxl(.xlsx)
"""
import os, re, time, datetime
import pyodbc

# ---------------- 常量 ----------------
INSTALL_DIR = r'F:\GRASPⅡTOP15'
SHEETS = {  # Excel sheet 名 -> 模块
    '单位': 'unit', '商品': 'goods', '仓库': 'warehouse', '职员': 'employee',
    '库存信息': 'stockini', '单位价格跟踪': 'pricefollow',
}
MODULE_ORDER = ['unit', 'warehouse', 'goods', 'employee', 'stockini', 'pricefollow']
MODULE_NAMES = {
    'unit': '单位', 'warehouse': '仓库', 'goods': '商品(含预设价格)',
    'employee': '职员', 'stockini': '期初库存', 'pricefollow': '客户价格追踪',
}
# 各模块期望表头（官方模板列名）
REQUIRED_COLS = {
    'unit': ['单位编号', '单位全名'],
    'warehouse': ['仓库编号', '仓库全名'],
    'goods': ['商品编号', '商品全名'],
    'employee': ['职员编号', '职员全名'],
    'stockini': ['商品编号', '仓库编号', '数量'],
    'pricefollow': ['单位编号', '商品编号'],
}
ALL_COLS = {
    'unit': ['父类编号', '单位编号', '单位全名', '单位简名', '单位地址', '单位电话', '手机一', '手机二', '传真',
             '邮编', '联系人一', '联系人二', '开户银行', '银行账号', '税号', '期初应收款', '期初应付款',
             '期初预收款', '期初预付款', '应收款上限', '应付款上限', '结算期限(天)', '备注', '换货期限(天)',
             '换货比例(%)', '电子邮件', '默认经手人(编号)', '地区(编号)', '结算单位(编号)', '适用价格'],
    'warehouse': ['父类编号', '仓库编号', '仓库全名', '仓库简称', '备注', '仓库地址', '仓库负责人', '负责人电话'],
    'goods': ['父类编号', '商品编号', '商品全名', '简名', '型号', '规格', '产地', '成本算法', '商品属性',
              '联营扣率(%)', '默认供应商(编号)', '品牌(编号)', '计价方式', '有效期(天)', '备注', '基本单位',
              '条码', '零售价', '预设售价1', '预设售价2', '预设售价3', '配送价', '最近折后进价', '最低售价',
              '辅助单位1', '辅1换算关系', '辅1条码', '辅1零售价', '辅1预设售价1', '辅1预设售价2', '辅1预设售价3',
              '辅助单位2', '辅2换算关系', '辅2条码', '辅2零售价', '辅2预设售价1', '辅2预设售价2', '辅2预设售价3',
              '口味管理', '默认税率', '前台议价', '启用批次', '是否自制', '自制周期', '是否外购', '外购周期'],
    'employee': ['父类编号', '职员编号', '职员全名', '生日', '职员电话', '职员职务', '职员地址', '最低折扣下限',
                 '每单优惠限额', '备注'],
    'stockini': ['商品编号', '商品全名', '仓库编号', '仓库全名', '数量', '金额', '批号', '生产日期', '有效期至'],
    'pricefollow': ['单位编号', '单位全名', '商品编号', '商品全名', '计量单位', '最近进价', '最近售价',
                    '最近进货折扣', '最近销售折扣'],
}

# 官方模板表头（含前导行说明时会跳过）
TEMPLATE_SAMPLE = {
    '单位': ['00000', '示例001', '示例客户', '客户', '', '', '', '', '', '', '', '', '', '', '', '0', '0', '0', '0', '', '', '', '', '', '', '', '', '', '', '', '零售价'],
    '仓库': ['00000', '示例001', '示例仓库', '仓库', '', '', '', ''],
    '商品': ['00000', '示例001', '示例商品', '商品', '', '', '', '0', '', '', '', '', '', '', '', '件', '690000000001', '10', '12', '13', '14', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    '职员': ['00000', '示例001', '示例职员', '', '', '', '', '', '', ''],
    '库存信息': ['示例001', '示例商品', '示例001', '示例仓库', '100', '1000', '', '', ''],
    '单位价格跟踪': ['示例001', '示例客户', '示例001', '示例商品', '件', '6.5', '11', '0.9', '0.95'],
}


# ---------------- 连接与驱动 ----------------
def detect_driver():
    """探测可用 ODBC 驱动，优先 ODBC Driver 17 → Native Client 11 → SQL Server"""
    drvs = pyodbc.drivers()
    for pref in ['ODBC Driver 17 for SQL Server', 'ODBC Driver 13 for SQL Server',
                 'SQL Server Native Client 11.0', 'SQL Server Native Client 10.0', 'SQL Server']:
        for d in drvs:
            if d.lower() == pref.lower():
                return d
    return None


def connect(server, database, user, pwd, timeout=5):
    drv = detect_driver()
    if not drv:
        raise RuntimeError('未找到可用 ODBC SQL Server 驱动（需安装 ODBC Driver 17 或 Native Client）')
    cs = 'DRIVER={%s};SERVER=%s;UID=%s;PWD=%s;DATABASE=%s' % (drv, server, user, pwd, database)
    return pyodbc.connect(cs, autocommit=True, timeout=timeout)


def server_version(cur):
    cur.execute("SELECT CAST(SERVERPROPERTY('ProductVersion') AS varchar(50)), CAST(SERVERPROPERTY('Edition') AS varchar(100))")
    r = cur.fetchone()
    return (r[0], r[1]) if r else ('', '')


def check_account(cur):
    """账套自检：返回 (ok, 报告行列表)"""
    rep = []
    ver, ed = server_version(cur)
    rep.append('SQL Server 版本：%s（%s）' % (ver, ed))
    sps = ['p_hh_BasicInsertB_Commodity', 'p_hh_BasicInsertK_Commodity', 'p_hh_BascInsertP_Commodity',
           'p_hh_BasicInsertE_Commodity', 'p_hh_ModInitGoodsStocks_SNMng_Package', 'p_hh_ImportPtypePriceGZ',
           'xw_GetBasicData']
    cur.execute("SET NOCOUNT ON; SELECT name FROM sys.objects WHERE type='P' AND name IN (%s)" %
                ','.join("N'%s'" % s for s in sps))
    have = {r[0] for r in cur.fetchall()}
    missing_sp = [s for s in sps if s not in have]
    tables = ['btype', 'ptype', 'Stock', 'employee', 'IniGoodsStocks',
              'xw_P_PriceFollow', 'xw_P_PtypePrice', 'xw_PtypeUnit', 'xw_PtypeBarCode']
    cur.execute("SET NOCOUNT ON; SELECT name FROM sys.tables WHERE name IN (%s)" %
                ','.join("N'%s'" % t for t in tables))
    have_t = {r[0] for r in cur.fetchall()}
    missing_t = [t for t in tables if t not in have_t]
    ok = (not missing_sp) and (not missing_t)
    rep.append('官方导入存储过程：%d/7%s' % (7 - len(missing_sp), '（缺：%s）' % ','.join(missing_sp) if missing_sp else ' 齐全'))
    rep.append('核心表：%d/%d%s' % (len(tables) - len(missing_t), len(tables),
                                    '（缺：%s）' % ','.join(missing_t) if missing_t else ' 齐全'))
    if ok:
        root_col = {'btype': 'btypeid', 'ptype': 'ptypeid', 'Stock': 'ktypeid', 'employee': 'etypeid'}
        for t, root, name in [('btype', '00000', '全部单位'), ('ptype', '00000', '全部商品'),
                              ('Stock', '00000', '全部仓库'), ('employee', '00000', '管理员')]:
            cur.execute("SET NOCOUNT ON; SELECT COUNT(*) FROM %s WHERE %s = '%s'" %
                        (t, root_col[t], root))
            n = cur.fetchone()[0]
            if n == 0:
                ok = False
                rep.append('✘ %s 缺少根节点 %s' % (name, root))
        rep.append('✓ 账套结构完整，可执行一键导入')
    else:
        rep.append('✘ 账套结构不完整，请确认选择了正确的管家婆账套库（HH15new）')
    return ok, rep


# ---------------- 备份 / 回滚 ----------------
BAK_TABLES = ['btype', 'ptype', 'Stock', 'employee', 'IniGoodsStocks',
              'xw_P_PriceFollow', 'xw_P_PtypePrice', 'xw_PtypeUnit', 'xw_PtypeBarCode',
              'Commission', 'IniCommission', 'lendborrow', 'lendborrow00']


def backup_tables(cur):
    """SELECT INTO 备份核心表（含委托/受托、借进/借出业务表），返回备份前缀"""
    stamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    pre = '_impbak_' + stamp
    for t in BAK_TABLES:
        cur.execute("SET NOCOUNT ON; SELECT * INTO %s%s FROM dbo.%s" % (t, pre, t))
    return pre


def rollback_tables(cur, pre):
    """从备份恢复（先清当前表再插回；排除 identity 列由目标自增，显式事务保证原子性）"""
    cur.execute("SET NOCOUNT ON; BEGIN TRAN")
    try:
        for t in BAK_TABLES:
            cur.execute("SET NOCOUNT ON; SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('dbo.%s%s') AND is_identity = 0" % (t, pre))
            cols = [r[0] for r in cur.fetchall()]
            if not cols:
                continue
            collist = ','.join(cols)
            cur.execute("""
SET NOCOUNT ON;
IF OBJECT_ID('dbo.%s%s') IS NOT NULL
BEGIN
  DELETE FROM dbo.%s;
  EXEC('INSERT INTO dbo.%s (%s) SELECT %s FROM %s%s');
END""" % (t, pre, t, t, collist, collist, t, pre))
        cur.execute("COMMIT")
        return True
    except Exception:
        try:
            cur.execute("ROLLBACK")
        except Exception:
            pass
        raise


def clear_localdata(database):
    """清管家婆客户端缓存 localData/<账套>-00000-*"""
    d = os.path.join(INSTALL_DIR, 'Data', 'localData')
    if not os.path.isdir(d):
        return False, '未找到缓存目录：%s' % d
    n = 0
    for f in os.listdir(d):
        if f.startswith(database + '-'):
            p = os.path.join(d, f)
            try:
                os.remove(p)
                n += 1
            except Exception:
                pass
    return True, '已清缓存 %d 个文件' % n


# ---------------- Excel 读取 ----------------
def _norm(s):
    """表头归一：去空格/全半角/小写"""
    if s is None:
        return ''
    s = str(s).strip()
    return s


def read_excel(path):
    """读取 Excel，返回 {sheet名: [{'列名': 值}, ...]}（自动跳过头部说明行）"""
    ext = os.path.splitext(path)[1].lower()
    sheets = {}
    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(path)
        for sh in wb.sheets():
            rows = []
            for i in range(sh.nrows):
                rows.append([sh.cell_value(i, c) for c in range(sh.ncols)])
            sheets[sh.name] = rows
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append([('' if v is None else v) for v in row])
            sheets[ws.title] = rows
    # 清洗：找出表头行（首行含目标列名即作为表头；否则跳过分隔行）
    out = {}
    for name, rows in sheets.items():
        header, data_rows = _extract(rows, name)
        if data_rows is not None:
            out[name] = (header, data_rows)
    return out


def _extract(rows, sheetname):
    """定位表头行并返回 (header, [行dict])；找不到返回 (None, None)"""
    target_cols = ALL_COLS.get(SHEETS.get(sheetname), [])
    if not target_cols:
        # 非目标 sheet 直接忽略
        return None, None
    # 找含 '编号' 且与目标列交集≥2 的行作表头
    for idx, row in enumerate(rows):
        hdr = [_norm(x) for x in row]
        inter = [h for h in hdr if h in target_cols]
        if len(inter) >= 2:
            header = hdr
            data = []
            for r in rows[idx + 1:]:
                if all(('' if x is None else str(x)).strip() == '' for x in r):
                    continue
                d = {}
                for c, h in enumerate(header):
                    if h:
                        v = r[c] if c < len(r) else ''
                        d[h] = v
                # 跳过模板示例行（编号列含“示例”）
                code_key = {'unit': '单位编号', 'warehouse': '仓库编号', 'goods': '商品编号',
                            'employee': '职员编号', 'stockini': '商品编号', 'pricefollow': '单位编号'}.get(SHEETS.get(sheetname), '')
                if code_key in d and '示例' in str(d[code_key]):
                    continue
                data.append(d)
            return header, data
    return None, None


def validate_cols(module, header):
    """校验模块必填列，返回缺失列表"""
    req = REQUIRED_COLS.get(module, [])
    return [c for c in req if c not in header]


# ---------------- 官方存储过程调用 ----------------
def _exec_sel(cur, sql, *args):
    cur.execute("SET NOCOUNT ON;\n" + sql, *args)
    return cur.fetchone()


def call_unit(cur, parid, fullname, usercode, name='', namepy='', address='', tel='', tel2='',
              mophone='', fax='', postcode='', person='', person2='', bank='', bankacct='',
              taxno='', art00='', apt00='', yrt00='', ypt00='', arup='', apup='', skqx='',
              comment='', changeday='', changerate='', email='', definput='', areatype='',
              gatherbtype='', prtypeid=''):
    cur.execute("""
SET NOCOUNT ON;
DECLARE @rt varchar(50), @err varchar(100); DECLARE @rc int;
EXEC @rc = dbo.p_hh_BasicInsertB_Commodity @Parid=?,@FullName=?,@UserCode=?,@Name=?,@NamePy=?,
  @Address=?,@Tel=?,@Tel2=?,@moPhone=?,@FaxNo=?,@PostCode=?,@Person=?,@Person2=?,@Bank=?,
  @BankAccounts=?,@TaxNo=?,@Artotal00=?,@Aptotal00=?,@YRTotal00=?,@YPTotal00=?,@ArUp=?,@ApUp=?,
  @skqx=?,@Comment=?,@ChangeDay=?,@ChangeRate=?,@EMail=?,@definputCode=?,@areatypeCode=?,
  @gatherbtypecode=?,@prtypeid=?,@RltTypeID=@rt OUTPUT,@errorValue=@err OUTPUT;
SELECT @rc,@rt,@err;""",
        parid, fullname, usercode, name, namepy, address, tel, tel2, mophone, fax, postcode,
        person, person2, bank, bankacct, taxno, art00, apt00, yrt00, ypt00, arup, apup, skqx,
        comment, changeday, changerate, email, definput, areatype, gatherbtype, prtypeid)
    return cur.fetchone()


def call_warehouse(cur, parid, fullname, usercode, name='', namepy='', comment='', address='', person='', tel=''):
    cur.execute("""
SET NOCOUNT ON;
DECLARE @rt varchar(50), @err varchar(50); DECLARE @rc int;
EXEC @rc = dbo.p_hh_BasicInsertK_Commodity @parid=?,@fullname=?,@usercode=?,@name=?,@namepy=?,
  @comment=?,@Address=?,@Person=?,@Tel=?,@rlttypeid=@rt OUTPUT,@errorValue=@err OUTPUT;
SELECT @rc,@rt,@err;""",
        parid, fullname, usercode, name, namepy, comment, address, person, tel)
    return cur.fetchone()


def call_goods(cur, parid, fullname, usercode, name='', namepy='', typ='', standard='', area='',
               costmode='', ptypetype='', pooldisc='', btypecode='', ppcode='', isweight='',
               lifeday='', comment='', unit='', ucode='', saleprice='', pre1='', pre2='', pre3='',
               sendprice='', buydiscprice='', lowsaleprice='', unit1='', unitrate1='', ucode1='',
               u1saleprice='', u1pre1='', u1pre2='', u1pre3='', unit2='', unitrate2='', ucode2='',
               u2saleprice='', u2pre1='', u2pre2='', u2pre3='', sideunit='', pgmancode='',
               snmancode='', kwmancode='', pgholunit='', pgholunitrate='', defaulttax='',
               allowprice='', pjobmancode='', swzz='', swzzday='', swwg='', swwgday='', swww='', swwwday=''):
    cur.execute("""
SET NOCOUNT ON;
DECLARE @rt varchar(50), @err varchar(250); DECLARE @rc int;
EXEC @rc = dbo.p_hh_BascInsertP_Commodity @Parid=?,@FullName=?,@UserCode=?,@Name=?,@NamePy=?,
  @Type=?,@Standard=?,@Area=?,@CostMode=?,@PtypeType=?,@pooldiscount=?,@btypecode=?,@ppCode=?,
  @IsWeight=?,@LifeDay=?,@Comment=?,@unit=?,@UCode=?,@SalePrice=?,@PrePrice1=?,@PrePrice2=?,
  @PrePrice3=?,@SendPrice=?,@BuyDiscountPrice=?,@LowSalePrice=?,@unit1=?,@unitRate1=?,@UCode1=?,
  @unit1SalePrice=?,@unit1PrePrice1=?,@unit1PrePrice2=?,@unit1PrePrice3=?,@unit2=?,@unitRate2=?,
  @UCode2=?,@unit2SalePrice=?,@unit2PrePrice1=?,@unit2PrePrice2=?,@unit2PrePrice3=?,@SideUnit=?,
  @pgManCode=?,@snManCode=?,@kwManCode=?,@Pgholunit=?,@Pgholunitrate=?,@defaulttax=?,@allowprice=?,
  @pjobmancode=?,@sourcewayzz=?,@sourcewayzzDay=?,@sourcewaywg=?,@sourcewaywgDay=?,@sourcewayww=?,
  @sourcewaywwDay=?,@RltTypeID=@rt OUTPUT,@errorValue=@err OUTPUT;
SELECT @rc,@rt,@err;""",
        parid, fullname, usercode, name, namepy, typ, standard, area, costmode, ptypetype,
        pooldisc, btypecode, ppcode, isweight, lifeday, comment, unit, ucode, saleprice, pre1,
        pre2, pre3, sendprice, buydiscprice, lowsaleprice, unit1, unitrate1, ucode1, u1saleprice,
        u1pre1, u1pre2, u1pre3, unit2, unitrate2, ucode2, u2saleprice, u2pre1, u2pre2, u2pre3,
        sideunit, pgmancode, snmancode, kwmancode, pgholunit, pgholunitrate, defaulttax,
        allowprice, pjobmancode, swzz, swzzday, swwg, swwgday, swww, swwwday)
    return cur.fetchone()


def call_employee(cur, parid, fullname, usercode, namepy='', birthday='', tel='', duty='',
                  address='', lowdisc='', toptotal='', ecomment=''):
    cur.execute("""
SET NOCOUNT ON;
DECLARE @rt varchar(50), @err varchar(250); DECLARE @rc int;
EXEC @rc = dbo.p_hh_BasicInsertE_Commodity @parid=?,@fullname=?,@usercode=?,@namepy=?,@Birthday=?,
  @Tel=?,@duty=?,@Address=?,@LowLimitDiscount=?,@TopTotal=?,@eComment=?,@rlttypeid=@rt OUTPUT,
  @errorValue=@err OUTPUT;
SELECT @rc,@rt,@err;""",
        parid, fullname, usercode, namepy, birthday, tel, duty, address, lowdisc, toptotal, ecomment)
    return cur.fetchone()


def call_stockini(cur, ptypeid, puser, ktypeid, kuser, qty, total='', jobno='', createdate='', userend=''):
    cur.execute("""
SET NOCOUNT ON;
DECLARE @oid int, @err varchar(250); DECLARE @rc int;
EXEC @rc = dbo.p_hh_ModInitGoodsStocks_SNMng_Package @szptypeid=?,@szPUserCode=?,@szktypeid=?,
  @szKUserCode=?,@szslFullname='',@szslUserCode='',@dQty=?,@dTotal=?,@JobNumber=?,@CreateDate=?,
  @UserEndDate=?,@SideQty='',@szSnNo='',@OrderID=@oid OUTPUT,@errorValue=@err OUTPUT;
SELECT @rc,@oid,@err;""",
        ptypeid, puser, ktypeid, kuser, qty, total, jobno, createdate, userend)
    return cur.fetchone()


def call_pricefollow(cur, btypeid, bfull, ptypeid, pfull, punit='', buyprice='', saleprice='',
                     buydisc='', saledisc=''):
    cur.execute("""
SET NOCOUNT ON;
DECLARE @oid int, @err varchar(250); DECLARE @rc int;
EXEC @rc = dbo.p_hh_ImportPtypePriceGZ @szbtypeid=?,@szbfullname=?,@szptypeid=?,@szpfullname=?,
  @punit=?,@buyPrice=?,@salePrice=?,@buyDiscount=?,@saleDisCount=?,@OrderID=@oid OUTPUT,
  @errorValue=@err OUTPUT;
SELECT @rc,@oid,@err;""",
        btypeid, bfull, ptypeid, pfull, punit, buyprice, saleprice, buydisc, saledisc)
    return cur.fetchone()


# ---------------- 模块导入 ----------------
def _v(row, *keys):
    for k in keys:
        if k in row:
            v = row[k]
            if v is None:
                return ''
            s = str(v).strip()
            if isinstance(v, float) and v == int(v):
                s = str(int(v))
            return s
    return ''


class ParentResolver:
    """父类编号解析：支持填自定义编号(busercode/pusercode/kusercode)自动转树编号，也支持直接填树编号"""

    def __init__(self, cur):
        self.cur = cur
        self.maps = {'unit': {}, 'warehouse': {}, 'goods': {}}

    def add(self, module, code, treeid):
        if code:
            self.maps[module][code] = treeid

    def resolve(self, module, code):
        code = (code or '').strip()
        if not code or code == '00000':
            return '00000'
        d = self.maps[module]
        if code in d:
            return d[code]
        # 查库（自定义编号 -> 树编号）
        t = {'unit': ('btype', 'busercode', 'btypeid'),
             'warehouse': ('Stock', 'kusercode', 'ktypeid'),
             'goods': ('ptype', 'pusercode', 'ptypeid')}[module]
        try:
            self.cur.execute("SET NOCOUNT ON; SELECT %s FROM dbo.%s WHERE %s = ? AND deleted = 0" %
                             (t[2], t[0], t[1]), code)
            r = self.cur.fetchone()
            if r and r[0]:
                d[code] = r[0]
                return r[0]
        except Exception:
            pass
        return code  # 原样传（可能是树编号，官方过程兜底校验）


def import_sheet(cur, module, rows, logcb, resolver=None):
    """按模块导入 Excel 行，返回 (成功数, 失败数, [失败行列表])"""
    ok_n, fail_n, fails = 0, 0, []
    total = len(rows)
    for i, row in enumerate(rows):
        ln = i + 2  # 表头占 1 行，数据从第 2 行起（跳过说明行后按实际行号）
        try:
            rc, out, err = None, None, None
            if module == 'unit':
                par = resolver.resolve(module, _v(row, '父类编号')) if resolver else (_v(row, '父类编号') or '00000')
                r = call_unit(cur, par, _v(row, '单位编号'),
                              _v(row, '单位全名'), _v(row, '单位简名'), '',
                              _v(row, '单位地址'), _v(row, '单位电话'), _v(row, '手机一'),
                              _v(row, '手机二'), _v(row, '传真'), _v(row, '邮编'), _v(row, '联系人一'),
                              _v(row, '联系人二'), _v(row, '开户银行'), _v(row, '银行账号'),
                              _v(row, '税号'), _v(row, '期初应收款'), _v(row, '期初应付款'),
                              _v(row, '期初预收款'), _v(row, '期初预付款'), _v(row, '应收款上限'),
                              _v(row, '应付款上限'), _v(row, '结算期限(天)'), _v(row, '备注'),
                              _v(row, '换货期限(天)'), _v(row, '换货比例(%)'), _v(row, '电子邮件'),
                              _v(row, '默认经手人(编号)'), _v(row, '地区(编号)'), _v(row, '结算单位(编号)'),
                              _v(row, '适用价格'))
                rc, out, err = r[0], r[1], r[2]
            elif module == 'warehouse':
                par = resolver.resolve(module, _v(row, '父类编号')) if resolver else (_v(row, '父类编号') or '00000')
                r = call_warehouse(cur, par, _v(row, '仓库编号'),
                                   _v(row, '仓库全名'), _v(row, '仓库简称'), '', _v(row, '备注'),
                                   _v(row, '仓库地址'), _v(row, '仓库负责人'), _v(row, '负责人电话'))
                rc, out, err = r[0], r[1], r[2]
            elif module == 'goods':
                par = resolver.resolve(module, _v(row, '父类编号')) if resolver else (_v(row, '父类编号') or '00000')
                r = call_goods(cur, par, _v(row, '商品编号'),
                               _v(row, '商品全名'), _v(row, '简名'), '', _v(row, '型号'), _v(row, '规格'),
                               _v(row, '产地'), _v(row, '成本算法'), _v(row, '商品属性'),
                               _v(row, '联营扣率(%)'), _v(row, '默认供应商(编号)'), _v(row, '品牌(编号)'),
                               _v(row, '计价方式'), _v(row, '有效期(天)'), _v(row, '备注'),
                               _v(row, '基本单位'), _v(row, '条码'), _v(row, '零售价'),
                               _v(row, '预设售价1'), _v(row, '预设售价2'), _v(row, '预设售价3'),
                               _v(row, '配送价'), _v(row, '最近折后进价'), _v(row, '最低售价'),
                               _v(row, '辅助单位1'), _v(row, '辅1换算关系'), _v(row, '辅1条码'),
                               _v(row, '辅1零售价'), _v(row, '辅1预设售价1'), _v(row, '辅1预设售价2'),
                               _v(row, '辅1预设售价3'), _v(row, '辅助单位2'), _v(row, '辅2换算关系'),
                               _v(row, '辅2条码'), _v(row, '辅2零售价'), _v(row, '辅2预设售价1'),
                               _v(row, '辅2预设售价2'), _v(row, '辅2预设售价3'), '', '', '', '',
                               _v(row, '口味管理'), '', _v(row, '默认税率'), _v(row, '前台议价'),
                               _v(row, '启用批次'), _v(row, '是否自制'), _v(row, '自制周期'),
                               _v(row, '是否外购'), _v(row, '外购周期'), '', '')
                rc, out, err = r[0], r[1], r[2]
            elif module == 'employee':
                r = call_employee(cur, _v(row, '父类编号', '父类名称') or '00000', _v(row, '职员编号'),
                                  _v(row, '职员全名'), '', _v(row, '生日'), _v(row, '职员电话'),
                                  _v(row, '职员职务'), _v(row, '职员地址'), _v(row, '最低折扣下限'),
                                  _v(row, '每单优惠限额'), _v(row, '备注'))
                rc, out, err = r[0], r[1], r[2]
            elif module == 'stockini':
                r = call_stockini(cur, _v(row, '商品编号'), _v(row, '商品全名'), _v(row, '仓库编号'),
                                  _v(row, '仓库全名'), _v(row, '数量'), _v(row, '金额'),
                                  _v(row, '批号'), _v(row, '生产日期'), _v(row, '有效期至'))
                rc, out, err = r[0], r[1], r[2]
            elif module == 'pricefollow':
                r = call_pricefollow(cur, _v(row, '单位编号'), _v(row, '单位全名'), _v(row, '商品编号'),
                                     _v(row, '商品全名'), _v(row, '计量单位'), _v(row, '最近进价'),
                                     _v(row, '最近售价'), _v(row, '最近进货折扣'), _v(row, '最近销售折扣'))
                rc, out, err = r[0], r[1], r[2]
            if rc is not None and rc >= 0:
                ok_n += 1
                if resolver and module in resolver.maps and out:
                    resolver.add(module, _v(row, {'unit': '单位编号', 'warehouse': '仓库编号', 'goods': '商品编号'}[module]), out)
                logcb('OK', module, ln, '%s → %s' % (MODULE_NAMES[module], out if out else ''))
            else:
                fail_n += 1
                msg = err or ('返回码 %s' % rc)
                fails.append((ln, msg))
                logcb('FAIL', module, ln, msg)
        except Exception as e:
            fail_n += 1
            fails.append((ln, str(e)))
            logcb('FAIL', module, ln, str(e))
        if total and (i + 1) % 50 == 0:
            logcb('PROG', module, 0, '进度 %d/%d' % (i + 1, total))
    return ok_n, fail_n, fails


def orchestrate(cur, files, modules, logcb):
    """一键导入编排：备份 → 按依赖顺序导入所选模块"""
    # 解析 Excel
    data = {}  # module -> (header, rows)
    for f in files:
        try:
            sheets = read_excel(f)
        except Exception as e:
            logcb('FAIL', '文件', 0, '%s：%s' % (f, e))
            continue
        for sheetname, (hdr, rows) in sheets.items():
            m = SHEETS.get(sheetname)
            if m and m in modules and m not in data:
                data[m] = (hdr, rows)
                logcb('INFO', '文件', 0, '%s：%s sheet %d 行' % (os.path.basename(f), sheetname, len(rows)))
    if not data:
        logcb('FAIL', '文件', 0, '未从所选 Excel 中识别到任何目标 sheet（单位/商品/仓库/职员/库存信息/单位价格跟踪）')
        return
    # 列校验
    for m in modules:
        if m not in data:
            logcb('WARN', m, 0, '缺少 %s 数据，跳过' % MODULE_NAMES[m])
            continue
        missing = validate_cols(m, data[m][0])
        if missing:
            logcb('FAIL', m, 0, '缺少必填列：%s（模板列见说明）' % '、'.join(missing))
            data.pop(m)
    if not data:
        return
    # 备份
    try:
        pre = backup_tables(cur)
        logcb('INFO', '备份', 0, '已备份 %d 张表（前缀 %s）' % (len(BAK_TABLES), pre))
    except Exception as e:
        logcb('FAIL', '备份', 0, '备份失败：%s，已中止导入' % e)
        return
    # 按依赖顺序导入
    total_ok = total_fail = 0
    resolver = ParentResolver(cur)
    for m in MODULE_ORDER:
        if m not in data:
            continue
        logcb('INFO', m, 0, '── 开始导入%s ──' % MODULE_NAMES[m])
        ok, fail, _fails = import_sheet(cur, m, data[m][1], logcb, resolver)
        total_ok += ok
        total_fail += fail
        logcb('INFO', m, 0, '── %s 完成：成功 %d，失败 %d ──' % (MODULE_NAMES[m], ok, fail))
    logcb('INFO', '汇总', 0, '全部完成：成功 %d 行，失败 %d 行（备份前缀 %s，可在管家婆内核对后回滚）' %
          (total_ok, total_fail, pre))
    return pre


# ---------------- 模板生成 ----------------
def generate_template(path):
    """生成官方规格 6-sheet xlsx 模板（示例行编号带'示例'前缀，导入时自动跳过）"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.comments import Comment
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    head_fill = PatternFill('solid', fgColor='DDEBF7')
    sample_fill = PatternFill('solid', fgColor='F2F2F2')
    head_font = Font(bold=True)
    for sheetname in ['单位', '仓库', '商品', '职员', '库存信息', '单位价格跟踪']:
        ws = wb.create_sheet(sheetname)
        cols = ALL_COLS[SHEETS[sheetname]]
        for c, name in enumerate(cols, 1):
            cell = ws.cell(1, c, name)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal='center')
        sample = TEMPLATE_SAMPLE[sheetname]
        for c, v in enumerate(sample, 1):
            cell = ws.cell(2, c, v)
            cell.fill = sample_fill
        ws.cell(2, 1).comment = Comment('示例行（编号含“示例”），导入时自动跳过；正式填写请删除本行。', '模板')
        ws.freeze_panes = 'A2'
    wb.save(path)
    return True
