# -*- coding: utf-8 -*-
"""
管家婆辉煌 TOP15 双向直导工具 v3.0（GUI）

设计目标（针对 v2.0 实机验收反馈）：
  ① 零预填：所有连接参数（服务器/数据库/用户/密码）初始全部留空，由用户按实际环境填写，不内置任何实例/账套。
  ② 每个功能按钮旁配“?”帮助按钮，点击弹出该功能的用途、前置条件与操作说明。
  ③ 步骤引导式：每个页签内以“第 1 步 / 第 2 步 …”卡片组织，按顺序操作即可。
  ④ 功能融合 + 界面美化：统一专业配色、卡片式布局、清晰字体层级，去掉难理解的零散堆砌。
四个页签：
  1. Excel 导入   2. 千业→管家婆（正向）   3. 管家婆→千业（反向）   4. 维护工具
所有耗时任务走独立线程 + 独立连接，全局 busy 锁，线程安全日志。
"""
import os, sys, queue, threading, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import gpj_top15_core as core
import gpj_top15_srcimport as srcimp
import gpj_top15_srcexport as srcexp

# ---------------- 视觉主题 ----------------
C_PRIMARY = '#1F4E79'      # 主深蓝
C_PRIMARY2 = '#2E75B6'     # 中蓝
C_BG = '#EEF2F7'           # 窗口浅底
C_CARD = '#FFFFFF'         # 卡片白
C_TEXT = '#25313F'         # 主文字
C_MUTED = '#697586'        # 次要文字
C_BORDER = '#D3DCE6'       # 边框
C_ACCENT = '#E7F1FB'       # 浅蓝高亮
C_GREEN = '#2E7D32'
C_RED = '#C62828'
FONT_TITLE = ('Microsoft YaHei UI', 15, 'bold')
FONT_H = ('Microsoft YaHei UI', 10, 'bold')
FONT_BODY = ('Microsoft YaHei UI', 9)
FONT_MONO = ('Consolas', 9)

# ---------------- 帮助文本（集中管理） ----------------
HELP = {
    'overview': ('使用引导',
                 '本工具用于「千业管理通/商贸通」与「管家婆辉煌 TOP15」之间的双向数据迁移。\n\n'
                 '通用流程：\n'
                 '1) 先在对应页签填写连接参数（服务器、数据库、用户名、密码），点“测试连接”确认可连通；\n'
                 '2) 按“第 1 步、第 2 步……”顺序选择文件或勾选模块；\n'
                 '3) 点带★的主按钮执行，过程中请看下方“执行日志”；\n'
                 '4) 完成后建议清客户端缓存并重新登录管家婆核对。\n\n'
                 '每个功能按钮旁的“?”都可查看该按钮的详细说明。'),
    'test': ('测试连接',
             '用当前填写的服务器/数据库/用户名/密码尝试连接 SQL Server。\n\n'
             '· 成功：弹窗显示 SQL Server 版本与版本类型（如 Express）；\n'
             '· 失败：弹窗显示具体原因，常见为实例名错误、服务未启动、用户名或密码错误。\n'
             '请先测试成功，再执行后续导入/导出。'),
    'pick_excel': ('选择 Excel',
                   '选择一个或多个按本工具模板整理的 Excel 文件（.xls/.xlsx）。\n'
                   '选择后会自动扫描，识别每个工作表对应的数据模块并在下方勾选。\n'
                   '若没有模板，可先点“生成模板”。'),
    'gen_tpl': ('生成模板',
                '导出标准 Excel 模板，包含单位/仓库/商品/职员/期初库存/价格跟踪等工作表及表头。\n'
                '请按模板列名整理数据，带“编号”的列填纯数字编号，父子关系通过“父类编号”体现。'),
    'scan': ('扫描识别',
             '重新读取已选 Excel，识别其中各工作表对应的模块，并自动勾选/取消下方模块。\n'
             '用于手动增删文件或修改文件后刷新识别结果。'),
    'excel_import': ('一键导入 Excel',
                     '把已勾选模块的数据，通过管家婆官方存储过程导入目标账套。\n'
                     '执行前会自动备份相关表；编号自动数字化、父子引用自动映射。\n'
                     '请确保已正确选择文件并测试连接成功。'),
    'fwd_modules': ('正向同步模块',
                    '勾选需要从千业 Bas 库同步到管家婆的模块：\n'
                    '单位/仓库/商品/职员走官方存储过程；价格、条码、辅助单位、价格跟踪、期初库存走批量直写。\n'
                    '默认全部勾选，可按需取消。'),
    'forward': ('一键正向直导',
                '把千业 Bas 库（如 CCC、CCC2）的数据迁移到管家婆标准账套。\n\n'
                '· 自动级联数字化编号并保留层级；\n'
                '· 16 个价格方案（含用户自定义预设价格列）完整对应，不遗漏；\n'
                '· 库存取千业“当前结存”（StorageGoodsInfo），不是历史期初表；\n'
                '· 已开账账套会在无业务单据的前提下临时反开账录入期初，完成后自动恢复。\n'
                '执行前自动备份，可在维护工具中回滚。'),
    'backward': ('一键反向导出',
                 '把管家婆标准账套的数据回写到千业 Bas 库，按软件显示语义重建：\n\n'
                 '· 重建 BasePath 层级路径（不依赖编号大小）；\n'
                 '· 辅助单位换算率、16 个价格槽位（含自定义价格）一一对应；\n'
                 '· 补全仓库、职员、当前结存库存；\n'
                 '· 委托/受托、借进/借出等业务表同步对应。\n'
                 '执行前自动备份目标表，可回滚。'),
    'check': ('账套自检',
              '检查目标库是否具备导入所需的官方存储过程和核心表，以及根节点是否完整。\n'
              '用于在导入前确认选择的是正确、完整的管家婆账套库。'),
    'backup': ('立即备份',
               '对核心表（含主表与明细表）执行 SELECT INTO 备份，生成带时间戳的备份前缀。\n'
               '备份表保留在同一数据库中，可用于回滚恢复。'),
    'rollback': ('回滚恢复',
                 '用指定的备份前缀把核心表恢复到备份时的状态（覆盖当前数据）。\n'
                 '请在“回滚前缀”框准确填写备份前缀，执行前会再次确认。'),
    'clearcache': ('清客户端缓存',
                   '删除管家婆客户端 localData 中该账套的本地缓存。\n'
                   '改库后若界面显示旧数据/树空白，清缓存并重新登录即可刷新。'),
    'compare': ('差异对比',
                '只读对比源库与目标库核心表的行数，以及商品 BaseNo 集合差异。\n'
                '用于增量同步前发现“源独有/目标独有/共有”的商品，不修改任何数据。'),
}


def setup_style(root):
    st = ttk.Style(root)
    st.theme_use('clam')
    root.configure(bg=C_BG)
    st.configure('.', background=C_BG, foreground=C_TEXT, font=FONT_BODY)
    st.configure('TFrame', background=C_BG)
    st.configure('Card.TFrame', background=C_CARD)
    st.configure('TLabel', background=C_BG, foreground=C_TEXT, font=FONT_BODY)
    st.configure('Card.TLabel', background=C_CARD, foreground=C_TEXT, font=FONT_BODY)
    st.configure('Muted.Card.TLabel', background=C_CARD, foreground=C_MUTED, font=FONT_BODY)
    st.configure('H.Card.TLabel', background=C_CARD, foreground=C_PRIMARY, font=FONT_H)
    st.configure('TLabelframe', background=C_CARD, bordercolor=C_BORDER, relief='solid')
    st.configure('TLabelframe.Label', background=C_CARD, foreground=C_PRIMARY, font=FONT_H)
    st.configure('TButton', font=FONT_BODY, padding=4, background=C_CARD, bordercolor=C_BORDER)
    st.map('TButton', background=[('active', C_ACCENT), ('disabled', '#ECEFF3')])
    st.configure('Primary.TButton', font=('Microsoft YaHei UI', 10, 'bold'),
                 foreground='white', background=C_PRIMARY, padding=8, bordercolor=C_PRIMARY)
    st.map('Primary.TButton', background=[('active', C_PRIMARY2), ('disabled', '#9DB4CC')])
    st.configure('Help.TButton', font=('Microsoft YaHei UI', 8, 'bold'),
                 foreground='white', background=C_PRIMARY2, padding=0, width=2)
    st.map('Help.TButton', background=[('active', C_PRIMARY)])
    st.configure('TCheckbutton', background=C_CARD, foreground=C_TEXT, font=FONT_BODY)
    st.map('TCheckbutton', background=[('active', C_CARD)])
    st.configure('TEntry', fieldbackground='white', bordercolor=C_BORDER, padding=2)
    st.configure('TNotebook', background=C_BG, borderwidth=0)
    st.configure('TNotebook.Tab', font=('Microsoft YaHei UI', 10), padding=(18, 7),
                 background=C_BG, foreground=C_MUTED)
    st.map('TNotebook.Tab', background=[('selected', C_CARD), ('active', C_ACCENT)],
           foreground=[('selected', C_PRIMARY)])
    st.configure('Horizontal.TProgressbar', background=C_PRIMARY2,
                 troughcolor=C_BORDER, bordercolor=C_BORDER)


class HelpButton(ttk.Button):
    """圆形“?”帮助按钮，点击弹出对应帮助。"""

    def __init__(self, parent, key):
        super().__init__(parent, text='?', style='Help.TButton',
                         command=lambda: show_help(key, parent))


def show_help(key, parent=None):
    title, body = HELP[key]
    messagebox.showinfo(title, body, parent=parent)


class StepHeader(tk.Frame):
    """带蓝色圆形序号的步骤标题条。"""

    def __init__(self, parent, number, title, hint=''):
        super().__init__(parent, bg=C_CARD)
        cv = tk.Canvas(self, width=26, height=26, bg=C_CARD, highlightthickness=0)
        cv.create_oval(2, 2, 24, 24, fill=C_PRIMARY, outline=C_PRIMARY)
        cv.create_text(13, 13, text=str(number), fill='white',
                       font=('Microsoft YaHei UI', 10, 'bold'))
        cv.pack(side='left', padx=(2, 8), pady=4)
        ttk.Label(self, text=title, style='H.Card.TLabel').pack(side='left')
        if hint:
            ttk.Label(self, text='   ' + hint, style='Muted.Card.TLabel').pack(side='left')


class ConnPanel(ttk.LabelFrame):
    """可复用连接配置（零预填：服务器/数据库/用户/密码初始全部为空）。"""

    def __init__(self, parent, title, app):
        super().__init__(parent, text=' %s ' % title)
        self.app = app
        self.var_server = tk.StringVar(value='')
        self.var_db = tk.StringVar(value='')
        self.var_user = tk.StringVar(value='')
        self.var_pwd = tk.StringVar(value='')
        # 第 1 行：服务器 / 数据库
        ttk.Label(self, text='服务器实例', style='Card.TLabel').grid(
            row=0, column=0, sticky='e', padx=(8, 4), pady=5)
        ttk.Entry(self, textvariable=self.var_server, width=34).grid(
            row=0, column=1, sticky='w', padx=4)
        ttk.Label(self, text='数据库', style='Card.TLabel').grid(
            row=0, column=2, sticky='e', padx=(10, 4))
        ttk.Entry(self, textvariable=self.var_db, width=16).grid(
            row=0, column=3, sticky='w', padx=4)
        # 第 2 行：用户名 / 密码 / 测试 / 帮助
        ttk.Label(self, text='用户名', style='Card.TLabel').grid(
            row=1, column=0, sticky='e', padx=(8, 4), pady=(0, 6))
        ttk.Entry(self, textvariable=self.var_user, width=20).grid(
            row=1, column=1, sticky='w', padx=4)
        ttk.Label(self, text='密码', style='Card.TLabel').grid(
            row=1, column=2, sticky='e', padx=(10, 4))
        ttk.Entry(self, textvariable=self.var_pwd, width=16, show='*').grid(
            row=1, column=3, sticky='w', padx=4)
        ttk.Button(self, text='测试连接', style='Primary.TButton',
                   command=self.test).grid(row=1, column=4, padx=8)
        HelpButton(self, 'test').grid(row=1, column=5, padx=(0, 8))

    def connect(self):
        return core.connect(self.var_server.get().strip(), self.var_db.get().strip(),
                            self.var_user.get().strip(), self.var_pwd.get())

    def test(self):
        if not self.var_server.get().strip():
            messagebox.showwarning('提示', '请先填写服务器实例')
            return
        try:
            cn = self.connect()
            ver, ed = core.server_version(cn.cursor())
            cn.close()
            self.app.log('INFO', '连接', 0, '%s 连接成功：%s（%s）' % (self.var_db.get(), ver, ed))
            messagebox.showinfo('连接成功', '%s\nSQL Server %s（%s）' % (self.var_db.get(), ver, ed))
        except Exception as e:
            self.app.log('FAIL', '连接', 0, '%s 连接失败：%s' % (self.var_db.get(), e))
            messagebox.showerror('连接失败', str(e))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('管家婆辉煌 TOP15 双向直导工具 v3.0')
        self.geometry('1060x820')
        self.minsize(960, 700)
        self.log_q = queue.Queue()
        self.busy = False
        self.files = []
        self._all_buttons = []
        setup_style(self)
        self._build()
        self.after(100, self._drain_log)

    # ---------------- 主界面 ----------------
    def _build(self):
        # 顶部横幅
        banner = tk.Frame(self, bg=C_PRIMARY, height=58)
        banner.pack(fill='x', padx=8, pady=(8, 4))
        banner.pack_propagate(False)
        ttk.Label(banner, text='管家婆辉煌 TOP15  双向直导工具',
                  background=C_PRIMARY, foreground='white',
                  font=FONT_TITLE).pack(side='left', padx=16)
        ttk.Button(banner, text='使用引导', command=lambda: show_help('overview', self)).pack(
            side='right', padx=14)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True, padx=8, pady=4)
        self._tab_excel()
        self._tab_forward()
        self._tab_backward()
        self._tab_maint()

        frm_log = ttk.LabelFrame(self, text=' 执行日志 ')
        frm_log.pack(fill='both', expand=True, padx=8, pady=4)
        self.txt = scrolledtext.ScrolledText(frm_log, height=10, state='disabled',
                                             font=FONT_MONO, bg='#FBFCFE', relief='flat')
        self.txt.pack(fill='both', expand=True, padx=6, pady=6)

        frm_bar = ttk.Frame(self)
        frm_bar.pack(fill='x', padx=8, pady=(0, 8))
        self.prog = ttk.Progressbar(frm_bar, mode='indeterminate')
        self.prog.pack(side='left', fill='x', expand=True, padx=4)
        self.var_stat = tk.StringVar(value='就绪')
        ttk.Label(frm_bar, textvariable=self.var_stat).pack(side='left', padx=4)

    def _track(self, b):
        self._all_buttons.append(b)
        return b

    def _card(self, parent):
        c = ttk.Frame(parent, style='Card.TFrame')
        c.pack(fill='x', padx=6, pady=5)
        return c

    # ---------------- 页签 1：Excel 导入 ----------------
    def _tab_excel(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='Excel 导入')

        # 步骤 1：连接
        self.excel_conn = ConnPanel(f, '第 1 步 · 连接目标管家婆账套', self)
        self.excel_conn.pack(fill='x', padx=6, pady=6)

        # 步骤 2：选择文件
        card = self._card(f)
        StepHeader(card, 2, '选择 Excel 数据文件').pack(anchor='w', padx=8, pady=(6, 0))
        row = ttk.Frame(card, style='Card.TFrame')
        row.pack(anchor='w', padx=10, pady=6)
        self._track(ttk.Button(row, text='选择 Excel…', command=self.pick_files)).pack(side='left', padx=3)
        HelpButton(row, 'pick_excel').pack(side='left', padx=(0, 8))
        self._track(ttk.Button(row, text='生成模板', command=self.gen_tpl)).pack(side='left', padx=3)
        HelpButton(row, 'gen_tpl').pack(side='left', padx=(0, 8))
        self._track(ttk.Button(row, text='扫描识别', command=self.scan_files)).pack(side='left', padx=3)
        HelpButton(row, 'scan').pack(side='left')
        self.lbl_files = ttk.Label(card, text='未选择文件', style='Muted.Card.TLabel')
        self.lbl_files.pack(anchor='w', padx=10)
        self.lst_files = tk.Listbox(card, height=3, relief='flat', bg='#F4F8FC',
                                    highlightthickness=1, highlightcolor=C_BORDER)
        self.lst_files.pack(fill='x', padx=10, pady=(4, 8))

        # 步骤 3：模块 + 执行
        card3 = self._card(f)
        StepHeader(card3, 3, '确认导入模块并执行').pack(anchor='w', padx=8, pady=(6, 0))
        frm_mod = ttk.Frame(card3, style='Card.TFrame')
        frm_mod.pack(anchor='w', padx=10, pady=4)
        self.excel_chk = {}
        for i, (m, name) in enumerate(core.MODULE_NAMES.items()):
            v = tk.BooleanVar(value=True)
            self.excel_chk[m] = v
            ttk.Checkbutton(frm_mod, text=name, variable=v).grid(
                row=0, column=i, sticky='w', padx=6)
        ra = ttk.Frame(card3, style='Card.TFrame')
        ra.pack(anchor='w', padx=10, pady=(2, 10))
        self._track(ttk.Button(ra, text='★ 一键导入 Excel', style='Primary.TButton',
                               command=self.run_excel)).pack(side='left', padx=3)
        HelpButton(ra, 'excel_import').pack(side='left')

    # ---------------- 页签 2：正向直连 ----------------
    def _tab_forward(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='千业 → 管家婆（正向）')

        self.fwd_src = ConnPanel(f, '第 1 步 · 源：千业 Bas 库（CCC / CCC2 等）', self)
        self.fwd_src.pack(fill='x', padx=6, pady=6)
        self.fwd_dst = ConnPanel(f, '第 2 步 · 目标：管家婆标准账套', self)
        self.fwd_dst.pack(fill='x', padx=6, pady=(0, 6))

        card = self._card(f)
        StepHeader(card, 3, '勾选同步模块', '默认全部，可按需取消').pack(anchor='w', padx=8, pady=(6, 0))
        hb = ttk.Frame(card, style='Card.TFrame')
        hb.pack(anchor='w', padx=6, pady=2)
        HelpButton(hb, 'fwd_modules').pack(side='left', padx=(4, 8))
        frm_mod = ttk.Frame(card, style='Card.TFrame')
        frm_mod.pack(anchor='w', padx=10, pady=4)
        self.fwd_chk = {}
        for i, m in enumerate(srcimp.MODULES):
            v = tk.BooleanVar(value=True)
            self.fwd_chk[m] = v
            ttk.Checkbutton(frm_mod, text=srcimp.MODULE_CN[m], variable=v).grid(
                row=i // 5, column=i % 5, sticky='w', padx=6, pady=2)

        ra = ttk.Frame(card, style='Card.TFrame')
        ra.pack(anchor='w', padx=10, pady=(2, 10))
        self._track(ttk.Button(ra, text='★ 一键正向直导', style='Primary.TButton',
                               command=self.run_forward)).pack(side='left', padx=3)
        HelpButton(ra, 'forward').pack(side='left')

    # ---------------- 页签 3：反向导出 ----------------
    def _tab_backward(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='管家婆 → 千业（反向）')

        self.bwd_src = ConnPanel(f, '第 1 步 · 源：管家婆标准账套', self)
        self.bwd_src.pack(fill='x', padx=6, pady=6)
        self.bwd_dst = ConnPanel(f, '第 2 步 · 目标：千业 Bas 库', self)
        self.bwd_dst.pack(fill='x', padx=6, pady=(0, 6))

        card = self._card(f)
        StepHeader(card, 3, '执行反向导出').pack(anchor='w', padx=8, pady=(6, 0))
        ttk.Label(card,
                  text='自动重建 BasePath 层级、辅助单位、16 价格槽位（含自定义价格）、仓库职员、当前结存，以及委托/受托、借进/借出业务表；执行前自动备份。',
                  style='Muted.Card.TLabel', wraplength=900, justify='left').pack(
            anchor='w', padx=10, pady=4)
        ra = ttk.Frame(card, style='Card.TFrame')
        ra.pack(anchor='w', padx=10, pady=(0, 10))
        self._track(ttk.Button(ra, text='★ 一键反向导出', style='Primary.TButton',
                               command=self.run_backward)).pack(side='left', padx=3)
        HelpButton(ra, 'backward').pack(side='left')

    # ---------------- 页签 4：维护工具 ----------------
    def _tab_maint(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='维护工具')

        self.mt_conn = ConnPanel(f, '第 1 步 · 维护连接（填写需维护的账套）', self)
        self.mt_conn.pack(fill='x', padx=6, pady=6)

        card = self._card(f)
        StepHeader(card, 2, '账套维护').pack(anchor='w', padx=8, pady=(6, 0))
        r1 = ttk.Frame(card, style='Card.TFrame')
        r1.pack(anchor='w', padx=10, pady=6)
        for txt, key, cmd in [
            ('账套自检', 'check', self.do_check),
            ('立即备份', 'backup', self.do_backup),
            ('清客户端缓存', 'clearcache', self.do_clearcache)]:
            self._track(ttk.Button(r1, text=txt, command=cmd)).pack(side='left', padx=3)
            HelpButton(r1, key).pack(side='left', padx=(0, 10))

        r2 = ttk.Frame(card, style='Card.TFrame')
        r2.pack(anchor='w', padx=10, pady=(0, 8))
        ttk.Label(r2, text='回滚前缀：', style='Card.TLabel').pack(side='left')
        self.var_roll = tk.StringVar()
        ttk.Entry(r2, textvariable=self.var_roll, width=28).pack(side='left', padx=4)
        self._track(ttk.Button(r2, text='回滚恢复', command=self.do_rollback)).pack(side='left', padx=3)
        HelpButton(r2, 'rollback').pack(side='left')

        # 差异对比
        cmpf = ttk.LabelFrame(f, text=' 第 3 步 · 源-目标差异对比（只读） ')
        cmpf.pack(fill='x', padx=6, pady=6)
        self.cmp_src = ConnPanel(cmpf, '对比源', self)
        self.cmp_src.pack(fill='x', padx=6, pady=6)
        self.cmp_dst = ConnPanel(cmpf, '对比目标', self)
        self.cmp_dst.pack(fill='x', padx=6, pady=(0, 6))
        r3 = ttk.Frame(cmpf)
        r3.pack(anchor='w', padx=10, pady=(0, 8))
        self._track(ttk.Button(r3, text='对比核心表行数 + 商品 BaseNo 差异',
                               command=self.do_compare)).pack(side='left', padx=3)
        HelpButton(r3, 'compare').pack(side='left')

    # ---------------- 日志 ----------------
    def log(self, kind, module, line, msg):
        now = datetime.datetime.now().strftime('%H:%M:%S')
        tag = {'OK': '✓', 'FAIL': '✘', 'WARN': '⚠', 'INFO': '·', 'PROG': '…'}.get(kind, '·')
        s = '[%s] %s %s' % (now, tag, msg) if module in ('连接',) else \
            '[%s] %s %s | %s' % (now, tag, module, msg)
        self.log_q.put(s)

    def _drain_log(self):
        try:
            while True:
                s = self.log_q.get_nowait()
                self.txt.configure(state='normal')
                self.txt.insert('end', s + '\n')
                self.txt.see('end')
                self.txt.configure(state='disabled')
        except queue.Empty:
            pass
        self.after(100, self._drain_log)

    # ---------------- 异步执行器 ----------------
    def submit(self, name, worker):
        if self.busy:
            messagebox.showwarning('提示', '已有任务正在执行，请等待完成')
            return
        self.busy = True
        self.var_stat.set(name + '…')
        for b in self._all_buttons:
            b.configure(state='disabled')
        self.prog.start(12)

        def wrap():
            try:
                worker()
            except Exception as e:
                self.log('FAIL', '任务', 0, '%s 中止：%s' % (name, e))
            finally:
                self.after(0, self._done)

        threading.Thread(target=wrap, daemon=True).start()

    def _done(self):
        self.prog.stop()
        self.busy = False
        self.var_stat.set('完成')
        for b in self._all_buttons:
            b.configure(state='normal')

    # ---------------- Excel 动作 ----------------
    def pick_files(self):
        fl = filedialog.askopenfilenames(title='选择 Excel（多选）',
                                         filetypes=[('Excel', '*.xls *.xlsx'), ('所有文件', '*.*')])
        if fl:
            self.files = list(fl)
            self.lst_files.delete(0, 'end')
            for x in self.files:
                self.lst_files.insert('end', os.path.basename(x))
            self.lbl_files.config(text='%d 个文件' % len(self.files))
            self.scan_files()

    def scan_files(self):
        if not self.files:
            return
        found = {}
        for fp in self.files:
            try:
                sheets = core.read_excel(fp)
            except Exception as e:
                self.log('FAIL', '扫描', 0, '%s：%s' % (os.path.basename(fp), e))
                continue
            for sn, (hdr, rows) in sheets.items():
                m = core.SHEETS.get(sn)
                if m:
                    found.setdefault(m, []).append('%s（%d 行）' % (os.path.basename(fp), len(rows)))
        for m, v in self.excel_chk.items():
            v.set(m in found)
        if found:
            self.log('INFO', '扫描', 0, '识别到：' + '；'.join(
                '%s %s' % (core.MODULE_NAMES[m], ','.join(x)) for m, x in found.items()))

    def gen_tpl(self):
        p = filedialog.asksaveasfilename(title='保存模板', defaultextension='.xlsx',
                                         initialfile='管家婆TOP15一键导入模板.xlsx',
                                         filetypes=[('Excel', '*.xlsx')])
        if p:
            try:
                core.generate_template(p)
                self.log('INFO', '模板', 0, '已生成：%s' % p)
                messagebox.showinfo('完成', '模板已生成：\n%s' % p)
            except Exception as e:
                messagebox.showerror('失败', str(e))

    def run_excel(self):
        if not self.files:
            messagebox.showwarning('提示', '请先选择 Excel 文件')
            return
        modules = [m for m, v in self.excel_chk.items() if v.get()]
        if not modules:
            messagebox.showwarning('提示', '请至少勾选一个模块')
            return
        files = list(self.files)
        panel = self.excel_conn

        def worker():
            cn = panel.connect()
            core.orchestrate(cn.cursor(), files, modules,
                             lambda k, m, l, s: self.log(k, m, l, s))
            cn.close()
            self.log('INFO', '缓存', 0, '建议清客户端缓存后重登查看')
        self.submit('Excel 导入', worker)

    # ---------------- 正向动作 ----------------
    def run_forward(self):
        mods = [m for m, v in self.fwd_chk.items() if v.get()]
        if not mods:
            messagebox.showwarning('提示', '请至少勾选一个模块')
            return
        sp, dp = self.fwd_src, self.fwd_dst

        def worker():
            s, d = sp.connect(), dp.connect()
            srcimp.import_from_db(s.cursor(), d.cursor(), sp.var_db.get(), mods,
                                  lambda k, m, l, x: self.log(k, m, l, x))
            s.close(); d.close()
        self.submit('正向直导', worker)

    # ---------------- 反向动作 ----------------
    def run_backward(self):
        sp, dp = self.bwd_src, self.bwd_dst

        def worker():
            s, d = sp.connect(), dp.connect()
            pre, stats = srcexp.export_to_db(s.cursor(), d.cursor(),
                                             lambda k, m, l, x: self.log(k, m, l, x))
            for t, n in stats.items():
                self.log('INFO', '导出', 0, '%s = %d 行' % (t, n))
            s.close(); d.close()
            self.log('INFO', '导出', 0, '完成，目标备份前缀 %s' % pre)
        self.submit('反向导出', worker)

    # ---------------- 维护动作 ----------------
    def do_check(self):
        panel = self.mt_conn

        def worker():
            cn = panel.connect()
            ok, rep = core.check_account(cn.cursor())
            cn.close()
            for line in rep:
                self.log('INFO' if ok else 'WARN', '自检', 0, line)
        self.submit('账套自检', worker)

    def do_backup(self):
        panel = self.mt_conn

        def worker():
            cn = panel.connect()
            pre = core.backup_tables(cn.cursor())
            cn.close()
            self.var_roll.set(pre)
            self.log('INFO', '备份', 0, '已备份核心表，前缀 %s' % pre)
        self.submit('备份', worker)

    def do_rollback(self):
        pre = self.var_roll.get().strip()
        if not pre:
            messagebox.showwarning('提示', '请填写备份前缀')
            return
        if not messagebox.askyesno('确认', '用备份 %s 覆盖当前数据，确定？' % pre):
            return
        panel = self.mt_conn

        def worker():
            cn = panel.connect()
            core.rollback_tables(cn.cursor(), pre)
            cn.close()
            self.log('INFO', '回滚', 0, '已从 %s 恢复' % pre)
        self.submit('回滚', worker)

    def do_clearcache(self):
        ok, msg = core.clear_localdata(self.mt_conn.var_db.get().strip())
        self.log('INFO' if ok else 'FAIL', '缓存', 0, msg)

    def do_compare(self):
        sp, dp = self.cmp_src, self.cmp_dst
        cmp_tables = ['BasUnitTypeInner', 'BasPtypeInner', 'BasStockInner', 'BasPersonInner',
                      'BasPTypeBarMx', 'BasPTypePriceMx', 'price', 'StorageGoodsInfo']

        def worker():
            s, d = sp.connect(), dp.connect()
            sc, dc = s.cursor(), d.cursor()
            self.log('INFO', '对比', 0, '核心表行数：表名 源→目标')
            for t in cmp_tables:
                try:
                    sc.execute("SELECT COUNT(*) FROM dbo.%s" % t)
                    x = sc.fetchone()[0]
                    dc.execute("SELECT COUNT(*) FROM dbo.%s" % t)
                    y = sc.fetchone()[0]
                    mark = '✓' if x == y else '≠'
                    self.log('INFO' if x == y else 'WARN', '对比', 0,
                             '%s %s：%d → %d %s' % (mark, t, x, y, mark))
                except Exception as e:
                    self.log('WARN', '对比', 0, '%s 跳过：%s' % (t, e))
            try:
                sc.execute("SELECT BaseNo FROM dbo.BasPtypeInner")
                sx = set(r[0] for r in sc.fetchall())
                dc.execute("SELECT BaseNo FROM dbo.BasPtypeInner")
                dx = set(r[0] for r in sc.fetchall())
                self.log('INFO', '对比', 0, '商品 BaseNo：源独有 %d（待新增），目标独有 %d，共有 %d' %
                         (len(sx - dx), len(dx - dx), len(sx & dx)))
            except Exception as e:
                self.log('WARN', '对比', 0, 'BaseNo 差异跳过：%s' % e)
            s.close(); d.close()
        self.submit('差异对比', worker)

    def run(self):
        self.mainloop()


if __name__ == '__main__':
    App().run()
