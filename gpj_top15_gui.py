# -*- coding: utf-8 -*-
"""
管家婆辉煌 TOP15 双向直导工具 v2.0（GUI）
四个页签：
  1. Excel 通道（官方存储过程，v1.0 能力）
  2. 千业→管家婆 正向直连（Bas 源库 → 标准账套，16 方案完整价格 + 真实换算率 + 当前结存）
  3. 管家婆→千业 反向导出（标准账套 → Bas 源库，主键 100% 还原）
  4. 维护工具（账套自检 / 备份 / 回滚 / 清缓存 / 源-目标差异对比）
所有耗时任务走独立线程 + 独立连接，全局 busy 锁，线程安全日志。
"""
import os, sys, queue, threading, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import gpj_top15_core as core
import gpj_top15_srcimport as srcimp
import gpj_top15_srcexport as srcexp

# 默认实例（可在界面修改，适配任意 SQL Server / 账套）
DST_SERVER = r'DESKTOP-8P5G8VQ\SQLSERVER'
SRC_SERVER = r'DESKTOP-8P5G8VQ\SQLEXPRESS'
DEFAULT_USER = 'sa'


class ConnPanel(ttk.LabelFrame):
    """可复用连接配置：服务器/数据库/用户名/密码 + 测试连接"""

    def __init__(self, parent, title, default_server, default_db, app):
        super().__init__(parent, text=' %s ' % title)
        self.app = app
        pad = {'padx': 4, 'pady': 2}
        self.var_server = tk.StringVar(value=default_server)
        self.var_db = tk.StringVar(value=default_db)
        self.var_user = tk.StringVar(value=DEFAULT_USER)
        self.var_pwd = tk.StringVar(value='')
        ttk.Label(self, text='服务器').grid(row=0, column=0, sticky='e', **pad)
        ttk.Entry(self, textvariable=self.var_server, width=26).grid(row=0, column=1, **pad)
        ttk.Label(self, text='数据库').grid(row=0, column=2, sticky='e', **pad)
        ttk.Entry(self, textvariable=self.var_db, width=14).grid(row=0, column=3, **pad)
        ttk.Label(self, text='用户').grid(row=0, column=4, sticky='e', **pad)
        ttk.Entry(self, textvariable=self.var_user, width=8).grid(row=0, column=5, **pad)
        ttk.Label(self, text='密码').grid(row=0, column=6, sticky='e', **pad)
        ttk.Entry(self, textvariable=self.var_pwd, width=10, show='*').grid(row=0, column=7, **pad)
        self.btn_test = ttk.Button(self, text='测试', width=6, command=self.test)
        self.btn_test.grid(row=0, column=8, **pad)

    def connect(self):
        return core.connect(self.var_server.get().strip(), self.var_db.get().strip(),
                            self.var_user.get().strip(), self.var_pwd.get())

    def test(self):
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
        self.title('管家婆辉煌 TOP15 双向直导工具 v2.0')
        self.geometry('1040x780')
        self.minsize(920, 680)
        self.log_q = queue.Queue()
        self.busy = False
        self.files = []
        self._all_buttons = []
        self._build()
        self.after(100, self._drain_log)

    # ---------------- 主界面 ----------------
    def _build(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='x', padx=8, pady=(8, 4))
        self._tab_excel()
        self._tab_forward()
        self._tab_backward()
        self._tab_maint()

        frm_log = ttk.LabelFrame(self, text=' 执行日志 ')
        frm_log.pack(fill='both', expand=True, padx=8, pady=4)
        self.txt = scrolledtext.ScrolledText(frm_log, height=12, state='disabled', font=('Consolas', 9))
        self.txt.pack(fill='both', expand=True, padx=6, pady=6)

        frm_bar = ttk.Frame(self)
        frm_bar.pack(fill='x', padx=8, pady=(0, 8))
        self.prog = ttk.Progressbar(frm_bar, mode='indeterminate')
        self.prog.pack(side='left', fill='x', expand=True, padx=4)
        self.var_stat = tk.StringVar(value='就绪')
        ttk.Label(frm_bar, textvariable=self.var_stat).pack(side='left', padx=4)

    def _btn(self, parent, text, cmd):
        b = ttk.Button(parent, text=text, command=cmd)
        b.pack(side='left', padx=4, pady=3)
        self._all_buttons.append(b)
        return b

    # ---------------- 页签 1：Excel ----------------
    def _tab_excel(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='Excel 通道')
        self.excel_conn = ConnPanel(f, '目标管家婆账套', DST_SERVER, 'HH15new', self)
        self.excel_conn.pack(fill='x', padx=6, pady=6)

        frm_file = ttk.LabelFrame(f, text=' 数据文件（Excel，支持多选） ')
        frm_file.pack(fill='x', padx=6, pady=4)
        row = ttk.Frame(frm_file); row.pack(fill='x')
        self._btn(row, '选择 Excel…', self.pick_files)
        self._btn(row, '生成模板', self.gen_tpl)
        self._btn(row, '扫描识别', self.scan_files)
        self.lbl_files = ttk.Label(row, text='未选择文件', foreground='#888')
        self.lbl_files.pack(side='left', padx=6)
        self.lst_files = tk.Listbox(frm_file, height=3)
        self.lst_files.pack(fill='x', padx=6, pady=(0, 4))

        frm_mod = ttk.LabelFrame(f, text=' 导入模块 ')
        frm_mod.pack(fill='x', padx=6, pady=4)
        self.excel_chk = {}
        for i, (m, name) in enumerate(core.MODULE_NAMES.items()):
            v = tk.BooleanVar(value=True)
            self.excel_chk[m] = v
            ttk.Checkbutton(frm_mod, text=name, variable=v).grid(row=0, column=i, sticky='w', padx=6, pady=3)

        row_act = ttk.Frame(f); row_act.pack(fill='x', padx=6, pady=4)
        self._btn(row_act, '★ 一键导入 Excel', self.run_excel)

    # ---------------- 页签 2：正向直连 ----------------
    def _tab_forward(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='千业→管家婆 正向直连')
        self.fwd_src = ConnPanel(f, '源：千业 Bas 库（CCC / CCC2 等）', SRC_SERVER, 'CCC', self)
        self.fwd_src.pack(fill='x', padx=6, pady=6)
        self.fwd_dst = ConnPanel(f, '目标：管家婆标准账套', DST_SERVER, 'HH15new', self)
        self.fwd_dst.pack(fill='x', padx=6, pady=(0, 6))

        frm_mod = ttk.LabelFrame(f, text=' 同步模块（级联数字化 + 官方过程 + 明细直写） ')
        frm_mod.pack(fill='x', padx=6, pady=4)
        self.fwd_chk = {}
        for i, m in enumerate(srcimp.MODULES):
            v = tk.BooleanVar(value=True)
            self.fwd_chk[m] = v
            ttk.Checkbutton(frm, text=srcimp.MODULE_CN[m], variable=v).grid(
                row=i // 5, column=i % 5, sticky='w', padx=6, pady=3)

        row_act = ttk.Frame(f); row_act.pack(fill='x', padx=6, pady=4)
        self._btn(row_act, '★ 一键正向直导', self.run_forward)

    # ---------------- 页签 3：反向导出 ----------------
    def _tab_backward(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='管家婆→千业 反向导出')
        self.bwd_src = ConnPanel(f, '源：管家婆标准账套', DST_SERVER, 'HH15new', self)
        self.bwd_src.pack(fill='x', padx=6, pady=6)
        self.bwd_dst = ConnPanel(f, '目标：千业 Bas 库', SRC_SERVER, 'qysrc_test', self)
        self.bwd_dst.pack(fill='x', padx=6, pady=(0, 6))
        ttk.Label(f, text='说明：自动备份目标 8 表；按备注“源BaseNo”100% 还原主键（含历史缺口）；库存写当前结存。',
                  foreground='#666').pack(anchor='w', padx=10, pady=4)
        row_act = ttk.Frame(f); row_act.pack(fill='x', padx=6, pady=4)
        self._btn(row_act, '★ 一键反向导出', self.run_backward)

    # ---------------- 页签 4：维护 ----------------
    def _tab_maint(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='维护工具')
        self.mt_conn = ConnPanel(f, '维护连接', DST_SERVER, 'HH15new', self)
        self.mt_conn.pack(fill='x', padx=6, pady=6)

        frm = ttk.LabelFrame(f, text=' 账套维护 ')
        frm.pack(fill='x', padx=6, pady=4)
        r1 = ttk.Frame(frm); r1.pack(fill='x')
        self._btn(r1, '账套自检', self.do_check)
        self._btn(r1, '立即备份', self.do_backup)
        self._btn(r1, '清客户端缓存', self.do_clearcache)

        r2 = ttk.Frame(frm); r2.pack(fill='x')
        ttk.Label(r2, text='回滚前缀：').pack(side='left', padx=(6, 2), pady=3)
        self.var_roll = tk.StringVar()
        ttk.Entry(r2, textvariable=self.var_roll, width=28).pack(side='left', padx=2)
        self._btn(r2, '回滚恢复', self.do_rollback)

        # 差异对比（增量同步的差异发现）
        frm_cmp = ttk.LabelFrame(f, text=' 源-目标差异对比（增量发现，只读） ')
        frm_cmp.pack(fill='x', padx=6, pady=4)
        self.cmp_src = ConnPanel(frm_cmp, '对比源', SRC_SERVER, 'CCC', self)
        self.cmp_src.pack(fill='x', padx=6, pady=6)
        self.cmp_dst = ConnPanel(frm_cmp, '对比目标', SRC_SERVER, 'qysrc_test', self)
        self.cmp_dst.pack(fill='x', padx=6, pady=(0, 6))
        r3 = ttk.Frame(frm_cmp); r3.pack(fill='x', padx=6, pady=4)
        self._btn(r3, '对比核心表行数 + 商品 BaseNo 差异', self.do_compare)

    # ---------------- 日志 ----------------
    def log(self, kind, module, line, msg):
        now = datetime.datetime.now().strftime('%H:%M:%S')
        tag = {'OK': '✓', 'FAIL': '✘', 'WARN': '⚠', 'INFO': '·', 'PROG': '…'}.get(kind, '·')
        s = '[%s] %s %s' % (now, tag, msg) if module in ('连接',) else '[%s] %s %s | %s' % (now, tag, module, msg)
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
        files = list(self.files); panel = self.excel_conn

        def worker():
            cn = panel.connect(); cur = cn.cursor()
            core.orchestrate(cur, files, modules, lambda k, m, l, s: self.log(k, m, l, s))
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
            s = sp.connect(); d = dp.connect()
            srcimp.import_from_db(s.cursor(), d.cursor(), sp.var_db.get(), mods,
                                  lambda k, m, l, x: self.log(k, m, l, x))
            s.close(); d.close()
        self.submit('正向直导', worker)

    # ---------------- 反向动作 ----------------
    def run_backward(self):
        sp, dp = self.bwd_src, self.bwd_dst

        def worker():
            s = sp.connect(); d = dp.connect()
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
            pre = core.backup_tables(cursor())
            cn.close()
            self.var_roll.set(pre)
            self.log('INFO', '备份', 0, '已备份 9 表，前缀 %s' % pre)
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
            s = sp.connect(); d = dp.connect()
            sc, dc = s.cursor(), d.cursor()
            self.log('INFO', '对比', 0, '核心表行数：表名 源→目标')
            for t in cmp_tables:
                try:
                    sc.execute("SELECT COUNT(*) FROM dbo.%s" % t); x = sc.fetchone()[0]
                    dc.execute("SELECT COUNT(*) FROM dbo.%s" % t); y = dc.fetchone()[0]
                    mark = '✓' if x == y else '≠'
                    self.log('INFO' if x == y else 'WARN', '对比', 0, '%s %s：%d → %d %s' % (mark, t, x, y, mark))
                except Exception as e:
                    self.log('WARN', '对比', 0, '%s 跳过：%s' % (t, e))
            # 商品 BaseNo 集合差异
            try:
                sc.execute("SELECT BaseNo FROM dbo.BasPtypeInner"); sx = set(r[0] for r in sc.fetchall())
                dc.execute("SELECT BaseNo FROM dbo.BasPtypeInner"); dx = set(r[0] for r in dc.fetchall())
                self.log('INFO', '对比', 0, '商品 BaseNo：源独有 %d（目标待新增），目标独有 %d，共有 %d' %
                         (len(sx - dx), len(dx - sx), len(sx & dx)))
            except Exception as e:
                self.log('WARN', '对比', 0, 'BaseNo 差异跳过：%s' % e)
            s.close(); d.close()
        self.submit('差异对比', worker)

    def run(self):
        self.mainloop()


if __name__ == '__main__':
    App().run()
