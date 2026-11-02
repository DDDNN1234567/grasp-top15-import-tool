# -*- coding: utf-8 -*-
"""
管家婆辉煌 TOP15 一键导入工具 v1.0（GUI）
功能：一键导入 基本资料(单位/仓库/商品含预设价格/职员) + 期初库存 + 客户价格追踪
通道：官方存储过程（自动维护编号/拼音/层级/updatetag）
"""
import os, sys, queue, threading, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import gpj_top15_core as core

DEFAULT_SERVER = 'DESKTOP-8P5G8VQ\\SQLSERVER'
DEFAULT_DB = 'HH15new'
DEFAULT_USER = 'sa'
DEFAULT_PWD = ''  # 发布版：不预填密码，由使用者自行输入


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('管家婆辉煌 TOP15 一键导入工具 v1.0')
        self.geometry('980x700')
        self.minsize(860, 620)
        self.log_q = queue.Queue()
        self.worker = None
        self.busy = False
        self.files = []
        self._build()
        self.after(100, self._drain_log)

    # ---------- 界面 ----------
    def _build(self):
        pad = {'padx': 6, 'pady': 3}
        frm_cfg = ttk.LabelFrame(self, text=' 连接配置 ')
        frm_cfg.pack(fill='x', padx=8, pady=(8, 4))
        ttk.Label(frm_cfg, text='服务器').grid(row=0, column=0, sticky='e', **pad)
        self.var_server = tk.StringVar(value=DEFAULT_SERVER)
        ttk.Entry(frm_cfg, textvariable=self.var_server, width=28).grid(row=0, column=1, **pad)
        ttk.Label(frm_cfg, text='数据库').grid(row=0, column=2, sticky='e', **pad)
        self.var_db = tk.StringVar(value=DEFAULT_DB)
        ttk.Entry(frm_cfg, textvariable=self.var_db, width=16).grid(row=0, column=3, **pad)
        ttk.Label(frm_cfg, text='用户名').grid(row=0, column=4, sticky='e', **pad)
        self.var_user = tk.StringVar(value=DEFAULT_USER)
        ttk.Entry(frm_cfg, textvariable=self.var_user, width=10).grid(row=0, column=5, **pad)
        ttk.Label(frm_cfg, text='密码').grid(row=0, column=6, sticky='e', **pad)
        self.var_pwd = tk.StringVar(value=DEFAULT_PWD)
        ttk.Entry(frm_cfg, textvariable=self.var_pwd, width=12, show='*').grid(row=0, column=7, **pad)
        self.btn_test = ttk.Button(frm_cfg, text='测试连接', command=self.test_conn)
        self.btn_test.grid(row=0, column=8, **pad)
        self.btn_check = ttk.Button(frm_cfg, text='账套自检', command=self.check_acct)
        self.btn_check.grid(row=0, column=9, **pad)
        self.lbl_cfg = ttk.Label(frm_cfg, text='就绪', foreground='#666')
        self.lbl_cfg.grid(row=0, column=10, **pad)

        frm_file = ttk.LabelFrame(self, text=' 数据文件（Excel，支持多选） ')
        frm_file.pack(fill='x', padx=8, pady=4)
        row0 = ttk.Frame(frm_file)
        row0.pack(fill='x', **pad)
        self.btn_add = ttk.Button(row0, text='选择 Excel…', command=self.pick_files)
        self.btn_add.pack(side='left', **pad)
        self.btn_tpl = ttk.Button(row0, text='生成模板', command=self.gen_tpl)
        self.btn_tpl.pack(side='left', **pad)
        self.btn_scan = ttk.Button(row0, text='扫描识别', command=self.scan_files)
        self.btn_scan.pack(side='left', **pad)
        self.lbl_files = ttk.Label(row0, text='未选择文件', foreground='#888')
        self.lbl_files.pack(side='left', **pad)
        self.lst_files = tk.Listbox(frm_file, height=3)
        self.lst_files.pack(fill='x', padx=6, pady=(0, 4))

        frm_mod = ttk.LabelFrame(self, text=' 导入模块 ')
        frm_mod.pack(fill='x', padx=8, pady=4)
        self.chk_vars = {}
        for i, (m, name) in enumerate(core.MODULE_NAMES.items()):
            v = tk.BooleanVar(value=True)
            self.chk_vars[m] = v
            ttk.Checkbutton(frm_mod, text=name, variable=v).grid(row=0, column=i, sticky='w', **pad)
        ttk.Label(frm_mod, text='  导入顺序：单位 → 仓库 → 商品 → 职员 → 期初库存 → 客户价格追踪',
                  foreground='#666').grid(row=1, column=0, columnspan=6, sticky='w', **pad)

        frm_act = ttk.Frame(self)
        frm_act.pack(fill='x', padx=8, pady=4)
        self.btn_run = ttk.Button(frm_act, text='★ 一键导入', command=self.run_import)
        self.btn_run.pack(side='left', **pad)
        self.btn_bak = ttk.Button(frm_act, text='立即备份', command=self.do_backup)
        self.btn_bak.pack(side='left', **pad)
        self.btn_roll = ttk.Button(frm_act, text='回滚(恢复备份)', command=self.do_rollback)
        self.btn_roll.pack(side='left', **pad)
        self.btn_cache = ttk.Button(frm_act, text='清客户端缓存', command=self.do_clearcache)
        self.btn_cache.pack(side='left', **pad)
        self.lbl_bak = ttk.Label(frm_act, text='', foreground='#888')
        self.lbl_bak.pack(side='left', **pad)
        self.var_bak = tk.StringVar()
        self.lbl_bak.config(textvariable=self.var_bak)

        frm_log = ttk.LabelFrame(self, text=' 执行日志 ')
        frm_log.pack(fill='both', expand=True, padx=8, pady=4)
        self.txt = scrolledtext.ScrolledText(frm_log, height=14, state='disabled',
                                             font=('Consolas', 9))
        self.txt.pack(fill='both', expand=True, padx=6, pady=6)

        frm_bar = ttk.Frame(self)
        frm_bar.pack(fill='x', padx=8, pady=(0, 8))
        self.prog = ttk.Progressbar(frm_bar, mode='indeterminate')
        self.prog.pack(side='left', fill='x', expand=True, **pad)
        self.var_stat = tk.StringVar(value='就绪')
        ttk.Label(frm_bar, textvariable=self.var_stat).pack(side='left', **pad)

    # ---------- 日志 ----------
    def log(self, kind, module, line, msg):
        now = datetime.datetime.now().strftime('%H:%M:%S')
        tag = {'OK': '✓', 'FAIL': '✘', 'WARN': '⚠', 'INFO': '·', 'PROG': '…'}.get(kind, '·')
        s = '[%s] %s %s' % (now, tag, msg)
        if module and module != '文件':
            s = '[%s] %s %s | %s' % (now, tag, module, msg)
            if line:
                s += ' | 行 %d' % line
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

    # ---------- 连接 ----------
    def _connect(self):
        return core.connect(self.var_server.get().strip(), self.var_db.get().strip(),
                            self.var_user.get().strip(), self.var_pwd.get())

    def test_conn(self):
        try:
            cn = self._connect()
            ver, ed = core.server_version(cn.cursor())
            cn.close()
            self.lbl_cfg.config(text='连接成功：%s（%s）' % (ver, ed), foreground='#080')
            messagebox.showinfo('连接成功', 'SQL Server %s（%s）\n数据库：%s' % (ver, ed, self.var_db.get()))
        except Exception as e:
            self.lbl_cfg.config(text='连接失败', foreground='#c00')
            messagebox.showerror('连接失败', str(e))

    def check_acct(self):
        try:
            cn = self._connect()
            ok, rep = core.check_account(cn.cursor())
            cn.close()
            for line in rep:
                self.log('INFO' if ok else 'WARN', '自检', 0, line)
            self.lbl_cfg.config(text='自检%s' % ('通过' if ok else '未通过'), foreground='#080' if ok else '#c00')
        except Exception as e:
            self.lbl_cfg.config(text='自检异常', foreground='#c00')
            self.log('FAIL', '自检', 0, str(e))

    # ---------- 文件 ----------
    def pick_files(self):
        files = filedialog.askopenfilenames(
            title='选择管家婆导入 Excel（支持多选）',
            filetypes=[('Excel 文件', '*.xls *.xlsx'), ('所有文件', '*.*')])
        if files:
            self.files = list(files)
            self.lst_files.delete(0, 'end')
            for f in self.files:
                self.lst_files.insert('end', os.path.basename(f))
            self.lbl_files.config(text='%d 个文件' % len(self.files))
            self.scan_files()

    def scan_files(self):
        if not self.files:
            return
        found = {}
        for f in self.files:
            try:
                sheets = core.read_excel(f)
            except Exception as e:
                self.log('FAIL', '扫描', 0, '%s：%s' % (os.path.basename(f), e))
                continue
            for sheetname, (hdr, rows) in sheets.items():
                m = core.SHEETS.get(sheetname)
                if m:
                    found.setdefault(m, []).append('%s（%d 行）' % (os.path.basename(f), len(rows)))
        for m, v in self.chk_vars.items():
            v.set(m in found)
        if found:
            desc = '；'.join('%s：%s' % (core.MODULE_NAMES[m], ', '.join(x)) for m, x in found.items())
            self.log('INFO', '扫描', 0, '识别到：' + desc)
            self.lbl_cfg.config(text='已识别 %d 个模块' % len(found), foreground='#080')
        else:
            self.log('WARN', '扫描', 0, '未识别到目标 sheet（单位/商品/仓库/职员/库存信息/单位价格跟踪）')

    def gen_tpl(self):
        path = filedialog.asksaveasfilename(title='保存导入模板',
                                            defaultextension='.xlsx',
                                            initialfile='管家婆TOP15一键导入模板.xlsx',
                                            filetypes=[('Excel 文件', '*.xlsx')])
        if not path:
            return
        try:
            core.generate_template(path)
            self.log('INFO', '模板', 0, '已生成模板：%s（含示例行，可删除后按列填写）' % path)
            messagebox.showinfo('完成', '模板已生成：\n%s' % path)
        except Exception as e:
            messagebox.showerror('失败', str(e))

    # ---------- 操作 ----------
    def _guard(self):
        if self.busy:
            messagebox.showwarning('提示', '已有任务正在执行，请等待完成')
            return False
        return True

    def _set_busy(self, busy):
        self.busy = busy
        for b in [self.btn_run, self.btn_bak, self.btn_roll, self.btn_cache,
                  self.btn_test, self.btn_check, self.btn_add, self.btn_scan]:
            b.configure(state='disabled' if busy else 'normal')

    def run_import(self):
        if not self._guard():
            return
        if not self.files:
            messagebox.showwarning('提示', '请先选择 Excel 数据文件')
            return
        modules = [m for m, v in self.chk_vars.items() if v.get()]
        if not modules:
            messagebox.showwarning('提示', '请至少勾选一个导入模块')
            return
        files = list(self.files)
        self._set_busy(True)
        self.prog.start(12)
        self.var_stat.set('导入中…')

        def work():
            try:
                cn = core.connect(self.var_server.get().strip(), self.var_db.get().strip(),
                                  self.var_user.get().strip(), self.var_pwd.get())
                cur = cn.cursor()
                core.orchestrate(cur, files, modules, lambda k, m, l, s: self.log(k, m, l, s))
                self.log('INFO', '缓存', 0, '建议清客户端缓存后重新登录查看（点击“清客户端缓存”）')
                cn.close()
            except Exception as e:
                self.log('FAIL', '连接', 0, str(e))
            finally:
                self.after(0, self._done)

        threading.Thread(target=work, daemon=True).start()

    def _done(self):
        self.prog.stop()
        self.var_stat.set('完成')
        self._set_busy(False)

    def do_backup(self):
        if not self._guard():
            return
        try:
            cn = self._connect()
            pre = core.backup_tables(cn.cursor())
            cn.close()
            self.var_bak.set('备份前缀：' + pre)
            self.log('INFO', '备份', 0, '已备份 9 张表，前缀 %s（回滚按钮可恢复）' % pre)
        except Exception as e:
            self.log('FAIL', '备份', 0, str(e))

    def do_rollback(self):
        if not self._guard():
            return
        pre = self.var_bak.get().replace('备份前缀：', '').strip()
        if not pre or not pre.startswith('_impbak_'):
            messagebox.showwarning('提示', '请先点击“立即备份”生成备份，再执行回滚')
            return
        if not messagebox.askyesno('确认回滚', '将用备份（%s）覆盖当前数据，确定继续？' % pre):
            return
        self._set_busy(True)
        self.prog.start(12)
        self.var_stat.set('回滚中…')

        def work():
            try:
                cn = self._connect()
                core.rollback_tables(cn.cursor(), pre)
                cn.close()
                self.log('INFO', '回滚', 0, '已从备份 %s 恢复 9 张表' % pre)
            except Exception as e:
                self.log('FAIL', '回滚', 0, str(e))
            finally:
                self.after(0, self._done)

        threading.Thread(target=work, daemon=True).start()

    def do_clearcache(self):
        if not self._guard():
            return
        ok, msg = core.clear_localdata(self.var_db.get().strip())
        self.log('INFO' if ok else 'FAIL', '缓存', 0, msg)
        if ok:
            messagebox.showinfo('完成', msg + '\n请重新打开管家婆登录账套查看')

    # ---------- 快捷键 ----------
    def run(self):
        self.mainloop()


if __name__ == '__main__':
    App().run()
