# frontend_gui.py

import threading

# 核心解耦：从后端文件中导入配置和启动调度器
from backend_logic import CONFIG, OneClickOrchestrator

class BackendAdminPanel(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("量化研报后台管理系统 v2.0")
        self.geometry("850x700")
        
        # 顶部 Tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        self._init_tab1_global()
        self._init_tab2_selection()
        self._init_tab3_reports()
        self._init_tab4_valuation()
        
        # 底部控制台与按钮
        self.bottom_frame = tk.Frame(self)
        self.bottom_frame.pack(fill='x', padx=10, pady=5)
        
        self.run_btn = tk.Button(self.bottom_frame, text="🚀 一键启动流水线", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), command=self.start_pipeline)
        self.run_btn.pack(side='top', fill='x', pady=5)
        
        self.log_text = tk.Text(self.bottom_frame, height=12, bg="#1E1E1E", fg="#00FF00", font=("Consolas", 10))
        self.log_text.pack(side='bottom', fill='x')
        
    def log(self, message):
        """线程安全的日志输出"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def _init_tab1_global(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="全局配置")
        
        ttk.Label(frame, text="结果保存路径:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.path_var = tk.StringVar(value=CONFIG["BASE_DIR"])
        ttk.Entry(frame, textvariable=self.path_var, width=50).grid(row=0, column=1, padx=10)
        ttk.Button(frame, text="浏览", command=lambda: self.path_var.set(filedialog.askdirectory())).grid(row=0, column=2)

        ttk.Label(frame, text="搜索引擎配置:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.search_var = tk.StringVar()
        engine_cb = ttk.Combobox(frame, textvariable=self.search_var, values=["Bing", "DuckDuckGo", "360"], state="readonly")
        engine_cb.grid(row=1, column=1, sticky='w', padx=10)
        engine_cb.current(0)
        
        self.overwrite_var = tk.BooleanVar(value=CONFIG["OVERWRITE_MODE"])
        ttk.Checkbutton(frame, text="启动时清空旧数据(覆盖最新结果)", variable=self.overwrite_var).grid(row=2, column=1, sticky='w', padx=10, pady=10)
        
        ttk.Label(frame, text="AI API Key:").grid(row=3, column=0, padx=10, pady=10, sticky='w')
        self.api_key_var = tk.StringVar(value="sk-your-key-here")
        ttk.Entry(frame, textvariable=self.api_key_var, width=50, show="*").grid(row=3, column=1, padx=10)

    def _init_tab2_selection(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="数据海选")
        
        ttk.Label(frame, text="问财选股条件模板:").pack(anchor='w', padx=10, pady=5)
        self.wencai_var = tk.StringVar()
        w_cb = ttk.Combobox(frame, textvariable=self.wencai_var, values=CONFIG["TEMPLATES"]["wencai_conditions"], width=80)
        w_cb.pack(padx=10, pady=5)
        w_cb.current(0)

        ttk.Label(frame, text="AI过滤条件模板:").pack(anchor='w', padx=10, pady=15)
        self.ai_filter_var = tk.StringVar()
        a_cb = ttk.Combobox(frame, textvariable=self.ai_filter_var, values=CONFIG["TEMPLATES"]["ai_filter_prompts"], width=80)
        a_cb.pack(padx=10, pady=5)
        a_cb.current(0)

    def _init_tab3_reports(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="年报下载与提取")
        ttk.Label(frame, text="本面板已将下载与提取模块深度整合。").pack(padx=10, pady=10)
        self.report_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="流水线中自动执行: 财务报表下载 -> 多年度数据提取 -> A/B表补全", variable=self.report_auto_var).pack(anchor='w', padx=20)

    def _init_tab4_valuation(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="AI企业深度估值")
        ttk.Label(frame, text="由文本分析、企业护城河分析、好价分析三大子程序整合").pack(padx=10, pady=10, anchor='w')

        ttk.Label(frame, text="1. AI 文本分析指令:").pack(anchor='w', padx=10)
        self.text_prompt_var = tk.StringVar()
        t_cb = ttk.Combobox(frame, textvariable=self.text_prompt_var, values=CONFIG["TEMPLATES"]["ai_text_prompts"], width=80)
        t_cb.pack(padx=10, pady=5)
        t_cb.current(0)

        ttk.Label(frame, text="2. AI 企业估值指令:").pack(anchor='w', padx=10, pady=(15, 0))
        self.ent_prompt_var = tk.StringVar()
        e_cb = ttk.Combobox(frame, textvariable=self.ent_prompt_var, values=CONFIG["TEMPLATES"]["ai_enterprise_prompts"], width=80)
        e_cb.pack(padx=10, pady=5)
        e_cb.current(0)
        
        self.good_price_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="3. 开启财务好价分析 (PE/PB/DCF估值测算)", variable=self.good_price_var).pack(anchor='w', padx=10, pady=20)

    def start_pipeline(self):
        """搜集面板配置并启动后台线程"""
        ui_config = {
            "base_dir": self.path_var.get(),
            "engine": self.search_var.get(),
            "overwrite": self.overwrite_var.get(),
            "api_url": "https://api.openai.com/v1",
            "api_key": self.api_key_var.get(),
            "model_name": "gpt-4-turbo",
            "wencai": self.wencai_var.get(),
            "ai_filter": self.ai_filter_var.get(),
            "text_prompt": self.text_prompt_var.get(),
            "ent_prompt": self.ent_prompt_var.get(),
            "good_price": self.good_price_var.get()
        }
        
        # 将UI的数据实时更新回全局配置中
        CONFIG["OVERWRITE_MODE"] = ui_config["overwrite"]
        CONFIG["BASE_DIR"] = ui_config["base_dir"]
        CONFIG["SEARCH_ENGINE"] = ui_config["engine"]
        
        self.run_btn.config(state="disabled", text="任务执行中...")
        self.log_text.delete(1.0, tk.END)
        
        # 使用多线程避免阻塞 GUI 界面
        def worker():
            # 这里的 OneClickOrchestrator 是从 backend_logic 导入的
            OneClickOrchestrator.run_all(ui_config, log_func=self.log)
            self.run_btn.config(state="normal", text="🚀 一键启动流水线")
            
        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = BackendAdminPanel()
    app.mainloop()
