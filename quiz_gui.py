import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
import os
from quiz_reader import QuizReader  # 导入原有的QuizReader类
import re
import random
import json
import time
import hashlib

class QuizApp:
    def __init__(self, root):
        """初始化答题应用"""
        self.root = root
        self.root.title("智能题库系统")
        self.root.geometry("1000x900")  # 增加窗口默认大小
        self.root.minsize(900, 700)     # 设置最小窗口大小
        self.root.configure(bg="#f5f6f7")  # 更现代的背景色
        
        # 初始化变量
        self.quiz = Quiz()  # 初始化 self.quiz
        self.current_mode = None  # normal, exam, review
        self.quiz_dir = None  # 存储选择的题库文件夹路径
        self.quiz_files = []  # 存储选择的题库文件列表
        self.all_questions = []  # 存储所有题目
        self.available_questions = {  # 存储每种类型的可用题目数量
            '单选题': 0,
            '多选题': 0,
            '判断题': 0
        }
        
        # 初始化答题记
        self.answered_questions = {}  # 存储已答题目的答案
        self.question_feedback = {}   # 存储题目的反馈信息
        self.question_status = {}     # 存储题目的答题状态(正确/错误)
        
        # 题目类型顺序
        self.question_type_order = ['单选题', '多选题', '判断题']
        
        # 考试计时器相关变量
        self.exam_start_time = None  # 考试开始时间
        self.exam_timer = None  # 考试计时器
        self.exam_duration = 0  # 考试持续时间(秒)
        
        # 错题本相关
        self.wrong_questions = {
            '单选题': {},  # {question_hash: {'question': question_dict, 'correct_count': 0}}
            '多选题': {},
            '判断题': {}
        }
        self.remove_threshold = 2  # 默认做对2次从错题本移除
        
        # 尝试加载已保存的错题本
        self.load_wrong_questions_from_json()
        
        # 初始化主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建欢迎页面
        self.create_welcome_page()
        self.create_file_select_page()
        self.create_quiz_page()
        
        # 设置样式
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f5f6f7")
        self.style.configure("TLabelframe", background="#f5f6f7")
        self.style.configure("TLabelframe.Label", font=('Microsoft YaHei', 10, 'bold'), background="#f5f6f7")
        self.style.configure("TButton", padding=8, font=('Microsoft YaHei', 10, 'bold'))
        self.style.configure("TLabel", padding=5, font=('Microsoft YaHei', 10), background="#f5f6f7")
        self.style.configure("Header.TLabel", font=('Microsoft YaHei', 16, 'bold'), background="#f5f6f7")
        self.style.configure("Score.TLabel", font=('Microsoft YaHei', 12, 'bold'), foreground='#0066cc', background="#f5f6f7")
        self.style.configure("Question.TLabel", font=('Microsoft YaHei', 11), background="#f5f6f7")
        self.style.configure("Feedback.TLabel", font=('Microsoft YaHei', 11), padding=10, background="#ffffff")
        self.style.configure("Correct.TRadiobutton", font=('Microsoft YaHei', 12, 'bold'), foreground='green')
        self.style.configure("Wrong.TRadiobutton", font=('Microsoft YaHei', 12, 'bold'), foreground='red')
        self.style.configure("Correct.TCheckbutton", font=('Microsoft YaHei', 12, 'bold'), foreground='green')
        self.style.configure("Wrong.TCheckbutton", font=('Microsoft YaHei', 12, 'bold'), foreground='red')
        
        # 配置一些颜色变量
        self.colors = {
            'primary': '#0066cc',
            'success': '#28a745',
            'danger': '#dc3545',
            'warning': '#ffc107',
            'light': '#f8f9fa',
            'dark': '#343a40'
        }
        
        # 初始显示欢迎页
        self.show_welcome_page()
        
        # 添加题库路径保存相关
        self.quiz_dir = None  # 存储选择的题库文件夹路径
        self.load_last_quiz_dir()  # 加载上次的题库路径
        
    def create_welcome_page(self):
        """创建欢迎页面"""
        self.welcome_frame = ttk.Frame(self.main_frame)
        
        # 欢迎标题
        title_label = ttk.Label(self.welcome_frame, 
                              text="欢迎使用智能题库系统",
                              style="Header.TLabel")
        title_label.pack(pady=50)
        
        # 选择题库按钮
        start_btn = ttk.Button(self.welcome_frame, 
                             text="进入题库系统",
                             command=self.show_file_select_page,
                             style="TButton")
        start_btn.pack(pady=20)

    def create_file_select_page(self):
        """创建文件选择页面"""
        self.file_select_frame = ttk.Frame(self.main_frame)
        
        # 标题
        title_label = ttk.Label(self.file_select_frame,
                              text="选择要答题的文档(可多选)",
                              style="Header.TLabel")
        title_label.pack(pady=(30, 20))
        
        # 选择文件夹按钮
        select_dir_btn = ttk.Button(self.file_select_frame,
                                  text="选择题库文件夹",
                                  command=self.select_quiz_directory)
        select_dir_btn.pack(pady=(0, 20))
        
        # 文件列表框架
        list_frame = ttk.Frame(self.file_select_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=20)
        
        # 创建文件列表(使用Treeview支持多选)
        self.file_list = ttk.Treeview(list_frame,
                                    selectmode='extended',
                                    columns=('name', 'questions'),
                                    show='headings',
                                    height=15)  # 设置显示行数
        
        # 设置列
        self.file_list.heading('name', text='文件名')
        self.file_list.heading('questions', text='题目数量')
        self.file_list.column('name', width=300, anchor='w')
        self.file_list.column('questions', width=300, anchor='w')
        
        # 添加滚动条
        y_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                                  command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=y_scrollbar.set)
        
        # 添加水平滚动条
        x_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL,
                                  command=self.file_list.xview)
        self.file_list.configure(xscrollcommand=x_scrollbar.set)
        
        # 布局列表和滚动条
        self.file_list.grid(row=0, column=0, sticky='nsew')
        y_scrollbar.grid(row=0, column=1, sticky='ns')
        x_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # 配置grid权重使列表可以扩展
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # 添加全选按钮
        select_all_btn = ttk.Button(self.file_select_frame,
                                  text="全选",
                                  command=self.select_all_files)
        select_all_btn.pack(pady=(0, 10))
        
        # 底部按钮框架
        button_frame = ttk.Frame(self.file_select_frame)
        button_frame.pack(pady=30)
        
        # 返回按钮
        back_btn = ttk.Button(button_frame,
                            text="返回",
                            command=self.show_welcome_page)
        back_btn.pack(side=tk.LEFT, padx=10)
        
        # 练习模式按钮
        self.practice_btn = ttk.Button(button_frame,
                                   text="练习模式",
                                   command=lambda: self.start_quiz("normal"),
                                   state='disabled')
        self.practice_btn.pack(side=tk.LEFT, padx=10)
        
        # 考试模式按钮
        self.exam_btn = ttk.Button(button_frame,
                                text="考试模式",
                                command=self.show_exam_config,
                                state='disabled')
        self.exam_btn.pack(side=tk.LEFT, padx=10)
        
        # 错题重做按钮
        total_wrong = sum(len(questions) for questions in self.wrong_questions.values())
        self.review_btn = ttk.Button(button_frame,
                                   text=f"错题重做({total_wrong}题)",
                                   command=self.show_wrong_questions_config,
                                   state='normal' if total_wrong > 0 else 'disabled')
        self.review_btn.pack(side=tk.LEFT, padx=10)

    def select_all_files(self):
        """全选文件列表中的所有文件"""
        for item in self.file_list.get_children():
            self.file_list.selection_add(item)
        self.on_file_select()

    def on_file_select(self, event=None):
        """处理文件选择事件"""
        selected = self.file_list.selection()
        if selected:
            self.practice_btn.state(['!disabled'])
            self.exam_btn.state(['!disabled'])
        else:
            self.practice_btn.state(['disabled'])
            self.exam_btn.state(['disabled'])

    def show_exam_config(self):
        """显示考试配置窗口"""
        # 创建配置窗口
        config_window = tk.Toplevel(self.root)
        config_window.title("考试模式配置")
        config_window.geometry("400x500")
        config_window.transient(self.root)  # 设置为主窗口的子窗口
        
        # 统计所有可用题目
        self.count_available_questions()
        
        # 加载上次考试配置
        last_config = self.load_last_exam_config()
        
        # 创建配置框架
        config_frame = ttk.Frame(config_window, padding="20")
        config_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(config_frame,
                              text="请选择各类型题目数量",
                              style="Header.TLabel")
        title_label.pack(pady=(0, 20))
        
        # 创建各类型题目数量选择
        spinbox_vars = {}
        for q_type in self.question_type_order:  # 按固定顺序添加题目
            # 创建框架
            type_frame = ttk.Frame(config_frame)
            type_frame.pack(fill='x', pady=10)
            
            # 标签
            ttk.Label(type_frame,
                     text=f"{q_type}(可用:{self.available_questions[q_type]}题):").pack(side=tk.LEFT)
            
            # 数量选择框
            var = tk.StringVar(value=str(last_config[q_type]))
            spinbox = ttk.Spinbox(type_frame,
                                from_=0,
                                to=self.available_questions[q_type],
                                width=5,
                                textvariable=var)
            spinbox.pack(side=tk.RIGHT)
            spinbox_vars[q_type] = var
        
        # 确认按钮
        ttk.Button(config_frame,
                  text="开始考试",
                  command=lambda: self.start_exam(spinbox_vars, config_window)).pack(pady=20)
        
        # 设置模态
        config_window.grab_set()
        config_window.focus_set()

    def load_last_exam_config(self):
        """加载上次考试配置"""
        try:
            config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'exam_config.json')
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading exam config: {str(e)}")
        return {'单选题': 0, '多选题': 0, '判断题': 0}

    def save_exam_config(self, config):
        """保存考试配置"""
        try:
            config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'exam_config.json')
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving exam config: {str(e)}")

    def count_available_questions(self):
        """统计所有可用题目"""
        self.all_questions = []  # 清空之前的题目
        self.available_questions = {'单选题': 0, '多选题': 0, '判断题': 0}
        
        # 获取选中的文件
        selected_indices = self.file_list.selection()
        for idx in selected_indices:
            file_name = self.file_list.item(idx)['values'][0]
            file_path = os.path.join(self.quiz_dir, file_name)
            
            # 加载题目
            temp_quiz = QuizReader(file_path)
            self.all_questions.extend(temp_quiz.questions)
            
            # 统计各类型题目数量
            for question in temp_quiz.questions:
                if question['type'] in self.available_questions:
                    self.available_questions[question['type']] += 1

    def start_exam(self, spinbox_vars, config_window):
        """开始考试模式"""
        # 获取用户选择的题目数量
        selected_counts = {
            q_type: int(var.get())
            for q_type, var in spinbox_vars.items()
        }
        
        # 保存考试配置
        self.save_exam_config(selected_counts)
        
        # 检查是否选择了题目
        total_selected = sum(selected_counts.values())
        if total_selected == 0:
            messagebox.showwarning("警告", "请至少选择一道题目!")
            return
        
        # 按类型和顺序选择题目
        selected_questions = []
        for q_type in self.question_type_order:  # 按固定顺序添加题目
            count = selected_counts[q_type]
            if count > 0:
                # 获取该类型的所有题目
                type_questions = [q for q in self.all_questions if q['type'] == q_type]
                # 随机选择指定数量的题目
                selected = random.sample(type_questions, min(count, len(type_questions)))
                selected_questions.extend(selected)
        
        # 创建新的QuizReader实例
        self.quiz = QuizReader(None)
        self.quiz.questions = selected_questions
        self.quiz.current_question = 0
        self.quiz.score = 0
        self.quiz.wrong_questions = []
        
        # 设置为考试模式
        self.current_mode = "exam"
        
        # 重置答题记录
        self.answered_questions = {}
        self.question_feedback = {}
        self.question_status = {}
        
        # 开始计时
        self.exam_start_time = time.time()
        self.exam_duration = 0
        self.update_exam_timer()
        
        # 关闭配置窗口
        config_window.destroy()
        
        # 显示答题页面
        self.show_quiz_page()
        self.display_question()

    def create_quiz_page(self):
        """创建答题页面"""
        self.quiz_frame = ttk.Frame(self.main_frame)
        
        # 顶部工具栏
        toolbar_frame = ttk.Frame(self.quiz_frame)
        toolbar_frame.pack(fill="x", pady=(0, 15))
        
        # 返回按钮
        back_btn = ttk.Button(toolbar_frame, text="返回选择文档",
                           command=self.confirm_return_to_select,
                           style="TButton")
        back_btn.pack(side=tk.LEFT, padx=20)
        
        # 题目导航按钮
        nav_btn = ttk.Button(toolbar_frame, text="题目导航",
                          command=self.show_question_navigator,
                          style="TButton")
        nav_btn.pack(side=tk.LEFT, padx=5)
        
        # 信息显示区
        stats_frame = ttk.Frame(toolbar_frame)
        stats_frame.pack(side=tk.RIGHT, padx=20)
        
        self.progress_label = ttk.Label(stats_frame, text="题目进度:0/0",
                                      style="Score.TLabel")
        self.progress_label.pack(side=tk.LEFT, padx=(0, 20))
        
        self.score_label = ttk.Label(stats_frame, text="当前得分:0",
                                   style="Score.TLabel")
        self.score_label.pack(side=tk.LEFT, padx=(0, 20))
        
        self.time_label = ttk.Label(stats_frame, text="用时:00:00:00",
                                  style="Score.TLabel")
        self.time_label.pack(side=tk.LEFT)
        
        # 分隔线
        ttk.Separator(self.quiz_frame, orient='horizontal').pack(fill='x')
        
        # 内容区域使用Grid布局
        content_frame = ttk.Frame(self.quiz_frame)
        content_frame.pack(fill="both", expand=True, padx=20)
        content_frame.grid_columnconfigure(0, weight=1)
        
        # 题目显示区
        question_frame = ttk.LabelFrame(content_frame, text="题目内容", padding=15)
        question_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 15))
        
        self.question_text = tk.Text(question_frame, wrap=tk.WORD, height=5,
                                   font=('Microsoft YaHei', 11),
                                   relief="flat", padx=15, pady=15,
                                   bg='#ffffff', border=0)
        self.question_text.pack(fill="both", expand=True)
        
        # 选项区域
        options_frame = ttk.LabelFrame(content_frame, text="选择答案", padding=15)
        options_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 15))
        
        self.option_var = tk.StringVar()
        self.option_buttons = []
        self.options_frame = options_frame
        
        # 答案反馈区域
        feedback_frame = ttk.Frame(content_frame)
        feedback_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 15))
        
        self.feedback_text = tk.Text(feedback_frame, wrap="word", height=5, bg="#ffffff", font=('Microsoft YaHei', 11))
        feedback_scrollbar = ttk.Scrollbar(feedback_frame, command=self.feedback_text.yview)
        self.feedback_text['yscrollcommand'] = feedback_scrollbar.set
        
        self.feedback_text.pack(side=tk.LEFT, fill="both", expand=True)
        feedback_scrollbar.pack(side=tk.RIGHT, fill="y")
        
        # 底部按钮区域
        button_frame = ttk.Frame(content_frame)
        button_frame.grid(row=3, column=0, sticky="ew", pady=15)
        button_frame.grid_columnconfigure(1, weight=1)  # 中间空白区域可伸缩
        
        # 导航按钮(左侧)
        nav_frame = ttk.Frame(button_frame)
        nav_frame.grid(row=0, column=0, sticky="w")
        
        self.prev_btn = ttk.Button(nav_frame, text="上一题",
                                 command=self.prev_question,
                                 style="TButton")
        self.prev_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.next_btn = ttk.Button(nav_frame, text="下一题",
                                 command=self.next_question,
                                 style="TButton")
        self.next_btn.pack(side=tk.LEFT)
        
        # 提交按钮(右侧)
        self.submit_btn = ttk.Button(button_frame, text="提交答案",
                                   command=self.handle_answer,
                                   style="TButton")
        self.submit_btn.grid(row=0, column=2, sticky="e")
        
        # 交卷按钮(右侧)
        self.finish_exam_btn = ttk.Button(button_frame, text="交卷",
                                        command=self.finish_exam,
                                        style="TButton")
        self.finish_exam_btn.grid(row=0, column=3, sticky="e", padx=(10, 0))
        self.finish_exam_btn.grid_remove()  # 默认隐藏

    def get_question_type_counts(self, questions):
        """统计各类型题目数量"""
        counts = {
            '单选题': 0,
            '多选题': 0,
            '判断题': 0
        }
        for question in questions:
            if question['type'] in counts:
                counts[question['type']] += 1
        return counts

    def select_quiz_directory(self):
        """选择题库文件夹"""
        new_dir = filedialog.askdirectory(title="选择题库文件夹")
        if new_dir:
            self.quiz_dir = new_dir
            self.save_quiz_dir()  # 保存新选择的路径
            self.load_quiz_files()

    def load_quiz_files(self):
        """加载文件夹中的题库文件"""
        # 清空文件列表
        for item in self.file_list.get_children():
            self.file_list.delete(item)
        
        # 获取文件夹中的所有.docx文件
        self.quiz_files = []
        for file in os.listdir(self.quiz_dir):
            if file.endswith('.docx'):
                file_path = os.path.join(self.quiz_dir, file)
                # 临时加载文件获取题目数量
                temp_quiz = QuizReader(file_path)
                
                # 获取各类型题目数量
                type_counts = self.get_question_type_counts(temp_quiz.questions)
                total_count = sum(type_counts.values())
                
                # 格式化题目统计信息
                type_info = []
                if type_counts['单选题'] > 0:
                    type_info.append(f"单选题:{type_counts['单选题']}")
                if type_counts['多选题'] > 0:
                    type_info.append(f"多选题:{type_counts['多选题']}")
                if type_counts['判断题'] > 0:
                    type_info.append(f"判断题:{type_counts['判断题']}")
                
                questions_info = f"{total_count}题 ({', '.join(type_info)})"
                
                # 添加到列表
                self.quiz_files.append(file_path)
                self.file_list.insert('', 'end', values=(file, questions_info))
        
        # 绑定选择事件
        self.file_list.bind('<<TreeviewSelect>>', self.on_file_select)

    def start_quiz(self, mode):
        """开始答题"""
        selected_indices = self.file_list.selection()
        if not selected_indices:
            return
        
        # 获取选中的文件路径
        selected_files = []
        for idx in selected_indices:
            file_name = self.file_list.item(idx)['values'][0]
            file_path = os.path.join(self.quiz_dir, file_name)
            selected_files.append(file_path)
        
        # 合并所有选中文件的题目
        all_questions = []
        for file_path in selected_files:
            temp_quiz = QuizReader(file_path)
            all_questions.extend(temp_quiz.questions)
        
        # 创建新的QuizReader实例
        self.quiz = QuizReader(selected_files[0])  # 使用第一个文件初始化
        self.quiz.questions = all_questions  # 替换为合并后的题目
        self.quiz.current_question = 0
        self.quiz.score = 0
        self.quiz.wrong_questions = []  # 重置错题列表
        
        # 重置答题记录
        self.answered_questions = {}
        self.question_feedback = {}
        self.question_status = {}
        
        # 设置模式
        self.current_mode = mode
        
        # 显示答题页面
        self.show_quiz_page()
        self.display_question()

    def show_welcome_page(self):
        """显示欢迎页面"""
        self.quiz_frame.pack_forget() if hasattr(self, 'quiz_frame') else None
        self.file_select_frame.pack_forget() if hasattr(self, 'file_select_frame') else None
        self.welcome_frame.pack(fill=tk.BOTH, expand=True)
        
        # 停止计时器
        if self.exam_timer:
            self.root.after_cancel(self.exam_timer)
            self.exam_timer = None
            self.exam_start_time = None

    def show_file_select_page(self):
        """显示文件选择页面"""
        self.welcome_frame.pack_forget()
        self.quiz_frame.pack_forget() if hasattr(self, 'quiz_frame') else None
        self.file_select_frame.pack(fill=tk.BOTH, expand=True)

    def show_quiz_page(self):
        """显示答题页面"""
        self.welcome_frame.pack_forget()
        self.file_select_frame.pack_forget()
        self.quiz_frame.pack(fill=tk.BOTH, expand=True)
        
    def display_question(self):
        """显示当前题目"""
        if self.current_mode == "review":
            question = self.quiz.questions[self.quiz.current_question]
            total = len(self.quiz.questions)
        else:
            question = self.quiz.questions[self.quiz.current_question]
            total = len(self.quiz.questions)
            
        # 更新进度和分数
        self.progress_label.config(
            text=f"题目进度:{self.quiz.current_question + 1}/{total}")
        self.score_label.config(
            text=f"当前得分:{self.quiz.score}/{total}")
            
        # 格式化并显示题目
        question_text = question['question'].strip()
        if not question_text.endswith(('?', '?', '.', '.')):
            question_text += '.'
        
        # 对于多选题,添加提示
        if question['type'] == '多选题':
            question_text = "[多选题] " + question_text
        
        self.question_text.delete('1.0', tk.END)
        self.question_text.insert('1.0', question_text)
        
        # 清除旧选项和框架
        for widget in self.options_frame.winfo_children():
            widget.destroy()
        self.option_buttons.clear()
        
        # 创建新选项
        if question['type'] == "判断题":
            options = [('T', '对'), ('F', '错')]
            self.option_var = tk.StringVar()  # 单选
            widget_class = ttk.Radiobutton
        elif question['type'] == "多选题":
            options = [(opt[0], opt) for opt in question['options']]
            self.option_vars = []  # 多选用多个变量
            widget_class = ttk.Checkbutton
        else:  # 单选题
            options = [(opt[0], opt) for opt in question['options']]
            self.option_var = tk.StringVar()  # 单选
            widget_class = ttk.Radiobutton
            
        # 创建选项框架
        options_inner_frame = ttk.Frame(self.options_frame)
        options_inner_frame.pack(fill="both", expand=True, padx=20, pady=10)
            
        for i, (value, text) in enumerate(options):
            if question['type'] == "多选题":
                var = tk.BooleanVar()
                self.option_vars.append(var)
                btn = widget_class(
                    options_inner_frame,
                    text=text.strip(),
                    variable=var,
                    style="TCheckbutton"
                )
                # 保存选项值用于后续判断
                btn.option_value = value
            else:
                btn = widget_class(
                    options_inner_frame,
                    text=text.strip(),
                    value=value,
                    variable=self.option_var,
                    style="TRadiobutton"
                )
            btn.pack(anchor="w", pady=8)
            self.option_buttons.append(btn)
        
        # 恢复之前的答案和反馈(如果有)
        question_id = f"{self.current_mode}_{self.quiz.current_question}"
        if question_id in self.answered_questions:
            # 恢复选择的答案
            if question['type'] == "多选题":
                selected_answers = self.answered_questions[question_id]
                for i, var in enumerate(self.option_vars):
                    var.set(options[i][0] in selected_answers)
            else:
                self.option_var.set(self.answered_questions[question_id])
            
            self.submit_btn.state(['disabled'])
            
            # 恢复答案反馈
            if question_id in self.question_feedback:
                self.feedback_text.delete('1.0', tk.END)
                self.feedback_text.insert(tk.END, self.question_feedback[question_id])
                self.feedback_text.configure(foreground='green' if self.question_status[question_id] else 'red')
                
                # 恢复选项颜色
                correct_answer = question['answer'].upper()
                if question['type'] == "多选题":
                    correct_answers = ''.join(c for c in correct_answer if c.isalpha())
                    for btn in self.option_buttons:
                        if btn.option_value in correct_answers:
                            btn.configure(style="Correct.TCheckbutton")
                        elif btn.option_value in selected_answers and btn.option_value not in correct_answers:
                            btn.configure(style="Wrong.TCheckbutton")
                else:
                    for btn in self.option_buttons:
                        if btn.cget('value') == correct_answer:
                            btn.configure(style="Correct.TRadiobutton")
                        elif btn.cget('value') == self.answered_questions[question_id] and btn.cget('value') != correct_answer:
                            btn.configure(style="Wrong.TRadiobutton")
        else:
            # 重置选择
            if question['type'] == "多选题":
                for var in self.option_vars:
                    var.set(False)
            else:
                self.option_var.set('')
                
            self.submit_btn.state(['!disabled'])
            self.feedback_text.delete('1.0', tk.END)
            
            # 重置所有选项样式
            for btn in self.option_buttons:
                if question['type'] == "多选题":
                    btn.configure(style="TCheckbutton")
                else:
                    btn.configure(style="TRadiobutton")
        
        # 检查题目状态并调整按钮布局
        question_id = f"{self.current_mode}_{self.quiz.current_question}"
        if self.current_mode == "exam" and question_id in self.question_status and not self.question_status[question_id]:
            # 如果是考试模式且题目答错,保留"回答错误"按钮布局
            self.next_btn.pack_forget()  # 隐藏原来的下一题按钮
            self.submit_btn.config(text="下一题", command=self.next_question)
            # 确保导航按钮可用
            self.prev_btn.state(['!disabled'])
            if self.quiz.current_question < len(self.quiz.questions) - 1:
                self.submit_btn.state(['!disabled'])
        else:
            # 练习模式或其他情况,保持常规按钮布局
            self.submit_btn.config(text="提交答案", command=self.handle_answer)
            self.next_btn.pack()  # 显示下一题按钮
            # 如果题目已经回答过,启用导航按钮
            if question_id in self.question_status:
                self.prev_btn.state(['!disabled'])
                self.next_btn.state(['!disabled'])
        
        # 检查是否为考试模式的最后一题
        if self.current_mode == "exam" and self.quiz.current_question == len(self.quiz.questions) - 1:
            # 如果已经回答过这题,显示交卷按钮
            if question_id in self.question_status:
                self.next_btn.pack_forget()  # 隐藏下一题按钮
                self.finish_exam_btn.grid()  # 显示交卷按钮
        
        # 更新导航按钮状态
        self.update_navigation_buttons()

    def prev_question(self):
        """显示上一题"""
        if self.quiz.current_question > 0:
            self.quiz.current_question -= 1
            self.display_question()
            self.update_navigation_buttons()
    
    def next_question(self):
        """显示下一题"""
        if self.current_mode == "review":
            total = len(self.quiz.questions)
        else:
            total = len(self.quiz.questions)
            
        if self.quiz.current_question < total - 1:
            self.quiz.current_question += 1
            self.display_question()
            self.update_navigation_buttons()
        else:
            self.show_quiz_complete()
        
        # 重置按钮为"提交答案"
        self.submit_btn.config(text="提交答案", command=self.handle_answer)
        self.submit_btn.state(['!disabled'])

    def update_navigation_buttons(self):
        """更新导航按钮状态"""
        # 处理上一题按钮
        if self.quiz.current_question == 0:
            self.prev_btn.state(['disabled'])
        else:
            self.prev_btn.state(['!disabled'])
        
        # 处理下一题按钮
        if self.current_mode == "review":
            total = len(self.quiz.questions)
        else:
            total = len(self.quiz.questions)
            
        if self.quiz.current_question >= total - 1:
            self.next_btn.state(['disabled'])
        else:
            self.next_btn.state(['!disabled'])
    
    def handle_answer(self):
        """处理答案提交"""
        # 获取当前题目
        if self.current_mode == "review":
            question = self.quiz.questions[self.quiz.current_question]
            total = len(self.quiz.questions)
        else:
            question = self.quiz.questions[self.quiz.current_question]
            total = len(self.quiz.questions)
            
        # 获取答案
        if question['type'] == "多选题":
            # 收集所有选中的选项
            selected_values = []
            for i, var in enumerate(self.option_vars):
                if var.get():
                    selected_values.append(question['options'][i][0])
            
            if not selected_values:
                messagebox.showwarning("警告", "请至少选择一个选项!")
                return
                
            # 对选中的选项排序
            selected_values.sort()
            answer = ''.join(selected_values)  # 多选题答案直接连接,不使用逗号
            
            # 处理正确答案,移除所有空格和逗号
            correct_answer = ''.join(c for c in question['answer'].upper() if c.isalpha())
        else:
            answer = self.option_var.get()
            if not answer:
                messagebox.showwarning("警告", "请选择一个答案!")
                return
            correct_answer = question['answer'].upper()

        # 检查答案
        is_correct = answer.upper() == correct_answer

        # 更新选项颜色
        if question['type'] == "多选题":
            for btn in self.option_buttons:
                if btn.option_value in correct_answer:
                    btn.configure(style="Correct.TCheckbutton")
                elif btn.option_value in answer and btn.option_value not in correct_answer:
                    btn.configure(style="Wrong.TCheckbutton")
                else:
                    btn.configure(style="TCheckbutton")
        else:
            for btn in self.option_buttons:
                if btn.cget('value') == correct_answer:
                    btn.configure(style="Correct.TRadiobutton")
                elif btn.cget('value') == answer and btn.cget('value') != correct_answer:
                    btn.configure(style="Wrong.TRadiobutton")
                else:
                    btn.configure(style="TRadiobutton")

        # 获取选项文本
        if question['type'] == "判断题":
            selected_text = "对" if answer == "T" else "错"
            correct_text = "对" if correct_answer == "T" else "错"
        elif question['type'] == "多选题":
            # 获取选中的选项文本
            selected_text = []
            correct_text = []
            
            # 处理用户选择的答案
            for opt in question['options']:
                if opt[0] in answer:
                    # 提取选项内容(去掉选项标记和点)
                    opt_content = re.sub(r'^[A-Z][.、\s]', '', opt).strip()
                    selected_text.append(f"{opt[0]}. {opt_content}")
                    
            # 处理正确答案
            for opt in question['options']:
                if opt[0] in correct_answer:
                    # 提取选项内容(去掉选项标记和点)
                    opt_content = re.sub(r'^[A-Z][.、\s]', '', opt).strip()
                    correct_text.append(f"{opt[0]}. {opt_content}")
                    
            selected_text = "\n".join(selected_text)  # 每个选项单独一行
            correct_text = "\n".join(correct_text)    # 每个选项单独一行
        else:
            # 提取选项内容(去掉选项标记和点)
            selected_text = ""
            correct_text = ""
            for opt in question['options']:
                if opt.startswith(answer):
                    selected_text = re.sub(r'^[A-Z][.、\s]', '', opt).strip()
                    selected_text = f"{opt[0]}. {selected_text}"
                if opt.startswith(correct_answer):
                    correct_text = re.sub(r'^[A-Z][.、\s]', '', opt).strip()
                    correct_text = f"{opt[0]}. {correct_text}"
        
        # 更新分数和错题本
        if is_correct:
            if not self.quiz.is_review_mode:
                self.quiz.score += 1
            feedback = f"✓ 回答正确!\n你的答案:\n{selected_text}"
            self.feedback_text.configure(foreground='green')
            
            # 答对自动跳转到下一题
            if self.quiz.current_question < total - 1:
                self.root.after(1000, self.next_question)  # 延迟1秒后跳转
            else:
                self.root.after(1000, self.show_quiz_complete)  # 如果是最后一题,显示完成信息
        else:
            if not self.quiz.is_review_mode:
                self.quiz.wrong_questions.append(question)
            feedback = f"✗ 回答错误!\n你的答案:\n{selected_text}\n\n正确答案:\n{correct_text}"
            self.feedback_text.configure(foreground='red')
            # 答错不自动跳转,改为下一题按钮
            self.next_btn.pack_forget()  # 隐藏下一题按钮
            self.submit_btn.config(text="下一题", command=self.next_question)
        
        # 显示答案反馈
        self.feedback_text.delete('1.0', tk.END)
        self.feedback_text.insert(tk.END, feedback)
        
        # 保存答案和反馈状态
        question_id = f"{self.current_mode}_{self.quiz.current_question}"
        self.answered_questions[question_id] = answer
        self.question_feedback[question_id] = feedback
        self.question_status[question_id] = is_correct
               
        # 禁用提交按钮(在答对情况下)
        if is_correct:
            self.submit_btn.state(['disabled'])
        
        # 更新分数显示
        self.score_label.config(text=f"当前得分:{self.quiz.score}/{total}")
        
        # 更新错题本
        if self.current_mode == "normal":  # 只在练习模式下记录错题
            question_hash = self.get_question_hash(question)
            if is_correct:
                # 如果在错题本中且答对了,增加正确次数
                if question['type'] in self.wrong_questions and \
                   question_hash in self.wrong_questions[question['type']]:
                    self.wrong_questions[question['type']][question_hash]['correct_count'] += 1
                    # 检查是否达到移除阈值
                    if self.wrong_questions[question['type']][question_hash]['correct_count'] >= self.remove_threshold:
                        del self.wrong_questions[question['type']][question_hash]
            else:
                # 答错了,添加到错题本或重置正确次数
                if question_hash not in self.wrong_questions[question['type']]:
                    self.wrong_questions[question['type']][question_hash] = {
                        'question': question,
                        'correct_count': 0
                    }
                else:
                    self.wrong_questions[question['type']][question_hash]['correct_count'] = 0
            
            # 保存错题本
            self.save_wrong_questions()

    def show_quiz_complete(self):
        """显示测验完成信息"""
        if self.current_mode == "normal":
            message = f"""
测验完成!最终统计:
总题数:{self.quiz.total_score}
答对题数:{self.quiz.score}
答错题数:{self.quiz.total_score - self.quiz.score}
正确率:{(self.quiz.score/self.quiz.total_score)*100:.1f}%

是否要重新开始?
            """
        else:
            message = "错题重做完成!是否要重新开始?"
            
        if messagebox.askyesno("完成", message):
            self.show_welcome_page()
        else:
            self.root.quit()

    def confirm_return_to_select(self):
        """确认是否返回选择文档页面"""
        if messagebox.askyesno("确认返回", 
                             "返回选择文档页面将丢失当前答题进度,确定要返回吗?"):
            self.return_to_select()

    def return_to_select(self):
        """返回选择文档页面"""
        # 重置答题状态
        self.answered_questions = {}
        self.question_feedback = {}
        self.question_status = {}
        
        # 显示文件选择页面
        self.show_file_select_page()

    def show_question_navigator(self):
        """显示题目导航器"""
        # 创建一个新的顶层窗口
        self.nav_window = tk.Toplevel(self.root)
        self.nav_window.title("题目导航")
        self.nav_window.geometry("800x600")
        self.nav_window.transient(self.root)  # 设置为主窗口的子窗口
        
        # 创建主框架
        main_frame = ttk.Frame(self.nav_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Label(title_frame, text="题目导航",
                 style="Header.TLabel").pack(side=tk.LEFT)
        
        # 题目类型统计
        if self.current_mode == "review":
            total = len(self.quiz.questions)
        else:
            total = len(self.quiz.questions)
            type_counts = self.get_question_type_counts(self.quiz.questions)
            stats = []
            if type_counts['单选题'] > 0:
                stats.append(f"单选题:{type_counts['单选题']}")
            if type_counts['多选题'] > 0:
                stats.append(f"多选题:{type_counts['多选题']}")
            if type_counts['判断题'] > 0:
                stats.append(f"判断题:{type_counts['判断题']}")
            ttk.Label(title_frame, 
                     text=f"共{total}题 ({', '.join(stats)})",
                     style="Score.TLabel").pack(side=tk.RIGHT)
        
        # 创建滚动区域
        canvas = tk.Canvas(main_frame, bg='#ffffff')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", 
                                command=canvas.yview)
        
        # 创建网格框架
        self.grid_frame = ttk.Frame(canvas)
        
        # 配置画布
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 打包滚动组件
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 创建窗口
        canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        
        # 创建题目按钮网格
        self.create_question_grid()
        
        # 更新滚动区域
        self.grid_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        
        # 绑定鼠标滚轮
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(
            int(-1*(e.delta/120)), "units"))

    def create_question_grid(self):
        """创建题目按钮网格"""
        # 每行显示8个按钮
        COLS = 8
        
        if self.current_mode == "review":
            questions = self.quiz.questions
        else:
            questions = self.quiz.questions
        
        for i, question in enumerate(questions):
            row = i // COLS
            col = i % COLS
            
            # 创建按钮框架
            btn_frame = ttk.Frame(self.grid_frame)
            btn_frame.grid(row=row, column=col, padx=5, pady=5)
            
            # 获取题目状态
            question_id = f"{self.current_mode}_{i}"
            if question_id in self.answered_questions:
                if self.question_status[question_id]:
                    bg_color = '#28a745'  # 正确
                    fg_color = 'white'
                else:
                    bg_color = '#dc3545'  # 错误
                    fg_color = 'white'
            else:
                bg_color = '#f8f9fa'  # 未答
                fg_color = 'black'
            
            # 创建按钮
            btn = tk.Button(btn_frame,
                          text=f"{i + 1}",
                          width=6,
                          height=2,
                          bg=bg_color,
                          fg=fg_color,
                          relief='flat',
                          command=lambda idx=i: self.jump_to_question(idx))
            btn.pack(expand=True, fill=tk.BOTH)
            
            # 添加题目类型提示
            ttk.Label(btn_frame,
                     text=question['type'][0], # 显示类型首字
                     font=('Microsoft YaHei', 8)).pack()

    def jump_to_question(self, index):
        """跳转到指定题目"""
        if self.current_mode == "review":
            total = len(self.quiz.questions)
        else:
            total = len(self.quiz.questions)
            
        if 0 <= index < total:
            self.quiz.current_question = index
            self.display_question()
            self.nav_window.destroy()

    def update_exam_timer(self):
        """更新考试计时器"""
        if self.current_mode == "exam" and self.exam_start_time is not None:
            self.exam_duration = int(time.time() - self.exam_start_time)
            hours = self.exam_duration // 3600
            minutes = (self.exam_duration % 3600) // 60
            seconds = self.exam_duration % 60
            self.time_label.config(text=f"用时:{hours:02d}:{minutes:02d}:{seconds:02d}")
            # 每秒更新一次
            self.exam_timer = self.root.after(1000, self.update_exam_timer)

    def finish_exam(self):
        """交卷并显示考试结果"""
        # 停止计时器
        if self.exam_timer:
            self.root.after_cancel(self.exam_timer)
            self.exam_timer = None
        
        # 计算得分
        total_questions = len(self.quiz.questions)
        correct_answers = sum(1 for status in self.question_status.values() if status)
        score = int((correct_answers / total_questions) * 100)
        
        # 创建结果窗口
        result_window = tk.Toplevel(self.root)
        result_window.title("考试结果")
        result_window.geometry("400x300")
        result_window.transient(self.root)
        
        # 创建结果框架
        result_frame = ttk.Frame(result_window, padding="20")
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(result_frame, 
                 text="考试完成!",
                 style="Header.TLabel").pack(pady=(0, 20))
        
        # 详细信息
        info_frame = ttk.Frame(result_frame)
        info_frame.pack(fill=tk.X, pady=10)
        
        # 计算用时
        hours = self.exam_duration // 3600
        minutes = (self.exam_duration % 3600) // 60
        seconds = self.exam_duration % 60
        time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
        # 显示统计信息
        ttk.Label(info_frame, 
                 text=f"总题数:{total_questions}题",
                 style="Score.TLabel").pack(pady=5)
        ttk.Label(info_frame,
                 text=f"正确题数:{correct_answers}题",
                 style="Score.TLabel").pack(pady=5)
        ttk.Label(info_frame,
                 text=f"最终得分:{score}分",
                 style="Score.TLabel").pack(pady=5)
        ttk.Label(info_frame,
                 text=f"答题用时:{time_str}",
                 style="Score.TLabel").pack(pady=5)
        
        # 返回按钮
        ttk.Button(result_frame,
                  text="返回主页",
                  command=lambda: [result_window.destroy(), self.show_welcome_page()],
                  style="TButton").pack(pady=20)
        
        # 设置模态
        result_window.grab_set()
        result_window.focus_set()

    def show_wrong_questions_config(self):
        """显示错题重做配置窗口"""
        config_window = tk.Toplevel(self.root)
        config_window.title("错题重做设置")
        config_window.geometry("400x500")
        config_window.transient(self.root)
        
        # 创建配置框架
        config_frame = ttk.Frame(config_window, padding="20")
        config_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(config_frame,
                 text="错题重做设置",
                 style="Header.TLabel").pack(pady=(0, 20))
        
        # 显示各类型错题数量
        for q_type in self.question_type_order:
            count = len(self.wrong_questions[q_type])
            ttk.Label(config_frame,
                     text=f"{q_type}:{count}题",
                     style="Score.TLabel").pack(pady=5)
        
        # 设置移除阈值
        threshold_frame = ttk.Frame(config_frame)
        threshold_frame.pack(fill=tk.X, pady=20)
        
        ttk.Label(threshold_frame,
                 text="连续答对次数达到后移除:",
                 style="Score.TLabel").pack(side=tk.LEFT)
        
        threshold_var = tk.StringVar(value=str(self.remove_threshold))
        threshold_spinbox = ttk.Spinbox(threshold_frame,
                                      from_=1,
                                      to=10,
                                      width=5,
                                      textvariable=threshold_var)
        threshold_spinbox.pack(side=tk.LEFT, padx=5)
        
        # 开始按钮
        ttk.Button(config_frame,
                  text="开始练习",
                  command=lambda: self.start_wrong_questions_review(int(threshold_var.get()), config_window),
                  style="TButton").pack(pady=20)
        
        # 设置模态
        config_window.grab_set()
        config_window.focus_set()

    def start_wrong_questions_review(self, threshold=None, config_window=None):
        """开始错题重做"""
        if threshold is not None:
            self.remove_threshold = threshold
        
        # 检查是否有错题可供重做
        total_wrong_questions = sum(len(questions) for questions in self.wrong_questions.values())
        if total_wrong_questions == 0:
            messagebox.showinfo("提示", "当前没有错题")
            if config_window:
                config_window.destroy()
            return
        
        # 获取所有错题
        all_wrong_questions = []
        for q_type, questions in self.wrong_questions.items():
            for q_hash, q_data in questions.items():
                question = q_data['question']  # 直接使用保存的题目数据
                all_wrong_questions.append(question)
        
        # 随机打乱错题顺序
        random.shuffle(all_wrong_questions)
        
        # 初始化错题练习
        self.current_mode = "review"
        self.quiz = Quiz()
        self.quiz.questions = all_wrong_questions
        self.quiz.current_question = 0
        self.quiz.score = 0
        self.quiz.wrong_questions = []
        self.quiz.is_review_mode = True
        
        self.question_status = {}
        
        # 关闭配置窗口并显示答题页面
        if config_window:
            config_window.destroy()
        self.show_quiz_page()
        self.display_question()
        
        # 添加调试信息
        print(f"加载了 {len(all_wrong_questions)} 道错题")
        print(f"当前题目索引: {self.quiz.current_question}")
        print(f"题目列表长度: {len(self.quiz.questions)}")

    def get_question_hash(self, question):
        """生成题目的唯一标识"""
        # 使用题目内容和选项生成哈希值
        question_text = question['question']
        options = sorted(question['options'])  # 排序选项以确保相同选项不同顺序的题目有相同的哈希值
        content = question_text + ''.join(options)
        return hashlib.md5(content.encode()).hexdigest()
    
    def save_wrong_questions(self):
        """保存错题本到文件"""
        save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'wrong_questions.json')
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                data = {
                    'questions': self.wrong_questions,
                    'threshold': self.remove_threshold
                }
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存错题本时出错:{e}")
    
    def load_wrong_questions_from_json(self):
        """从JSON文件加载错题"""
        save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'wrong_questions.json')
        try:
            if os.path.exists(save_path):
                with open(save_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.wrong_questions = data['questions']
                    self.remove_threshold = data['threshold']
        except Exception as e:
            print(f"加载错题本时出错:{e}")

    def load_last_quiz_dir(self):
        """加载上次使用的题库路径"""
        try:
            config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'quiz_config.json')
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'quiz_dir' in config and os.path.exists(config['quiz_dir']):
                        self.quiz_dir = config['quiz_dir']
                        # 如果有保存的路径,自动加载题库
                        self.show_file_select_page()
                        self.load_quiz_files()
        except Exception as e:
            print(f"Error loading quiz directory: {str(e)}")

    def save_quiz_dir(self):
        """保存当前题库路径"""
        try:
            config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'quiz_config.json')
            config = {'quiz_dir': self.quiz_dir}
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving quiz directory: {str(e)}")

class Quiz:
    def __init__(self, files=None, type_counts=None, questions=None):
        """初始化测验"""
        self.current_question = 0
        self.score = 0
        self.questions = []  # 初始化为空列表
        self.is_review_mode = False
        
        if questions is not None:
            # 直接使用传入的题目列表(用于错题重做)
            self.questions = questions
            random.shuffle(self.questions)  # 随机打乱顺序
        elif files and type_counts:
            # 从文件加载题目
            self.load_questions(files, type_counts)

    def load_questions(self, files, type_counts):
        """从文件加载题目"""
        all_questions = []
        
        # 读取所有文件中的题目
        for file in files:
            try:
                doc = Document(file)
                questions = self.parse_questions(doc)
                all_questions.extend(questions)
            except Exception as e:
                print(f"读取文件 {file} 时出错:{e}")
        
        # 按类型筛选题目
        for q_type, count in type_counts.items():
            type_questions = [q for q in all_questions if q['type'] == q_type]
            if count > 0:
                # 随机选择指定数量的题目
                selected = random.sample(type_questions, min(count, len(type_questions)))
                self.questions.extend(selected)
        
        # 随机打乱题目顺序
        random.shuffle(self.questions)

    def parse_questions(self, doc):
        """解析文档中的题目"""
        questions = []
        current_question = None
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            
            # 检查是否是新题目
            if re.match(r'^\d+[.、]', text):
                # 保存上一个题目(如果有)
                if current_question:
                    questions.append(current_question)
                
                # 开始新题目
                question_text = re.sub(r'^\d+[.、]\s*', '', text)
                current_question = {
                    'question': question_text,
                    'options': [],
                    'correct': [],
                    'type': '单选题'  # 默认为单选题
                }
            
            # 检查选项   
            elif text.startswith(('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H')):
                if current_question:
                    option_text = text[1:].strip('.::、 ')
                    current_question['options'].append(option_text)
                    
                    # 检查是否是正确答案(通过落样式或特殊标记)
                    if any(run.bold or run.underline for run in para.runs):
                        current_question['correct'].append(option_text)
            
            # 检查答案标记
            elif text.startswith('答案') and current_question:
                answer_text = text.split(':')[-1].strip()
                # 处理答案格式
                if answer_text in ['对', '错']:
                    current_question['type'] = '判断题'
                    current_question['correct'] = [answer_text]
                else:
                    # 解析选项答案
                    selected_options = []
                    for option in answer_text:
                        idx = ord(option.upper()) - ord('A')
                        if 0 <= idx < len(current_question['options']):
                            selected_options.append(current_question['options'][idx])
                    if len(selected_options) > 1:
                        current_question['type'] = '多选题'
                    current_question['correct'] = selected_options
        
        # 添加最后一个题目
        if current_question:
            questions.append(current_question)
        
        return questions

def main():
    root = tk.Tk()
    app = QuizApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()


