from docx import Document
import re
import os

class QuizReader:
    def __init__(self, docx_path):
        self.document = Document(docx_path)
        self.questions = []
        self.current_question = 0
        self.score = 0
        self.total_score = 0
        self.wrong_questions = []  # 存储错题
        self.is_review_mode = False  # 是否是错题重做模式
        self.parse_questions()

    def parse_questions(self):
        current_question = None
        current_options = []
        current_answer = None
        
        for paragraph in self.document.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue

            # 检查是否是答案
            if text.startswith(('答案：', 'Answer:', '答案:', 'Answer：', '正确答案:', '正确答案：')):
                if text.startswith(('正确答案:', '正确答案：')):
                    current_answer = text.split(':')[-1].strip()
                    if current_answer == "对":
                        current_answer = "T"
                    elif current_answer == "错":
                        current_answer = "F"
                else:
                    current_answer = text.split('：')[-1].split(':')[-1].strip().upper()
                
                # 处理多选题答案，将中文逗号转换为英文逗号
                current_answer = current_answer.replace('，', ',')
                
                # 如果已有题目和答案，保存题目
                if current_question and current_answer:
                    question_type = self.determine_question_type(current_options, current_answer)
                    self.questions.append({
                        'question': current_question,
                        'options': current_options.copy(),
                        'answer': current_answer,
                        'type': question_type
                    })
                    self.total_score += 1
                    current_question = None
                    current_options = []
                    current_answer = None
                continue

            # 检查是否是新题目
            if re.match(r'^[0-9一二三四五六七八九十]+[.、]', text):
                # 如果已有题目和答案，保存之前的题目
                if current_question and current_answer:
                    question_type = self.determine_question_type(current_options, current_answer)
                    self.questions.append({
                        'question': current_question,
                        'options': current_options.copy(),
                        'answer': current_answer,
                        'type': question_type
                    })
                    self.total_score += 1
                
                current_question = text
                current_options = []
                current_answer = None
            
            # 如果是选项（以A-Z开头）
            elif re.match(r'^[A-Z][.、]', text) or re.match(r'^[A-Z]\s', text):
                current_options.append(text)

        # 添加最后一个题目
        if current_question and current_answer:
            question_type = self.determine_question_type(current_options, current_answer)
            self.questions.append({
                'question': current_question,
                'options': current_options,
                'answer': current_answer,
                'type': question_type
            })
            self.total_score += 1

    def determine_question_type(self, options, answer):
        """根据题目格式判断题目类型"""
        # 首先检查是否是判断题
        if not options:
            return "判断题"
        elif len(options) == 2 and all(opt.endswith(('对', '错')) for opt in options):
            return "判断题"
            
        # 检查题目文本中是否包含题型标识
        question_text = self.questions[-1]['question'] if self.questions else ""
        if '多选题' in question_text or '(多选题' in question_text or '（多选题' in question_text:
            return "多选题"
        
        # 根据答案格式判断
        if ',' in answer or len(answer) > 1:
            return "多选题"
        
        # 默认为单选题
        return "单选题"

    def display_current_question(self):
        if self.is_review_mode:
            if self.current_question >= len(self.wrong_questions):
                return False
            question = self.wrong_questions[self.current_question]
            total = len(self.wrong_questions)
        else:
            if self.current_question >= len(self.questions):
                return False
            question = self.questions[self.current_question]
            total = len(self.questions)

        print("\n" + "="*50)
        print("错题重做模式" if self.is_review_mode else "正常答题模式")
        print(f"进度：第 {self.current_question + 1}/{total} 题")
        if not self.is_review_mode:
            print(f"当前正确率：{(self.score/max(1, self.current_question))*100:.1f}%" if self.current_question > 0 else "")
            print(f"已答对：{self.score} 题")
            print(f"已答错：{self.current_question - self.score} 题")
        print("="*50)
        print(question['question'])
        for option in question['options']:
            print(option)
        
        if question['type'] == "判断题":
            print("\n请输入你的答案(T/F):")
        elif question['type'] == "单选题":
            print("\n请输入你的答案(A/B/C/D):")
        else:  # 多选题
            print("\n请输入你的答案(多个选项用逗号分隔，如A,B,C):")
        return True

    def check_answer(self, user_answer):
        if self.is_review_mode:
            if self.current_question >= len(self.wrong_questions):
                return False
            question = self.wrong_questions[self.current_question]
        else:
            if self.current_question >= len(self.questions):
                return False
            question = self.questions[self.current_question]

        user_answer = user_answer.upper().strip()
        correct_answer = question['answer'].upper()

        # 根据题型进行不同的答案处理
        if question['type'] == "判断题":
            if user_answer == "对" or user_answer == "T":
                user_answer = "T"
            elif user_answer == "错" or user_answer == "F":
                user_answer = "F"
        elif question['type'] == "多选题":
            # 将答案转换为排序后的列表进行比较
            user_answers = sorted(user_answer.replace(" ", "").split(","))
            correct_answers = sorted(correct_answer.replace(" ", "").split(","))
            is_correct = user_answers == correct_answers
        else:  # 单选题
            is_correct = user_answer == correct_answer

        if is_correct:
            print("✓ 回答正确！")
            if not self.is_review_mode:
                self.score += 1
        else:
            print(f"✗ 回答错误。正确答案是：{correct_answer}")
            if not self.is_review_mode:
                self.wrong_questions.append(question)

        self.current_question += 1
        
        if not self.is_review_mode and self.current_question >= len(self.questions):
            print("\n" + "="*50)
            print("测验完成！最终统计：")
            print(f"总题数：{self.total_score}")
            print(f"答对题数：{self.score}")
            print(f"答错题数：{self.total_score - self.score}")
            print(f"正确率：{(self.score/self.total_score)*100:.1f}%")
            if len(self.wrong_questions) > 0:
                print(f"\n你有 {len(self.wrong_questions)} 道题答错了，是否要重做错题？(Y/N)")
                if input().upper().strip() == 'Y':
                    self.start_review_mode()
            print("="*50)
        elif self.is_review_mode and self.current_question >= len(self.wrong_questions):
            print("\n" + "="*50)
            print("错题重做完成！")
            print("="*50)
        
        return True

    def start_review_mode(self):
        """开始错题重做模式"""
        self.is_review_mode = True
        self.current_question = 0
        print("\n开始错题重做...")

def list_docx_files(folder_path):
    """列出指定文件夹中的所有Word文档"""
    docx_files = []
    try:
        for file in os.listdir(folder_path):
            if file.endswith('.docx'):
                docx_files.append(file)
    except Exception as e:
        print(f"读取文件夹失败：{e}")
        return []
    return docx_files

def main():
    while True:
        folder_path = input("请输入题库文件夹路径：").strip()
        if not os.path.exists(folder_path):
            print("文件夹不存在，请重新输入！")
            continue

        docx_files = list_docx_files(folder_path)
        if not docx_files:
            print("文件夹中没有Word文档，请检查路径！")
            continue

        print("\n可用的题库文件：")
        for i, file in enumerate(docx_files, 1):
            print(f"{i}. {file}")

        while True:
            try:
                choice = int(input("\n请选择题库编号（输入数字）："))
                if 1 <= choice <= len(docx_files):
                    selected_file = docx_files[choice - 1]
                    docx_path = os.path.join(folder_path, selected_file)
                    break
                else:
                    print("无效的编号，请重新选择！")
            except ValueError:
                print("请输入有效的数字！")
            except Exception as e:
                print(f"发生错误：{e}")

        try:
            print(f"\n正在加载题库：{selected_file}")
            quiz = QuizReader(docx_path)
            break
        except Exception as e:
            print(f"打开文件失败：{e}")
            print("请检查文件是否为有效的Word文档格式\n")

    # 选择答题模式
    while True:
        print("\n请选择答题模式：")
        print("1. 正常答题")
        if quiz.wrong_questions:
            print("2. 错题重做")
            valid_choices = ['1', '2']
        else:
            valid_choices = ['1']
            
        mode_choice = input("请输入模式编号：").strip()
        if mode_choice in valid_choices:
            break
        print("无效的选择，请重新输入！")

    # 根据选择进入不同模式
    if mode_choice == '2':
        quiz.start_review_mode()
    
    # 开始测验
    print("\n开始测验！")
    while quiz.display_current_question():
        user_answer = input().strip()
        quiz.check_answer(user_answer)

    # 询问是否继续
    while True:
        print("\n是否继续答题？")
        print("1. 继续答题")
        print("2. 退出程序")
        choice = input("请选择（1/2）：").strip()
        if choice == '1':
            main()  # 重新开始
            break
        elif choice == '2':
            print("\n感谢使用！再见！")
            break
        else:
            print("无效的选择，请重新输入！")

if __name__ == "__main__":
    main()
