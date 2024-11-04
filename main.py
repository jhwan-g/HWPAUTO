import tkinter as tk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import time
from hwpio import Hwp
from constants import *
import json
import os


class RuleModify:
    '''프로그램 규칙 수정 화면 클래스입니다

    화면 구성 & 수정까지 진행
    '''
    def __init__(self, window):
        self.window = window
        self.window.title('HwpAuto-규칙 수정')
        self.window.geometry('300x500')
        self.first_showed = True
        self.rule_path = BASE_DATA_PATH + RULE_FILE_NAME

        self.__get_rule()
        self.__show_window()
        self.__auto_save()

    # view
    def __show_rules(self):
        self.__get_rule()
        # self.label_rule_texts = list()
        # self.button_rule_delete = list()
        # self.checkbutton_rule_apply = list()
        # self.frame_button_and_box = list()
        for rule in self.rule['eqn_edit'].keys():
            label = tk.Label(self.frame_rule_describe, text=self.rule_texts[rule])
            frame = tk.Frame(self.frame_rule_describe, width=100)
            checkbutton = tk.Checkbutton(frame, text='적용', variable=self.rule_apply_stat[rule])
            button = tk.Button(frame, text='삭제', command=self.__delete_rule(rule))
            
            # self.label_rule_texts.append(label)
            # self.button_rule_delete.append(button)
            # self.checkbutton_rule_apply.append(checkbutton)
            # self.frame_button_and_box.append(frame)

            label.pack()
            frame.pack()
            checkbutton.pack(side='left')
            button.pack(side='right')

    def __show_window(self):
        if not self.first_showed:
            self.frame_rule_describe.destroy()
            self.frame_rule_describe = tk.Frame(self.window, width = 100)
            self.frame_rule_describe.place(x=150, y=0)

        if self.first_showed:
            self.frame_addrule = tk.Frame(self.window, width=100, height=100)
            self.new_rule_var_init = tk.StringVar()
            self.new_rule_var_fin = tk.StringVar()

            self.label = tk.Label(self.frame_addrule, text='원래 문자열')
            self.entry_new_rule = tk.Entry(self.frame_addrule, textvariable=self.new_rule_var_init)
            self.label2 = tk.Label(self.frame_addrule, text='바꿀 문자열')
            self.entry_new_rule2 = tk.Entry(self.frame_addrule, textvariable=self.new_rule_var_fin)
            self.button_add = tk.Button(self.frame_addrule, text="추가", command=self.__add_rule)

            self.frame_rule_describe = tk.Frame(self.window, width=100)

            self.label.grid()
            self.entry_new_rule.grid()
            self.label2.grid()
            self.entry_new_rule2.grid()
            self.button_add.grid()
            self.frame_addrule.place(x=0,y=0)

            self.frame_rule_describe.place(x=150, y=0)
            self.first_showed = False

        self.__show_rules()

    # controller(rule json 파일 관련)
    def __add_rule(self):
        '''규칙을 추가한다
        
        Args:
            None
        
        Returns:
            None
        
        Note:
            entry(self.entry_new_rule, self.entry_new_rule2)에 입력한 값을 기반으로 새 규칙을 추가한다.
            self.rule 수정 -> 저장 -> 저장된 것 load -> window 업데이트 순으로 진행
        '''
        new_rule_var_init = self.new_rule_var_init.get()
        self.new_rule_var_init.set('')
        new_rule_var_fin = self.new_rule_var_fin.get()
        self.new_rule_var_fin.set('')
        rule_name = f'eqn_rule_{new_rule_var_init}_to_{new_rule_var_fin}'
        self.rule_apply_stat[rule_name] = tk.BooleanVar(value=True)
        self.rule['eqn_edit'][rule_name] = {
            'init_str' : new_rule_var_init,
            'fin_str' : new_rule_var_fin,
            'apply' : True}
        self.__save_rule()
        self.__get_rule()
        self.__show_window()

    def __get_rule(self):
        '''규칙들을 불러온다

        Args:
            None
        
        Returns:
            None
        
        Note:
            rule관련 정보 저장하는 인스턴스 변수를 현재 json 파일 상태 기반으로 맞춤
        '''
        try:
            with open(self.rule_path, mode='r') as rule_data:
                self.rule = json.load(rule_data)
        except:
            if not os.path.exists(BASE_DATA_PATH):
                os.makedirs(BASE_DATA_PATH)
            with open(self.rule_path, 'w') as rule_data:
                json.dump(EMPTY_RULE_JSON, rule_data, indent=2)
            with open(self.rule_path, mode='r') as rule_data:
                self.rule = json.load(rule_data)

        self.rule_apply_stat = {}
        self.rule_texts = {}
        for rulename in self.rule['eqn_edit'].keys():
            self.rule_apply_stat[rulename] = tk.BooleanVar(value = self.rule['eqn_edit'][rulename]['apply'])
            self.rule_texts[rulename] = (f'"{self.rule["eqn_edit"][rulename]["init_str"]}"->"{self.rule["eqn_edit"][rulename]["fin_str"]}"')

    def __delete_rule(self, rule_name):
        '''rule_name에 해당하는 규칙을 없애는 함수를 리턴
        '''
        def delete_rule_func():
            del self.rule['eqn_edit'][rule_name]
            self.__save_rule()
            self.__show_window()
        return delete_rule_func
    
    def __save_rule(self):
        '''규칙을 저장한다
        
        Args:
            None
        
        Returns:
            None
        
        Note:
            현재 rule 상태와 동일하게 json 파일 수정
            규칙은 self.rule 기반, 실행 여부는 체크박스 기반    
        '''
        for i in self.rule['eqn_edit'].keys():
            self.rule['eqn_edit'][i]['apply'] = self.rule_apply_stat[i].get()
        with open(self.rule_path, mode = 'w') as rule_data:
            json.dump(self.rule, rule_data, indent=2)
    
    def __auto_save(self):
        '''자동저장 METHOD(1초 마다)
        '''
        self.__save_rule()
        self.window.after(1000, self.__auto_save)

            

class HwpAuto:
    '''프로그램 메인 화면 클래스입니다
    '''
    def __init__(self, window):
        '''프로그램 화면 및 필수 구성 요소의 구성을 진행합니다.

        Args:
            window (tkinter window): 화면이 구성될 윈도우 객체

        Returns:
            None
        '''
        self.window = window
        self.window.title('HwpAuto')
        self.window.geometry('300x900')

        # using vars
        self.rule_path = BASE_DATA_PATH + RULE_FILE_NAME
        self.files = list()
        self.saveasvar = tk.IntVar()
        self.statusstr = tk.StringVar()
        self.statusstr.set('')
        
        # make main label & buttons
        self.label1 = tk.Label(self.window, text = '원하는 파일(.hwp)를 선택해 주세요')
        self.button1 = tk.Button(self.window, text = '파일 추가', font = 12, command = self.file_pressed)
        
        self.button1.drop_target_register(DND_FILES)
        self.button1.dnd_bind('<<Drop>>', lambda e: self.add_file_path([e.data[1:-1]]))

        self.button_editrule = tk.Button(self.window, text='규칙 수정', font = 12, command = self.show_rule_modify)
        self.saveas_check = tk.Checkbutton(self.window, text="다른 이름으로 저장", variable=self.saveasvar)
        self.label2 = tk.Label(self.window, text = '선택된 파일 리스트')
        self.selected_file_label = tk.Label(self.window, text = ' ')

        # self.selected_file_label.drop_target_register(DND_FILES)
        # self.selected_file_label.bind('<<Drop>>', lambda e:print(e.data))
        self.button2 = tk.Button(self.window, text = '실행', font = 12, command = self.run)
        self.status_label = tk.Label(self.window, textvariable = self.statusstr, width=300)
        
        # pack
        self.label1.grid(columnspan=2, ipadx=40)
        self.button1.grid(column = 0, row = 1)
        self.button_editrule.grid(column= 1, row = 1)
        self.saveas_check.grid(columnspan=2)
        self.selected_file_label.grid(columnspan=2)
        self.button2.grid(columnspan=2)


        self.selected_file_label.configure(text = self.selected_file_text())

    # controller
    def file_pressed(self):
        '''파일 찾기 버튼을 눌렀을때 실행

        Args:
            None

        Return:
            None
        '''
        self.add_file_path(filedialog.askopenfilenames())
        # self.files.append(filedialog.askopenfilename(title='파일 선택창'))
        
    
    def add_file_path(self, path):
        '''파일 리스트에 path 추가하고 리스트 refresh

        Args:
            path(list) : 파일 경로 문자열의 리스트

        Return:
            None
        '''
        self.files += list(path)
        self.selected_file_label.configure(text = self.selected_file_text())

    def selected_file_text(self):
        ret = ''
        text_per_row_num = 25
        for i in self.files:
            for j in range((len(i)-1)//text_per_row_num + 1):
                ret += i[j*text_per_row_num : j*text_per_row_num + text_per_row_num] + '\n'
            ret += '\n\n'
            #ret += str(sum(str(i[j*100:j*100+100]) + '\n' for j in range((len(i)-1)//100+1))) + '\n\n'
        return ret
    
    def run(self):
        '''지정한 규칙에 따라 한글 파일의 수정 진행

        Args:
            None
        
        Returns:
            None
        '''
        for now_file in self.files:
            hwp = Hwp(now_file)
            time.sleep(1)

            rule = None
            try:
                with open(self.rule_path, mode='r') as rule_data:
                    rule = json.load(rule_data)
                    eqn_edit_tuple = tuple((rule['eqn_edit'][i]['init_str'], rule['eqn_edit'][i]['fin_str']) \
                                            for i in rule['eqn_edit'].keys() if rule['eqn_edit'][i]['apply'])
                    hwp.eqn_edit(eqn_edit_tuple)
            except:
                pass

            hwp.save(self.saveasvar.get())
            self.files.pop(0)
            self.selected_file_label.configure(text = self.selected_file_text())

            del hwp

        self.statusstr.set('')
    

    # connection to other window
    def show_rule_modify(self):
        rule_modify_window = tk.Toplevel()
        rulemodify = RuleModify(rule_modify_window)
        del rulemodify


    
if __name__ == '__main__':
    window_root = TkinterDnD.Tk()
    HwpAuto(window_root)
    window_root.mainloop()