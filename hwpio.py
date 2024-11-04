'''한글 파일 관련 작업을 담당하는 모듈
'''

import win32com.client as win32

def replace_multiple_str(string, restr):
    '''하나의 문자열에 대해 여러개의 규칙을 입력받아 수정합니다.

    Args:
        string (str): 바꿀 문자열
        restr (tuple): 문자열 수정 규칙 ((바꿀 문자열, 바뀔 문자열), (), ()) 형식으로 입력

    Returns:
        수정된 문자열

    Note:
        앞쪽에 위치한 튜플에 해당하는 규칙을 먼저 실행
    '''
    for i in restr:
        string = string.replace(*i)
    return string

# model
class Hwp:
    '''한글 파일 하나에 해당하는 객체

    객체 생성시 자동으로 지정한 한글 파일이 열립니다.
    객체 제거시 자동으로 한글이 닫힙니다.

    '__init__' 에서 파일 경로 + 이름을 file_name에 string 형태로 받습니다
    '''
    def __init__(self, file_name):
        self.file_name = file_name
        self.open()
    
    def open(self):
        '''자신의 한글 파일을 엽니다

        Args:
            None

        Returns:
            None
        '''
        self.hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
        self.hwp.XHwpWindows.Item(0).Visible = True
        self.hwp.Open(self.file_name)
        self.hwp.XHwpWindows.Item(0).Visible = False

    def eqn_edit(self, replace_eqn):
        '''한글 파일의 수식을 탐색하면서 해당하는 문자열을 변경합니다.

        Args: 
            equations (tuple): ((바꿀 문자열, 바뀔 문자열), (), )
        
        Returns:
            None
        '''
        ctrl = self.hwp.HeadCtrl
        eq_num = 0
        while ctrl != None:
            if ctrl.CtrlID == "eqed":
                eq_num += 1
            ctrl = ctrl.Next
        
        ctrl = self.hwp.HeadCtrl
        cnt = 0
        while ctrl != None: # 끝까지 탐색을 마치면 ctrl이 None을 리턴함:
            nextctrl = ctrl.Next
            if ctrl.CtrlID == "eqed": # 현재 컨트롤이 "수식eqed"인 경우
                # if cnt % 50 == 0:
                #     self.statusstr.set(f'{x + 1}번째 문서 {cnt}/{eq_num} 번째 수식')
                cnt += 1

                position = ctrl.GetAnchorPos(0)
                position = position.Item("List"), position.Item("Para"), position.Item("Pos")
                self.hwp.SetPos(*position)
                self.hwp.FindCtrl()

                Act = self.hwp.CreateAction("EquationModify")
                Set = Act.CreateSet()
                Pset = Set.CreateItemSet("EqEdit", "EqEdit")
                Act.GetDefault(Pset)

                if type(Pset.Item("String"))==type('str'): # 수식 변경
                    equation = replace_multiple_str(Pset.Item("String"),replace_eqn)
                    Pset.SetItem("String", equation)
                    Act.Execute(Pset)
                self.hwp.UnSelectCtrl()
                del Act
                del Set 
                del Pset
                del position

            ctrl = nextctrl

    
    def save(self, saveas = False):
        '''파일의 저장을 진행합니다

        Args:
            saveas (int or bool): 다른 이름으로 저장 여부
        
        Returns:
            None
        
        Note:
            파일을 저장하는 부수효과가 존재합니다.
        '''
        print(saveas)
        if saveas:
            self.hwp.SaveAs(Path = self.hwp.Path.replace(".hwp", "(1).hwp"))    
        else:
            self.hwp.Save()
        

    def __del__(self):
        self.hwp.Clear(1)
        self.hwp.Application.Quit()

