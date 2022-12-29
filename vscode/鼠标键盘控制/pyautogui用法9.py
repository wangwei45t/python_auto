import time
import cv2
import pyautogui
import xlrd
import selenium
from selenium import webdriver

print("开始执行")

def dataCheck (sheet1):
    checkCmd=True
    #行数检查
    if sheet1.nrows<2:
        print("没数据啊哥")
        checkCmd=False
    #每行数据检查
    i=1
    while i<sheet1.nrows:
        #第1列操作类型检查
        cmdType=sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
            and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print ('第',i+1,"行,第1列数据有毛病")
            checkCmd=False
        #第2列内容检查
        cmdValue=sheet1.row(i)[1]
        #读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype!=1:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd=False
        #输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype==0:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd=False
        #等待类型，内容必须为数字
        if cmdType.value==5.0:
            if cmdValue.ctype!=2:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd=False
        #滚轮事件，内容必须为数字
        if cmdType.value==6.0:
            if cmdValue.ctype!=2:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd=False
        i+=1
    return checkCmd

def get_xy (name_1,mode_1):
    if  mode_1 <= 4 :
        aaa="D:\dist\\"+name_1        
        #将屏幕截图保存
        pyautogui.screenshot(r"D:\dist\pingmujietu.png")
        #载入截图
        img=cv2.imread(r"D:\dist\pingmujietu.png")
        #图像模板
        img_terminal  =cv2.imread(aaa)
        #读取模板的宽度和高度
        height,width,channel=img_terminal.shape
        #进行模板匹配
        result=cv2.matchTemplate(img,img_terminal,cv2.TM_SQDIFF_NORMED)
        #解析出匹配区域的左上角的坐标(resule的第三个值是左上角坐标)
        upper_left=cv2.minMaxLoc(result)[2]
        #计算匹配区域右下角坐标
        lower_right=(upper_left[0]+width,upper_left[1]+height)
        #计算中心区域的坐标并返回
        avg=(int((upper_left[0]+lower_right[0])/2),int((upper_left[1]+lower_right[1])/2))
        return avg
        
    if mode_1==5:  
        print('开始等待')
    if mode_1==6:  
        print('移动滑轮')
    
def auto_click_left_one(var_avg):
    '''左键单击'''
    pyautogui.click(var_avg[0],var_avg[1],button='left',clicks=1)
    time.sleep(1)

def auto_click_left_too(var_avg):
    '''左键双击'''
    pyautogui.click(var_avg[0],var_avg[1],button='left',clicks=2)
    time.sleep(1)

def auto_click_right_one(var_avg):
    '''右键单击'''
    pyautogui.click(var_avg[0],var_avg[1],button='right',clicks=1)
    time.sleep(1)

def auto_click_right_too(var_avg):
    '''右键双击'''
    pyautogui.click(var_avg[0],var_avg[1],button='right',clicks=2)
    time.sleep(1)

def process_mode(node_date):
    '''处理数据'''
    node_date=str(node_date)
    node_date=node_date[7:]
    mode=float(node_date)
    return mode

def process_name(name_date,node_dete_1):
    if node_dete_1<=4:
        '''点击---返回图片名称'''
        name_date=str(name_date)
        name=name_date[6:-1]
        return name
        
    if node_dete_1==5 :
        '''等待---返回浮点数'''
        name_date=str(name_date)
        name_date=name_date[7:]
        name=float(name_date)
        return name

    if node_dete_1== 6 :
        '''滑轮---返回整数'''
        name_date=str(name_date)
        name_date=name_date[7:]
        name=float(name_date)
        name=int(name)
        return name

    if node_dete_1== 7 :
        '''键盘输入---返回字符串'''
        name_date=str(name_date)
        name=name_date[6:-1]
        return name

    if node_dete_1== 8 :
        '''粘贴输入---返回字符串'''
        name_date=str(name_date)
        name=name_date[6:-1]
        return name

    if node_dete_1== 9 :
        '''键盘按键---返回字符串'''
        name_date=str(name_date)
        name=name_date[6:-1]
        return name



def process_reTry(reTry_date):
    '''处理重复次数'''
    reTry_date=str(reTry_date)
    reTry_date=reTry_date[7:]
    reTry=float(reTry_date)
    return reTry

def process_tips(tips_dete):
    '''获取提示信息'''
    tips_dete=str(tips_dete)
    tips=tips_dete[6:-1]
    return tips

def mainWork(sheet):
    '''1-左键单击2左键双击3-右键单击4右键双击,重复次数,图片位置,名字'''
    #i代表xls表中的行数
    i=1
    while i < sheet1.nrows:
        '''获取xls表格中的数据'''
        mode1=sheet1.row(i)[0]
        name1=sheet1.row(i)[1]
        reTry1=sheet1.row(i)[2]
        tips1=sheet1.row(i)[3]
        '''处理xls表格中的数据'''
        mode=process_mode(mode1)
        name=process_name(name1,mode)
        reTry=process_reTry(reTry1)
        tips=process_tips(tips1)
        '''获得数据所在屏幕中的位置'''
        avg = get_xy(name,mode)
        '''根据重复次数,执行对应次数'''
        if reTry>=1.0:
            j=1
            while j < reTry+1:
                if mode==1:
                    print(f"正在左键单击:{name}",tips)
                    auto_click_left_one(avg)
                if mode==2:
                    print(f"正在左键双击:{name}",tips)
                    auto_click_left_too(avg)
                if mode==3:
                    print(f"正在右键单击:{name}",tips)
                    auto_click_right_one(avg)
                if mode==4:
                    print(f"正在右键双击:{name}",tips)
                    auto_click_right_too(avg)
                if mode==5:
                    print(f"正在等待:{name}秒后结束,请稍后",tips)
                    time.sleep(name)
                if mode==6:
                    pyautogui.scroll(name)
                    print(f"滚轮滑动",name,"距离",tips)     
                if mode==7:
                    pyautogui.typewrite(name)
                    print(f"按键输入",name,"字符",tips)  
                if mode==8: 
                    pyautogui.hotkey('ctrl','v')
                    print(f"粘贴输入",name,"字符",tips)   
                if mode==9: 
                    pyautogui.press(name)
                    print(f"键盘按键",name,"字符",tips)                                                                                
                j+=1
                time.sleep(0.1)
        i+=1

if __name__== '__main__' :
    #打开文件
    file='D:\dist\cmd.xls'
    wb=xlrd.open_workbook(filename=file)
    #获取xls文件中第一个sheet
    sheet1 = wb.sheet_by_index(0)
    #数据检查
    #checkCmd=dataCheck(sheet1)

    mainWork(sheet1)
  
    