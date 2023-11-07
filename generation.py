'''
按照名单、跑量表、训练表格式，指定输入输出文件名即可
同名同姓的张鑫手动标一下（航院张鑫无法识别）
加了一列不合格，不满足要求的为1，新人保护手动处理一下
未添加性别列，所以不合格按照小于40公里计算，可能会有错误
'''
import pandas as pd
import PySimpleGUI as sg

def generate(memberFile, trainingFile, distanceFile, month):
    # memberFile = '10月名单.xlsx'
    # distanceFile = '10月跑量.xlsx'
    # trainingFile = '10月例跑训练名单.xlsx'
    outputFile = '跑团'+str(month)+'月考核名单.xlsx'

    nana = float('nan')
    memberDF = pd.read_excel(memberFile)
    distanceDF = pd.read_excel(distanceFile)
    trainingDF = pd.read_excel(trainingFile)
    distance = []
    for memberName in memberDF['名字']:
        if '-' in memberName:
            memberNamePure = memberName.split('-')[1]
        else:
            memberNamePure = memberName
        for i in range(len(distanceDF)):
            disName = distanceDF['名字'][i]
            if disName.find(memberNamePure) != -1:
                distance.append(distanceDF['月里程（KM）'][i])
                break
            if i==len(distanceDF)-1:
                distance.append(nana)
    memberDF['跑量'] = distance
    for col in trainingDF:
        trainingPar = []
        for memberName in memberDF['名字']:
            if '-' in memberName:
                memberNamePure = memberName.split('-')[1]
            else:
                memberNamePure = memberName
            for i in range(len(trainingDF)):
                trainName = trainingDF[col][i]
                try:
                    if trainName.find(memberNamePure) != -1:
                        trainingPar.append(1)
                        break
                except:
                    trainingPar.append(nana)
                    break
                if i == len(trainingDF)-1:
                    trainingPar.append(nana)
        memberDF[col] = trainingPar

    unqualified = []
    for i in range(len(memberDF)):
        if (memberDF['性别'][i]=='男' and memberDF['跑量'][i] < 40)\
                or (memberDF['性别'][i]=='女' and memberDF['跑量'][i] < 30)\
                or pd.isna(memberDF['跑量'][i]):  # 跑量不达标
            if pd.isna(memberDF['特殊情况'][i]):  # 无特殊情况
                isTrain = False  # 看是否参加了例跑
                for col in memberDF:
                    index = memberDF[col][i]
                    if index == 1.0:
                        isTrain = True
                if not isTrain:  # 没有参加例跑
                    unqualified.append(1)  # 不合格标记
                    continue
        unqualified.append(nana)
    memberDF['不达标'] = unqualified
    print(memberDF)
    memberDF.to_excel(outputFile)

def main():
    # 选择主题
    sg.theme('DarkTeal12')

    sg.set_global_icon("RunnerClub.ico")

    layout = [
        [sg.Text('西交跑者俱乐部考核小程序', font=('微软雅黑', 18)),], \
        [sg.Text('全员名单路径：', font=('微软雅黑', 10), text_color='cyan'), sg.Text('', key='memberFile', size=(50, 1), font=('微软雅黑', 10), text_color='cyan')], \
        [sg.Text('例跑名单路径：', font=('微软雅黑', 10), text_color='cyan'), sg.Text('', key='trainingFile', size=(50, 1), font=('微软雅黑', 10), text_color='cyan')], \
        [sg.Text('跑量数据路径：', font=('微软雅黑', 10), text_color='cyan'), sg.Text('', key='distanceFile', size=(50, 1), font=('微软雅黑', 10), text_color='cyan')], \
        [sg.Output(size=(70, 10), font=('微软雅黑', 10))], \
        [sg.FilesBrowse('团员名单', key='member', target='memberFile', file_types=(("Excel表格", "*.xlsx"),)), \
         sg.FilesBrowse('例跑名单', key='training', target='trainingFile', file_types=(("Excel表格", "*.xlsx"),)), \
         sg.FilesBrowse('跑量表格', key='distance', target='distanceFile', file_types=(("Excel表格", "*.xlsx"),)), \
         sg.Text('月份：', font=('微软雅黑', 18)), \
         sg.Combo([i for i in range(1, 13)], default_value=1, key='month'), \
         sg.Button('转换'), \
         sg.Button('退出'),],
        ]
    # 创建窗口
    window = sg.Window("Renz的小工具系列", layout, font=("微软雅黑", 15), default_element_size=(50, 1))
    # 事件循环
    while True:
        # 窗口的读取，有两个返回值（1.事件；2.值）
        event, values = window.read()
        print(event, values)

        if event == '转换':
            if values['member']:
                if values['training']:
                    if values['distance']:
                        generate(values['member'], values['training'], values['distance'], values['month'])
                        print('------------------------------')
                        print('输出完成！')
                    else:
                        print('请输入跑量表格！')
                else:
                    print('请输入例跑名单！')
            else:
                print('请输入团员名单！')

        if event in (None, '退出'):
            break

    window.close()

if __name__ == '__main__':
    main()

