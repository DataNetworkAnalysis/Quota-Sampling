import pandas as pd # 데이터프레임
import numpy as np # 행렬처리
from tkinter import filedialog
from tkinter import messagebox
import tkinter as tk
import tkinter.ttk as ttk
from winreg import *
import os


def central_box(root):
    # Gets the requested values of the height and widht.
    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()

    # Gets both half the screen width/height and window width/height
    positionRight = int(root.winfo_screenwidth() / 2 - windowWidth / 2)
    positionDown = int(root.winfo_screenheight() / 2 - windowHeight / 2)

    # Positions the window in the center of the page.
    root.geometry("+{}+{}".format(positionRight, positionDown))

    return root

def make_quota(filename, levels, level, num, grouping, condition_name, filtering=None):
    # Load Data
    df = pd.read_excel("{}".format(filename), sheet_name=1)
    if filtering==None:
        df = df[df.구분=='광역시도']
    elif filtering[0]=='세종특별자치시':
        df = df[df.구분==level]
    df = df.drop(levels, axis=1)

    # 필터링
    if filtering!=None:
        if (filtering[1]=='구 지역') | (filtering[0]=='세종특별자치시'):
            df = df[df.광역시도==filtering[0]].sum()
            df = pd.DataFrame(df).T
            df['광역시도'] = filtering[0]
            df['시군구'] = filtering[1]
        else:
            df = df[(df.광역시도==filtering[0])&(df.시군구==filtering[1])]
        df = df.rename(columns={'광역시도': '전체'})
        df = df.drop('구분', axis=1)

    # 그룹화
    if grouping:
        # 그룹화 리스트
        gp_name_lst = [('경기도','인천광역시'),('대전광역시','세종특별자치시','충청남도','충청북도'),
                       ('광주광역시','전라남도','전라북도'),('대구광역시','경상북도'),
                       ('부산광역시','울산광역시','경상남도'),('강원도','제주특별자치도')]
        # 그룹화할 이름을 반복문으로 처리
        for gp_names in gp_name_lst:
            # 충청남도와 세종특별자치시만 추출후 합계
            new = df[df.광역시도.isin(gp_names)].sum(axis=0)
            # 충남/세종 합계 새로운 행으로 추가
            df = df.append(new, ignore_index=True)
            # 이름 변경
            df.iloc[-1,0] = '광역시도'
            df.iloc[-1,1] = '/'.join(gp_names) # /으로 지역들을 묶어줌
            # 충남, 세종 제외
            df = df[~df.광역시도.isin(gp_names)]
    elif (level=='광역시도') and (filtering==None):
        # 그룹화할 이름
        gp_names =['충청남도','세종특별자치시']
        # 충청남도와 세종특별자치시만 추출후 합계
        new = df[df.광역시도.isin(gp_names)].sum(axis=0)
        # 충남/세종 합계 새로운 행으로 추가
        df = df.append(new, ignore_index=True)
        # 이름 변경
        df.iloc[-1,0] = '광역시도'
        df.iloc[-1,1] = '충청남도/세종특별자치시'
        # 충남, 세종 제외
        df = df[~df.광역시도.isin(gp_names)]

    # Define Features
    male_cols = ['남 19-29세', '남 30대', '남 40대', '남 50대', '남 60세 이상']
    female_cols = ['여 19-29세', '여 30대', '여 40대', '여 50대', '여 60세이상']
    total_cols = male_cols + female_cols

    # Total Population
    try:
        total_pop = df[total_cols].sum().sum()
    except:
        messagebox.showerror("메세지 박스","해당 파일의 기준 변수명이 다릅니다.")
        exit()

    # 2단계 반올림 전
    before_df = df.copy()
    before_df[total_cols] = (df[total_cols] / total_pop) * num # 각 셀값을 전체 인구로 나누고 정해진 값으로 곱ㅎ
    before_df['남 합계'] = before_df[male_cols].sum(axis=1)
    before_df['여 합계'] = before_df[female_cols].sum(axis=1)
    before_df['총계'] = before_df[['남 합계' ,'여 합계']].sum(axis=1)
    # 2단계 남여 각각 합계의 반올림
    before_sex_sum = before_df[['남 합계' ,'여 합계']].sum().round()

    # 3단계 반올림 후
    after_df = df.copy()
    after_df[total_cols] = (df[total_cols] / total_pop) * num # 각 셀값을 전체 인구로 나누고 정해진 값으로 곱ㅎ
    after_df[total_cols] = after_df[total_cols].astype(float).round().astype(int) # 각 셀을 반올림
    after_df['남 합계'] = after_df[male_cols].sum(axis=1)
    after_df['여 합계'] = after_df[female_cols].sum(axis=1)
    after_df['총계'] = after_df[['남 합계' ,'여 합계']].sum(axis=1)
    # 3단계 남여 각각 합계의 반올림
    after_sex_sum = after_df[['남 합계' ,'여 합계']].sum()

    # 2,3단계 남여 합계의 차이
    '''
    차이는 세 가지 경우로 나뉜다: 남여 각각 차이가 1. 0이거나 / 2. 0보다 크거나 / 3. 0보다 작거나 
    1. 0인 경우는 추가적인 일 없이 표 완성
    2. 만약 차이가 0보다 큰 경우 : xx.5 보다 작고 xx.5에 가장 가까운 값인 반올림 값에 + 1 
    - Why? 반올림하여 내림이 된 값 중 가장 올림에 가까운 값에 1을 더하는 것이 이상적이기 때문
        ex) 2.49999 -> round(2.49999) + 1
    3. 만약 차이가 0보다 작은 경우 : xx.5 이상이고 xx.5에 가장 가까운 값인 반올림 값에 - 1
    - Why? 반올림하여 올림이 된 값 중 가장 내림에 가까운 값에 1을 빼는 것이 이상적이기 때문
        ex) 2.50001 -> round(2.50001) - 1
    '''
    sex_diff = before_sex_sum - after_sex_sum

    # 성별 합계 차이를 매꾸는 단계
    sex_cols_lst = [male_cols, female_cols]
    sex_idx = ['남 합계' ,'여 합계']
    for i in range(len(sex_idx)):
        if sex_diff.loc[sex_idx[i]] > 0:
            # 차이가 0보다 큰 경우
            '''
            1. 2단계 반올림 전 값을 모두 내림 한 후 0.5를 더 한다.
            2. 1번에서 한 값과 2단계 반올림 전 값을 뺀다.
            3. 음수로 나오는 값은 모두 1로 변환. 1이 가장 큰 값이기 때문.
            ex) 13.45 -> 13 으로 내림 후 (13 + 0.5) - 13.45 = 0.05
            '''
            temp = (before_df[sex_cols_lst[i]].astype(int) + 0.5) - before_df[sex_cols_lst[i]] # 1,2번
            temp = temp[temp >0].fillna(1) # 3번
            v = 1
        elif sex_diff.loc[sex_idx[i]] < 0:
            # 차이가 0보다 작은 경우
            '''
            1. 2단계 반올림 전 값을 모두 내림 한 후 0.5를 더 한다.
            2. 1번에서 한 값과 2단계 반올림 전 값을 뺀다.
            3. 음수로 나오는 값은 모두 1로 변환. 1이 가장 큰 값이기 때문.
            ex) 13.54 -> 13 으로 내림 후 13.54 - (13 + 0.5) = 0.04
            '''
            temp = before_df[sex_cols_lst[i]] - (before_df[sex_cols_lst[i]].astype(int) + 0.5) # 1,2번
            temp = temp[temp >0].fillna(1) # 3번에 해당
            v = -1
        else:
            # 차이가 0인 경우는 이후 과정 생략하고 그냥 통과
            continue

        # 실제합계와의 차이: 절대값을 통해서 음수를 변환하고 정수로 타입을 변환
        cnt = int(abs(sex_diff.loc[sex_idx[i]]))
        row_col = np.unravel_index(np.argsort(temp.values.ravel())[:cnt], temp.shape)
        rows = row_col[0]
        cols = row_col[1]
        # 각 (행,열) 좌표값에 v를 더함
        for r in range(len(rows)):
            temp = after_df[sex_cols_lst[i]].copy()
            temp.iloc[rows[r] ,cols[r]] = temp.iloc[rows[r] ,cols[r]] + v
            after_df[sex_cols_lst[i]] = temp
        print()


    # 부족한 부분이 채워졌으면 합계 계산
    after_df['남 합계'] = after_df[male_cols].sum(axis=1)
    after_df['여 합계'] = after_df[female_cols].sum(axis=1)
    after_df['총계'] = after_df[['남 합계' ,'여 합계']].sum(axis=1)
    final_sex_sum = after_df[['남 합계' ,'여 합계']].sum()

    # 다운로드 폴더 경로 찾기
    with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
        Downloads = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]

    # 완료 메세지
    if final_sex_sum.sum() != num:
        messagebox.showerror("메세지 상자","합계가 0이 아닙니다. 문제를 확인해주세요.")
    else:
        save_name = filename.split('/')[-1]
        file_path = '{}/{}{}'.format(Downloads, condition_name, save_name)
        if filtering==None:
            messagebox.showinfo("메세지 상자", "다운로드 폴더에 저장되었습니다.")
            after_df.to_excel(file_path, index=False, encoding='cp949')
        else:
            if os.path.isfile(file_path):
                saved_df = pd.read_excel(file_path)
                after_df = pd.concat([saved_df,after_df], axis=0)
            after_df.to_excel(file_path, index=False, encoding='cp949')

if __name__=='__main__':
    # 변경할 파일 이름을 입력받기 위한 코드
    upload = tk.Tk()
    upload = central_box(upload)

    # 파일 선택하기
    upload.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                 filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
    filename = upload.filename

    # 업로드 종료
    upload.destroy()

    starting=True

    # 계속 실행 반복
    while starting:
        # 버튼 만들기
        def start_click():
            # 전역 변수 설정하기
            global starting
            starting = True
            start_check.destroy()

        def quit_click():
            # 전역 변수 설정하기
            global starting
            starting = False
            start_check.destroy()

        # start_check 생성
        start_check = tk.Tk()
        start_check = central_box(start_check)

        # 시작 여부
        btn1 = ttk.Button(start_check, text='시작', command=start_click)
        btn2 = ttk.Button(start_check, text='종료', command=quit_click)
        btn1.pack()
        btn2.pack()

        start_check.mainloop()

        if not(starting):
            exit()

        def simple():
            # 전역 변수 설정하기
            global work
            work = True
            work_check.destroy()

        def multiplt():
            # 전역 변수 설정하기
            global work
            work = False
            work_check.destroy()

        # work_check 생성
        work_check = tk.Tk()
        work_check = central_box(work_check)

        # 작업 종류
        btn3 = ttk.Button(work_check, text='단순작업', command=simple)
        btn4 = ttk.Button(work_check, text='쿼터표 불러오기', command=multiplt)
        btn3.pack()
        btn4.pack()

        work_check.mainloop()

        if work:
            # 구분 / 크기를 입력받기 위한 코드
            root = tk.Tk()
            root = central_box(root)

            # 구분과 크기 입력 창 만들기
            ttk.Label(root, text="구분과 크기를 정하세요").grid(column=0, row=0)

            # 구분 리스트
            levels = ['광역시도', '시군구', '읍면동']
            level_lst_Chosen = ttk.Combobox(root, width=12, values=levels)
            level_lst_Chosen.grid(column=0, row=1)

            # 크기 정하기
            var = tk.IntVar().set(1000)
            cnt = ttk.Entry(root, width=10, text=var)
            cnt.grid(column=1, row=1)

            # 버튼 만들기
            def level_num_click():
                # 전역 변수 설정하기
                global level
                global num

                level = level_lst_Chosen.get()
                if not (isinstance(int(cnt.get()), int)):
                    messagebox.showerror("메세지박스", "크기는 숫자만 입력하세요.")
                    exit()
                num = int(cnt.get())
                root.destroy()


            action = ttk.Button(root, text="시작", command=level_num_click)
            action.grid(column=2, row=1)

            # 윈도우 창 실행하기
            root.mainloop()

            # 그룹화 확인 버튼
            if level == '광역시도':
                # root 생성
                gp_check = tk.Tk()
                gp_check = central_box(gp_check)
                # 그룹화 확인 문구
                ttk.Label(gp_check, text="그룹화 여부").grid(column=0, row=0)

                # 그룹화 여부
                radVar = tk.BooleanVar()
                r1 = ttk.Radiobutton(gp_check, text="Yes", variable=radVar, value=True)
                r1.grid(column=0, row=1)
                r2 = ttk.Radiobutton(gp_check, text="No", variable=radVar, value=False)
                r2.grid(column=1, row=1)

                # 버튼 만들기
                def gp_click():
                    # 전역 변수 설정하기
                    global grouping
                    grouping = radVar.get()
                    gp_check.destroy()

                # 버튼 생성
                action = ttk.Button(gp_check, text="시작", command=gp_click)
                action.grid(column=2, row=1)

                gp_check.mainloop()
            else:
                grouping = False

            # 저장할 파일명
            if grouping:
                group_name = '그룹'
            else:
                group_name = ''
            condition_name = '[{}_{}_{}]'.format(level, num, grouping)

            # levels에서 선택된 level은 제외
            levels.remove(level)

            # 실행
            make_quota(filename, levels, level, num, grouping, condition_name)
        else:
            # 필터링할 파일 이름을 입력받기 위한 코드
            filtering_upload = tk.Tk()
            filtering_upload = central_box(filtering_upload)

            # 파일 선택하기
            filtering_upload.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                         filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
            filtering_filename = filtering_upload.filename

            # 업로드 종료
            filtering_upload.destroy()

            filtering_df = pd.read_excel("{}".format(filtering_filename))
            for col in ['전체','시군구','쿼터 합계']:
                if col not in filtering_df:
                    print('컬럼명이 잘못되었습니다.')
                    print("['전체','시군구','쿼터합계']로 입력해주세요")
                    exit()

            # 초기값
            levels = ['읍면동']
            level = '시군구'
            grouping = False
            # 저장할 파일명
            condition_name = '[{}개_쿼터표]'.format(filtering_df.shape[0])


            # 진행정도
            process = tk.Tk()
            process = central_box(process)
            progress = ttk.Progressbar(process, orient='horizontal', length=286, mode='determinate')
            progress.pack()
            progress['maximum'] = filtering_df.shape[0]
            progress['value'] = 0

            # 실행
            for i in range(filtering_df.shape[0]):
                gwang = filtering_df.iloc[i,0]
                sigoongoo = filtering_df.iloc[i,1]
                num = filtering_df.iloc[i,2]
                filtering = [gwang, sigoongoo]

                make_quota(filename, levels, level, num, grouping, condition_name, filtering)
                progress['value'] += 1
                progress.update()

            progress.mainloop()
            messagebox.showinfo("메세지 상자", "다운로드 폴더에 저장되었습니다.")





