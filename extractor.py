import fitz
import pandas as pd
import os
import matplotlib.pyplot as plt


class ExtractSchoolProfile:
    def __init__(self, file_path):
        self.doc = fitz.open(file_path)
        self.contents = self._extract_contents()
        self.grade = grade_info
        self.directory = name_info+'/'
        self.name = name_info
        try:
            os.mkdir(name_info)
        except:
            kj = 1

    def _extract_contents(self):
        contents = []
        for page in self.doc:
            text = list(map(str, page.get_text().split("\n")))
            for con in text:
                if '풍산고등학교' not in con and '년  ' not in con:
                    contents.append(con)
        return contents

    def extract_basic_info(self):
        profile = self.contents
        name_count = 0
        name = ''
        while name_count >= 0:
            if profile[name_count] == '성명':
                name = profile[name_count + 1]
                name_count = -1
            else:
                name_count += 1
        num_count = 0
        num = 0
        while num_count >= 0:
            if profile[num_count] == '번호' and profile[num_count + 1] in list(map(str, range(0, 51))):
                num = profile[num_count + 1]
                num_count = -1
            else:
                num_count += 1
        class_count = 0
        class_num = 0
        while class_count >= 0:
            if profile[class_count] == '반' and profile[class_count + 1] in list(map(str, range(0, 11))):
                class_num = profile[class_count + 1]
                class_count = -1
            else:
                class_count += 1
        return (self.grade, int(class_num), int(num), name)




    def extract_details(self):
        while '/' in self.contents:
            slash = self.contents.index('/')
            for i in range(10):
                del self.contents[slash - 1]
        sections = ['창 의 적 체 험 활 동 상 황', '봉 사 활 동 실 적', '6. 교과학습발달상황', '7. 독서활동상황', '행 동 특 성 및 종 합 의 견']
        extracts = []
        for i in range(len(sections) - 1):
            start = self.contents.index(sections[i])
            end = self.contents.index(sections[i + 1])
            extracts.append(self.contents[start:end])
        extracts.append(self.contents[self.contents.index(sections[-1]):])
        return extracts

    def extract_changche(self):
        changche = self.extract_details()[0]
        ret_jayul = []
        ret_club = []
        ret_jinro = []
        for i in range(self.grade):
            start_jayul = changche.index('자율활동')
            end_jayul = changche.index('동아리활동')
            jayul_list = changche[start_jayul:end_jayul]
            changche = changche[end_jayul:]
            ret_jayul.append(jayul_list)

            start_club = changche.index('동아리활동')
            end_club = changche.index('진로활동')
            club_list = changche[start_club:end_club]
            ret_club.append(club_list)

            start_jinro = changche.index('진로활동')

            if '자율활동' in changche:
                end_jinro = changche.index('자율활동')
                jinro_list = changche[start_jinro:end_jinro]
                changche = changche[end_jinro:]
                ret_jinro.append(jinro_list)
            else:
                jinro_list = changche[start_jinro:]
                ret_jinro.append(jinro_list)
        ret = [ret_jinro, ret_club, ret_jayul]
        return(ret)

    def extract_jinro(self):
        ret = []
        all_jinro = self.extract_changche()[0]

        for i in range(len(all_jinro)):
            if all_jinro[i][-1] in list(map(str, range(0, 4))):
                del all_jinro[i][-1]
            while len(all_jinro[i][0]) < 8 and len(all_jinro[i]) != 0:
                del all_jinro[i][0]
                all_jinro[i].append('          ')


            while '창 의 적 체 험 활 동 상 황' in all_jinro[i]:
                slice = all_jinro[i].index('창 의 적 체 험 활 동 상 황')
                for k in range(7):
                    del all_jinro[i][slice-1]
                if len(all_jinro[i][slice-1]) <= 4:
                    del all_jinro[i][slice - 1]

            con = ''
            for j in all_jinro[i]:
                if j != '          ':
                    con += j
            ret.append(con)
        return(ret)

    def extract_club(self):
        ret = []
        all_club = self.extract_changche()[1]

        for i in range(len(all_club)):

            if all_club[i][-1] in list(map(str, range(0, 4))):
                del all_club[i][-1]

            while len(all_club[i][0]) < 8 and len(all_club[i]) != 0:
                del all_club[i][0]
                all_club[i].append('          ')


            while '창 의 적 체 험 활 동 상 황' in all_club[i]:
                slice = all_club[i].index('창 의 적 체 험 활 동 상 황')
                for k in range(7):
                    del all_club[i][slice-1]
                try:
                    if len(all_club[i][slice-1]) <= 4:
                        del all_club[i][slice - 1]
                except:
                    p=1

            con = ''
            for j in all_club[i]:
                if j != '          ':
                    con += j
            ret.append(con)
        return (ret)


    def extract_jayul(self):
        ret = []
        all_jayul = self.extract_changche()[2]

        for i in range(len(all_jayul)):
            if all_jayul[i][-1] in list(map(str, range(0, 4))):
                del all_jayul[i][-1]

            while len(all_jayul[i][0]) < 8 and len(all_jayul[i]) != 0:
                del all_jayul[i][0]
                all_jayul[i].append('          ')


            while '창 의 적 체 험 활 동 상 황' in all_jayul[i]:
                slice = all_jayul[i].index('창 의 적 체 험 활 동 상 황')
                for k in range(7):
                    del all_jayul[i][slice-1]
                    all_jayul[i].append('          ')
                if len(all_jayul[i][slice-1]) <= 4:
                    del all_jayul[i][slice - 1]
                    all_jayul[i].append('          ')

            con = ''
            for j in all_jayul[i]:
                if j != '          ':
                    con += j
            ret.append(con)
        return (ret)

    def extract_behave(self):
        behave = self.extract_details()[4]
        ret1 = ['']
        ret2 = ['']
        ret3 = ['']

        con1 = ''
        con2 = ''
        con3 = ''

        while '행 동 특 성 및 종 합 의 견' in behave:
            pos = behave.index('행 동 특 성 및 종 합 의 견')
            del behave[pos]
        while '학년' in behave:
            pos = behave.index('학년')
            del behave[pos]
        if '2' in behave:
            if '3' in behave:
                ret1 = behave[behave.index('1')+1:behave.index('2')]
                ret2 = behave[behave.index('2')+1:behave.index('3')]
                ret3 = behave[behave.index('3')+1:]
            else:
                ret1 = behave[behave.index('1') + 1:behave.index('2')]
                ret2 = behave[behave.index('2') + 1:]
        else:
            ret1 = behave[behave.index('1') + 1:]

        for i in ret1:
            con1 += i

        for j in ret2:
            con2 += j

        for k in ret3:
            con3 += k

        ret = [con1, con2, con3]
        return(ret)


    def extract_seteuk(self):

        all_seteuk = self.extract_details()[2]
        print(all_seteuk)
        lecture_list = ['국어', '통합사회', '한국사', '수학', '통합과학', '과학탐구실험', '영어',
                        '화법과 작문', '독서', '언어와 매체', '문학',
                        '실용 국어', '심화 국어', '고전 읽기',
                        '수학Ⅰ', '수학Ⅱ', '미적분', '확률과 통계',
                        '기본 수학', '실용 수학', '인공지능 수학', '기하', '경제 수학', '수학과제 탐구',
                        '영어Ⅰ', '영어Ⅱ', '영어 회화', '영어 독해와 작문',
                        '기본 영어', '실용 영어', '영어권 문화', '진로 영어', '영미 문학 읽기',
                        '한국지리', '세계지리', '세계사', '동아시아사', '경제', '정치와 법', '사회·문화', '생활과 윤리', '윤리와 사상',
                        '여행지리', '사회문제 탐구', '고전과 윤리',
                        '물리학Ⅰ', '화학Ⅰ', '생명과학Ⅰ', '지구과학Ⅰ',
                        '물리학Ⅱ', '화학Ⅱ', '생명과학Ⅱ', '지구과학Ⅱ', '과학사', '생활과 과학', '융합과학',
                        '체육', '운동과 건강',
                        '스포츠 생활', '체육 탐구',
                        '음악', '미술','연극',
                        '음악 연주', '음악 감상과 비평', '미술 창작', '미술 감상과 비평',
                        '기술·가정', '정보',
                        '농업 생명 과학', '공학 일반', '창의 경영', '해양 문화와 기술', '가정과학', '지식 재산 일반', '인공지능 기초',
                        '독일어Ⅰ', '프랑스어Ⅰ', '스페인어Ⅰ', '중국어Ⅰ', '일본어Ⅰ', '러시아어Ⅰ', '아랍어Ⅰ', '베트남어Ⅰ',
                        '독일어Ⅱ', '프랑스어Ⅱ', '스페인어Ⅱ', '중국어Ⅱ', '일본어Ⅱ', '러시아어Ⅱ', '아랍어Ⅱ', '베트남어Ⅱ',
                        '한문Ⅰ', '한문Ⅱ',
                        '철학', '논리학', '심리학', '교육학', '종교학', '진로와 직업', '보건', '환경', '실용 경제', '논술',
                        '심화 수학Ⅰ', '심화 수학Ⅱ', '고급 수학Ⅰ', '고급 수학Ⅱ', '고급 물리학', '고급 화학', '고급 생명과학', '고급 지구과학',
                        '물리학 실험', '화학 실험', '생명과학 실험', '지구과학 실험', '정보과학', '융합과학 탐구', '과학과제 연구', '생태와 환경',
                        '심화 영어 회화Ⅰ', '심화 영어 회화Ⅱ', '심화 영어Ⅰ', '심화 영어Ⅱ', '심화 영어 독해Ⅰ', '심화 영어 독해Ⅱ', '심화 영어 작문Ⅰ', '심화 영어 작문Ⅱ',
                        '국제 정치', '국제 경제', '국제법','지역 이해', '한국 사회의 이해', '비교 문화', '세계 문제와 미래 사회', '국제 관계와 국제기구', '현대 세계의 변화', '사회 탐구 방법', '사회과제 연구',
                        '프로그래밍'
                        ]
        lecture_list_in = list(map(str, [i +':' for i in lecture_list])) + list(map(str, [i + ' :' for i in lecture_list]))


        seteuk_table = []

        for i in range(len(all_seteuk)):
            if '세 부 능 력 및 특 기 사 항' in all_seteuk[i]:
                seteuk_table.append(i)
                all_seteuk[i] = '세특 구간'


        written_grade = self.grade
        if self.grade != 1 and all_seteuk[-1] == '해당 학년의 자료가 없습니다':
            written_grade -= 1

        seteuk_ilban_all = []
        seteuk_jinro_all = []
        seteuk_yeche_all = []


        ret_ilban = []
        ret_jinro = []
        ret_yeche = []

        ret = []

        print(written_grade)

        for i in range(written_grade):

            next_grade = '['+str(i+2)+'학년]'





            start_ilban = all_seteuk.index('세특 구간')
            end_ilban = all_seteuk.index('<진로 선택 과목>')

            ilban_seteuk = all_seteuk[start_ilban: end_ilban]
            all_seteuk = all_seteuk[end_ilban:]

            seteuk_ilban_all.append(ilban_seteuk)

            while '세특 구간' in ilban_seteuk:
                i = ilban_seteuk.index('세특 구간')
                del ilban_seteuk[i]

            start_jinro = all_seteuk.index('세특 구간')
            end_jinro = all_seteuk.index('<체육ㆍ예술>')

            jinro_seteuk = all_seteuk[start_jinro: end_jinro]

            while '세특 구간' in jinro_seteuk:
                i = jinro_seteuk.index('세특 구간')
                del jinro_seteuk[i]


            all_seteuk = all_seteuk[end_jinro:]

            seteuk_jinro_all.append(jinro_seteuk)

            if next_grade in all_seteuk:
                start_yeche = all_seteuk.index('세특 구간')
                end_yeche = all_seteuk.index(next_grade)

                yeche_seteuk = all_seteuk[start_yeche: end_yeche]
                all_seteuk = all_seteuk[end_yeche:]

                while '세특 구간' in yeche_seteuk:
                    i = yeche_seteuk.index('세특 구간')
                    del yeche_seteuk[i]

                seteuk_yeche_all.append(yeche_seteuk)
            else:

                start_yeche = all_seteuk.index('세특 구간')

                yeche_seteuk = all_seteuk[start_yeche:]

                while '세특 구간' in yeche_seteuk:
                    i = yeche_seteuk.index('세특 구간')
                    del yeche_seteuk[i]

            seteuk_yeche_all.append(yeche_seteuk)

        for i in seteuk_ilban_all:
            print(i)

            grade_all_ilban = []

            index_list = []
            for j in i:
                for k in lecture_list_in:
                    if k in j:
                        index_list.append(i.index(j))



            index_list.append(-1)


            for j in range(len(index_list)-1):
                contents = ''
                start = index_list[j]
                end =  index_list[j+1]
                for k in i[start:end]:
                    contents += k
                if end == -1:
                    contents += i[-1]
                grade_all_ilban.append(contents)
            ret_ilban.append(grade_all_ilban)

        for i in seteuk_yeche_all:

            grade_all_yeche = []

            index_list = []
            for j in i:
                for k in lecture_list_in:
                    if k in j:
                        index_list.append(i.index(j))



            index_list.append(-1)



            for j in range(len(index_list)-1):
                contents = ''
                start = index_list[j]
                end =  index_list[j+1]
                for k in i[start:end]:
                    contents += k
                if end == -1:
                    contents += i[-1]
                grade_all_yeche.append(contents)
            ret_yeche.append(grade_all_yeche)

        for i in seteuk_jinro_all:

            grade_all_jinro = []

            index_list = []
            for j in i:
                for k in lecture_list_in:
                    if k in j:
                        index_list.append(i.index(j))


            index_list.append(-1)


            for j in range(len(index_list)-1):
                contents = ''
                start = index_list[j]
                end =  index_list[j+1]
                for k in i[start:end]:
                    contents += k
                if end == -1:
                    contents += i[-1]
                grade_all_jinro.append(contents)
            ret_jinro.append(grade_all_jinro)







        print(seteuk_ilban_all)
        print(seteuk_jinro_all)
        print(seteuk_yeche_all)
        '''print(ret_ilban)
        print(ret_yeche)
        print(ret_jinro)'''

        for j in range(len(ret_ilban)-2):
            if ret_ilban[j] == ret_ilban[j+1]:
                del ret_ilban[j]
                ret_ilban.append('error')
        for j in range(len(ret_jinro)-2):
            if ret_jinro[j] == ret_jinro[j+1]:
                del ret_jinro[j]
                ret_jinro.append('error')
        for j in range(len(ret_yeche)-2):
            if ret_jinro[j] == ret_yeche[j+1]:
                del ret_yeche[j]
                ret_yeche.append('error')
        while "erorr" in ret_ilban:
            i = ret_ilban.index('error')
            del ret_ilban[i]
        while "erorr" in ret_jinro:
            i = ret_jinro.index('error')
            del ret_jinro[i]
        while "erorr" in ret_ilban:
            i = ret_jinro.index('error')
            del ret_jinro[i]




        ret = [ret_ilban, ret_jinro, ret_yeche] # 일반, 진로, 예체능


        return ret

    def extract_naesin(self):

        all_seteuk = self.extract_details()[2]


        lecture_list = ['국어', '통합사회', '한국사', '수학', '통합과학', '과학탐구실험', '영어',
                        '화법과 작문', '독서', '언어와 매체', '문학',
                        '실용 국어', '심화 국어', '고전 읽기',
                        '수학Ⅰ', '수학Ⅱ', '미적분', '확률과 통계',
                        '기본 수학', '실용 수학', '인공지능 수학', '기하', '경제 수학', '수학과제 탐구',
                        '영어Ⅰ', '영어Ⅱ', '영어 회화', '영어 독해와 작문',
                        '기본 영어', '실용 영어', '영어권 문화', '진로 영어', '영미 문학 읽기',
                        '한국지리', '세계지리', '세계사', '동아시아사', '경제', '정치와 법', '사회·문화', '생활과 윤리', '윤리와 사상',
                        '여행지리', '사회문제 탐구', '고전과 윤리',
                        '물리학Ⅰ', '화학Ⅰ', '생명과학Ⅰ', '지구과학Ⅰ',
                        '물리학Ⅱ', '화학Ⅱ', '생명과학Ⅱ', '지구과학Ⅱ', '과학사', '생활과 과학', '융합과학',
                        '체육', '운동과 건강',
                        '스포츠 생활', '체육 탐구',
                        '음악', '미술', '연극',
                        '음악 연주', '음악 감상과 비평', '미술 창작', '미술 감상과 비평',
                        '기술·가정', '정보',
                        '농업 생명 과학', '공학 일반', '창의 경영', '해양 문화와 기술', '가정과학', '지식 재산 일반', '인공지능 기초',
                        '독일어Ⅰ', '프랑스어Ⅰ', '스페인어Ⅰ', '중국어Ⅰ', '일본어Ⅰ', '러시아어Ⅰ', '아랍어Ⅰ', '베트남어Ⅰ',
                        '독일어Ⅱ', '프랑스어Ⅱ', '스페인어Ⅱ', '중국어Ⅱ', '일본어Ⅱ', '러시아어Ⅱ', '아랍어Ⅱ', '베트남어Ⅱ',
                        '한문Ⅰ', '한문Ⅱ',
                        '철학', '논리학', '심리학', '교육학', '종교학', '진로와 직업', '보건', '환경', '실용 경제', '논술',
                        '심화 수학Ⅰ', '심화 수학Ⅱ', '고급 수학Ⅰ', '고급 수학Ⅱ', '고급 물리학', '고급 화학', '고급 생명과학', '고급 지구과학',
                        '물리학 실험', '화학 실험', '생명과학 실험', '지구과학 실험', '정보과학', '융합과학 탐구', '과학과제 연구', '생태와 환경',
                        '심화 영어 회화Ⅰ', '심화 영어 회화Ⅱ', '심화 영어Ⅰ', '심화 영어Ⅱ', '심화 영어 독해Ⅰ', '심화 영어 독해Ⅱ', '심화 영어 작문Ⅰ',
                        '심화 영어 작문Ⅱ',
                        '국제 정치', '국제 경제', '국제법', '지역 이해', '한국 사회의 이해', '비교 문화', '세계 문제와 미래 사회', '국제 관계와 국제기구',
                        '현대 세계의 변화', '사회 탐구 방법', '사회과제 연구',
                        '프로그래밍'
                        ]
        lecture_list_in = list(map(str, [i + ' :' for i in lecture_list]))

        seteuk_table = []

        for i in range(len(all_seteuk)):
            if '세 부 능 력 및 특 기 사 항' in all_seteuk[i]:
                seteuk_table.append(i)
                all_seteuk[i] = '세특 구간'

        written_grade = self.grade
        if self.grade != 1 and all_seteuk[-1] == '해당 학년의 자료가 없습니다':
            written_grade -= 1

        seteuk_ilban_all = []
        seteuk_jinro_all = []
        seteuk_yeche_all = []

        ret_ilban = []
        ret_jinro = []
        ret_yeche = []

        ret = []

        for i in range(written_grade):

            next_grade = '[' + str(i + 2) + '학년]'

            start_ilban = all_seteuk.index('세특 구간')
            end_ilban = all_seteuk.index('<진로 선택 과목>')

            ilban_seteuk = all_seteuk[start_ilban: end_ilban]
            all_seteuk = all_seteuk[end_ilban:]

            seteuk_ilban_all.append(ilban_seteuk)

            while '세특 구간' in ilban_seteuk:
                i = ilban_seteuk.index('세특 구간')
                del ilban_seteuk[i]

            start_jinro = all_seteuk.index('세특 구간')
            end_jinro = all_seteuk.index('<체육ㆍ예술>')

            jinro_seteuk = all_seteuk[start_jinro: end_jinro]

            while '세특 구간' in jinro_seteuk:
                i = jinro_seteuk.index('세특 구간')
                del jinro_seteuk[i]

            all_seteuk = all_seteuk[end_jinro:]

            seteuk_jinro_all.append(jinro_seteuk)

            if next_grade in all_seteuk:
                start_yeche = all_seteuk.index('세특 구간')
                end_yeche = all_seteuk.index(next_grade)

                yeche_seteuk = all_seteuk[start_yeche: end_yeche]
                all_seteuk = all_seteuk[end_yeche:]

                while '세특 구간' in yeche_seteuk:
                    i = yeche_seteuk.index('세특 구간')
                    del yeche_seteuk[i]

                seteuk_yeche_all.append(yeche_seteuk)
            else:

                start_yeche = all_seteuk.index('세특 구간')

                yeche_seteuk = all_seteuk[start_yeche:]

                while '세특 구간' in yeche_seteuk:
                    i = yeche_seteuk.index('세특 구간')
                    del yeche_seteuk[i]

            seteuk_yeche_all.append(yeche_seteuk)


        st_list = [seteuk_ilban_all, seteuk_jinro_all,seteuk_yeche_all ]
        st_contents = []#세특 내용 다 적혀있음

        for i in st_list:
            for j in i:
                for k in j:
                    st_contents.append(k)


        all_naesin = [i for i in self.extract_details()[2] if i not in st_contents and i != '해당 학년의 자료가 없습니다' and '세 부 능 력 및 특 기 사 항' not in i and i not in ['6. 교과학습발달상황', '학기', '교과', '과목', '단위수', '원점수/과목평균', '(표준편차)', '성취도', '(수강자수)', '석차등급', '비고','단위수원점수/과목평균', '성취도별', '분포비율', '이수단위 합계'] ]

        print(len(all_naesin))


        for i in range(len(all_naesin)-1):
            if all_naesin[i] in '기술가정/제2외국어/한문/교양' and all_naesin[i+1] in '기술가정/제2외국어/한문/교양' and all_naesin[i] != '국어' and all_naesin[i] !='2':
                all_naesin[i] = '기술가정/제2외국어/한문/교양'
                all_naesin[i+1]='null_value'

        for i in range(len(all_naesin)-1):
            if all_naesin[i] in '사회(역사/도덕포함)' and all_naesin[i+1] in '사회(역사/도덕포함)' and all_naesin[i] != '사회':
                all_naesin[i] = '사회(역사/도덕포함)'
                all_naesin[i+1]='null_value'

        all_naesin_new = [i for i in all_naesin if i != 'null_value']
        all_naesin = all_naesin_new



        naesin_list = []
        if self.grade ==1:
            naesin_1 = all_naesin[all_naesin.index('[1학년]'):]
            naesin_list = [naesin_1]
        if self.grade ==2:
            naesin_1 = all_naesin[all_naesin.index('[1학년]'):all_naesin.index('[2학년]')]
            naesin_2 = all_naesin[all_naesin.index('[2학년]'):]
            naesin_list = [naesin_1, naesin_2]
        if self.grade == 3:
            naesin_1 = all_naesin[all_naesin.index('[1학년]'):all_naesin.index('[2학년]')]
            naesin_2 = all_naesin[all_naesin.index('[2학년]'):all_naesin.index('[3학년]')]
            naesin_3 = all_naesin[all_naesin.index('[3학년]'):]
            naesin_list = [naesin_1, naesin_2, naesin_3]

        naesin_ilban_list = []

        for i in naesin_list:
            if '<진로 선택 과목>' in i:
                k = i.index('<진로 선택 과목>')
                j = i[:k]
                naesin_ilban_list.append(j)
            else:
                naesin_ilban_list.append(i)

        return naesin_list



    def save_jajindong(self):
        jayul = self.extract_jayul()
        jinro = self.extract_jinro()
        club = self.extract_club()

        df_jayul = pd.DataFrame({
            'grade':range(len(jayul)+1)[1:],
            'contents':jayul
        })

        df_jinro = pd.DataFrame({
            'grade': range(len(jinro) + 1)[1:],
            'contents': jinro
        })

        df_club = pd.DataFrame({
            'grade': range(len(club) + 1)[1:],
            'contents': club
        })

        df_jayul.to_excel(self.directory+'jayul.xlsx')
        df_jinro.to_excel(self.directory+'jinro.xlsx')
        df_club.to_excel(self.directory+'club.xlsx')

    def save_seteuk(self):
        seteuk = self.extract_seteuk()
        ilban_seteuk = seteuk[0]
        jinro_seteuk = seteuk[1]
        yeche_seteuk = seteuk[2]

        print(yeche_seteuk)

        grade_list = []
        seteuk_list = []

        print(ilban_seteuk)
        g = 1


        for i in range(len(ilban_seteuk)):
            print(len(jinro_seteuk))
            list1 = []
            for k in ilban_seteuk[i]:
                list1.append(k)
            for k in jinro_seteuk[i]:
                list1.append(k)
            for k in yeche_seteuk[i]:
                list1.append(k)



            for j in range(len(list1)):

                grade_list.append(g)
                seteuk_list.append(list1[j])
            g+= 1

        df_seteuk = pd.DataFrame({
            'grade':grade_list,
            'contents':seteuk_list
        })

        df_seteuk.to_excel(self.directory+'seteuk.xlsx')

    def save_behave(self):
        behave = self.extract_behave()

        df_behave = pd.DataFrame({
            'grade': range(len(behave) + 1)[1:],
            'contents': behave
        })

        df_behave.to_excel(self.directory + 'behave.xlsx')
    def save_naesin(self):
        lecture_list = ['국어', '통합사회', '한국사', '수학', '통합과학', '영어', '화법과 작문', '독서', '언어와 매체', '문학','수학Ⅰ', '수학Ⅱ', '미적분', '확률과 통계',
                        '영어Ⅰ', '영어Ⅱ','한국지리', '세계지리', '세계사', '동아시아사', '경제', '정치와 법', '사회·문화', '생활과 윤리', '윤리와 사상','물리학Ⅰ', '화학Ⅰ',
                        '생명과학Ⅰ', '지구과학Ⅰ','정보','독일어Ⅰ', '프랑스어Ⅰ', '스페인어Ⅰ', '중국어Ⅰ', '일본어Ⅰ', '러시아어Ⅰ', '아랍어Ⅰ', '베트남어Ⅰ','한문Ⅰ', '영어 회화'
                        ]

        class_list = ['국어', '영어', '수학', '과학', '사회', '사회(역사/도덕포함)', '기술가정/제2외국어/한문/교양']
        grade_list = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '·']
        real_grade_list = ['1', '2', '3', '4', '5', '6', '7', '8', '9']
        naesin = self.extract_naesin()
        print(naesin)
        #학기 구분
        list_naesin = []#학기별로 구분

        for i in naesin:
            for j in range(len(i)-2):
                if i[j] in grade_list and i[j+1] == '2' and i[j+2] in class_list:
                    li1 = i[:j+1]
                    li2 = i[j+1:]
                    list_naesin.append(li1)
                    list_naesin.append(li2)
                    break



        print(list_naesin)

        naesin_list = []


        for i in list_naesin:#학기별로 가공
            lecture_df = []
            time_df = []
            grade_df = []
            a = len(i)
            i.append([123, 234, 345, 456, 567, 678, 789])
            for j in range(a-1):
                if i[j]==i[j+1]:
                    del i[j]
                    i.append('null')
            for k in range(a):
                if i[k] in lecture_list and i[k+1] in grade_list and i[k+4] in real_grade_list:
                    lecture = i[k]
                    time = i[k+1]
                    grade = i[k+4]

                    lecture_df.append(lecture)
                    time_df.append(time)
                    grade_df.append(grade)
            naesin_list.append([lecture_df, time_df, grade_df])
        print(naesin_list)
        all_lecture_list = []
        all_time_list = []
        all_grade_list = []
        all_semester_list = []
        s = 1
        for i in naesin_list:
            print(i[0])
            all_lecture_list = all_lecture_list + i[0]
            all_time_list = all_time_list + i[1]
            all_grade_list = all_grade_list + i[2]
            for j in range(len(i[0])):
                all_semester_list.append(s)
            s+=1

        df_grade = pd.DataFrame({
            'sem':all_semester_list,
            'lecture':all_lecture_list,
            'time':all_time_list,
            'grade':all_grade_list
        })

        df_grade.to_excel(self.directory+'grade.xlsx')

    def graph_naesin(self):
        korengmath = ['국어', '수학', '영어', '화법과 작문', '독서', '언어와 매체', '문학', '수학Ⅰ', '수학Ⅱ', '미적분', '확률과 통계', '영어Ⅰ', '영어Ⅱ', '영어 회화',
                      'korengmath']
        science = ['국어', '수학', '통합과학', '영어', '화법과 작문', '독서', '언어와 매체', '문학', '수학Ⅰ', '수학Ⅱ', '미적분', '확률과 통계', '영어Ⅰ', '영어 회화',
                   '영어Ⅱ', '물리학Ⅰ', '화학Ⅰ', '생명과학Ⅰ', '지구과학Ⅰ', 'science']
        social = ['국어', '통합사회', '한국사', '수학', '영어', '화법과 작문', '독서', '언어와 매체', '문학', '수학Ⅰ', '수학Ⅱ', '미적분', '확률과 통계', '영어 회화',
                  '영어Ⅰ', '영어Ⅱ', '한국지리', '세계지리', '세계사', '동아시아사', '경제', '정치와 법', '사회·문화', '생활과 윤리', '윤리와 사상', 'social']

        all_1, all_2, all_3, all_4 = 0, 0, 0, 0
        all_time_1, all_time_2, all_time_3, all_time_4 = 0, 0, 0, 0
        all_naesin = 0
        all_time = 0

        path = self.directory

        grade_df = pd.read_excel(path + 'grade.xlsx')

        for i in range(len(grade_df['sem'])):
            all_time += grade_df['time'][i]
            all_naesin += grade_df['time'][i] * grade_df['grade'][i]
        all_naesin = all_naesin / all_time

        for i in range(len(grade_df['sem'])):
            if grade_df['sem'][i] == 1:
                all_time_1 += grade_df['time'][i]
                all_1 += grade_df['time'][i] * grade_df['grade'][i]
        if all_time_1 == 0:
            all_time_1 += 1
            all_1 = 9
        all_1 = all_1 / all_time_1

        for i in range(len(grade_df['sem'])):
            if grade_df['sem'][i] == 2:
                all_time_2 += grade_df['time'][i]
                all_2 += grade_df['time'][i] * grade_df['grade'][i]
        if all_time_2 == 0:
            all_time_2 += 1
            all_2 = 9
        all_2 = all_2 / all_time_2

        for i in range(len(grade_df['sem'])):
            if grade_df['sem'][i] == 3:
                all_time_3 += grade_df['time'][i]
                all_3 += grade_df['time'][i] * grade_df['grade'][i]
        if all_time_3 == 0:
            all_time_3 += 1
            all_3 = 9
        all_3 = all_3 / all_time_3

        for i in range(len(grade_df['sem'])):
            if grade_df['sem'][i] == 4:
                all_time_4 += grade_df['time'][i]
                all_4 += grade_df['time'][i] * grade_df['grade'][i]
        if all_time_4 == 0:
            all_time_4 += 1
            all_4 = 9
        all_4 = all_4 / all_time_4

        fig, ax = plt.subplots()

        ax.plot(['1-1', '1-2', '2-1', '2-2'], [all_1, all_2, all_3, all_4])
        ax.plot(['1-1', '2-2'], [all_naesin, all_naesin])
        ax.scatter(['1-1', '1-2', '2-1', '2-2'], [all_1, all_2, all_3, all_4])
        ax.text('1-1', 7, str(all_1)[:4])
        ax.text('1-2', 7, str(all_2)[:4])
        ax.text('2-1', 7, str(all_3)[:4])
        ax.text('2-2', 7, str(all_4)[:4])
        ax.text('1-1', 8, 'mean: ' + str(all_naesin)[:4])
        ax.set_ylim(9, 1)
        plt.savefig(path + 'all_naesin.png')

        for x in (korengmath, science, social):

            all_1, all_2, all_3, all_4 = 0, 0, 0, 0
            all_time_1, all_time_2, all_time_3, all_time_4 = 0, 0, 0, 0
            all_naesin = 0
            all_time = 0

            for i in range(len(grade_df['sem'])):
                if grade_df['lecture'][i] in x:
                    all_time += grade_df['time'][i]
                    all_naesin += grade_df['time'][i] * grade_df['grade'][i]
            all_naesin = all_naesin / all_time

            for i in range(len(grade_df['sem'])):
                if grade_df['sem'][i] == 1 and grade_df['lecture'][i] in x:
                    all_time_1 += grade_df['time'][i]
                    all_1 += grade_df['time'][i] * grade_df['grade'][i]
            if all_time_1 == 0:
                all_time_1 += 1
                all_1 = 9
            all_1 = all_1 / all_time_1

            for i in range(len(grade_df['sem'])):
                if grade_df['sem'][i] == 2 and grade_df['lecture'][i] in x:
                    all_time_2 += grade_df['time'][i]
                    all_2 += grade_df['time'][i] * grade_df['grade'][i]
            if all_time_2 == 0:
                all_time_2 += 1
                all_2 = 9
            all_2 = all_2 / all_time_2

            for i in range(len(grade_df['sem'])):
                if grade_df['sem'][i] == 3 and grade_df['lecture'][i] in x:
                    all_time_3 += grade_df['time'][i]
                    all_3 += grade_df['time'][i] * grade_df['grade'][i]
            if all_time_3 == 0:
                all_time_3 += 1
                all_3 = 9
            all_3 = all_3 / all_time_3

            for i in range(len(grade_df['sem'])):
                if grade_df['sem'][i] == 4 and grade_df['lecture'][i] in x:
                    all_time_4 += grade_df['time'][i]
                    all_4 += grade_df['time'][i] * grade_df['grade'][i]
            if all_time_4 == 0:
                all_time_4 += 1
                all_4 = 9
            all_4 = all_4 / all_time_4

            fig, ax = plt.subplots()

            ax.plot(['1-1', '1-2', '2-1', '2-2'], [all_1, all_2, all_3, all_4])
            ax.plot(['1-1', '2-2'], [all_naesin, all_naesin])
            ax.scatter(['1-1', '1-2', '2-1', '2-2'], [all_1, all_2, all_3, all_4])
            ax.text('1-1', 7, str(all_1)[:4])
            ax.text('1-2', 7, str(all_2)[:4])
            ax.text('2-1', 7, str(all_3)[:4])
            ax.text('2-2', 7, str(all_4)[:4])
            ax.text('1-1', 8, 'mean: ' + str(all_naesin)[:4])
            ax.set_ylim(9, 1)
            plt.savefig(path + x[-1] + '_naesin.png')

    def name_save(self):
        name_df = pd.DataFrame({
            'name':[self.name],
            'grade':[self.grade]
        })

        name_df.to_excel(self.directory+'basic_info.xlsx')


file_path = ''

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def open_file_dialog():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", ".pdf")])
    if file_path:
        messagebox.showinfo("Selected File", f"File selected: {file_path}")




def start():
    global name_info
    global grade_info
    name_info = entry.get()
    grade_info = int(grade_combobox.get())

    # 예외 처리: 학년이 선택되지 않았을 경우
    if not grade_info:
        messagebox.showerror("오류", "학년을 선택해주세요.")
        return

    profile = ExtractSchoolProfile(file_path)
    profile.save_seteuk()
    profile.save_naesin()
    profile.save_jajindong()
    profile.graph_naesin()
    profile.name_save()
    profile.save_behave()
    messagebox.showinfo("추출 완료", "추출이 완료되었습니다.")


root = tk.Tk()
root.title("생활기록부 정보 추출기 made by 김현동")

label = tk.Label(root, text="하단에 학생 이름 입력")
label.pack(pady=5)

entry = tk.Entry(root, width=30)
entry.pack(pady=0)

label2 = tk.Label(root, text="학년 선택")
label2.pack(pady=5)

grade_combobox = ttk.Combobox(root, values=["1", "2", "3"], state="readonly")
grade_combobox.pack(pady=5)

file_button = tk.Button(root, text="파일 선택", command=open_file_dialog)
file_button.pack(pady=5)

text_button = tk.Button(root, text="추출 시작", command=start)
text_button.pack(pady=5)

root.mainloop()