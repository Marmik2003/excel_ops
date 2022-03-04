import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


class GetExcel:
    def __init__(self, data):
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.ws.title = 'Main'
        self.data = data
        self.topics_sheet = self.wb.create_sheet(title='Topics')

    def _create_topics_sheet(self):
        max_subtopic_length = 0
        tl = len(self.data)
        for topic_obj in self.data:
            subtopic_list = [i['name'] for i in topic_obj['subtopics']]
            if len(subtopic_list) > max_subtopic_length:
                max_subtopic_length = len(subtopic_list)
            self.topics_sheet.append([topic_obj['name']] + subtopic_list)
        max_subtopic_length += 2
        self.max_subtopic_length = max_subtopic_length
        self.col_letter = get_column_letter(max_subtopic_length)

        col_letter2 = get_column_letter(max_subtopic_length - 1)
        for i in range(1, 5000):
            self.topics_sheet.cell(row=i, column=max_subtopic_length).value = '=INDEX($B1:$' + col_letter2 + str(tl)\
                                                                              + ';MATCH(\'Main\'!$A' + str(i+1) + \
                                                                              ',$A1:$A' + str(tl) + '))'
        self.topics_sheet.sheet_state = 'hidden'
        self.topics_sheet[get_column_letter(max_subtopic_length*2) + '1'] = ' '

    def create(self):
        self._create_topics_sheet()
        self.ws.append(
            ['Topic', 'Subtopic', 'Question', 'Type of Answer', 'Number of options'] +
            ['Option {}'.format(i) for i in range(1, 9)] +
            ['Answer',  'Explanation', 'Difficulty', 'Marks', 'Negative marks']
        )
        topic_dv = DataValidation(
            type="list", formula1='\'Topics\'!$A$1:$A$' + str(len(self.data) + 1), allow_blank=False
        )
        self.ws.add_data_validation(topic_dv)
        topic_dv.add('A2:A5000')
        for i in range(2, 5001):
            subtopic_dv = DataValidation(
                type="list",
                formula1=f'\'Topics\'!${self.col_letter}${i-1}:${get_column_letter(self.max_subtopic_length*2)}${i-1}',
                allow_blank=False
            )
            self.ws.add_data_validation(subtopic_dv)
            subtopic_dv.add(f'B{i}:B{i}')

        anstype_dv = DataValidation(
            type="list",
            formula1='"Single, Multiple"',
            allow_blank=False
        )
        self.ws.add_data_validation(anstype_dv)
        anstype_dv.add('D2:D5000')

        numopts_dv = DataValidation(
            type="whole",
            formula1='1:8',
            allow_blank=False
        )
        self.ws.add_data_validation(numopts_dv)
        numopts_dv.add('E2:E5000')

        difficulty_dv = DataValidation(
            type="list",
            formula1='"Easy", "Medium", "Hard"',
            allow_blank=False
        )
        self.ws.add_data_validation(difficulty_dv)
        difficulty_dv.add('P2:P5000')

        return self.wb
