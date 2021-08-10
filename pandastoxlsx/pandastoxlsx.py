import string
import re
import xlsxwriter


class PandasToXLSX:
    """Класс для экспортирования датафрейма в XLSX-файл."""

    group_col = None

    group_blank_rows = 3

    _colors = ['#006400',
               '#8B008B',
               '#8B0000',
               '#2F4F4F',
               '#000080',
               '#DC143C',
               '#800000',
               '#008080']

    # Конфигурация по умолчанию для эспорта в XLSX-файл
    _config = {'sheet': {'name': 'Manual control'},
               'table_cell': {'format': {'border': 1}},
               'table_header': {'format': {'bold': True,
                                           'text_wrap': True,
                                           'align': 'center',
                                           'valign': 'top',
                                           'fg_color': '#dddddd',
                                           'bottom': 8,
                                           'top': 8,
                                           'font_size': 12,
                                           'border': 1}},
               'group_header': {'format': {'bold': True,
                                           'border': 1,
                                           'align': 'center',
                                           'valign': 'vcenter',
                                           'fg_color': '#D7E4BC',
                                           'font_size': 12
                                           },
                                'merge': True},
               'columns': None
               }

    _group_name_rules = ["number", "text"]

    def __init__(self, df_long, xlsx_filename, group_column="group", config=None, **kwargs):
        self.df = df_long  # датафрейм с данными для экспорта
        self.xlsx_filename = xlsx_filename  # имя XLSX-файла
        self.group_col = group_column
        self.group_name_rule = kwargs.get("group_name_rule", self._group_name_rules[0])
        self.group_name_rule = self.group_name_rule \
            if self.group_name_rule in self._group_name_rules \
            else self._group_name_rules[0]
        # конфигурация для XLSX-файла
        self._config.update(config)
        # Инициализация переменных
        self.writer, self.workbook, self.worksheet = None, None, None
        self._formats, self._columns_formats = None, None
        # Инициализация XLSX-файла
        self.init_workbook()
        self.init_formats()

    def init_workbook(self):
        """Инициализация книги."""

        self.workbook = xlsxwriter.Workbook(self.xlsx_filename)  # объект книги
        self.worksheet = self.workbook.add_worksheet(self.config['sheet']['name'])  # объет листа
        self.worksheet.set_zoom(100)
        self.worksheet.set_tab_color('green')

    def init_formats(self):
        """Инициализация форматов книги из конфигурации."""

        self._formats = {}
        for item in ['table_header', 'group_header', 'table_cell']:
            try:
                self._formats[item] = self.workbook.add_format(self.config[item]['format'])
            except:
                self._formats[item] = None

        self.init_columns_formats()
        self.set_columns_format()
        self.set_result_format()

    def export(self):
        """Процесс экспорта данных в XLSX-файл"""

        self.write_header()  # пишем заголовок
        self.write_data()  # пишем данные
        self.workbook.close()  # сохраняем в файл

    def write_header(self):
        """Пишет заголовок данных."""

        for col_num, value in enumerate(self.df.columns.values):
            self.worksheet.write(0, col_num, value, self._formats['table_header'])

    def _get_group_name(self, group_ind, group):
        if self.group_name_rule == "number":
            return "GROUP #" + str(group_ind + 1)
        return group

    def write_data(self):
        """Записывает данные на лист."""

        cur_row = 1 + self.group_blank_rows
        for group_ind, group in enumerate(self.groups):
            self.worksheet.merge_range(cur_row, 0, cur_row, self.prop_len - 1,
                                       self._get_group_name(group_ind, group),
                                       self._formats['group_header'])
            cur_row += 1
            for row in self.df[self.df[self.group_col] == group].itertuples(index=False):
                for col_num, value in enumerate(row):
                    try:
                        self.worksheet.write(cur_row, col_num, value,
                                             self._columns_formats[list(self.df.columns)[col_num]]['format_cell'])
                    except:
                        pass
                cur_row += 1
            cur_row += self.group_blank_rows

    def init_columns_formats(self):
        """Собирает форматы для столбцов из переданного словаря конфигурации."""

        col_formats = {col: {'format': None,
                             'format_cell': None,
                             'options': None,
                             'width': None} for col in self.df.columns}
        for item in self.config['columns']:
            col_names = re.split(r",\s*", item)
            for col_name in col_names:
                if 'format' in self.config['columns'][item]:
                    col_formats[col_name]['format'] = self.workbook.add_format(self.config['columns'][item]['format'])
                    _format_cell = self.config['columns'][item]['format']
                    _format_cell['border'] = 1
                    col_formats[col_name]['format_cell'] = self.workbook.add_format(_format_cell)
                if 'options' in self.config['columns'][item]:
                    col_formats[col_name]['options'] = self.config['columns'][item]['options']
                if 'width' in self.config['columns'][item]:
                    col_formats[col_name]['width'] = self.config['columns'][item]['width']
        self._columns_formats = col_formats

    def set_columns_format(self):
        # устанавливаем конфигурации для столбцов
        for item in self._columns_formats:
            self.worksheet.set_column(self.df.columns.get_loc(item),
                                      self.df.columns.get_loc(item),
                                      self._columns_formats[item]['width'],
                                      None,
                                      self._columns_formats[item]['options'])

    def set_result_format(self):
        """Устанавливает условное форматирование для столбца ввода результата проверки."""

        for i in range(len(self._colors)):
            l_ = self.columns_letters['result']
            col_range = l_ + "2:" + l_ + str(self.data_len + len(self.groups) * (1 + self.group_blank_rows) + 1)
            self.worksheet.conditional_format(col_range,
                                              {'type': 'formula',
                                               'criteria': '=$' + l_ + '2=' + str(i + 1),
                                               'format': self.workbook.add_format({'border': 1,
                                                                                   'bold': True,
                                                                                   'color': '#FFFFFF',
                                                                                   'bg_color': self._colors[i]})})

    @property
    def columns_letters(self):
        """
        Возращает словарь, в котором ключи -- названия столбцов из
        датафрейма, а значения -- имена столбцов на листе книги.
        """

        cols_letters = dict()
        for i, col in enumerate(self.df.columns):
            cols_letters[col] = string.ascii_uppercase[i]
        return cols_letters

    def letter(self, column_name):
        """Буква на листе, соответствующая имени столбца column_name датафрейма."""

        return self.columns_letters[column_name]

    @property
    def config(self):
        return self._config

    @property
    def data_len(self):
        return self.df.shape[0]

    @property
    def prop_len(self):
        return self.df.shape[1]

    @property
    def groups(self):
        return self.df[self.group_col].unique()
