import time
import pandas as pd
import os

from openpyxl.styles import PatternFill, Border, Side
from openpyxl import load_workbook
from openpyxl.styles import Font


def add_mean_value(dataframe: pd.DataFrame) -> pd.DataFrame:
    value_for_group = dataframe.iloc[:, 0].unique()
    result_df = pd.DataFrame()
    for value in value_for_group:
        df = dataframe[dataframe.iloc[:, 0] == value]
        mean_value = df.groupby(df.iloc[:, 0]).mean(numeric_only=True).reset_index()
        mean_value.iloc[:, 0] = mean_value.iloc[:, 0] + ' (среднее значение)'
        result_df = pd.concat([result_df, df])
        if df.shape[0] != 1:
            result_df = pd.concat([result_df, mean_value])
    return result_df


def join_sheet_files(files: list[list], sheet_name: str) -> pd.DataFrame:
    dataframe = pd.DataFrame()
    for url, file in files:
        analiz_spam_df = pd.read_excel(file, sheet_name=sheet_name)
        analiz_spam_df['URL'] = url
        dataframe = pd.concat([dataframe, analiz_spam_df])
    return dataframe


def re_spam(url_way_file: list[list]) -> pd.DataFrame:
    columns = ['Слово (самая популярная словоформа)',
               'Повторы у Вас',
               'Минимум повторов (норм.)',
               'Максимум повторов (норм.)',
               'Переспам, %',
               'Переспам * IDF, %',
               'IDF',
               'Количество повторений',
               'URL']
    dataframe = join_sheet_files(url_way_file, 'Переспам')
    count_replay = dataframe.groupby('Слово (самая популярная словоформа)', as_index=False).agg(
        {'URL': 'count'}).rename(columns={'URL': 'Количество повторений'})
    dataframe = dataframe.merge(count_replay).sort_values(by=['Количество повторений',
                                                                                'Слово (самая популярная словоформа)'],
                                                                            ascending=False)
    return add_mean_value(dataframe[columns])


def replay_word(url_way_file: list[list]) -> pd.DataFrame:
    columns = ['Слово (самая популярная словоформа)',
               'Повторы у Вас',
               'Минимум повторов (норм.)',
               'Максимум повторов (норм.)',
               'Количество повторений',
               'URL']
    dataframe = join_sheet_files(url_way_file, 'Повторы слов')
    count_replay = dataframe.groupby('Слово (самая популярная словоформа)', as_index=False).agg(
        {'URL': 'count'}).rename(columns={'URL': 'Количество повторений'})
    dataframe = dataframe.merge(count_replay).sort_values(by=['Количество повторений',
                                                                      'Слово (самая популярная словоформа)'],
                                                                  ascending=False)
    return add_mean_value(dataframe[columns])


def add_common_word(url_way_file: list[list]) -> pd.DataFrame:
    columns = ['Слово (самая популярная словоформа)',
               'Важные словоформы',
               'Все словоформы у конкурентов',
               'Количество повторений',
               'URL']
    dataframe = join_sheet_files(url_way_file, 'Добавить важные слова')
    count_replay = dataframe.groupby('Слово (самая популярная словоформа)', as_index=False).agg(
        {'URL': 'count'}).rename(columns={'URL': 'Количество повторений'})
    dataframe = dataframe.merge(count_replay).sort_values(by=['Количество повторений',
                                                                      'Слово (самая популярная словоформа)'],
                                                                  ascending=False)
    return add_mean_value(dataframe[columns])


def dop_word(url_way_file: list[list]) -> pd.DataFrame:
    columns = ['Дополнительные слова',
               'Количество повторений',
               'URL']

    dataframe = join_sheet_files(url_way_file, 'Доп. слова')
    count_replay = dataframe.groupby('Дополнительные слова',
                                         as_index=False).agg({'URL': 'count'}).rename(
        columns={'URL': 'Количество повторений'})
    dataframe = dataframe.merge(count_replay).sort_values(by=['Количество повторений',
                                                                      'Дополнительные слова'],
                                                                  ascending=False)
    return add_mean_value(dataframe[columns])


def title(url_way_file: list[list]) -> pd.DataFrame:
    columns = ['Можно добавить слова',
               'Количество повторений',
               'URL']
    dataframe = join_sheet_files(url_way_file, 'title')
    count_replay = dataframe.groupby('Можно добавить слова', as_index=False).agg({'URL': 'count'}).rename(
        columns={'URL': 'Количество повторений'})
    dataframe = dataframe.merge(count_replay).sort_values(by=['Количество повторений',
                                                                      'Можно добавить слова'],
                                                                  ascending=False)
    return add_mean_value(dataframe[columns])


def wight_row(path):
    wb = load_workbook(path)
    # размер шрифта документа
    font_size = 10

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        # словарь с размерами столбцов
        cols_dict = {}
        list_null_value = [1]

        # проходимся по всем строкам документа
        for row in ws.rows:
            if row[-1].value is None:
                list_null_value.append(row[0].row)
            # теперь по ячейкам каждой строки
            for cell in row:
                # получаем букву текущего столбца
                letter = cell.column_letter
                name_letter = cell.coordinate

                thins = Side(border_style="medium", color="000000")
                ws[name_letter].border = Border(left=thins, right=thins)

                if row[-1].value is None:
                    ws[name_letter].fill = PatternFill('solid', fgColor="F5FFFA")
                    ws[name_letter].border = Border(top=thins, bottom=thins, left=thins, right=thins)
                if cell.row == 1:
                    ws[name_letter].fill = PatternFill('solid', fgColor="9FB6CD")
                    ws[name_letter].border = Border(top=thins, bottom=thins, left=thins, right=thins)

                # если в ячейке записаны данные
                if cell.value:
                    # устанавливаем в ячейке размер шрифта
                    cell.font = Font(name='Calibri', size=font_size)
                    # вычисляем количество символов, записанных в ячейку
                    len_cell = len(str(cell.value))
                    # длинна колонки по умолчанию, если буква
                    # текущего столбца отсутствует в словаре `cols_dict`
                    len_cell_dict = 0
                    # смотрим в словарь c длинами столбцов
                    if letter in cols_dict:
                        # если в словаре есть буква текущего столбца
                        # то извлекаем соответствующую длину
                        len_cell_dict = cols_dict[letter]

                    # если текущая длина данных в ячейке
                    #  больше чем длинна из словаря
                    if len_cell > len_cell_dict:
                        # записываем новое значение ширины этого столбца
                        cols_dict[letter] = len_cell
                        ###!!! ПРОБЛЕМА АВТОМАТИЧЕСКОЙ ПОДГОНКИ !!!###
                        ###!!! расчет новой ширины колонки (здесь надо подгонять) !!!###
                        new_width_col = len_cell * font_size ** (font_size * 0.009)
                        # применение новой ширины столбца
                        ws.column_dimensions[cell.column_letter].width = new_width_col

        for count in range(len(list_null_value) - 1):
            min_row = int(list_null_value[count] + 1)
            max_row = int(list_null_value[count + 1] - 1)
            ws.row_dimensions.group(min_row, max_row, hidden=True)

    wb.save(path)


def main():
    # PATH = input("Введите путь: ")
    print('Формирую отчет...')
    PATH = r'C:\Users\Gennady\Documents\Relevantus_data_analysis\files\Report_14_02_2023__21_55'
    FILE_RESULT = os.path.join(PATH, 'Result.xlsx')
    RELEVANTUS_DATA_ANALYSIS = os.path.join(PATH, 'Relevantus_data_analysis.xlsx')
    result_df = pd.read_excel(FILE_RESULT, sheet_name='Результаты')
    analiz_spam = result_df[['URL', 'Анализ переспама']].values.tolist()
    recommendations = result_df[['URL', 'Рекомендации по улучшению релевантности']].values.tolist()

    with pd.ExcelWriter(RELEVANTUS_DATA_ANALYSIS, engine='xlsxwriter') as writer:
        re_spam(analiz_spam).to_excel(writer, sheet_name='Переспам', index=False)
        replay_word(recommendations).to_excel(writer, sheet_name='Повторы слов', index=False)
        add_common_word(recommendations).to_excel(writer, sheet_name='Добавить важные слова', index=False)
        dop_word(recommendations).to_excel(writer, sheet_name='Доп. слова', index=False)
        title(recommendations).to_excel(writer, sheet_name='Title', index=False)

    wight_row(RELEVANTUS_DATA_ANALYSIS)
    print(f'Отчет сформирован, доступен по следующему пути: \n{RELEVANTUS_DATA_ANALYSIS}')
    time.sleep(5)


if __name__ == '__main__':
    main()
