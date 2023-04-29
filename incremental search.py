import xlwings, rapidfuzz
from excel import Excel_regex
from progress.bar import IncrementalBar


что_сравниваем = []
с_чем_сравниваем = []
excel_regex = Excel_regex()

def load_data() -> list:

    var = excel_regex.source_file_and_cells()
    var = excel_regex.get_data(var[2],
                                var[1],
                                var[0])
    return var

def compare(что_сравниваем: list, с_чем_сравниваем: list) -> list:
    
    '''Сравнивает два списка с автораспаковкой, только если в сравнении
    более одного столбца! (можно во второй столбец добавить любой символ)'''
    result = []
    bar = IncrementalBar('Обработка', max = len(что_сравниваем))
    
    for source in что_сравниваем:

        tmp = []
        bar.next()
        
        for compare_str in с_чем_сравниваем:
                
                tmp_compare = rapidfuzz.fuzz.token_sort_ratio(source, compare_str[0])
                
                if type(compare_str) == list:
                    tmp.append([tmp_compare, source, *compare_str])
                else:
                    tmp.append([tmp_compare, source, compare_str])
                
        tmp.sort()
        
        result.append(tmp[-1])
        
    bar.finish()
    return result

def add_columns(source: list) -> list:

    needed_len = len(source[0]) - 3
    colums = ['% схожести', 'Источник', 'Луший результат']

    if needed_len > 0:
        
        number = 1
        
        while number <= needed_len:
            colums.append(f'Столбец {number}')
            number += 1

    source.insert(0, colums)

    return source

def make_book(source: list) -> None:

    new_book = xlwings.Book()
    new_sheet = new_book.sheets[0]
    new_sheet.range('A1').expand().value = source

def delete_nonetype(source: list) -> list:

    for i in range(len(source)):

        tmp = source[i]

        if str(type(tmp)) == "<class 'NoneType'>":
            source[i] = ''

    return source


if __name__ == '__main__':
    while True:
    
        input('\nНажмите ENTER, чтобы загрузить "что сравниваем"\n')

        что_сравниваем = load_data()
        что_сравниваем = delete_nonetype(что_сравниваем)

        input('\nНажмите ENTER, чтобы загрузить "с чем сравниваем"\n')

        с_чем_сравниваем = load_data()
        с_чем_сравниваем = delete_nonetype(с_чем_сравниваем)

        result = compare(что_сравниваем, с_чем_сравниваем)
        result = add_columns(result)

        make_book(result)