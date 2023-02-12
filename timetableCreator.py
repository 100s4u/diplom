def listsToTimetables(lists):
    [lessons_list, ordered_lessons_list] = lists
    # парсим лист в список пар
    lessons = list_to_lessons(lessons_list)
    # формируем расписание
    return lessons_to_timetables(lessons, ordered_lessons_list)

def lessons_to_timetables(lessons, ordered_lessons):
    timetables = {} # массив расписаний, 1 элемент - 1 расписание для группы

    # формируем список групп в timetables
    for lesson in lessons:
        # добавляем группу в timetables если ее там еще нет
        if (not lesson['group'] in timetables):
            #print("Добавляем группу " + lesson['group'])
            timetables[lesson['group']] = {
                'even': [[], [], [], [], [], [], []], # четная неделя
                'odd': [[], [], [], [], [], [], []] # нечетная неделя
            }

    # в дни добавляем пустые предметы
    for group in timetables:
        timetable = timetables[group]
        for week in timetable:
            week = timetable[week]
            for day in week:
                for i in range(20):
                    day.append('')
    
    # создаем массив пар, которые не удалось установить
    failed_lessons = []
    # работаем отдельно по каждой группе
    for group in timetables.keys():
        week = 'even' # начинаем с четной недели
        
        # получаем список запланированных занятий у этой группы
        lessons_of_group = get_lessons_of_group(lessons, group)

        # формируем расписание
        fill_timetable(timetables, group, lessons_of_group)
    
    return timetables

'''
Алгоритм заполнения расписания:
1. Проверить, что пара свободна
2. Внести пару
3. Перейти на следующий день
Повторять пока не кончатся пары
'''
def fill_timetable(timetables, group, week_lessons):
    day = 0 # начинаем с понедельника
    lesson_num = 0 # начинаем с 1 пары
    week = 'even' #  начинаем с четной недели

    while len(week_lessons) > 0:
        lesson = week_lessons[0]

        # проверяем, что на эту пару не установлена другая
        if timetables[group][week][day][lesson_num] == '':
            # проверяем свободен ли урок
            if is_lesson_free(timetables, lesson, week == 'even', day, lesson_num, lesson['teacher']):
                # размещение пары на своем месте
                timetables[group][week][day][lesson_num] = lesson

                # удаляем пару из week_lessons
                if (lesson['hours'] < 2):
                    week_lessons.remove(lesson)
                else:
                    lesson['hours'] -= 1

                # возвращаемся на первую пару первого дня
                day = 0
                lesson_num = 0
            else:
                # пробуем поставить в другое место
                day += 1
        
        day += 1
        if (day >= 5):
            if (week == 'even'):
                week = 'odd'
                day = 0
            else:
                week = 'even'
                day = 0
                lesson_num += 1

# из списка общего списка занятий возвращает занатия только для определенной группы
def get_lessons_of_group(lessons, target_group):
    result = []
    for lesson in lessons:
        if lesson['group'] == target_group:
            result.append(lesson)
    return result

# функция проверяет, свободен ли определенный предмет на конкретную неделю на конкретной паре
def is_lesson_free(timetables, lesson, is_week_even, day, lesson_num, teacher_name = '') -> bool:
    if is_week_even:
        week = 'even'
    else:
        week = 'odd'
    # проверяем что пара не занята
    for timetable in timetables:
        timetable = timetables[timetable]
        try:
            # если эта пара уже занята
            if timetable[week][day][lesson_num]['lesson'] == lesson['lesson']:
                # проверяем, занята ли она тем же преподом
                if timetable[week][day][lesson_num]['teacher'] == teacher_name:
                    return False
                else:
                    continue
            else:
                continue
        except:
            continue
    # проверяем что кабинет не занят
    if lesson['cabinet'] == '':
        return True
    for timetable in timetables:
        timetable = timetables[timetable]
        # проверяем каждую группу на определенный день определенную пару
        try:
            if timetable[week][day][lesson_num]['cabinet'] == lesson['cabinet']:
                return False
            else:
                continue
        except:
            continue
    return True

# функция проверяет, свободена конкретная пара, т.е. можно ли туда вотнкуть пару или она занята
#def is_lesson_free2():
#    timetables[group][week][day][lesson_num]

def list_to_lessons(lessons_list):
    # забираем значения таблицы
    table = lessons_list.values

    # ищем номер первой строки с парой
    first_row_num = None
    for idx, row in enumerate(table):
        temp = str(row[4])
        is_nan = temp == 'nan'
        if is_nan:
            continue
        if temp.isnumeric():
            first_row_num = idx
            break

    # от первой строки парсим все последующие, получая в результате массив с данными по парам
    lessons = []
    for i in range(first_row_num, len(table)):
        if str(table[i][0]) == 'nan':
            table[i][0] = ''
        if str(table[i][6]) == 'nan':
            table[i][6] = ''
        new_lesson = {
            'teacher': str(table[i][0]),
            'lesson': str(table[i][1]),
            'group': str(table[i][2]),
            'hours': int(table[i][4]),
            'cabinet': str(table[i][6]),
        }
        lessons.append(new_lesson)

    return lessons