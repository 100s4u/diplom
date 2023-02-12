import sys
import tableConverter
import timetableCreator

def main():
    # DEBUG
    sys.argv.append( '.\Задание на ВКР.xls' )

    selected_list = tableConverter.getListFromFile()

    # формируем расписания
    timetables = timetableCreator.listToTimetables(selected_list)

    # записываем их в файл
    tableConverter.saveTimetableToFile(timetables)

    # DEBUG
    f = open("result.txt", "w")
    f.write( str(timetables) )
    f.close()

if __name__ == '__main__':
    main()