import sys
import tableConverter
import timetableCreator

def main():
    # DEBUG
    sys.argv.append( '.\start_file.xls' )

    selected_lists = tableConverter.getListsFromFile()

    # формируем расписания
    timetables = timetableCreator.listsToTimetables(selected_lists)

    # записываем их в файл
    tableConverter.saveTimetableToFile(timetables)

    # DEBUG
    f = open("result.txt", "w")
    f.write( str(timetables) )
    f.close()

if __name__ == '__main__':
    main()