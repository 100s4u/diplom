import sys
import tableConverter
import timetableCreator
import result_table_creator


jsonData = './build/json_data.json'
startFile = './start_file.xls'

def generate():
    # DEBUG
    sys.argv.append( '.\start_file.xls' )

    selected_lists = tableConverter.getListsFromFile()

    # формируем расписания
    timetables, failed_lessons, teachers_timetables, cabinets_timetables = timetableCreator.listsToTimetables(selected_lists)

    # конвертация расписания в json
    timetables_json = tableConverter.saveTimetableToJson(timetables, failed_lessons, teachers_timetables, cabinets_timetables)

    # запись json'а в файл
    with open(jsonData, 'w') as outfile:
        outfile.write(timetables_json)
    
    #print(timetables_json)

    # записываем их в файл
    result_table_creator.main()
    # tableConverter.saveTimetableToFile(timetables) 

    # DEBUG
    f = open("result.txt", "w")
    f.write( str(timetables) )
    f.close()

if __name__ == '__main__':
    generate()