import sys
import tableConverter
import timetableCreator
import result_table_creator

def generate():
    # DEBUG
    sys.argv.append( '.\start_file.xls' )

    selected_lists = tableConverter.getListsFromFile()

    # формируем расписания
    timetables, failed_lessons = timetableCreator.listsToTimetables(selected_lists)

    # конвертация расписания в json
    timetables_json = tableConverter.saveTimetableToJson(timetables, failed_lessons)

    # запись json'а в файл
    with open('json_data.json', 'w') as outfile:
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