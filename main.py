from flask import *
from flask_cors import CORS, cross_origin
import webview
import sys
import threading

import tableConverter
import timetableCreator
import result_table_creator


jsonData = './build/json_data.json'
startFile = './start_file.xls'


def generate(file):
    # DEBUG
    sys.argv.append(file)

    selected_lists = tableConverter.getListsFromFile()

    # формируем расписания
    timetables, failed_lessons = timetableCreator.listsToTimetables(selected_lists)

    # конвертация расписания в json
    timetables_json = tableConverter.saveTimetableToJson(timetables, failed_lessons)

    # запись json'а в файл
    with open(jsonData, 'w') as outfile:
        outfile.write(timetables_json)
    
    #print(timetables_json)

    # записываем их в файл
    result_table_creator.main()
    #tableConverter.saveTimetableToFile(timetables) 

    # DEBUG
    f = open("result.txt", "w")
    f.write( str(timetables) )
    f.close()


app = Flask(__name__, static_url_path='', static_folder='build', template_folder='build')
CORS(app, support_credentials=True)


@app.route('/upload', methods = ['POST']) 
def upload(): 
    if request.method == 'POST': 
        f = request.files['file'] 
        f.save(startFile)
        generate(startFile)
        result = ''
        with open(jsonData, 'r') as f:
            json_data = json.load(f)
        with open(jsonData, 'w') as f:
            result = json_data
            json.dump(json_data, f, indent=2)
        f.close()
        return result
@cross_origin(supports_credentials=True)

@app.route('/changeLesson', methods = ['POST']) 
def change(): 
    if request.method == 'POST': 
        data = json.loads(request.data) # string to json
        result = ''
        with open(jsonData, 'r') as f:
            json_data = json.load(f)
            oldPare = json_data['timetables'][data['group']][data['_week']][int(data['_weekDay'])][int(data['_lesson'])]
            # копируем пару
            json_data['timetables'][data['group']][data['_week']][int(data['_weekDay'])][int(data['_lesson'])] = json_data['timetables'][data['group']][data['week']][int(data['weekDay'])][int(data['lesson'])]
            # устанавливаем кабинет, если он должен быть изменен
            if(len(data['cabinet'])>0):
                json_data['timetables'][data['group']][data['_week']][int(data['_weekDay'])][int(data['_lesson'])]['cabinet'] = data['cabinet']
            #меняем старую пару
            json_data['timetables'][data['group']][data['week']][int(data['weekDay'])][int(data['lesson'])] = oldPare
        with open(jsonData, 'w') as f:
            result = json_data
            json.dump(json_data, f, indent=2)
        f.close()
        return result
@cross_origin(supports_credentials=True)

@app.route('/pasteLesson', methods = ['POST']) 
def paste(): 
        data = json.loads(request.data) # string to json
        print(request.data)
        result = ''
        with open(jsonData, 'r') as f:
            json_data = json.load(f)
            json_data['timetables'][data['group']][data['week']][int(data['weekDay'])][int(data['lesson'])] = data['cabinet']
            json_data['timetables'][data['group']][data['week']][int(data['weekDay'])][int(data['lesson'])] = {
                'teacher': data['teacher'],
                'lesson': data['subject'],
                'group': data['group'],
                'hours': 1,
                'cabinet': data['cabinet']
            }
        with open(jsonData, 'w') as f:
            result = json_data
            json.dump(json_data, f, indent=2)
        f.close()
        return result
@cross_origin(supports_credentials=True)

@app.route('/save', methods = ['POST']) 
def save():
        print(request.data)
        result_table_creator.main()
        result = ''
        with open(jsonData, 'r') as f:
            json_data = json.load(f)
            result = json_data
        f.close()
        return result
@cross_origin(supports_credentials=True)


@app.route('/')
def home():
   return render_template('index.html')

def start_server():
    app.run(host='127.0.0.1', port=5000)

if __name__ == '__main__':
    t = threading.Thread(target=start_server)
    t.daemon = True
    t.start()

    window = webview.create_window('Генератор расписания', 'build/index.html', width=1700, height=1000)
    webview.start()
    sys.exit()