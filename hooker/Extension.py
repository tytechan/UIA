# -*- coding: utf-8 -*-
import requests
import json
from flask import Flask, jsonify, request, make_response
from config.DirAndTime import *
from config import Globals as cf
from hooker.Compile import recordIntoProject

parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
logPath = createCurrentDateDir(parentDirPath + "\\log\\")

if not os.path.exists(logPath + "\\HookLog.txt"):
    createTXT(logPath, "HookLog.txt")
hookLogPath = logPath + "\\HookLog.txt"


app = Flask(__name__)
request_counter = 0

@app.errorhandler(404)
def not_found(error):
    return make_response(jsonify({'error': 404}), 404)

@app.route('/rpa/xpath', methods=['POST'])
def handler():
    global request_counter
    # print("recv data = {}".format(str(request.json)))
    # print(request.form.get('expression'))
    expression = request.json.get('expression')
    print("xpath:", expression)

    isRecord = recordIntoProject(expression)
    if isRecord:
        fp = open(hookLogPath, "a", encoding='utf-8')
        fp.write('\n-----')
        fp.write('\n' + '[捕获时间] %s' % getCurrentTime(date=True))
        fp.write('\n' + '[捕获对象类型] %s' % cf.get_value("autoType"))
        fp.write('\n' + '[层级结构]' + '\n')
        fp.write(expression)

    response = make_response(
        jsonify({'error': 0, 'msg': 'success', 'data': {'counter': request_counter, 'request': request.json}}))
    # print(response._status_code)
    return response

def run():
    app.run(debug=True,
            port=54321,
            use_reloader=False,
            threaded=True)

if __name__ == '__main__':
    app.run(debug=True, port=54321, use_reloader=False, threaded=True)