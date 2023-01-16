from flask import Flask
from flask import request
import json
import requests
import openpyxl
from datetime import datetime

app = Flask(__name__)


@app.route("/")
def hello_world():
    return "hello, world!"


@app.route("/expert/review", methods=['POST'])
def review():
    excel_file = request.files['file']
    # 엑셀 파일 받아서 파싱
    # 검증 필요한 열들 플랫폼에 검증 요청
    excel_item,excel_rows = __parse_excel(excel_file)

    # 플랫폼에서 검증된 열들의 결과값 저장
    review_results = []
    for row in excel_rows:
        validate_result = __request_part_to_platform(row)
        json_object = json.loads(validate_result)
        review_results.append(json_object['verification'])

    # 검증된 열들의 결과값을 종합 처리후 응답
    passvalidate_results=[]
    result=True
    for verification, excel_row in zip(review_results,excel_rows):
        if verification == False:
          result=False
        passvalidate_result={
            "partName": excel_row[0],
            "designValue": excel_row[1],
            "verification": verification,
        }
        passvalidate_results.append(passvalidate_result)
    
    response = {
        excel_item[0][0]: excel_item[0][1],
        excel_item[1][0]: excel_item[1][1],
        excel_item[2][0]: excel_item[2][1],
        "passReview": result,
        "reviewResults": passvalidate_results
    }
          

    # response = review_results
    # 응답 예시

    # response = {
    #     "partNo": "LM2576HVSX-ADJ/NOPB",
    #     "type": "step_down",
    #     "manufacturer_name": "TI",
    #     "passReview": False,
    #     "reviewResults": [
    #         {
    #             "partName": "oprating_temperature_min",
    #             "designValue": -20,
    #             "verification": True,
    #         },
    #         {
    #             "partName": "oprating_temperature_max",
    #             "designValue": 100,
    #             "verification": False,
    #         },
    #         {
    #             "partName": "storage_temperature_min",
    #             "designValue": -45,
    #             "verification": True,
    #         },
    #         {
    #             "partName": "storage_temperature_max",
    #             "designValue": 110,
    #             "verification": False,
    #         }
    #     ]
    # }

    # return response = {
    #     "partNo": "LM2576HVSX-ADJ/NOPB",
    #     "type": "step_down",
    #     "manufacturer_name": "TI",
    #     "passReview": false,
    #     "reviewResults": [
    #         {
    #             "partName": "oprating_temperature_min",
    #             "designValue": -20,
    #             "verification": true,
    #         },
    #         {
    #             "partName": "oprating_temperature_max",
    #             "designValue": 100,
    #             "verification": false,
    #         },
    #         {
    #             "partName": "storage_temperature_min",
    #             "designValue": -45,
    #             "verification": true,
    #         },
    #         {
    #             "partName": "storage_temperature_max",
    #             "designValue": 110,
    #             "verification": false,
    #         }
    #     ]
    # }

    return json.dumps(response)


def __parse_excel(excel_file):
    # 엑셀 파싱후 검증 필요한 rows 반환
    wb=openpyxl.load_workbook(excel_file)
    snames=wb.sheetnames
    sheet=wb[snames[0]]
    excel_rows=[]
    item=[]
    for row in range(1,sheet.max_row):
      cols=[]
      for col in sheet.iter_cols(min_col=0, max_col=sheet.max_column):
          cols.append(col[row].value)
          del cols[1:3]
          if row < 4:
            del cols[2]
            item.append(cols)
          else:
            if cols[2] == 'o':
              del cols[2]
              excel_rows.append(cols)
    return item,excel_rows


def __request_part_to_platform(parsed_row):
    # dockerized url
    url = "http://platform/review/part"

    # local url
    # url = "http://localhost:8080/review/part"
    headers = {'Content-Type': 'application/json'}
    example_body = {
        "partName": parsed_row[0],
        "designValue": parsed_row[1]
    }

    # body = parsed_row
    body = example_body

    response = requests.post(url, json=body, headers=headers)
    return response.get_json()

    # response 응답 예시
    # response = {
    #    "verification": true
    # }


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001)
