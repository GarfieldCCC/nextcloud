import xlrd
import copy
import requests


class Header:
    def __init__(self):
        self.host = ""
        self.url = ""
        self.cookies = ""
        self.user_agent = ""
        self.request_token = ""

    def get_headers(self, filename):
        data = xlrd.open_workbook(filename)
        table = data.sheets()[0]
        self.host = table.cell_value(0, 1)
        self.url = table.cell_value(1, 1)
        self.cookies = table.cell_value(2, 1)
        self.user_agent = table.cell_value(3, 1)
        self.request_token = table.cell_value(4, 1)
        return {
            "HOST": self.host,
            "Referer": self.url,
            "User-Agent": self.user_agent,
            "Cookie": self.cookies,
            "requesttoken": self.request_token
        }


class User:
    def __init__(self):
        self.template = {"userid": "Test", "displayName": "测试用户", "password": "testpassword", "email": "test@new.com",
                         "groups": ["普通用户"], "subadmin": ["普通用户"], "quota": "1GB", "language": "zh_CN"}

    def get_user_data(self, filename):
        res = []
        data = xlrd.open_workbook(filename)
        table = data.sheets()[0]
        n_rows = table.nrows
        for i in range(1, n_rows):
            self.template['userid'] = table.cell_value(i, 0)
            self.template['displayName'] = table.cell_value(i, 1)
            self.template['password'] = table.cell_value(i, 2)
            self.template['email'] = table.cell_value(i, 3)
            if table.cell_value(i, 4).__len__() > 0:
                self.template['groups'] = table.cell_value(i, 4).split('、')
            else:
                self.template['groups'] = []
            if table.cell_value(i, 5).__len__() > 0:
                self.template['subadmin'] = table.cell_value(i, 5).split('、')
            else:
                self.template['subadmin'] = []
            self.template['quota'] = table.cell_value(i, 6)
            res.append(copy.copy(self.template))
        return res


if __name__ == '__main__':
    user_file = "users.xlsx"
    header_file = "headers.xlsx"
    user = User()
    header = Header()

    user_info = user.get_user_data(user_file)
    headers = header.get_headers(header_file)
    url = headers["Referer"]

    for info in user_info:
        print(info)
        response = requests.post(url, json=info, headers=headers)
        print(response.text)
