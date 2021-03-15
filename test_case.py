import time
import allure
import pytest
from tools.keywords import Keywords
# from tools.read_json import datas
from tools.read_ex import ReadEX
# from tools.test2 import ReadEX

rt = ReadEX('测试用例.xlsx', 'logintest')
# rt.read_excel('test01')
rt.read_excel('test02')


@allure.feature('qq功能测试')
class Test_yml:
    def setup_class(self):
        self.web = Keywords()
        # pass

    def teardown_class(self):
        time.sleep(2)
        self.web.br_quit()

    @allure.story('登录页测试')
    @pytest.mark.parametrize('case', rt.read_json('test02')['QQ登录'])
    def test_login(self, case):
        allure.dynamic.title(case['title'])
        allure.dynamic.description(case['title'])
        testcases = case['cases']
        time.sleep(1)
        for cases in testcases:
            # 将case（字典）中的值取出并强制转换为列表
            listcase = list(cases.values())
            try:
                random_userName = str(int(time.time()))
                listcase[5] = listcase[5].replace('{number}', random_userName)
            except Exception as f:
                pass
            try:
                with allure.step(listcase[0]):
                    func = getattr(self.web, listcase[3])
                    # # func = getattr(对象，函数名)，获取第三个字段的值作为函数名传入
                    # 从第4个字段开始，将值传入
                    values = listcase[4:]
                    func(*values)
                    # 将测试结果写入excel
                    rt.white_excel(listcase[1], 'pass')
                    allure.attach(self.web.br.get_screenshot_as_png(), '通过截图', allure.attachment_type.PNG)
            except Exception as e:
                rt.white_excel(listcase[1], 'fail')
                rt.white_excel(listcase[2], f'错误的原因是：{e}')
                allure.attach(self.web.br.get_screenshot_as_png(), '失败截图', allure.attachment_type.PNG)


if __name__ == '__main__':
    pytest.main(['-s'])
