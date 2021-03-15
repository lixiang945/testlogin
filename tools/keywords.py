from selenium import webdriver


class Keywords:
    def __init__(self):
        self.br = webdriver.Chrome(r'd://webdrivers/chromedriver.exe')
        self.br.maximize_window()
        self.br.implicitly_wait(5)

    def geturl(self, location):
        self.br.get(location)

    def find_ele(self, location=''):
        if location.startswith('/'):
            ele = self.br.find_element_by_xpath(location)
        else:
            ele = self.br.find_element_by_id(location)
        return ele

    def click_ele(self, location):
        self.find_ele(location).click()

    def input_text(self, location, value):
        self.find_ele(location).send_keys(value)

    def change_frame(self, location):
        self.br.switch_to.frame(location)

    def br_quit(self):
        self.br.quit()


if __name__ == '__main__':
    keys = Keywords()
