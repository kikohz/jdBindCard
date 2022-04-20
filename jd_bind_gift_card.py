# selenium 4
from ast import Not
from lib2to3.pytree import type_repr
from pkgutil import iter_modules
from time import sleep
from cv2 import log
from numpy import NaN, true_divide
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import pandas as pd
from loguru import logger
from PIL import Image
from jd_captcha import JDcaptcha_base64
from jd_yolo_captcha import JDyolocaptcha

user_agent = "Mozilla/5.0 (Linux; Android 11; M2007J3SC) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.210 Mobile Safari/537.36"
my_cookies = ""

class JDBindGiftCard(object):
    def __init__(self) -> None:
        # 初始化selenium配置
        self.browser = self.config_browser()
        self.wait = WebDriverWait(self.browser, 30)
        self.wait_check = WebDriverWait(self.browser, 3)
        # other
        self.config()
    # 配置浏览器
    def config_browser(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument("--incognito")  # 无痕模式
        chrome_options.add_experimental_option("excludeSwitches", ['enable-automation', 'enable-logging'])
        chrome_options.add_argument(f'user-agent={user_agent}')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),desired_capabilities={},options=chrome_options)
        return driver
    
    def config(self):
        self.JDyolo = JDyolocaptcha({"yolov4_weights":"yolov4/yolov4-tiny-custom.weights",
                                     "yolov4_cfg":"yolov4/yolov4-tiny-custom.cfg",
                                     "yolov4_net_size":512})
        self.image_captcha_type = 'local'
        
    
    # 获取cookie
    def getcookie(self):
        self.browser.get("https://plogin.m.jd.com/login/login")
        try:
            wait = WebDriverWait(self.browser, 135)
            print("请在网页端通过手机号码登录")
            wait.until(EC.presence_of_element_located((By.ID, 'msShortcutMenu')))
            self.browser.get("https://home.m.jd.com/myJd/newhome.action")
            username = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'my_header_name'))).text
            pt_key, pt_pin, cookie = "", "", ""
            for _ in self.browser.get_cookies():
                if _["name"] == "pt_key":
                    pt_key = _["value"]
                if _["name"] == "pt_pin":
                    pt_pin = _["value"]
                if pt_key and pt_pin:
                    break
            cookie = "pt_key=" + pt_key + ";pt_pin=" + pt_pin + ";"
            print("获取成功",username)
            return cookie
        except:
            print('获取失败')
            return ''
    # 读取xlsx 礼品卡文件
    def loadXlsx(self):
        df = pd.read_excel("card.xlsx")
        all_card = df.values[:,1]
        logger.info(all_card)
        return all_card

    # 本地识别图形验证码并模拟点击
    def local_auto_identify_captcha_click(self):
        for _ in range(4):
            cpc_img = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="cpc_img"]')))
            zoom = cpc_img.size['height'] / 170
            cpc_img_path_base64 = self.wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="cpc_img"]'))).get_attribute(
                'src').replace("data:image/jpeg;base64,", "")
            pcp_show_picture_path_base64 = self.wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//*[@class="pcp_showPicture"]'))).get_attribute('src')
            # 正在识别验证码
            if self.image_captcha_type == "local":
                logger.info("正在通过本地引擎识别")
                res = JDcaptcha_base64(cpc_img_path_base64, pcp_show_picture_path_base64)
            else:
                logger.info("正在通过深度学习引擎识别")
                res = self.JDyolo.JDyolo(cpc_img_path_base64, pcp_show_picture_path_base64)
            if res[0]:
                ActionChains(self.browser).move_to_element_with_offset(
                        cpc_img, int(res[1][0] * zoom),
                        int(res[1][1] * zoom)
                    ).click().perform()

            # 图形验证码坐标点击错误尝试重试
            # noinspection PyBroadException
                try:
                    WebDriverWait(self.browser, 3).until(EC.presence_of_element_located(
                         (By.XPATH, "//p[text()='验证失败，请重新验证']")
                        ))
                    sleep(1)
                    return False
                except Exception as _:
                    return True
            else:
                logger.info("识别未果")
                self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@class="jcap_refresh"]'))).click()
                sleep(1)
                return False
        return False
    def get_code_pic(self, name='code_pic.png'):
        """
        获取验证码图像
        :param name:
        :return:
        """
        # 确定验证码的左上角和右下角坐标
        code_img = self.wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='captcha_modal']//div")))
        location = code_img.location
        size = code_img.size
        _range_ = (int(location['x']), int(location['y']), (int(location['x']) + int(size['width'])),
                    (int(location['y']) + int(size['height'])))

        # 将整个页面截图
        self.browser.save_screenshot(name)

        # 获取浏览器大小
        window_size = self.wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='root']")))
        width, height = window_size.size['width'], window_size.size['height']

        # 图片根据窗口大小resize，避免高分辨率影响坐标
        i = Image.open(name)
        new_picture = i.resize((width, height))
        new_picture.save(name)

        # 剪裁图形验证码区域
        code_pic = new_picture.crop(_range_)
        code_pic.save(name)
        sleep(2)
        return code_img

    def bind_card(self,card):
        sleep(2)
        try:
            self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[text()="绑定新卡"]'))).click()
            sleep(2)
            self.browser.find_element_by_tag_name('input').send_keys(card)
            sleep(2)
            self.browser.find_element_by_class_name('button').click()
            if not self.local_auto_identify_captcha_click():
                logger.info('本地识别验证码位置点击错误，会再次尝试')
                sleep(3)
                if not self.local_auto_identify_captcha_click():
                    logger.info('本地识别验证码位置点击错误，跳过本次')
                    return "fail"
            sleep(4)
            #判断是否已经绑定卡了
            try:
                if self.browser.find_element_by_class_name('pop-card-pt').text == '您已绑定了该卡！':
                    sleep(2)
                    self.browser.find_element_by_class_name('Yep-dialog-button').click()
                    sleep(2)
                    self.browser.back()
                    return "used"
            except Exception as e:
                logger.debug(e)
            # 点击确定按钮
            self.browser.find_element_by_class_name('pop-card-ok').click()
            sleep(2)
            # 点击返回列表
            self.browser.find_element_by_class_name('outline-button').click()
            return "success"
        except Exception as e:
            logger.debug(e)
            return "fail"
        

    def start_bind(self):
        self.browser.get("https://m.jd.com/")
        self.browser.set_window_size(500,700)
        # 写入cookies
        self.browser.delete_all_cookies()
        for cookie in my_cookies.split(";", 1):
            self.browser.add_cookie(
                {"name": cookie.split("=")[0].strip(" "), "value": cookie.split("=")[1].strip(";"), "domain": ".jd.com"}
            )
        self.browser.refresh()
        sleep(2)
        cards = self.loadXlsx()
        logger.info(cards)
        for card in cards:
            self.browser.get("https://mygiftcard.jd.com/giftcardForM.html?source=JDM&sceneval=2&jxsid=16503716408399792128&appCode=ms0ca95114#/mygiftcard")
            if self.bind_card(card) == 'fail':
                logger.error(card)
            else:
                logger.success(card)
        

if __name__ == '__main__':
    jd = JDBindGiftCard()
    jd.start_bind()
    # jd.loadXlsx()