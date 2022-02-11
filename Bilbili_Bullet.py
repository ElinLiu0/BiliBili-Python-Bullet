from selenium.webdriver import ChromeOptions,Chrome
import time
import win32com.client
class BulletMachine:
    def __init__(self,live_url) -> None:
        self.chrome_options = '--headless'
        self.live_url = live_url
        self.hold = 3
        options = ChromeOptions()
        options.add_argument(self.chrome_options)
        global driver
        driver = Chrome(options=options)
        driver.get(self.live_url)
        driver.implicitly_wait(self.hold)
        time.sleep(self.hold)
        global speaker
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.caching = None
    def main(self):     
        driver.refresh()
        BulletMain = driver.find_elements_by_class_name('chat-items')
        for i in BulletMain:
            NewiestBulletUser = i.find_elements_by_tag_name('span')[-2]
            try:
                NewiestBullet = i.find_elements_by_class_name('open-menu')[0].get_attribute('alt')
                message = f"{NewiestBulletUser.text.replace(' :','')}说：{NewiestBullet.text}"
            except:
                NewiestBullet = i.find_elements_by_tag_name('span')[-1]
                message = f"{NewiestBulletUser.text.replace(' :','')}说：{NewiestBullet.text}"
            if self.caching != message:   
                speaker.Speak(message)
                self.caching = message
            else:
                print('重复弹幕，跳过！')
            time.sleep(2)
live_room_link = input('请输入BiliBili直播间地址：')
boot = BulletMachine(live_url=live_room_link)
if __name__ == "__main__":
    while True:
        boot.main()