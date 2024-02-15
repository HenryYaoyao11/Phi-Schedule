import os  
import ctypes  
import win32api  
import win32con  
import win32gui  
  
# 图片路径，这里假设图片位于当前工作目录下  
image_path = os.path.abspath('table\\pptjpg\\1\\幻灯片1.JPG')  
  
# SPI_SETDESKWALLPAPER 是Windows用来设置桌面壁纸的SPI值  
SPI_SETDESKWALLPAPER = 20  
  
# 调用SystemParametersInfo函数设置壁纸  
# 第一个参数是SPI动作，第二个参数是设置值（壁纸路径），第三个参数是指向设置值的类型（这里未用到），第四个参数是SPIF标志  
ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, image_path, win32con.SPIF_UPDATEINIFILE | win32con.SPIF_SENDWININICHANGE)  
  
# 刷新桌面  
win32gui.SendMessageTimeout(win32con.HWND_BROADCAST, win32con.WM_SETTINGCHANGE, 0, 'SPI_SETDESKWALLPAPER', win32con.SMTO_ABORTIFHUNG, 1000)
