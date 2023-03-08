
from win32com.client.gencache import EnsureDispatch

한컴=EnsureDispatch("HWPFrame.HwpObject")

한컴.XHwpWindows.Item(0).Visible=True
