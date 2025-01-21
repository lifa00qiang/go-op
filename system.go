package opsoft

import "github.com/go-ole/go-ole/oleutil"

//system
//命令
//
//接口目录
//RunApp: 运行可执行文
//WinExec: 运行可执行文件
//GetCmdStr: 运行命令行
//Delay: 延时
//Delays: 指定范围内随机毫秒数延迟
//SetClipboard: 设置剪贴板数据
//GetClipboard: 获取剪贴板数据

//接口方法

// RunApp
// 运行可执行文件,可指定模式
//
// long RunApp(app_path,mode)
// 参数	类型	描述
// app_path	string	指定的可执行程序全路径
// mode	int	取值如下
// mode
//
// 值	描述
// 0	普通模式
// 1	加强模式
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op.RunApp("c:\\windows\\notepad.exe",0)
func (o *Opsoft) RunApp(appPath string, mode int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "RunApp", appPath, mode)
	return ret.Value() == 1
}

//WinExec
//运行可执行文件，可指定显示模式
//
//long WinExec(cmdline, cmdshow)
//参数	类型	描述
//cmdline	string	指定的可执行程序全路径
//cmdshow	int	取值如下
//cmdshow
//
//值	描述
//0	隐藏
//1	用最近的大小和位置显示,激活
//返回值
//
//类型：int
//
//0: 失败
//1: 成功
//示例
//
//op.WinExec("c:\\windows\\notepad.exe",1)

func (o *Opsoft) WinExec(cmdline string, cmdshow int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "WinExec", cmdline, cmdshow)
	return ret.Value() == 1
}

//GetCmdStr
//运行命令行并返回结果
//
//string GetCmdStr(cmdline, millseconds)
//参数	类型	描述
//cmdline	string	指定的可执行程序全路径
//millseconds	int	等待的时间(毫秒)
//返回值
//
//类型：string
//
//cmd 输出的字符
//
//示例
//
//str = op.GetCmdStr("cmd.exe" 2000)

func (o *Opsoft) GetCmdStr(cmdline string, millseconds int64) string {
	ret, _ := oleutil.CallMethod(o.op, "GetCmdStr", cmdline, millseconds)
	return ret.ToString()
}

//SetClipboard
//设置剪贴板数据
//
//long SetClipboard(str)
//参数	类型	描述
//str	string	指设置剪贴板内容的字符串
//返回值
//
//类型：int
//
//0: 失败
//1: 成功
//示例
//
//ret = op.SetClipboard("Hello, World! This is a test string for the clipboard.")

func (o *Opsoft) SetClipboard(str string) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetClipboard", str)
	return ret.Value() == 1
}

//GetClipboard
//从系统剪贴板获取数据
//
//string GetClipboard()
//返回值
//
//类型：string
//
//成功则返回剪贴板数据
//
//示例
//
//str = op.GetClipboard()

func (o *Opsoft) GetClipboard() string {
	ret, _ := oleutil.CallMethod(o.op, "GetClipboard")
	return ret.ToString()
}

//Delay
//该函数旨在实现一个指定毫秒数的延迟，同时确保在此期间不会阻塞用户界面（UI）操作
//
//long Delay(mis)
//参数	类型	描述
//mis	int	指定延迟的时间，单位为毫秒
//返回值
//
//类型：int
//
//0: 失败
//1: 成功
//示例
//
//ret = op.Delay(2000) // 2秒

func (o *Opsoft) Delay(mis int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "Delay", mis)
	return ret.Value() == 1
}

//Delays
//该函数旨在实现一个指定毫秒数的延迟，同时确保在此期间不会阻塞用户界面（UI）操作；
//
//long Delays(mis_min,mis_max)
//参数	类型	描述
//mis_min	int	指定延迟时间的最小值，单位为毫秒
//mis_max	int	指定延迟时间的最大值，单位为毫秒
//返回值
//
//类型：int
//
//0: 失败
//1: 成功
//函数将随机选择一个介于 mis_min 和 mis_max 之间的延迟时间。
//
//示例
//
//ret = op.Delays(1000,5000) // 最小1秒, 最大5秒

func (o *Opsoft) Delays(misMin int64, misMax int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "Delays", misMin, misMax)
	return ret.Value() == 1
}
