package opsoft

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/lifa00qiang/go-op/tool"
	"strings"
)

//EnumWindow: 枚举窗口
//EnumWindowByProcess: 根据进程条件枚举窗口
//EnumProcess: 根据指定进程名称枚举进程 PID
//ClientToScreen: 窗口客户区坐标转换为屏幕坐标
//FindWindow: 根据窗口名称和类名寻找窗口
//FindWindowByProcess: 根据进程名称寻找窗口
//FindWindowByProcessId: 根据进程 Id 寻找窗口
//FindWindowEx: 寻找窗口
//GetClientRect: 获取窗口客户区域矩形
//GetClientSize: 获取窗口客户区域大小u
//GetForegroundFocus: 获取具有输入焦点的窗口句柄
//GetForegroundWindow: 获取顶层活动窗口
//GetMousePointWindow: 获取鼠标指定坐标窗口句柄
//GetPointWindow: 获取指定坐标窗口句柄
//GetProcessInfo: 获取进程信息
//GetSpecialWindow: 获取特殊窗口
//GetWindow: 获取窗口句柄
//GetWindowClass: 获取窗口的类名
//GetWindowProcessId: 获取指定窗口所在的进程 ID
//GetWindowProcessPath: 获取指定窗口所在的进程全路径
//GetWindowRect: 获取窗口矩形
//GetWindowState: 获取窗口的状态
//GetWindowTitle: 获取窗口的标题
//MoveWindow: 移动窗口
//ScreenToClient: 把屏幕坐标转换为窗口坐标
//SendPaste: 向指定窗口发送粘贴命令
//SetClientSize: 设置窗口客户区域大小
//SetWindowState: 设置窗口的状态
//SetWindowSize: 设置窗口的大小
//SetWindowText: 设置窗口的标题
//SetWindowTransparent: 设置窗口的透明度
//SendString: 向指定窗口发送文本数据
//SendStringIme: 向指定窗口发送文本数据

// 接口方法
// EnumWindow
// 根据指定条件,枚举系统中符合条件的窗口
//
// string EnumWindow(parent,title,class_name,filter)
// 参数	类型	描述
// parent		int		父窗口的句柄
// title		string	窗口的标题
// class_name	string	窗口的类名
// filter		int		表示窗口的过滤条件,可以是以下值
// filter
//
// 值	描述
// 1	匹配窗口标题，参数 title 有效
// 2	匹配窗口类名，参数 class_name 有效
// 4	只匹配指定父窗口的第一层子窗口
// 8	匹配所有者窗口为 0 的窗口，即顶级窗口
// 16	匹配可见的窗口
// 32	匹配出的窗口按照窗口打开顺序依次排列
// 这些值可以相加,比如 4+8+16 就是类似于任务管理器中的窗口列表
//
// 返回值
// 类型：string
// 返回所有匹配到的窗口句柄

// 无
func (o *Opsoft) EnumWindow(parent int64, title string, className string, filter int64) (out []int64) {
	ret, _ := oleutil.CallMethod(o.op, "EnumWindow", parent, title, className, filter)

	sp := strings.Split(ret.ToString(), ",")
	for _, v := range sp {
		out = append(out, tool.StrToInt64(v))
	}
	return
}

// EnumWindowByProcess
// 根据指定进程以及其它条件,枚举系统中符合条件的窗口
//
// string EnumWindowByProcess(process_name,title,class_name,filter)
// 参数	类型	描述
// process_name	string	进程名称
// title	string	窗口的标题
// class_name	string	窗口的类名
// filter	int	表示窗口的过滤条件,可以是以下值
// filter
//
// 值	描述
// 1	匹配窗口标题，参数 title 有效
// 2	匹配窗口类名，参数 class_name 有效
// 4	只匹配指定父窗口的第一层子窗口
// 8	匹配所有者窗口为 0 的窗口，即顶级窗口
// 16	匹配可见的窗口
// 32	匹配出的窗口按照窗口打开顺序依次排列
// 这些值可以相加,比如 4+8+16 就是类似于任务管理器中的窗口列表
//
// 返回值
//
// 类型：string
//
// 返回所有匹配到的窗口句柄
//
// 示例
//
// 无
func (o *Opsoft) EnumWindowByProcess(processName string, title string, className string, filter int64) (out []int64) {
	ret, _ := oleutil.CallMethod(o.op, "EnumWindowByProcess", processName, title, className, filter)

	sp := strings.Split(ret.ToString(), ",")
	for _, v := range sp {
		out = append(out, tool.StrToInt64(v))
	}
	return
}

// EnumProcess
// 根据指定进程名,枚举系统中符合条件的进程 PID
//
// string EnumProcess(name)
// 参数	类型	描述
// name	string	进程名称
// 返回值
//
// 类型：string
//
// 返回所有匹配的进程 PID,返回格式："10180,15352,15000,17620,19412"
//
// 示例
// ret=op.EnumProcess("qq.exe")
// println(ret)
func (o *Opsoft) EnumProcess(name string) (out []int64) {
	ret, _ := oleutil.CallMethod(o.op, "EnumProcess", name)

	sp := strings.Split(ret.ToString(), ",")
	for _, v := range sp {
		out = append(out, tool.StrToInt64(v))
	}

	return
}

//ClientToScreen
//把窗口坐标转换为屏幕坐标
//
//long ClientToScreen(hwnd,x,y)
//参数	类型	描述
//hwnd	int	指定的窗口句柄
//x	int*	变参指针: 接收窗口 X 坐标
//y	int*	变参指针: 接收窗口 Y 坐标
//返回值
//
//类型：int
//
//0：表示操作失败。
//1：表示操作成功。
//示例
//
//var x,y int
//ret = op.ClientToScreen(hwnd,x,y)
//println(x,y)

func (o *Opsoft) ClientToScreen(hwnd int64, intX *int64, intY *int64) bool {
	x := ole.NewVariant(ole.VT_CY, int64(*intX))
	y := ole.NewVariant(ole.VT_CY, int64(*intY))
	defer func() {
		x.Clear()
		y.Clear()
	}()
	ret := oleutil.MustCallMethod(o.op, "ClientToScreen", hwnd, &x, &y)
	*intX = (x.Val)
	*intY = (y.Val)
	return ret.Val == 1
}

//FindWindow
//查找符合类名或者标题名的顶层可见窗口
//
//long FindWindow(class,title)
//参数	类型	描述
//class	string	窗口类名,如果为空,则匹配所有,这里的匹配是模糊匹配
//title	string	窗口标题,如果为空,则匹配所有,这里的匹配是模糊匹配
//返回值
//
//类型：int
//
//返回窗口句柄,没找到则返回 0
//
//示例
//
//hwnd = op.FindWindow("","记事本")

func (o *Opsoft) FindWindow(class string, title string) (out int64) {
	ret, _ := oleutil.CallMethod(o.op, "FindWindow", class, title)
	out = ret.Val
	return
}

//FindWindowByProcess
//根据指定的进程名字，来查找可见窗口
//
//long FindWindowByProcess(process_name,class,title)
//参数	类型	描述
//process_name	string	进程名,比如(notepad.exe),这里是精确匹配,但不区分大小写
//class	string	窗口类名,如果为空,则匹配所有,这里的匹配是模糊匹配
//title	string	窗口标题,如果为空,则匹配所有,这里的匹配是模糊匹配
//返回值
//
//类型：int
//
//返回窗口句柄,没找到则返回 0
//
//示例
//
//hwnd = op.FindWindowByProcess("noteapd.exe","","记事本")

func (o *Opsoft) FindWindowByProcess(processName, class, title string) int64 {
	ret, _ := oleutil.CallMethod(o.op, "FindWindowByProcess", processName, class, title)
	return ret.Val
}

// FindWindowByProcessId
// 根据指定的进程 Id，来查找可见窗口
//
// long FindWindowByProcessId(process_id,class,title)
// 参数	类型	描述
// process_id	int	进程 id
// class	string	窗口类名,如果为空,则匹配所有,这里的匹配是模糊匹配
// title	string	窗口标题,如果为空,则匹配所有,这里的匹配是模糊匹配
// 返回值
//
// 类型：int
//
// 返回窗口句柄,没找到则返回 0
//
// 示例
//
// hwnd = op.FindWindowByProcessId(123456,"","记事本")
func (o *Opsoft) FindWindowByProcessId(processId int64, class, title string) int64 {
	ret, _ := oleutil.CallMethod(o.op, "FindWindowByProcessId", processId, class, title)
	return ret.Val
}

// FindWindowEx
// 查找符合类名或者标题名的顶层可见窗口,如果指定了 parent,则在 parent 的第一层子窗口中查找
//
// long FindWindowEx(parent,class,title)
// 参数	类型	描述
// parent	int	父窗口句柄，如果为空，则匹配所有顶层窗口
// class	string	窗口类名,如果为空,则匹配所有,这里的匹配是模糊匹配
// title	string	窗口标题,如果为空,则匹配所有,这里的匹配是模糊匹配
// 返回值
//
// 类型：int
//
// 返回窗口句柄,没找到则返回 0
//
// 示例
//
// hwnd = op.FindWindowEx(0,"","记事本")
func (o *Opsoft) FindWindowEx(parent int64, class, title string) int64 {
	ret, _ := oleutil.CallMethod(o.op, "FindWindowEx", parent, class, title)
	return ret.Val
}

// GetClientRect
// 获取窗口客户区域在屏幕上的位置
//
// long GetClientRect(hwnd,x1,y1,x2,y2)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// x1	int*	变参指针: 返回窗口客户区左上角 X 坐标
// y1	int*	变参指针: 返回窗口客户区左上角 Y 坐标
// x2	int*	变参指针: 返回窗口客户区右下角 X 坐标
// y2	int*	变参指针: 返回窗口客户区右下角 Y 坐标
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// ret = op.GetClientRect(hwnd,x1,y1,x2,y2)
func (o *Opsoft) GetClientRect(hwnd int64, x1, y1, x2, y2 *int64) bool {
	x1_ := ole.NewVariant(ole.VT_CY, int64(*x1))
	y1_ := ole.NewVariant(ole.VT_CY, int64(*y1))
	x2_ := ole.NewVariant(ole.VT_CY, int64(*x2))
	y2_ := ole.NewVariant(ole.VT_CY, int64(*y2))
	defer func() {
		x1_.Clear()
		y1_.Clear()
		x2_.Clear()
		y2_.Clear()
	}()
	ret := oleutil.MustCallMethod(o.op, "GetClientRect", hwnd, &x1_, &y1_, &x2_, &y2_)
	*x1 = x1_.Val
	*y1 = y1_.Val
	*x2 = x2_.Val
	*y2 = y2_.Val
	return ret.Val == 1
}

//GetClientSize
//获取窗口客户区域的宽度和高度
//
//long GetClientSize(hwnd,width,width)
//参数	类型	描述
//parent	int	指定的窗口句柄
//width	int*	变参指针: 窗口宽度
//width	int*	变参指针: 窗口高度
//返回值
//
//类型：int
//
//0: 失败
//1: 成功
//示例
//
//var w,h int
//op.GetClientSize(hwnd,&w,&h)
//printf("宽度:%d,高度:%d",w,h)

func (o *Opsoft) GetClientSize(hwnd int64, width, height *int64) bool {

	width_ := ole.NewVariant(ole.VT_CY, int64(*width))
	height_ := ole.NewVariant(ole.VT_CY, int64(*height))
	defer func() {
		width_.Clear()
		height_.Clear()
	}()
	ret := oleutil.MustCallMethod(o.op, "GetClientSize", hwnd, &width_, &height_)
	*width = width_.Val
	*height = height_.Val
	return ret.Val == 1

}

// GetForegroundFocus
// 获取顶层活动窗口中具有输入焦点的窗口句柄
//
// long GetForegroundFocus()
// 返回值
//
// 类型：int
//
// 返回窗口句柄
//
// 示例
//
// hwnd = op.GetForegroundFocus()
func (o *Opsoft) GetForegroundFocus() int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetForegroundFocus")
	return ret.Val
}

// GetForegroundWindow
// 获取顶层活动窗口,可以获取到按键自带插件无法获取到的句柄
//
// long GetForegroundWindow()
// 返回值
//
// 类型：int
//
// 返回窗口句柄
//
// 示例
//
// hwnd = op.GetForegroundWindow()
func (o *Opsoft) GetForegroundWindow() int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetForegroundWindow")
	return ret.Val
}

// GetMousePointWindow
// 获取鼠标指向的可见窗口句柄,可以获取到按键自带的插件无法获取到的句柄
//
// long GetMousePointWindow()
// 返回值
//
// 类型：int
//
// 返回窗口句柄
//
// 示例
//
// hwnd = op.GetMousePointWindow()
func (o *Opsoft) GetMousePointWindow() int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetMousePointWindow")
	return ret.Val

}

// GetPointWindow
// 获取给定坐标的可见窗口句柄
//
// long GetPointWindow(x,y)
// 参数	类型	描述
// x	int	屏幕 X 坐标
// y	int	屏幕 Y 坐标
// 返回值
//
// 类型：int
//
// 返回窗口句柄
//
// 示例
//
// hwnd = op.GetPointWindow(100,100)
func (o *Opsoft) GetPointWindow(x, y int64) int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetPointWindow", x, y)
	return ret.Val
}

type OpProcessInfo struct {
	Name   string
	Path   string
	CPU    string
	Memory string
}

// GetProcessInfo
// 根据指定的 pid 获取进程详细信息,(进程名,进程全路径,CPU 占用率(百分比),内存占用量(字节))
//
// string GetProcessInfo(pid)
// 参数	类型	描述
// pid	int	进程 pid
// 返回值
//
// 类型：string
//
// 返回格式"进程名|进程路径|cpu|内存"
//
// 示例
//
// infos = op.GetProcessInfo(12345)
// infos = split(infos,"|")
// TracePrint "进程名:"&infos(0)
// TracePrint "进程路径:"&infos(1)
// TracePrint "进程CPU占用率(百分比):"&infos(2)
// TracePrint "进程内存占用量(字节):"&infos(3)
func (o *Opsoft) GetProcessInfo(pid int64) (p OpProcessInfo) {
	ret, _ := oleutil.CallMethod(o.op, "GetProcessInfo", pid)

	sp := strings.Split(ret.ToString(), "|")
	p.Name = sp[0]
	p.Path = sp[1]
	p.CPU = sp[2]
	p.Memory = sp[3]
	return

}

// GetSpecialWindow
// 获取特殊窗口
//
// long GetSpecialWindow(flag)
// 参数	类型	描述
// flag	int	取值如下
// flag
//
// 值	描述
// 0	获取桌面窗口
// 1	获取任务栏窗口
// 返回值
//
// 类型：int
//
// 返回窗口句柄
//
// 示例
//
// desk_win = op.GetSpecialWindow(0)
func (o *Opsoft) GetSpecialWindow(flag int64) int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetSpecialWindow", flag)
	return ret.Val

}

// GetWindow
// 获取给定窗口相关的窗口句柄
//
// long GetWindow(hwnd,flag)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// flag	int	取值如下
// flag
//
// 值	描述
// 0	获取父窗口
// 1	获取第一个儿子窗口
// 2	获取 First 窗口
// 3	获取 Last 窗口
// 4	获取下一个窗口
// 5	获取上一个窗口
// 6	获取拥有者窗口
// 7	获取顶层窗口
// 返回值
//
// 类型：int
//
// 返回窗口句柄
//
// 示例
//
// hwnd = op.GetWindow(hwnd,6)
func (o *Opsoft) GetWindow(hwnd int64, flag int64) int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetWindow", hwnd, flag)
	return ret.Val
}

// GetWindowClass
// 获取窗口的类名
//
// string GetWindowClass(hwnd)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// 返回值
//
// 类型：string
//
// 窗口的类名
//
// 示例
//
// class_name = op.GetWindowClass(hwnd)
func (o *Opsoft) GetWindowClass(hwnd int64) string {
	ret, _ := oleutil.CallMethod(o.op, "GetWindowClass", hwnd)
	return ret.ToString()
}

// GetWindowProcessId
// 获取指定窗口所在的进程 ID
//
// long GetWindowProcessId(hwnd)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// 返回值
//
// 类型：int
//
// 返回进程 ID
//
// 示例
//
// process_id = op.GetWindowProcessId(hwnd)
func (o *Opsoft) GetWindowProcessId(hwnd int64) int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetWindowProcessId", hwnd)
	return ret.Val
}

// GetWindowProcessPath
// 获取指定窗口所在的进程的 exe 文件全路径
//
// string GetWindowProcessPath(hwnd)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// 返回值
//
// 类型：string
//
// 返回进程所在的全路径
//
// 示例
//
// process_path = op.GetWindowProcessPath(hwnd)
func (o *Opsoft) GetWindowProcessPath(hwnd int64) string {
	ret, _ := oleutil.CallMethod(o.op, "GetWindowProcessPath", hwnd)
	return ret.ToString()
}

// GetWindowRect
// 获取窗口在屏幕上的位置
//
// long GetWindowRect(hwnd,x1,y1,x2,y2)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// x1	int*	变参指针: 返回窗口左上角 X 坐标
// y1	int*	变参指针: 返回窗口左上角 Y 坐标
// x2	int*	变参指针: 返回窗口右下角 X 坐标
// y2	int*	变参指针: 返回窗口右下角 Y 坐标
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.GetWindowRect(hwnd,x1,y1,x2,y2)
func (o *Opsoft) GetWindowRect(hwnd int64, x1, y1, x2, y2 *int64) bool {
	x1_ := ole.NewVariant(ole.VT_CY, int64(*x1))
	y1_ := ole.NewVariant(ole.VT_CY, int64(*y1))
	x2_ := ole.NewVariant(ole.VT_CY, int64(*x2))
	y2_ := ole.NewVariant(ole.VT_CY, int64(*y2))
	defer func() {
		x1_.Clear()
		y1_.Clear()
		x2_.Clear()
		y2_.Clear()
	}()
	ret := oleutil.MustCallMethod(o.op, "GetWindowRect", hwnd, &x1_, &y1_, &x2_, &y2_)
	*x1 = x1_.Val
	*y1 = y1_.Val
	*x2 = x2_.Val
	*y2 = y2_.Val
	return ret.Val == 1
}

// GetWindowState
// 获取指定窗口的一些属性
//
// long GetWindowState(hwnd,flag)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// flag	int	取值如下
// flag
//
// 值	描述
// 0	判断窗口是否存在
// 1	判断窗口是否处于激活
// 2	判断窗口是否可见
// 3	判断窗口是否最小化
// 4	判断窗口是否最大化
// 5	判断窗口是否置顶
// 6	判断窗口是否无响应
// 7	判断窗口是否可用(灰色为不可用)
// 返回值
//
// 类型：int
//
// 0: 不满足条件
// 1: 满足条件
// 示例
//
// op_ret = op.GetWindowState(hwnd,3)
//
//	if op_ret ==1 {
//	 println("窗口已经最小化了")
//	}
func (o *Opsoft) GetWindowState(hwnd int64, flag int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "GetWindowState", hwnd, flag)
	return ret.Val == 1
}

// GetWindowTitle
// 获取窗口的标题
//
// string GetWindowTitle(hwnd)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// 返回值
//
// 类型：string
//
// 返回窗口的标题
//
// 示例
//
// title = op.GetWindowTitle(hwnd)
func (o *Opsoft) GetWindowTitle(hwnd int64) string {
	ret, _ := oleutil.CallMethod(o.op, "GetWindowTitle", hwnd)
	return ret.ToString()
}

// MoveWindow
// 移动指定窗口到指定位置
//
// long MoveWindow(hwnd,x,y)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// x	int	指定的 X 坐标
// y	int	指定的 Y 坐标
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op.MoveWindow(hwnd,100,200)
func (o *Opsoft) MoveWindow(hwnd int64, x, y int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "MoveWindow", hwnd, x, y)
	return ret.Val == 1
}

// ScreenToClient
// 把屏幕坐标转换为窗口坐标
//
// long ScreenToClient(hwnd,x,y)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// x	int*	变参指针: 屏幕 X 坐标
// y	int*	变参指针: 屏幕 Y 坐标
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// var x =100
// var y =100
// op_ret = op.ScreenToClient(hwnd,&x,&y)
func (o *Opsoft) ScreenToClient(hwnd int64, x, y *int64) bool {
	x_ := ole.NewVariant(ole.VT_CY, int64(*x))
	y_ := ole.NewVariant(ole.VT_CY, int64(*y))
	defer func() {
		x_.Clear()
		y_.Clear()
	}()
	ret := oleutil.MustCallMethod(o.op, "ScreenToClient", hwnd, &x_, &y_)
	*x = x_.Val
	*y = y_.Val
	return ret.Val == 1
}

// SetClientSize
// 设置窗口客户区域的宽度和高度
//
// long SetClientSize(hwnd,width,height)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// width	int	宽度
// height	int	高度
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.SetClientSize(hwnd,800,600)
func (o *Opsoft) SetClientSize(hwnd int64, width, height int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetClientSize", hwnd, width, height)
	return ret.Val == 1
}

// SetWindowState
// 设置窗口的状态
//
// long SetWindowState(hwnd,flag)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// flag	int	取值定义如下
// flag
//
// 值	描述
// 0	关闭指定窗口
// 1	激活指定窗口
// 2	最小化指定窗口,但不激活
// 3	最小化指定窗口,并释放内存,但同时也会激活窗口
// 4	最大化指定窗口,同时激活窗口
// 5	恢复指定窗口 ,但不激活
// 6	隐藏指定窗口
// 7	显示指定窗口
// 8	置顶指定窗口
// 9	取消置顶指定窗口
// 10	禁止指定窗口
// 11	取消禁止指定窗口
// 12	恢复并激活指定窗口
// 13	强制结束窗口所在进程
// 14	闪烁指定的窗口
// 15	使指定的窗口获取输入焦点
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.SetWindowState(hwnd,0)
func (o *Opsoft) SetWindowState(hwnd int64, flag int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetWindowState", hwnd, flag)
	return ret.Val == 1
}

// SetWindowSize
// 设置窗口客户区域的宽度和高度
//
// long SetWindowSize(hwnd,width,height)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// width	int	宽度
// height	int	高度
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.SetWindowSize(hwnd,300,400)
func (o *Opsoft) SetWindowSize(hwnd int64, width, height int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetWindowSize", hwnd, width, height)
	return ret.Val == 1
}

// SetWindowText
// 设置窗口的标题
//
// long SetWindowText(hwnd,title)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// title	string	标题
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.SetWindowText(hwnd,"test")
func (o *Opsoft) SetWindowText(hwnd int64, title string) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetWindowText", hwnd, title)
	return ret.Val == 1
}

// SetWindowTransparent
// 设置窗口的透明度
//
// long SetWindowTransparent(hwnd,trans)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// trans	int	透明度取值(0-255) 越小透明度越大 0 为完全透明(不可见) 255 为完全显示(不透明)
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.SetWindowTransparent(hwnd,200)
func (o *Opsoft) SetWindowTransparent(hwnd int64, trans int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetWindowTransparent", hwnd, trans)
	return ret.Val == 1
}

// SendString
// 向指定窗口发送文本数据
//
// long SendString(hwnd,str)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// str	string	发送的文本数据
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.SendString(hwnd,"我是来测试的")
func (o *Opsoft) SendString(hwnd int64, str string) bool {
	ret, _ := oleutil.CallMethod(o.op, "SendString", hwnd, str)
	return ret.Val == 1
}

// SendStringIme
// 向指定窗口发送文本数据-输入法
//
// long SendStringIme(hwnd,str)
// 参数	类型	描述
// hwnd	int	指定的窗口句柄
// str	string	发送的文本数据
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.SendStringIme(hwnd,"我是来测试的")
func (o *Opsoft) SendStringIme(hwnd int64, str string) bool {
	ret, _ := oleutil.CallMethod(o.op, "SendStringIme", hwnd, str)
	return ret.Val == 1
}
