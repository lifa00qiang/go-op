package opsoft

import "github.com/go-ole/go-ole/oleutil"

//操作键盘按键
//
//接口目录
//GetKeyState: 获取指定的按键状态
//KeyDown: 按住指定的虚拟键码
//KeyDownChar: 按住指定的键码名称
//KeyUp: 弹起来虚拟键
//KeyUpChar: 弹起来按键名称
//WaitKey: 等待指定的按键按下
//KeyPress: 按住指定的虚拟键码
//KeyPressChar: 按住指定的按键名称
//SetKeypadDelay: 设置按键时,键盘按下和弹起的时间间隔

//接口方法

// GetKeyState
// 获取指定的按键状态.(前台信息,不是后台)
//
// long GetKeyState(vk_code)
// 参数	类型	描述
// vk_code	int	虚拟按键码
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// op.GetKeyState(13)
func (o *Opsoft) GetKeyState(vkCode int) bool {
	ret, _ := oleutil.CallMethod(o.op, "GetKeyState", vkCode)
	return ret.Value() == 1
}

//KeyDown
//按住指定的虚拟键码
//
//long KeyDown(vk_code)
//参数	类型	描述
//vk_code	int	虚拟按键码
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.KeyDown(13)

func (o *Opsoft) KeyDown(vkCode int) bool {
	ret, _ := oleutil.CallMethod(o.op, "KeyDown", vkCode)
	return ret.Value() == 1
}

//KeyDownChar
//按住指定的虚拟键码,字符串形式
//
//long KeyDownChar(key_str)
//参数	类型	描述
//key_str	string	字符串描述的键码,大小写无所谓，按键具体对应关系按键码
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.KeyDownChar("enter")
//op.KeyDownChar("1")
//op.KeyDownChar("F1")
//op.KeyDownChar("a")
//op.KeyDownChar("B")

func (o *Opsoft) KeyDownChar(keyStr string) bool {
	ret, _ := oleutil.CallMethod(o.op, "KeyDownChar", keyStr)
	return ret.Value() == 1
}

//KeyUp
//弹起来虚拟键 vk_code
//
//long KeyUp(vk_code)
//参数	类型	描述
//vk_code	int	虚拟按键码
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.KeyUp(13)

func (o *Opsoft) KeyUp(vkCode int) bool {
	ret, _ := oleutil.CallMethod(o.op, "KeyUp", vkCode)
	return ret.Value() == 1
}

//KeyUpChar
//弹起来虚拟键,字符串形式
//
//long KeyUpChar(key_str)
//参数	类型	描述
//key_str	string	字符串描述的键码,大小写无所谓，按键具体对应关系查看按键码
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.KeyUpChar("enter")
//op.KeyUpChar("1")
//op.KeyUpChar("F1")
//op.KeyUpChar("a")
//op.KeyUpChar("B")

func (o *Opsoft) KeyUpChar(keyStr string) bool {
	ret, _ := oleutil.CallMethod(o.op, "KeyUpChar", keyStr)
	return ret.Value() == 1
}

//WaitKey
//等待指定的按键按下 (前台,不是后台)
//
//long WaitKey(vk_code,time_out)
//参数	类型	描述
//vk_code	int	虚拟按键码,当此值为：0，表示等待任意按键。 鼠标左键是：1,鼠标右键时：2,鼠标中键是：4
//time_out	string	等待多久,单位毫秒. 如果是 0，表示一直等待
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.WaitKey(66,0)

func (o *Opsoft) WaitKey(vkCode, timeOut int) bool {
	ret, _ := oleutil.CallMethod(o.op, "WaitKey", vkCode, timeOut)
	return ret.Value() == 1
}

//KeyPress
//按住指定的虚拟键码
//
//long KeyPress(vk_code)
//参数	类型	描述
//vk_code	int	虚拟按键码
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.KeyPress(13)

func (o *Opsoft) KeyPress(vkCode int) bool {
	ret, _ := oleutil.CallMethod(o.op, "KeyPress", vkCode)
	return ret.Value() == 1
}

//KeyPressChar
//按住指定的虚拟键码,字符串形式
//
//long KeyPressChar(key_str)
//参数	类型	描述
//key_str	string	字符串描述的键码,大小写无所谓，按键具体对应关系查看按键码
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.KeyPressChar("enter")
//op.KeyPressChar("1")
//op.KeyPressChar("F1")
//op.KeyPressChar("a")
//op.KeyPressChar("B")

func (o *Opsoft) KeyPressChar(keyStr string) bool {
	ret, _ := oleutil.CallMethod(o.op, "KeyPressChar", keyStr)
	return ret.Value() == 1
}

//SetKeypadDelay
//该函数旨在设置按键时，键盘按下和弹起之间的时间间隔。
//
//long SetKeypadDelay(type,delay)
//参数	类型	描述
//type	string	键盘类型，取值： "normal" | "normal2" | "windows" | "dx"
//delay	int	指定键盘按下和弹起之间的时间间隔，单位通常为毫秒（milliseconds）
//当取值为"normal",默认为: 30ms
//
//当取值为"normal2",默认为: 30ms
//
//当取值为"windows",默认为: 10ms
//
//当取值为"dx",默认为: 50ms
//
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.SetKeypadDelay("dx",1000)

func (o *Opsoft) SetKeypadDelay(keypadType string, delay int) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetKeypadDelay", keypadType, delay)
	return ret.Value() == 1
}
