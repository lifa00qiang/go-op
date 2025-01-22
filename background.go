package opsoft

import "github.com/go-ole/go-ole/oleutil"

//接口目录
//BindWindow: 绑定窗口
//UnBindWindow: 解除绑定窗口
//IsBind: 判断当前对象是否已绑定窗口
//GetBindWindow: 获取当前对象已经绑定的窗口句柄
//接口方法

const (
	Normal     string = "normal"      // 正常模式
	NormalDxgi string = "normal.dxgi" // dxgi 截图模式，这个速度更快，更省 CPU
	NormalWgc  string = "normal.wgc"  // wgc 截图模式，仅支持 win11 以上版本
	Gdi        string = "gdi"         // gdi 模式,用于窗口采用 GDI 方式刷新时,此模式占用 CPU 较大
	Gdi2       string = "gdi2"        // gdi2 模式,此模式兼容性较强,但是速度比 gdi 模式要慢许多
	Dx         string = "dx"          // dx 模式,等同于 dx.d3d9
	Dx2        string = "dx2"         // 等同于 dx2.d3d9
	DxD3D9     string = "dx.d3d9"     // d3d9 模式,使用 d3d9 渲染
	DxD3D10    string = "dx.d3d10"    // d3d10 模式,使用 d3d10 渲染
	DxD3D11    string = "dx.d3d11"    // d3d11 模式,使用 d3d11 渲染
	DxD3D12    string = "dx.d3d12"    // d3d12 模式,使用 d3d12 渲染
	OpenGL     string = "opengl"      // opengl 模式，使用 opengl 渲染的窗口
	OpenGLStd  string = "opengl.std"  //	测试中
	OpenGLNox  string = "opengl.nox"  // opengl 模式，针对最新夜神模拟器的渲染方式，测试中...
	OpenGLEs   string = "opengl.es"   // 测试中...
	OpenGLFi   string = "opengl.fi"   // 测试中...
)

const (
	NormalMouse   string = "normal"    // 正常模式
	NormalMouseHd string = "normal.hd" // hd 模式
	WindowsMouse  string = "windows"   // windows 模式
	DxMouse       string = "dx"        // dx 模式
)

const (
	NormalKeypad   string = "normal"    // 正常模式
	NormalKeypadHd string = "normal.hd" // hd 模式
	WindowsKeypad  string = "windows"   // windows 模式
	DxKeypad       string = "dx"        // dx 模式
)

// BindWindow
// 绑定指定的窗口,并指定这个窗口的屏幕颜色获取方式,鼠标仿真模式,键盘仿真模式,以及模式设定.
//
// long BindWindow(hwnd,display,mouse,keypad,mode)
// 参数		类型		描述
// hwnd		int		指定的窗口句柄
// display	string	屏幕显示模式,取值定义如下
// mouse		string	鼠标仿真模式,取值定义如下
// keypad	string	键盘仿真模式,取值定义如下
// mode		int		模式,取值 0、1
//
// display
// 值			描述
// normal		正常模式,平常我们用的前台截屏模式
// normal.dxgi	dxgi 截图模式，这个速度更快，更省 CPU
// normal.wgc	wgc 截图模式，仅支持 win11 以上版本
// gdi	gdi 	模式,用于窗口采用 GDI 方式刷新时,此模式占用 CPU 较大
// gdi2	gdi2	模式,此模式兼容性较强,但是速度比 gdi 模式要慢许多
// dx	dx 		模式,等同于 dx.d3d9
// dx2	dx2		模式,用于窗口采用 dx 模式刷新
// dx.d3d9		d3d9 模式,使用 d3d9 渲染
// dx.d3d10		d3d10 模式,使用 d3d10 渲染
// dx.d3d11		d3d11 模式,使用 d3d11 渲染
// dx.d3d12		d3d12 模式,使用 d3d12 渲染
// opengl		opengl 模式，使用 opengl 渲染的窗口
// opengl.std	测试中
// opengl.nox	opengl 模式，针对最新夜神模拟器的渲染方式，测试中...
// opengl.es		测试中...
// opengl.fi		测试中...
//
// mouse
// 值		描述
// normal	正常模式,平常我们用的前台鼠标模式
// normal.hd
// windows	Windows 模式,采取模拟 windows 消息方式
// dx		dx 模式
//
// keypad
// 值		描述
// normal	正常模式,平常我们用的前台键盘模式
// normal.hd
// windows	Windows 模式,采取模拟 windows 消息方式
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// // display: 前台 鼠标:前台键盘:前台 模式0
// op_ret = op.BindWindow(hwnd,"normal","normal","normal",0)
//
// // display: gdi 鼠标:前台 键盘:前台 模式1
// op_ret = op.BindWindow(hwnd,"gdi","normal","normal",1)
func (o *Opsoft) BindWindow(hwnd int64, display string, mouse string, keypad string, mode int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "BindWindow", hwnd, display, mouse, keypad, mode)
	return ret.Val == 1
}

// UnBindWindow
// 解除绑定窗口,并释放系统资源
//
// long UnBindWindow()
// 返回值
//
// 类型：int
//
// 0: 失败
// 1: 成功
// 示例
//
// op_ret = op.UnBindWindow()
func (o *Opsoft) UnBindWindow() bool {
	ret, _ := oleutil.CallMethod(o.op, "UnBindWindow")
	return ret.Val == 1
}

// GetBindWindow
// 获取当前对象已经绑定的窗口句柄. 无绑定返回:0
//
// long GetBindWindow()
// 返回值
//
// 类型：int
//
// 0: 没有绑定窗口
// 示例
//
// op_ret = op.GetBindWindow()
func (o *Opsoft) GetBindWindow() int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetBindWindow")
	return ret.Val

}

// IsBind
// 该函数旨在判断当前对象是否已绑定窗口
//
// long IsBind()
// 返回值
//
// 类型：int
//
// 0: 表示未绑定状态
// 1: 表示已绑定状态
// 示例
//
// op_ret = op.IsBind()
func (o *Opsoft) IsBind() bool {
	ret, _ := oleutil.CallMethod(o.op, "IsBind")
	return ret.Val == 1

}
