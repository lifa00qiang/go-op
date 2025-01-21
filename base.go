package opsoft

import "github.com/go-ole/go-ole/oleutil"

//接口目录
//Ver: 插件版本号
//SetPath: 设置全局路径
//GetPath: 获取全局路径
//GetBasePath: 获取插件所在目录
//GetID: 获取对象 ID
//GetLastError: 获取最后的错误
//SetShowErrorMsg: 设置是否弹出错误信息
//Sleep: 休眠
//InjectDll: 注入 DLL
//EnablePicCache: 是否启用图片缓存机制
//CapturePre: 取上次操作的图色区域
//SetScreenDataMode: 设置屏幕数据模式

// 接口方法
// Ver
// 获取当前 op 插件的版本号
//
// string Ver()
// 返回值
//
// 类型：string
//
// 返回 op 插件的版本号
func (o *Opsoft) Ver() string {
	ret, _ := oleutil.CallMethod(o.op, "Ver")
	return ret.ToString()
}

// SetPath
// 设置全局路径,设置了此路径后,所有接口调用中,相关的文件都相对于此路径. 比如图片,字库等.
//
// long SetPath(path)
// 参数	类型	描述
// path	string	指定的路径。
// 返回值
//
// 类型：int
//
// 0：表示操作失败
// 1：表示操作成功
// 示例
//
// 无
func (s *Opsoft) SetPath(path string) bool {
	ret, _ := oleutil.CallMethod(s.op, "SetPath", path)
	return ret.Val == 1
}

// GetPath
// 获取全局路径
//
// string GetPath()
// 返回值
//
// 类型：string
//
// 返回当前设置的全局路径
//
// 示例
//
// 无
func (o *Opsoft) GetPath() string {
	ret, _ := oleutil.CallMethod(o.op, "GetPath")
	return ret.ToString()
}

// GetBasePath
// 获取插件目录
//
// string GetBasePath()
// 返回值
//
// 类型：string
//
// 返回当前插件所在路径
//
// 示例
//
// 无
func (o *Opsoft) GetBasePath() string {
	ret, _ := oleutil.CallMethod(o.op, "GetBasePath")
	return ret.ToString()
}

// GetID
// 返回当前对象的 ID 值，这个值对于每个对象是唯一存在的。可以用来判定两个对象是否一致
//
// long GetID()
// 返回值
//
// 类型：int
//
// 当前对象的 ID 值.
//
// 示例
//
// 无
func (o *Opsoft) GetID() int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetID")
	return ret.Val
}

// GetLastError
// 获取最后的错误
//
// long GetLastError()
// 返回值
//
// 类型：int
//
// 0: 表示无错误
// 示例
//
// 无
func (o *Opsoft) GetLastError() int64 {
	ret, _ := oleutil.CallMethod(o.op, "GetLastError")
	return ret.Val
}

// SetShowErrorMsg
// 设置是否弹出错误信息,默认是打开
//
// long SetShowErrorMsg(show)
// 参数	类型	描述
// show	int	0:关闭，1:显示为信息框，2:保存到文件,3:输出到标准输出
// 返回值
//
// 类型：int
//
// 0：表示操作失败
// 1：表示操作成功
// 示例
//
// 无
func (o *Opsoft) SetShowErrorMsg(show int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetShowErrorMsg", show)
	return ret.Val == 1
}

// Sleep
// 设置休眠时间
//
// long Sleep(millisecond)
// 参数	类型	描述
// millisecond	int	休眠时间(毫秒)
// 返回值
//
// 类型：int
//
// 0：表示操作失败
// 1：表示操作成功
// 示例
//
// 无
func (s *Opsoft) Sleep(millisecond int64) bool {
	ret, _ := oleutil.CallMethod(s.op, "Sleep", millisecond)
	return ret.Val == 1
}

// InjectDll
// 将指定的 DLL 注入到指定的进程中
//
// void InjectDll(const wchar_t* process_name,const wchar_t* dll_name, long* ret);
// 参数	类型	描述
// process_name	string	指定要注入 DLL 的进程名称
// dll_name	string	注入的 DLL 名称
// 返回值
//
// 类型：int
//
// 0：表示操作失败
// 1：表示操作成功
// 示例
//
// 无
func (o *Opsoft) InjectDll(processName string, dllName string) bool {
	ret, _ := oleutil.CallMethod(o.op, "InjectDll", processName, dllName)
	return ret.Val == 1
}

// EnablePicCache
// 设置是否开启或者关闭插件内部的图片缓存机制
//
// long EnablePicCache(enable)
// 参数	类型	描述
// enable	int	0：关闭,1:打开
// 返回值
//
// 类型：int
//
// 0：表示操作失败
// 1：表示操作成功
// 示例
//
// 无
func (o *Opsoft) EnablePicCache(enable int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "EnablePicCache", enable)
	return ret.Val == 1
}

// CapturePre
// 取上次操作的图色区域，保存为 file(24 位位图)
//
// long CapturePre(file)
// 参数	类型	描述
// file	string	设置保存文件名,保存路径是 SetPath 设置的目录，也可以指定全路径
// 返回值
//
// 类型：int
//
// 0：表示操作失败
// 1：表示操作成功
// 示例
//
// op.CapturePre("screen.bmp")
func (o *Opsoft) CapturePre(filePath string) bool {
	ret, _ := oleutil.CallMethod(o.op, "CapturePre", filePath)
	return ret.Val == 1
}

// SetScreenDataMode
// 设置屏幕数据模式
//
// long SetScreenDataMode(mode)
// 参数	类型	描述
// mode	int	0:从上到下(默认),1:从下到上
// 返回值
//
// 类型：int
//
// 0：表示操作失败
// 1：表示操作成功
// 示例
//
// 无
func (o *Opsoft) SetScreenDataMode(mode int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetScreenDataMode", mode)
	return ret.Val == 1
}
