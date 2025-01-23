package opsoft

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

//接口目录
//GetCursorPos: 获取鼠标位置
//MoveR: 鼠标相对于上次坐标移动
//MoveTo: 鼠标移动到指定坐标
//MoveToEx: 鼠标移动到指定随机坐标
//LeftClick: 按下鼠标左键
//LeftDoubleClick: 双击鼠标左键
//LeftDown: 按住鼠标左键
//LeftUp: 弹起鼠标左键
//MiddleClick: 按下鼠标中键
//MiddleDown: 按住鼠标中键
//MiddleUp: 弹起鼠标中键
//RightClick: 按下鼠标右键
//RightDown: 按住鼠标右键
//RightUp: 弹起鼠标右键
//WheelDown: 滚轮向下滚
//WheelUp: 滚轮向上滚
//SetMouseDelay: 设置鼠标单击或者双击时,鼠标按下和弹起的时间间隔

//接口方法

// GetCursorPos
// 获取鼠标位置
//
// long GetCursorPos(x,y)
// 参数	类型	描述
// x	int*	变参指针: 返回 X 坐标
// y	int*	变参指针: 返回 Y 坐标
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// var x, y int
// op.GetCursorPos(&x,&y)
// printf("x:%d,y:%d",x,y)
func (o *Opsoft) GetCursorPos(x, y *int64) bool {
	x_ := ole.NewVariant(ole.VT_I4, *x)
	y_ := ole.NewVariant(ole.VT_I4, *y)
	defer func() {
		x_.Clear()
		y_.Clear()
	}()
	ret, err := oleutil.CallMethod(o.op, "GetCursorPos", &x_, &y_)
	if err != nil {
		fmt.Println("", err.Error())
	}

	*x = x_.Val
	*y = y_.Val
	return ret.Value() == 0

}

// MoveR
// 鼠标相对于上次的位置移动 rx,ry.
//
// long MoveR(rx,ry)
// 参数	类型	描述
// rx	int	相对于上次的 X 偏移
// ry	int	相对于上次的 Y 偏移
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// op.MoveR(rx,ry)
func (o *Opsoft) MoveR(rx, ry int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "MoveR", &rx, &ry)
	return ret.Value() == 1
}

// MoveTo
// 把鼠标移动到目的点(x,y)
//
// long MoveTo(x,y)
// 参数	类型	描述
// x	int	X 坐标
// y	int	Y 坐标
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// op.MoveTo(x,y)
func (o *Opsoft) MoveTo(x, y int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "MoveTo", &x, &y)
	return ret.Value() == 1
}

// MoveToEx
// 把鼠标移动到目的范围内的任意一点
//
// string MoveToEx(x,y,w,h)
// 参数	类型	描述
// x	int	X 坐标
// y	int	Y 坐标
// w	int	宽度(从 x 计算起)
// h	int	高度(从 y 计算起)
// 返回值
//
// 类型：string
//
// 返回要移动到的目标点. 格式为 x,y. 比如 MoveToEx 100,100,10,10,返回值可能是 101,102
//
// 示例
//
// // 移动鼠标到(100,100)到(110,110)这个矩形范围内的任意一点.
// op.MoveToEx(100,100,10,10)
// 注: 此函数的意思是移动鼠标到指定的范围(x,y,x+w,y+h)内的任意随机一点
func (o *Opsoft) MoveToEx(x, y, w, h int64) string {
	ret, _ := oleutil.CallMethod(o.op, "MoveToEx", &x, &y, &w, &h)
	return ret.ToString()
}

//LeftClick
//按下鼠标左键
//
//long LeftClick()
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.LeftClick()

func (o *Opsoft) LeftClick() bool {
	ret, _ := oleutil.CallMethod(o.op, "LeftClick")
	return ret.Value() == 1
}

// LeftDoubleClick
// 双击鼠标左键
//
// long LeftDoubleClick()
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// op.LeftDoubleClick()
func (o *Opsoft) LeftDoubleClick() bool {
	ret, _ := oleutil.CallMethod(o.op, "LeftDoubleClick")
	return ret.Value() == 1
}

// LeftDown
// 按住鼠标左键
//
// long LeftDown()
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// op.LeftDown()
func (o *Opsoft) LeftDown() bool {
	ret, _ := oleutil.CallMethod(o.op, "LeftDown")
	return ret.Value() == 1
}

// LeftUp
// 弹起鼠标左键
//
// long LeftUp()
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// op.LeftUp()
func (o *Opsoft) LeftUp() bool {
	ret, _ := oleutil.CallMethod(o.op, "LeftUp")
	return ret.Value() == 1
}

// MiddleClick
// 按下鼠标中键
//
// long MiddleClick()
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// op.MiddleClick()
func (o *Opsoft) MiddleClick() bool {
	ret, _ := oleutil.CallMethod(o.op, "MiddleClick")
	return ret.Value() == 1
}

//MiddleDown
//按住鼠标中键
//
//long MiddleDown()
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.MiddleDown()

func (o *Opsoft) MiddleDown() bool {
	ret, _ := oleutil.CallMethod(o.op, "MiddleDown")
	return ret.Value() == 1
}

//MiddleUp
//弹起鼠标中键
//
//long MiddleUp()
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.MiddleUp()

func (o *Opsoft) MiddleUp() bool {
	ret, _ := oleutil.CallMethod(o.op, "MiddleUp")
	return ret.Value() == 1
}

//RightClick
//按下鼠标右键
//
//long RightClick()
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.RightClick()

func (o *Opsoft) RightClick() bool {
	ret, _ := oleutil.CallMethod(o.op, "RightClick")
	return ret.Value() == 1
}

//RightDown
//按住鼠标右键
//
//long RightDown()
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.RightDown()

func (o *Opsoft) RightDown() bool {
	ret, _ := oleutil.CallMethod(o.op, "RightDown")
	return ret.Value() == 1
}

//RightUp
//弹起鼠标右键
//
//long RightUp()
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.RightUp()

func (o *Opsoft) RightUp() bool {
	ret, _ := oleutil.CallMethod(o.op, "RightUp")
	return ret.Value() == 1
}

//WheelDown
//滚轮向下滚
//
//long WheelDown()
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.WheelDown()

func (o *Opsoft) WheelDown() bool {
	ret, _ := oleutil.CallMethod(o.op, "WheelDown")
	return ret.Value() == 1
}

//WheelUp
//滚轮向下滚
//
//long WheelUp()
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op.WheelUp()

func (o *Opsoft) WheelUp() bool {
	ret, _ := oleutil.CallMethod(o.op, "WheelUp")
	return ret.Value() == 1
}

// SetMouseDelay
// 该函数旨在设置鼠标单击或双击时，鼠标按下和弹起之间的时间间隔。
//
// long SetMouseDelay(type,delay)
// 参数	类型	描述
// type	string	鼠标类型，取值： "normal" | "windows" | "dx"
// delay	int	指定鼠标按下和弹起之间的时间间隔，单位通常为毫秒（milliseconds）
// 当取值为"normal",默认为: 30ms
//
// 当取值为"windows",默认为: 10ms
//
// 当取值为"dx",默认为: 40ms
//
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// op.SetMouseDelay("dx",1000)
func (o *Opsoft) SetMouseDelay(t string, delay int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetMouseDelay", t, delay)
	return ret.Value() == 1
}
