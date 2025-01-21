package opsoft

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

//Image Proc
//图片处理
//
//接口目录
//Capture: 抓取指定区域图像
//CmpColor: 比较指定坐标点的颜色
//FindColor: 查找指定区域内的颜色
//FindColorEx: 查找指定区域内的所有颜色
//FindMultiColor: 根据指定的多点颜色查找颜色坐标
//FindMultiColorEx: 根据指定的多点颜色查找所有颜色坐标
//FindPic: 查找指定区域内的图片
//FindPicEx: 查找多个图片
//FindPicExS: 查找多个图片同 FindPicEx
//FindColorBlock: 查找指定区域内的颜色块
//FindColorBlockEx: 查找指定区域内的所有颜色块
//GetColor: 获取指定坐标的颜色
//SetDisplayInput: 设置图像输入方式
//LoadPic: 预加载图片
//FreePic: 释放图片
//GetScreenData: 获取指定区域的图像数据
//GetScreenDataBmp: 获取指定区域的图像
//GetScreenFrameInfo: 获取屏幕帧信息，未导出 COM
//MatchPicName: 使用通配符并获取文件集合

//接口方法

//Capture
//抓取指定区域(x1, y1, x2, y2)的图像, 保存为文件
//
//long Capture(x1, y1, x2, y2, file)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//file	string	文件名,保存在 SetPath 中设置的目录，也可以自定义路径
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op_ret = op.Capture(0,0,2000,2000,"screen.bmp")

func (o *Opsoft) Capture(x1 int64, y1 int64, x2 int64, y2 int64, file string) bool {
	ret, _ := oleutil.CallMethod(o.op, "Capture", x1, y1, x2, y2, file)
	return ret.Value() == 1
}

//CmpColor
//比较指定坐标点(x,y)的颜色
//
//long CmpColor(x,y,color,sim)
//参数	类型	描述
//x	int	X 坐标
//y	int	Y 坐标
//color	int	颜色字符串，例如"ffffff-202020|000000-000000"，，每种颜色用"|"分割，最多 10 种
//sim	double	相似度,取值范围 0.1-1.0 标
//返回值
//
//类型：int
//
//0：颜色不匹配
//1：颜色匹配
//示例
//
//op = op.CmpColor(200,300,"000000-000000|ff00ff-101010",0.9)
//If op = 0 Then
//    MessageBox "相等"
//End If

func (o *Opsoft) CmpColor(x int64, y int64, color string, sim float64) bool {
	ret, _ := oleutil.CallMethod(o.op, "CmpColor", x, y, color, sim)
	return ret.Value() == 1
}

//FindColor
//查找指定区域内的颜色
//
//long FindColor(x1, y1, x2, y2, color, sim, dir,intX,intY)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//color	string	颜色格式串，比如"FFFFFF-000000|CCCCCC-000000"每种颜色用"|"分割
//sim	double	相似度,取值范围 0.1-1.0
//dir	int	查找方向,取值如下
//intX	int*	变参指针: 返回 X 坐标
//intY	int*	变参指针: 返回 Y 坐标
//返回值
//
//类型：int
//
//0：未找到
//1：成功找到
//示例
//
//op = op.FindColor(0,0,2000,2000,"123456-000000|aabbcc-030303|ddeeff-202020",1.0,0,intX,intY)
//If intX >= 0 and intY >= 0 Then
//    MessageBox "找到"
//End If

func (o *Opsoft) FindColor(x1 int64, y1 int64, x2 int64, y2 int64, color string, sim float64, dir int64, intX *int64, intY *int64) bool {
	x_ := ole.NewVariant(ole.VT_I4, *intX)
	y_ := ole.NewVariant(ole.VT_I4, *intY)
	defer func() {
		x_.Clear()
		y_.Clear()
	}()
	ret, _ := oleutil.CallMethod(o.op, "FindColor", x1, y1, x2, y2, color, sim, dir, &x_, &y_)

	intX = &x_.Val
	intY = &y_.Val
	return ret.Value() == 1
}

//FindColorEx
//查找指定区域内的所有颜色
//
//string FindColorEx(x1, y1, x2, y2, color, sim, dir)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//color	string	颜色格式串，比如"FFFFFF-000000|CCCCCC-000000"每种颜色用"|"分割
//sim	double	相似度,取值范围 0.1-1.0
//dir	int	查找方向,取值如下
//返回值
//
//类型：string
//
//返回所有颜色信息的坐标值
//
//示例
//
//s = op.FindColorEx(0,0,2000,2000,"123456-000000|abcdef-202020",1.0,0)
//count = op.GetResultCount(s)
//index = 0
//Do While index < count
//    op = op.GetResultPos(s,index,intX,intY)
//    MessageBox intX&","&intY
//    index = index + 1
//Loop

func (o *Opsoft) FindColorEx(x1 int64, y1 int64, x2 int64, y2 int64, color string, sim float64, dir int64) string {
	ret, _ := oleutil.CallMethod(o.op, "FindColorEx", x1, y1, x2, y2, color, sim, dir)
	return ret.ToString()
}

//FindMultiColor
//根据指定的多点查找颜色坐标
//
//long FindMultiColor(x1, y1, x2, y2,first_color,offset_color,sim, dir,intX,intY)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//first_color	string	颜色格式串，比如"FFFFFF-000000|CCCCCC-000000"每种颜色用"|"分割
//offset_color	string	偏移颜色可以支持任意多个点,格式为"x1|y1|RRGGBB-DRDGDB|RRGGBB-DRDGDB……
//sim	double	相似度,取值范围 0.1-1.0
//dir	int	查找方向,取值如下Dir
//intX	int*	变参指针: 返回 X 坐标, 坐标为 first_color 所在坐标
//intY	int*	变参指针: 返回 Y 坐标, 坐标为 first_color 所在坐标
//返回值
//
//类型：int
//
//0：未找到
//1：成功找到
//示例
//
//op = op.FindMultiColor(0,0,2000,2000,"cc805b-020202|606060-010101","9|2|-00ff00|-ff0000,15|2|2dff1c-010101,6|11|a0d962|aabbcc,11|14|-ffffff",1.0,1,intX,intY)
//op.MoveTo intX,intY

func (o *Opsoft) FindMultiColor(x1 int64, y1 int64, x2 int64, y2 int64, firstColor string, offsetColor string, sim float64, dir int64, intX *int64, intY *int64) bool {
	x_ := ole.NewVariant(ole.VT_I4, *intX)
	y_ := ole.NewVariant(ole.VT_I4, *intY)
	defer func() {
		x_.Clear()
		y_.Clear()
	}()
	ret, _ := oleutil.CallMethod(o.op, "FindMultiColor", x1, y1, x2, y2, firstColor, offsetColor, sim, dir, &x_, &y_)
	intX = &x_.Val
	intY = &y_.Val
	return ret.Value() == 1
}

//FindMultiColorEx
//根据指定的多点查找所有颜色坐标
//
//string FindMultiColorEx(x1, y1, x2, y2,first_color,offset_color,sim, dir)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//first_color	string	颜色格式串，比如"FFFFFF-000000|CCCCCC-000000"每种颜色用"|"分割
//offset_color	string	偏移颜色可以支持任意多个点,格式为"x1|y1|RRGGBB-DRDGDB|RRGGBB-DRDGDB……
//sim	double	相似度,取值范围 0.1-1.0
//dir	int	查找方向,取值如下Dir
//返回值
//
//类型：string
//
//返回所有颜色信息的坐标,坐标是 first_color 所在的坐标
//
//示例
//
//op = op.FindMultiColorEx(0,0,2000,2000, "cc805b-020202|606060-010101","9|2|-00ff00|-ff0000,15|2|2dff1c-010101,6|11|a0d962|aabbcc,11|14|-ffffff",1.0,1)
//count = op.GetResultCount(op)
//index = 0
//Do While index < count
//   aa = op.GetResultPos(op,index,intX,intY)
//   op.MoveTo intX,intY
//   index = index + 1
//   Delay  1000
//Loop

func (o *Opsoft) FindMultiColorEx(x1 int64, y1 int64, x2 int64, y2 int64, firstColor string, offsetColor string, sim float64, dir int64) string {
	ret, _ := oleutil.CallMethod(o.op, "FindMultiColorEx", x1, y1, x2, y2, firstColor, offsetColor, sim, dir)
	return ret.ToString()
}

//FindPic
//查找指定区域内的图片
//
//long FindPic(x1, y1, x2, y2, pic_name, delta_color,sim, dir,intX, intY)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//pic_name	string	图片名,可以是多个图片,比如"test1.bmp|test2.bmp|test3.bmp"
//delta_color	string	颜色色偏,比如"203040"
//sim	double	相似度,取值范围 0.1-1.0
//dir	int	查找方向,取值如下Dir
//intX	int*	变参指针: 返回图片左上角的 X 坐标
//intY	int*	变参指针: 返回图片左上角的 Y 坐标
//返回值
//
//类型：int
//
//返回找到的图片的序号,从 0 开始索引.如果没找到返回-1
//
//示例
//
//op = op.FindPic(0,0,2000,2000,"1.bmp|2.bmp|3.bmp","000000",0.9,0,intX,intY)
//If intX >= 0 and intY >= 0 Then
//    MessageBox "找到"
//End If

func (o *Opsoft) FindPic(x1 int64, y1 int64, x2 int64, y2 int64, picName string, deltaColor string, sim float64, dir int64, intX *int64, intY *int64) bool {
	x_ := ole.NewVariant(ole.VT_I4, *intX)
	y_ := ole.NewVariant(ole.VT_I4, *intY)
	defer func() {
		x_.Clear()
		y_.Clear()
	}()
	ret, _ := oleutil.CallMethod(o.op, "FindPic", x1, y1, x2, y2, picName, deltaColor, sim, dir, &x_, &y_)
	intX = &x_.Val
	intY = &y_.Val
	return ret.Value() == 1
}

//FindPicEx
//查找多个图片,并且返回所有找到的图像的坐标
//
//string FindPicEx(x1, y1, x2, y2, pic_name, delta_color,sim, dir)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//pic_name	string	图片名,可以是多个图片,比如"test1.bmp|test2.bmp|test3.bmp"
//delta_color	string	颜色色偏,比如"203040"
//sim	double	相似度,取值范围 0.1-1.0
//dir	int	查找方向,取值如下Dir
//返回值
//
//类型：string
//
//返回的是所有找到的坐标格式如下:"id,x,y|id,x,y..|id,x,y";id 对应图片序号，x,y 图片左上角的坐标
//
//示例
//
//op = op.FindPicEx(0,0,2000,2000,"test.bmp|test2.bmp|test3.bmp|test4.bmp|test5.bmp","020202",1.0,0)
//If len(op) > 0 Then
//   ss = split(op,"|")
//   index = 0
//   count = UBound(ss) + 1
//   Do While index < count
//      TracePrint ss(index)
//      sss = split(ss(index),",")
//      id = int(sss(0))
//      x = int(sss(1))
//      y = int(sss(2))
//      op.MoveTo x,y
//      Delay 1000
//      index = index+1
//   Loop
//End If

func (o *Opsoft) FindPicEx(x1 int64, y1 int64, x2 int64, y2 int64, picName string, deltaColor string, sim float64, dir int64) string {
	ret, _ := oleutil.CallMethod(o.op, "FindPicEx", x1, y1, x2, y2, picName, deltaColor, sim, dir)
	return ret.ToString()
}

//FindPicExS
//这个函数可以查找多个图片, 并且返回所有找到的图像的坐标
//
//此函数同 FindPicEx, 只是返回值不同。
//
//string FindPicExS(x1, y1, x2, y2, pic_name, delta_color,sim, dir)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//pic_name	string	图片名,可以是多个图片,比如"test1.bmp|test2.bmp|test3.bmp"
//delta_color	string	颜色色偏,比如"203040"
//sim	double	相似度,取值范围 0.1-1.0
//dir	int	查找方向,取值如下Dir
//返回值
//
//类型：string
//
//返回的是所有找到的坐标格式如下:"file,x,y| file,x,y..| file,x,y" (图片左上角的坐标)
//
//比如"1.bmp,100,20|2.bmp,30,40" 表示找到了两个,第一个,对应的图片是 1.bmp,坐标是(100,20),第二个是 2.bmp,坐标(30,40)
//
//示例
//
//op = op.FindPicExS(0,0,2000,2000,"test.bmp|test2.bmp|test3.bmp|test4.bmp|test5.bmp","020202",1.0,0)
//If len(op) > 0 Then
//   ss = split(op,"|")
//   index = 0
//   count = UBound(ss) + 1
//   Do While index < count
//      TracePrint ss(index)
//      sss = split(ss(index),",")
//      f = sss(0)
//      x = int(sss(1))
//      y = int(sss(2))
//      op.MoveTo x,y
//      Delay 1000
//      index = index+1
//   Loop
//End If

func (o *Opsoft) FindPicExS(x1 int64, y1 int64, x2 int64, y2 int64, picName string, deltaColor string, sim float64, dir int64) string {
	ret, _ := oleutil.CallMethod(o.op, "FindPicExS", x1, y1, x2, y2, picName, deltaColor, sim, dir)
	return ret.ToString()
}

//FindColorBlock
//查找指定区域内的颜色块,颜色格式"RRGGBB-DRDGDB"
//
//long FindColorBlock(x1, y1, x2, y2, color, sim, count, width, height, intX, intY)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//color	string	颜色格式串，比如"FFFFFF-000000|CCCCCC-000000"每种颜色用"|"分割
//sim	double	相似度,取值范围 0.1-1.0
//count	int	在宽度为 width,高度为 height 的颜色块中，符合 color 颜色的最小数量,通过工具在二值化区域中查看
//width	int	颜色块的宽度
//height	int	颜色块的高度
//intX	int*	变参指针: 返回颜色块的左上角的 X 坐标
//intY	int*	变参指针: 返回颜色块的左上角的 Y 坐标
//返回值
//
//类型：int
//
//0:找到 1:没找到
//
//示例
//
//op = op.FindColorBlock(0,0,2000,2000,"123456-000000|aabbcc-030303|ddeeff-202020",1.0,350,100,200,intX,intY)
//If intX >= 0 and intY >= 0 Then
//    MessageBox "找到"
//End If

func (o *Opsoft) FindColorBlock(x1 int64, y1 int64, x2 int64, y2 int64, color string, sim float64, count int64, width int64, height int64, intX *int64, intY *int64) bool {
	x_ := ole.NewVariant(ole.VT_I4, *intX)
	y_ := ole.NewVariant(ole.VT_I4, *intY)
	defer func() {
		x_.Clear()
		y_.Clear()
	}()
	ret, _ := oleutil.CallMethod(o.op, "FindColorBlock", x1, y1, x2, y2, color, sim, count, width, height, &x_, &y_)
	intX = &x_.Val
	intY = &y_.Val
	return ret.Value() == 1
}

//FindColorBlockEx
//查找指定区域内的所有颜色块, 颜色格式"RRGGBB-DRDGDB"
//
//string FindColorBlockEx(x1, y1, x2, y2, color, sim, count, width, height)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//color	string	颜色格式串，比如"FFFFFF-000000|CCCCCC-000000"每种颜色用"|"分割
//sim	double	相似度,取值范围 0.1-1.0
//count	int	在宽度为 width,高度为 height 的颜色块中，符合 color 颜色的最小数量,通过工具在二值化区域中查看
//width	int	颜色块的宽度
//height	int	颜色块的高度
//返回值
//
//类型：string
//
//返回所有颜色块信息的坐标
//
//示例
//
//s = op.FindColorBlockEx(0,0,2000,2000,"123456-000000|abcdef-202020",1.0,350,100,200)
//count = op.GetResultCount(s)
//index = 0
//Do While index < count
//    op = op.GetResultPos(s,index,intX,intY)
//    MessageBox intX&","&intY
//    index = index + 1
//Loop

func (o *Opsoft) FindColorBlockEx(x1 int64, y1 int64, x2 int64, y2 int64, color string, sim float64, count int64, width int64, height int64) string {
	ret, _ := oleutil.CallMethod(o.op, "FindColorBlockEx", x1, y1, x2, y2, color, sim, count, width, height)
	return ret.ToString()
}

//GetColor
//获取(x,y)的颜色
//
//string GetColor(x,y)
//参数	类型	描述
//x	int	X 坐标
//y	int	Y 坐标
//返回值
//
//类型：string
//
//返回颜色字符串
//
//示例
//
//color = op.GetColor(30,30)
//If color = "ffffff" Then
//     MessageBox "是白色"
//End If

func (o *Opsoft) GetColor(x int64, y int64) string {
	ret, _ := oleutil.CallMethod(o.op, "GetColor", x, y)
	return ret.ToString()
}

//SetDisplayInput
//设置图像输入方式，默认窗口截图
//
//long SetDisplayInput(mode)
//参数	类型	描述
//mode	string	图色输入模式
//mode
//
//值	描述
//screen	默认的模式，表示使用显示器或者后台窗口
//pic	指定输入模式为指定的图片,可以是相对路径,相对于 SetPath 的路径：pic:test.bmp,也可以是绝对路径: pic:d:\test\test.bmp
//mem	指定输入模式为指定的图片,此图片在内存当中，格式：mem:addr 例如：mem:1230434
//pic
//
//指定输入模式为指定的图片,如果使用了这个模式，则所有和图色相关的函数
//
//均视为对此图片进行处理，比如文字识别查找图片 颜色 等等一切图色函数.
//
//需要注意的是，设定以后，此图片就已经加入了缓冲，如果更改了源图片内容，那么需要释放此缓冲，重新设置.
//
//mem
//
//指定输入模式为指定的图片,此图片在内存当中. addr 为图像内存地址
//
//如果使用了这个模式，则所有和图色相关的函数,均视为对此图片进行处理.
//
//比如文字识别 查找图片 颜色 等等一切图色函数.
//
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//// 设定为默认的模式
//op_ret = op.SetDisplayInput("screen")
//
//// 设定为图片模式 图片采用相对路径模式 相对于SetPath的路径
//op_ret = op.SetDisplayInput("pic:test.bmp")
//
//// 设为图片模式 图片采用绝对路径模式
//op_ret = op.SetDisplayInput("pic:d:\test\test.bmp")
//
//// 设为图片模式 但是每次设置前 先清除缓冲
//op_ret = op.FreePic("test.bmp")
//op_ret = op.SetDisplayInput("pic:test.bmp")
//
//// 设置为图片模式,图片从内存中获取
//op_ret = op.SetDisplayInput("mem:1230434")

func (o *Opsoft) SetDisplayInput(mode string) bool {
	ret, _ := oleutil.CallMethod(o.op, "SetDisplayInput", mode)
	return ret.Value() == 1
}

//LoadPic
//预加载指定的图片加入缓存
//
//该方法需要 SetPath 设置全局路径
//
//long LoadPic(pic_name)
//参数	类型	描述
//pic_name	string	文件名,比如"1.bmp|2.bmp|3.bmp" 等.
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//op_ret = op.SetPath("c:\test")
//all_pic = "1.bmp|2.bmp|3.bmp"
//op_ret = op.LoadPic(all_pic)

func (o *Opsoft) LoadPic(picName string) bool {
	ret, _ := oleutil.CallMethod(o.op, "LoadPic", picName)
	return ret.Value() == 1
}

//FreePic
//释放指定的图片
//
//long FreePic(pic_name)
//参数	类型	描述
//pic_name	string	文件名
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//all_pic = "1.bmp|2.bmp|3.bmp"
//op_ret = op.FreePic(all_pic)

func (o *Opsoft) FreePic(picName string) bool {
	ret, _ := oleutil.CallMethod(o.op, "FreePic", picName)
	return ret.Value() == 1
}

// LoadMemPic
// 从内存中加载图片，并将加载结果返回
//
// long LoadMemPic(file_name, data, size)
// 参数	类型	描述
// file_name	string	图片的文件名
// data	int	图像数据
// size	int	图像数据的大小
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// imgage = "test.jpg"
// data = imgage.tobytes // 得到图片数据
// op.LoadMemPic(pic_name, data, len(data))
func (o *Opsoft) LoadMemPic(fileName string, data int64, size int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "LoadMemPic", fileName, data, size)
	return ret.Value() == 1
}

// GetScreenData
// 获取指定区域的图像,用二进制数据的方式返回
//
// long GetScreenData(x1,y1,x2,y2)
// 参数	类型	描述
// x1	int	区域的左上 X 坐标
// y1	int	区域的左上 Y 坐标
// x2	int	区域的右下 X 坐标
// y2	int	区域的右下 Y 坐标
// 返回值
//
// 类型：int
//
// 返回的是指定区域的二进制图片颜色数据，每个颜色是 4 个字节,表示方式为(00RRGGBB)
//
// 示例
//
// op.GetScreenData(0, 0, 2000,2000)
func (o *Opsoft) GetScreenData(x1 int64, y1 int64, x2 int64, y2 int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "GetScreenData", x1, y1, x2, y2)
	return ret.Value() == 1
}

//GetScreenDataBmp
//获取指定区域的图像,用 24 位位图的数据格式返回
//
//long GetScreenDataBmp(x1,y1,x2,y2,data,size)
//参数	类型	描述
//x1	int	区域的左上 X 坐标
//y1	int	区域的左上 Y 坐标
//x2	int	区域的右下 X 坐标
//y2	int	区域的右下 Y 坐标
//data	int*	变参指针:返回图片的数据指针
//size	int*	变参指针:返回图片的数据长度
//返回值
//
//类型：int
//
//0：失败
//1：成功
//示例
//
//var data  int
//var size int
//op.GetScreenDataBmp(0, 0, 2000,2000,&data,&size)

func (o *Opsoft) GetScreenDataBmp(x1 int64, y1 int64, x2 int64, y2 int64, data *int64, size *int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "GetScreenDataBmp", x1, y1, x2, y2, data, size)
	return ret.Value() == 1
}

// GetScreenFrameInfo
// 获取屏幕帧信息,刚方法未导出到 com 库
//
// void GetScreenFrameInfo(long* frame_id, long* time);
// 参数	类型	描述
// frame_id	int	屏幕帧的 ID
// time	int	时间戳
// 返回值
//
// 类型：int
//
// 0：失败
// 1：成功
// 示例
//
// 无
func (o *Opsoft) GetScreenFrameInfo(frameId *int64, time *int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "GetScreenFrameInfo", frameId, time)
	return ret.Value() == 1
}

//MatchPicName
//根据通配符获取文件集合. 方便用于 FindPic 和 FindPicEx
//
//string MatchPicName(pic_name)
//参数	类型	描述
//pic_name	string	文件名,比如"1.bmp|2.bmp|3.bmp" 等
//pic_name 参数: 可以使用通配符
//
//可以使用通配符,比如：
//
//"*.bmp" 这个对应了所有的 bmp 文件
//
//"a?c*.bmp" 这个代表了所有第一个字母是 a 第三个字母是 c 第二个字母任意的所有 bmp 文件
//
//"abc???.bmp|1.bmp|aa??.bmp" 可以这样任意组合.
//
//返回值
//
//类型：string
//
//返回的是通配符对应的文件集合，每个图片以|分割
//
//示例
//
//all_pic = "abc*.bmp"
//pic_name = op.MatchPicName(all_pic)

func (o *Opsoft) MatchPicName(picName string) string {
	ret, _ := oleutil.CallMethod(o.op, "MatchPicName", picName)
	return ret.ToString()
}

//Dir
//值	描述
//0	从左到右,从上到下
//1	从左到右,从下到上
//2	从右到左,从上到下
//3	从右到左,从下到上
//4	从中心往外查找
//5	从上到下,从左到右
//6	从上到下,从右到左
//7	从下到上,从左到右
//8	从下到上,从右到左
