package opsoft

import "github.com/go-ole/go-ole/oleutil"

//memory
//内存操作
//
//接口目录
//WriteData: 向某进程写入数据
//ReadData: 读取数据

//接口方法

//WriteData
//向某进程写入数据
//
//long WriteData(hwnd,address,data,size)
//参数	类型	描述
//hwnd	int	窗口句柄，用于指定要在哪个窗口内写入数据
//address	string	写入数据的地址
//data	string	写入的数据
//size	int	写入的数据的大小
//返回值
//
//类型：int
//
//0: 失败
//1: 成功
//示例
//
//无

func (o *Opsoft) WriteData(hwnd int64, address string, data string, size int64) bool {
	ret, _ := oleutil.CallMethod(o.op, "WriteData", hwnd, address, data, size)
	return ret.Value() == 1
}

//ReadData
//读取数据
//
//string ReadData(hwnd,address,size)
//参数	类型	描述
//hwnd	int	窗口句柄，用于指定要从哪个窗口内读取数据
//address	string	表示要读取数据的地址
//size	int	要读取的数据的大小
//返回值
//
//类型：string
//
//读取到的数值
//
//示例
//
//无

func (o *Opsoft) ReadData(hwnd int64, address string, size int64) string {
	ret, _ := oleutil.CallMethod(o.op, "ReadData", hwnd, address, size)
	return ret.ToString()
}
