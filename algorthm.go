package opsoft

import "github.com/go-ole/go-ole/oleutil"

//算法
//
//接口目录
//AStarFindPath: A 星算法
//FindNearestPos: 查找最近的位置

//接口方法

//AStarFindPath
//根据 A 星算法，获取地图上从源坐标到目的坐标的一条最短路径
//
//string AStarFindPath(mapWidth, mapHeight, disable_points, beginX, beginY, endX, endY)
//参数	类型	描述
//mapWidth	int	区域的左上 X 坐标
//mapHeight	int	区域的左上 Y 坐标
//disable_points	string	不可通行的坐标，以"|"分割，例如:"10,15 | 20,30"
//beginX	int	源坐标 X
//beginY	int	源坐标 Y
//endX	int	目的坐标 X
//endY	int	目的坐标 Y
//返回值
//
//类型：string
//
//找到的路径结果
//
//示例
//
//path=op.AStarFindPath(10,10,"1,0|1,1|1,2|1,3",0,0,9,9)

func (o *Opsoft) AStarFindPath(mapWidth int64, mapHeight int64, disablePoints string, beginX int64, beginY int64, endX int64, endY int64) string {
	ret, _ := oleutil.CallMethod(o.op, "AStarFindPath", mapWidth, mapHeight, disablePoints, beginX, beginY, endX, endY)
	return ret.ToString()
}

// FindNearestPos
// 在一组位置中查找最近的位置
// void FindNearestPos(const wchar_t* all_pos, long type, long x, long y, std::wstring& ret);
// 参数	类型	描述
// all_pos	string	位置
// type	int	类型
// x	int	坐标 x
// y	int	坐标 y
// 返回值
//
// 类型：string
//
// 最接近指定坐标 (x, y) 的位置
//
// 示例
//
// 无
func (o *Opsoft) FindNearestPos(allPos string, type_ int64, x int64, y int64) string {
	ret, _ := oleutil.CallMethod(o.op, "FindNearestPos", allPos, type_, x, y)
	return ret.ToString()
}
