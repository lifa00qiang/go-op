package opsoft

import (
	"embed"
	"errors"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"os"
	"syscall"
	"unsafe"
)

var (
	opTool  *syscall.LazyDLL
	_setupA *syscall.LazyProc
	_setupW *syscall.LazyProc
)

type Opsoft struct {
	op       *ole.IDispatch
	IUnknown *ole.IUnknown
}

//go:embed dll
var dllFS embed.FS

func init() {
	tools, _ := dllFS.ReadFile("dll/tools.dll")
	opdll, _ := dllFS.ReadFile("dll/op_x64.dll")
	tf, err := os.Create("dll/tools.dll")
	if err != nil {
		return
	}
	tf.Write(tools)
	tf.Close()
	of, err := os.Create("dll/op_x64.dll")
	if err != nil {
		return
	}
	of.Write(opdll)
	of.Close()

	opTool = syscall.NewLazyDLL("dll/tools.dll")
	_setupA = opTool.NewProc("setupA")
	_setupW = opTool.NewProc("setupW")

}

func NewOpsoft() *Opsoft {
	//注册op
	if !SetupW("dll/op_x64.dll") {
		panic(errors.New("注册失败"))
	}

	//初始化op对象
	var o *ole.IDispatch
	var i *ole.IUnknown
	var err error
	_ = ole.CoInitialize(0)
	i, err = oleutil.CreateObject("op.opsoft")
	if err != nil {
		panic(err)
	}
	o, err = i.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		panic(err)
	}
	return &Opsoft{
		op:       o,
		IUnknown: i,
	}
}

func (o *Opsoft) Release() {
	o.op.Release()
	o.IUnknown.Release()
	ole.CoUninitialize()
}

func SetupA(path string) bool {
	var _p0 *uint16
	_p0, _ = syscall.UTF16PtrFromString(path)
	ret, _, _ := _setupA.Call(uintptr(unsafe.Pointer(_p0)))
	return ret != 0
}

func SetupW(path string) bool {
	var _p0 *uint16
	_p0, _ = syscall.UTF16PtrFromString(path)
	ret, _, _ := _setupW.Call(uintptr(unsafe.Pointer(_p0)))
	return ret != 0
}
