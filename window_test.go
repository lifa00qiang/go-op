package opsoft

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"reflect"
	"testing"
	"time"
)

var op *Opsoft

func init() {
	ole.CoInitialize(0)
	op = NewOpsoft()
}

func TestOpsoft_ClientToScreen(t *testing.T) {

	hwnd := op.FindWindow("", "计算器")

	var (
		x int64 = 1
		y int64 = 1
	)

	op.ClientToScreen(hwnd, &x, &y)
	fmt.Println("", x, y)

}

func TestOpsoft_EnumProcess(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		name string
	}
	tests := []struct {
		name    string
		fields  fields
		args    args
		wantOut []int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if gotOut := o.EnumProcess(tt.args.name); !reflect.DeepEqual(gotOut, tt.wantOut) {
				t.Errorf("EnumProcess() = %v, want %v", gotOut, tt.wantOut)
			}
		})
	}
}

func TestOpsoft_EnumWindow(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		parent    int64
		title     string
		className string
		filter    int64
	}
	tests := []struct {
		name    string
		fields  fields
		args    args
		wantOut []int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if gotOut := o.EnumWindow(tt.args.parent, tt.args.title, tt.args.className, tt.args.filter); !reflect.DeepEqual(gotOut, tt.wantOut) {
				t.Errorf("EnumWindow() = %v, want %v", gotOut, tt.wantOut)
			}
		})
	}
}

func TestOpsoft_EnumWindowByProcess(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		processName string
		title       string
		className   string
		filter      int64
	}
	tests := []struct {
		name    string
		fields  fields
		args    args
		wantOut []int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if gotOut := o.EnumWindowByProcess(tt.args.processName, tt.args.title, tt.args.className, tt.args.filter); !reflect.DeepEqual(gotOut, tt.wantOut) {
				t.Errorf("EnumWindowByProcess() = %v, want %v", gotOut, tt.wantOut)
			}
		})
	}
}

func TestOpsoft_FindWindow(t *testing.T) {
	hwnd := op.FindWindow("", "计算器")
	fmt.Println("", hwnd)
}

func TestOpsoft_FindWindowByProcess(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		processName string
		class       string
		title       string
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.FindWindowByProcess(tt.args.processName, tt.args.class, tt.args.title); got != tt.want {
				t.Errorf("FindWindowByProcess() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_FindWindowByProcessId(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		processId int64
		class     string
		title     string
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.FindWindowByProcessId(tt.args.processId, tt.args.class, tt.args.title); got != tt.want {
				t.Errorf("FindWindowByProcessId() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_FindWindowEx(t *testing.T) {
	hwnd := op.FindWindowEx(0, "", "计算器")
	fmt.Println("", hwnd)
}

func TestOpsoft_GetClientRect(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		x1   *int64
		y1   *int64
		x2   *int64
		y2   *int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetClientRect(tt.args.hwnd, tt.args.x1, tt.args.y1, tt.args.x2, tt.args.y2); got != tt.want {
				t.Errorf("GetClientRect() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetClientSize(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd   int64
		width  *int64
		height *int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetClientSize(tt.args.hwnd, tt.args.width, tt.args.height); got != tt.want {
				t.Errorf("GetClientSize() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetForegroundFocus(t *testing.T) {
	op.GetForegroundFocus()
}

func TestOpsoft_GetForegroundWindow(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	tests := []struct {
		name   string
		fields fields
		want   int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetForegroundWindow(); got != tt.want {
				t.Errorf("GetForegroundWindow() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetMousePointWindow(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	tests := []struct {
		name   string
		fields fields
		want   int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetMousePointWindow(); got != tt.want {
				t.Errorf("GetMousePointWindow() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetPointWindow(t *testing.T) {
	hwnd := op.GetPointWindow(315, 223)
	fmt.Println("", hwnd)
}

func TestOpsoft_GetProcessInfo(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		pid int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		wantP  OpProcessInfo
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if gotP := o.GetProcessInfo(tt.args.pid); !reflect.DeepEqual(gotP, tt.wantP) {
				t.Errorf("GetProcessInfo() = %v, want %v", gotP, tt.wantP)
			}
		})
	}
}

func TestOpsoft_GetSpecialWindow(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		flag int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetSpecialWindow(tt.args.flag); got != tt.want {
				t.Errorf("GetSpecialWindow() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetWindow(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		flag int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetWindow(tt.args.hwnd, tt.args.flag); got != tt.want {
				t.Errorf("GetWindow() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetWindowClass(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   string
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetWindowClass(tt.args.hwnd); got != tt.want {
				t.Errorf("GetWindowClass() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetWindowProcessId(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetWindowProcessId(tt.args.hwnd); got != tt.want {
				t.Errorf("GetWindowProcessId() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetWindowProcessPath(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   string
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetWindowProcessPath(tt.args.hwnd); got != tt.want {
				t.Errorf("GetWindowProcessPath() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetWindowRect(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		x1   *int64
		y1   *int64
		x2   *int64
		y2   *int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetWindowRect(tt.args.hwnd, tt.args.x1, tt.args.y1, tt.args.x2, tt.args.y2); got != tt.want {
				t.Errorf("GetWindowRect() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetWindowState(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		flag int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.GetWindowState(tt.args.hwnd, tt.args.flag); got != tt.want {
				t.Errorf("GetWindowState() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetWindowTitle(t *testing.T) {
	hwnds := op.EnumWindow(0, "计算器", "", 1+16)
	fmt.Println("", hwnds)
	for _, v := range hwnds {
		fmt.Println("", op.GetWindowTitle(v))

	}

	//fmt.Println("", op.GetWindowTitle(2823006))
}

func TestOpsoft_MoveWindow(t *testing.T) {
	op.SetWindowState(2823006, 1)
	op.Sleep(1000)
	op.MoveWindow(10749702, 1000, 10)
	time.Sleep(1000)
}

func TestOpsoft_ScreenToClient(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		x    *int64
		y    *int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.ScreenToClient(tt.args.hwnd, tt.args.x, tt.args.y); got != tt.want {
				t.Errorf("ScreenToClient() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_SendString(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		str  string
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.SendString(tt.args.hwnd, tt.args.str); got != tt.want {
				t.Errorf("SendString() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_SendStringIme(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		str  string
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.SendStringIme(tt.args.hwnd, tt.args.str); got != tt.want {
				t.Errorf("SendStringIme() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_SetClientSize(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd   int64
		width  int64
		height int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.SetClientSize(tt.args.hwnd, tt.args.width, tt.args.height); got != tt.want {
				t.Errorf("SetClientSize() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_SetWindowSize(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd   int64
		width  int64
		height int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.SetWindowSize(tt.args.hwnd, tt.args.width, tt.args.height); got != tt.want {
				t.Errorf("SetWindowSize() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_SetWindowState(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		flag int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.SetWindowState(tt.args.hwnd, tt.args.flag); got != tt.want {
				t.Errorf("SetWindowState() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_SetWindowText(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd  int64
		title string
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.SetWindowText(tt.args.hwnd, tt.args.title); got != tt.want {
				t.Errorf("SetWindowText() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_SetWindowTransparent(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd  int64
		trans int64
	}
	tests := []struct {
		name   string
		fields fields
		args   args
		want   bool
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if got := o.SetWindowTransparent(tt.args.hwnd, tt.args.trans); got != tt.want {
				t.Errorf("SetWindowTransparent() = %v, want %v", got, tt.want)
			}
		})
	}
}
