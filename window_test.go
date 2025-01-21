package opsoft

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"reflect"
	"testing"
)

var op *Opsoft

func init() {
	op = NewOpsoft()
	op.Ver()
}
func TestOpsoft_Ver(t *testing.T) {
	fmt.Println("", op.Ver())
}

func TestOpsoft_ClientToScreen(t *testing.T) {
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		hwnd int64
		intX *int64
		intY *int64
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
			if got := o.ClientToScreen(tt.args.hwnd, tt.args.intX, tt.args.intY); got != tt.want {
				t.Errorf("ClientToScreen() = %v, want %v", got, tt.want)
			}
		})
	}
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
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		class string
		title string
	}
	tests := []struct {
		name    string
		fields  fields
		args    args
		wantOut int64
	}{
		// TODO: Add test cases.
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			o := &Opsoft{
				op:       tt.fields.op,
				IUnknown: tt.fields.IUnknown,
			}
			if gotOut := o.FindWindow(tt.args.class, tt.args.title); gotOut != tt.wantOut {
				t.Errorf("FindWindow() = %v, want %v", gotOut, tt.wantOut)
			}
		})
	}
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
	type fields struct {
		op       *ole.IDispatch
		IUnknown *ole.IUnknown
	}
	type args struct {
		parent int64
		class  string
		title  string
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
			if got := o.FindWindowEx(tt.args.parent, tt.args.class, tt.args.title); got != tt.want {
				t.Errorf("FindWindowEx() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestOpsoft_GetClientRect(t *testing.T) {
	hwnd := op.FindWindow("", "微信")
	var x1, x2, y1, y2 int64
	op.GetClientRect(hwnd, &x1, &y1, &x2, &y2)
	fmt.Println("", x1, y1, x2, y2)
}

func TestOpsoft_GetClientSize(t *testing.T) {
	hwnd := op.FindWindow("", "微信")
	var w, h int64
	op.GetClientSize(hwnd, &w, &h)

	fmt.Println("", w, h)

}

func TestOpsoft_MoveWindow(t *testing.T) {
	hwnd := op.FindWindow("", "微信")

	if op.MoveWindow(hwnd, 0, 0) {
		fmt.Println("移动成功")
	} else {
		panic("移动失败")
	}
}
