package opsoft

import (
	"fmt"
	"testing"
)

func TestOpsoft_GetCursorPos(t *testing.T) {
	op := NewOpsoft()
	var x, y int64
	if !op.GetCursorPos(&x, &y) {
		fmt.Println("", 111111111111)
		op.SetShowErrorMsg(3)
		e := op.GetLastError()
		fmt.Println("", e)
	}
	fmt.Println("", x, y)
}

func TestOpsoft_WheelUp(t *testing.T) {
	op.SetWindowState(83261660, 1)
	op.WheelUp()
	op.Sleep(3000)
	op.WheelUp()
	op.Sleep(3000)
	op.WheelUp()
	op.Sleep(3000)
	op.WheelUp()
	op.Sleep(3000)
}

func TestOpsoft_MoveTo(t *testing.T) {
	op.MoveTo(100, 100)
	op.Sleep(1000)
}

func TestOpsoft_LeftClick(t *testing.T) {
	op.LeftClick()
	op.Sleep(1000)
}

func TestOpsoft_LeftDoubleClick(t *testing.T) {
	if !op.LeftDoubleClick() {
		fmt.Println("", 123)
	}
}
