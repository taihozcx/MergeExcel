package models

import (
	"github.com/google/uuid"
	"os"
	"os/exec"
)

// Clear 清除命令行中的信息
func Clear() {
	cmd := exec.Command("cmd", "/c", "cls") //Windows example, its tested
	cmd.Stdout = os.Stdout
	err := cmd.Run()
	if err != nil {
		return
	}
}

func SetTitle(title string) {
	cmd := exec.Command("cmd", "/c", "title "+title)
	err := cmd.Run()
	if err != nil {
		return
	}
}

func Getuuid() string {
	return uuid.New().String()
}
