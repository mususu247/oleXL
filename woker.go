package oleXL

// version 2026-01-26

import (
	"crypto/rand"
	"fmt"
	"log"
	"runtime"
	"sync"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type msg struct {
	ID   string
	Cmd  string
	Disp *ole.IDispatch
	Name string
	Args []any
}

type Worker struct {
	parent *Cores
	sendQ  chan msg
	recvQ  chan msg
	mux    sync.Mutex
	loop   bool
}

func (w *Worker) Loop() error {
	if w == nil {
		return fmt.Errorf("(w *Worker) is NULL\n")
	}
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	var recvMsg msg
	var err error
	var unknown *ole.IUnknown
	var app *ole.IDispatch
	var arg *ole.VARIANT

	for {
		sendMsg, ok := <-w.sendQ
		if !ok {
			if w.parent.debug {
				log.Printf("Loop.Exit\n")
			}
			break
		}

		switch sendMsg.Cmd {
		case "Create":
			recvMsg = sendMsg

			if app != nil {
				return fmt.Errorf("already created.")
			}

			err = ole.CoInitialize(0)
			if err != nil {
				recvMsg.Name = "CoInitialize"
				recvMsg.Args = []any{err}
				w.recvQ <- recvMsg
				continue
			}

			unknown, err = oleutil.CreateObject(sendMsg.Name)
			if err != nil {
				recvMsg.Name = "CreateObject"
				recvMsg.Args = []any{err}
				w.recvQ <- recvMsg
				continue
			}

			app, err = unknown.QueryInterface(ole.IID_IDispatch)
			if err != nil {
				recvMsg.Name = "QueryInterface"
				recvMsg.Args = []any{err}
				w.recvQ <- recvMsg
				continue
			}

			recvMsg.Args = []any{app}
			w.recvQ <- recvMsg
		case "Get":
			recvMsg = sendMsg

			arg, err = oleutil.GetProperty(sendMsg.Disp, sendMsg.Name, sendMsg.Args...)
			if err != nil {
				recvMsg.Args = []any{err}
				w.recvQ <- recvMsg
				continue
			}

			if arg.VT == ole.VT_ERROR {
				recvMsg.Args = []any{arg}
			} else {
				x := arg.Value()
				recvMsg.Args = []any{x}
			}

			w.recvQ <- recvMsg
		case "Put":
			recvMsg = sendMsg

			arg, err = oleutil.PutProperty(sendMsg.Disp, sendMsg.Name, sendMsg.Args...)
			if err != nil {
				recvMsg.Args = []any{err}
				w.recvQ <- recvMsg
				continue
			}

			x := arg.Value()
			recvMsg.Args = []any{x}
			w.recvQ <- recvMsg
		case "Method":
			recvMsg = sendMsg

			arg, err := oleutil.CallMethod(sendMsg.Disp, sendMsg.Name, sendMsg.Args...)
			if err != nil {
				recvMsg.Args = []any{err}
				w.recvQ <- recvMsg
				continue
			}

			x := arg.Value()
			recvMsg.Args = []any{x}
			w.recvQ <- recvMsg
		case "Release":
			recvMsg = sendMsg

			if app != sendMsg.Disp {
				ans := sendMsg.Disp.Release()
				recvMsg.Args = []any{ans}
			} else {
				ans1 := app.Release()
				ans2 := unknown.Release()
				app = nil
				unknown = nil
				ole.CoUninitialize()
				recvMsg.Args = []any{ans1, ans2}
			}

			w.recvQ <- recvMsg
		default:
			recvMsg = sendMsg

			err = fmt.Errorf("not found Cmd: %v", sendMsg.Cmd)
			recvMsg.Args = []any{err}
			w.recvQ <- recvMsg
		}
	}
	if w.parent.debug {
		log.Printf("Loop.Quit\n")
	}
	return nil
}

func (w *Worker) IsOpened() bool {
	return w.loop
}

func (w *Worker) Start() error {
	if w == nil {
		return fmt.Errorf("(w *Worker) is NULL\n")
	}

	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	if !w.loop {
		w.loop = true

		if w.parent.debug {
			log.Printf("queue.Open.\n")
		}
		w.sendQ = make(chan msg)
		w.recvQ = make(chan msg)

		go func() {
			runtime.LockOSThread()
			defer runtime.UnlockOSThread()

			if w.parent.debug {
				log.Printf("Loop.Start \n")
			}
			w.Loop()
			if w.parent.debug {
				log.Printf("Loop.Stop \n")
			}
		}()
	} else {
		return fmt.Errorf("queue.Opened.")
	}
	return nil
}

func (w *Worker) Stop() error {
	if w == nil {
		return fmt.Errorf("(w *Worker) is NULL\n")
	}

	if w.loop {
		w.loop = false

		w.mux.Lock()
		close(w.sendQ)
		close(w.recvQ)
		w.mux.Unlock()
		if w.parent.debug {
			log.Printf("queue.Close.\n")
		}
	}
	return nil
}

func (w *Worker) Send(cmd string, disp *ole.IDispatch, name string, args []any) []any {
	if w == nil {
		log.Printf("(w *Worker) is NULL\n")
		return nil
	}

	if cmd == "" {
		log.Printf("cmd is empty\n")
		return nil
	}

	if disp == nil {
		if cmd != "Create" {
			//log.Printf("disp is NULL\n")
			return nil
		}
	}

	if name == "" {
		log.Printf("name is empty\n")
		return nil
	}

	if !w.loop {
		log.Printf("queue is not opened\n")
		return nil
	}

	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	var result []any
	var sendMsg msg
	var recvMsg msg

	sendMsg.ID = rand.Text()
	sendMsg.Cmd = cmd
	sendMsg.Disp = disp
	sendMsg.Name = name
	sendMsg.Args = args

	w.mux.Lock()
	w.sendQ <- sendMsg
	for {
		recvMsg = <-w.recvQ
		if recvMsg.ID == sendMsg.ID {
			break
		}
	}
	result = recvMsg.Args
	w.mux.Unlock()

	return result
}
