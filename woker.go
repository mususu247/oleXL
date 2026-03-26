package oleXL

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

func (Q *Worker) Loop() error {
	if Q == nil {
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
		sendMsg, ok := <-Q.sendQ
		if !ok {
			if Q.parent.debug {
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
				Q.recvQ <- recvMsg
				continue
			}

			unknown, err = oleutil.CreateObject(sendMsg.Name)
			if err != nil {
				recvMsg.Name = "CreateObject"
				recvMsg.Args = []any{err}
				Q.recvQ <- recvMsg
				continue
			}

			app, err = unknown.QueryInterface(ole.IID_IDispatch)
			if err != nil {
				recvMsg.Name = "QueryInterface"
				recvMsg.Args = []any{err}
				Q.recvQ <- recvMsg
				continue
			}

			recvMsg.Args = []any{app}
			Q.recvQ <- recvMsg
		case "Get":
			recvMsg = sendMsg

			arg, err = oleutil.GetProperty(sendMsg.Disp, sendMsg.Name, sendMsg.Args...)
			if err != nil {
				recvMsg.Args = []any{err}
				Q.recvQ <- recvMsg
				continue
			}

			if arg.VT == ole.VT_ERROR {
				recvMsg.Args = []any{arg}
			} else {
				x := arg.Value()
				recvMsg.Args = []any{x}
			}

			Q.recvQ <- recvMsg
		case "Put":
			recvMsg = sendMsg

			arg, err = oleutil.PutProperty(sendMsg.Disp, sendMsg.Name, sendMsg.Args...)
			if err != nil {
				recvMsg.Args = []any{err}
				Q.recvQ <- recvMsg
				continue
			}

			x := arg.Value()
			recvMsg.Args = []any{x}
			Q.recvQ <- recvMsg
		case "Method":
			recvMsg = sendMsg

			arg, err := oleutil.CallMethod(sendMsg.Disp, sendMsg.Name, sendMsg.Args...)
			if err != nil {
				recvMsg.Args = []any{err}
				Q.recvQ <- recvMsg
				continue
			}

			x := arg.Value()
			recvMsg.Args = []any{x}
			Q.recvQ <- recvMsg
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

			Q.recvQ <- recvMsg
		default:
			recvMsg = sendMsg

			err = fmt.Errorf("not found Cmd: %v", sendMsg.Cmd)
			recvMsg.Args = []any{err}
			Q.recvQ <- recvMsg
		}
	}
	if Q.parent.debug {
		log.Printf("Loop.Quit\n")
	}
	return nil
}

func (Q *Worker) IsOpened() bool {
	return Q.loop
}

func (Q *Worker) Start() error {
	if Q == nil {
		return fmt.Errorf("(w *Worker) is NULL\n")
	}

	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	if !Q.loop {
		Q.loop = true

		if Q.parent.debug {
			log.Printf("queue.Open.\n")
		}
		Q.sendQ = make(chan msg)
		Q.recvQ = make(chan msg)

		go func() {
			runtime.LockOSThread()
			defer runtime.UnlockOSThread()

			if Q.parent.debug {
				log.Printf("Loop.Start \n")
			}
			Q.Loop()
			if Q.parent.debug {
				log.Printf("Loop.Stop \n")
			}
		}()
	} else {
		return fmt.Errorf("queue.Opened.")
	}
	return nil
}

func (Q *Worker) Stop() error {
	if Q == nil {
		return fmt.Errorf("(w *Worker) is NULL\n")
	}

	if Q.loop {
		Q.loop = false

		Q.mux.Lock()
		close(Q.sendQ)
		close(Q.recvQ)
		Q.mux.Unlock()
		if Q.parent.debug {
			log.Printf("queue.Close.\n")
		}
	}
	return nil
}

func (Q *Worker) Send(cmd string, disp *ole.IDispatch, name string, args []any) []any {
	if Q == nil {
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

	if !Q.loop {
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

	Q.mux.Lock()
	Q.sendQ <- sendMsg
	for {
		recvMsg = <-Q.recvQ
		if recvMsg.ID == sendMsg.ID {
			break
		}
	}
	result = recvMsg.Args
	Q.mux.Unlock()

	return result
}
