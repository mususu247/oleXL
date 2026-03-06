package oleXL

import (
	"bytes"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
)

// VBA style like: FileSystemObject(FSO)

func CopyFile(src string, dst string, overwrite ...bool) error {
	var ow bool = true
	if len(overwrite) > 0 {
		ow = overwrite[0]
	}

	var _copyFile func(src string, dst string) error

	_copyFile = func(src string, dst string) error {
		c, err := os.Create(dst)
		if err != nil {
			return err
		}
		defer c.Close()

		r, err := os.Open(src)
		if err != nil {
			return err
		}
		defer r.Close()

		_, err = io.Copy(c, r)
		if err != nil {
			return err
		}

		fi, _ := os.Stat(src)
		err = os.Chtimes(dst, fi.ModTime(), fi.ModTime())
		if err != nil {
			return err
		}
		return nil
	}

	if FileExists(src) {
		//mode Single FIle
		if FolderExists(dst) {
			fn := filepath.Base(src)
			dst = filepath.Join(dst, fn)
		}

		if FileExists(dst) {
			if ow {
				err := DeleteFile(dst)
				if err != nil {
					return err
				}
			} else {
				return fmt.Errorf("file already exists: %s", dst)
			}
		}

		err := _copyFile(src, dst)
		if err != nil {
			return err
		}
	} else {
		//mode Multi Files
		srcPath := GetFilePath(src)
		find := filepath.Base(src)
		files, err := FindFiles(find, srcPath, 0)
		if err != nil {
			return err
		}

		var errs error
		for i := range files {
			fn := filepath.Base(files[i])
			dstFile := filepath.Join(dst, fn)

			if FileExists(dstFile) {
				if ow {
					err := DeleteFile(dstFile)
					if err != nil {
						errs = err
					} else {
						err := _copyFile(files[i], dstFile)
						if err != nil {
							errs = err
						}
					}
				}
			} else {
				err := _copyFile(files[i], dstFile)
				if err != nil {
					errs = err
				}
			}
		}
		if errs != nil {
			return errs
		}
	}
	return nil
}

func FileExists(fileName string) bool {
	fi, err := os.Stat(fileName)
	if err != nil {
		return false
	} else {
		if fi.IsDir() {
			return false
		}
		return true
	}
}

func FolderExists(filePath string) bool {
	fi, err := os.Stat(filePath)
	if err == nil {
		if fi.IsDir() {
			return true
		}
	}
	return false
}

func DeleteFile(fileName string) error {
	return os.Remove(fileName)
}

func DeleteFolder(filePath string) error {
	return os.RemoveAll(filePath)
}

func GetFile(fileName string) (string, error) {
	return os.Executable()
}

func GetFilePath(fileName string) string {
	return filepath.Dir(fileName)
}

func GetExtensionName(fileName string) string {
	return filepath.Ext(fileName)
}

func GetBaseName(fileName string) string {
	ext := filepath.Ext(fileName)
	base := filepath.Base(fileName)
	return strings.ReplaceAll(base, ext, "")
}

func GetFileName(fileName string) string {
	return filepath.Base(fileName)
}

func GetAbsolutePathName(fileName string) (string, error) {
	return filepath.Abs(fileName)
}

func BuildPath(filePath string, fileName string) string {
	return filepath.Join(filePath, fileName)
}

func FindFiles(find string, dir string, deep int) ([]string, error) {
	var fulldir string
	var results []string

	fulldir, err := filepath.Abs(dir)
	if err != nil {
		return results, err
	}

	var _findFiles func(find string, dir string, deep int) ([]string, error)

	_findFiles = func(find string, dir string, deep int) ([]string, error) {
		var results []string

		fulldir, err := filepath.Abs(dir)
		if err != nil {
			return results, err
		}

		files, err := os.ReadDir(fulldir)
		if err != nil {
			return results, err
		}

		var childFind bool
		if deep != 0 {
			childFind = true
		}

		childDeep := deep
		if deep > 0 {
			childDeep--
		}

		for _, f := range files {
			fullName := filepath.Join(fulldir, f.Name())

			if f.IsDir() {
				if childFind {
					files, _ := _findFiles(find, fullName, childDeep)

					for i := range files {
						results = append(results, files[i])
					}
				}
			} else {
				if len(find) > 0 {
					if ok, _ := filepath.Match(find, f.Name()); ok {
						results = append(results, fullName)
					}
				} else {
					results = append(results, fullName)
				}
			}
		}

		return results, nil
	}

	results, err = _findFiles(find, fulldir, deep)
	if err != nil {
		return results, err
	}

	return results, nil
}

func AddBOM(fileName string) error {
	bom := []byte{0xEF, 0xBB, 0xBF}

	content, err := os.ReadFile(fileName)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}

	if len(content) >= 3 && bytes.Equal(content[:3], bom) {
		log.Printf("(Info) BOM already exists.")
		return nil
	}

	newContent := append(bom, content...)
	err = os.WriteFile(fileName, newContent, 0644)
	if err != nil {
		log.Printf("(Error) %v", err)
		return err
	}
	return nil
}
