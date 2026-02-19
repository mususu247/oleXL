package oleXL

import (
	"os"
	"path/filepath"
	"strings"
)

// version 2025-10-17
// VBA style like: FileSystemObject(FSO)

func FileExists(fileName string) bool {
	_, err := os.Stat(fileName)
	if err != nil {
		return false
	} else {
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
