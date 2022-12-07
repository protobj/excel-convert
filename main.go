package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/fs"
	"os"
	"path"
	"path/filepath"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

var (
	waitGroup   = sync.WaitGroup{}
	jsonDir     *string
	luaDir      *string
	supportType = map[string]func(value string) (interface{}, error){
		"string": func(value string) (interface{}, error) {
			return value, nil
		},
		"boolean": func(value string) (interface{}, error) {
			return strings.ToLower(value) == "true", nil
		},
		"number": func(value string) (interface{}, error) {
			return strconv.ParseInt(value, 10, 64)
		},
		"object": func(value string) (interface{}, error) {
			m := map[string]interface{}{}
			err := json.Unmarshal([]byte(value), &m)
			return m, err
		},
		"array": func(value string) (interface{}, error) {
			slice := make([]interface{}, 0)
			err := json.Unmarshal([]byte(value), &slice)
			return slice, err
		},
	}
	fileRecord = make(map[string]int64)
)

func main() {

	const recordFileName = "./fileRecord.json"
	file, err := os.ReadFile(recordFileName)
	if err == nil {
		json.Unmarshal(file, &fileRecord)
	}

	startTime := time.Now().UnixNano()
	sourceDir := flag.String("source_dir", "./excel", "excel文件路径")
	jsonDir = flag.String("json_dir", "./json", "json文件路径")
	luaDir = flag.String("lua_dir", "./lua", "lua文件路径")
	flag.Parse()

	filepath.Walk(*sourceDir, walkFunc)

	waitGroup.Wait()

	fileContent, _ := json.Marshal(fileRecord)
	os.WriteFile(recordFileName, fileContent, fs.ModePerm)
	endTime := time.Now().UnixNano()
	fmt.Printf("总耗时:%v毫秒\n", (endTime-startTime)/1000000)

}

func walkFunc(files string, info os.FileInfo, err error) error {
	_, fileName := filepath.Split(files)
	open, _ := os.Open(files)
	stat, _ := open.Stat()
	modified := stat.ModTime().UnixNano()
	if value, ok := fileRecord[files]; ok && value == modified {
		return nil
	}

	if path.Ext(files) == ".xlsx" && !strings.HasPrefix(fileName, "~$") && !strings.HasPrefix(fileName, "#") {
		waitGroup.Add(1)
		go parseXlsx(files, strings.Replace(fileName, ".xlsx", "", -1), modified)
	}
	return nil
}

// 解析xlsx
func parseXlsx(path string, fileName string, modified int64) {
	// 打开excel
	xlsx, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Errorf("%s %s", fileName, err)
		waitGroup.Done()
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := xlsx.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	var sheetCount = xlsx.SheetCount

	if sheetCount == 0 {
		waitGroup.Done()
		return
	}
	var wait = sync.WaitGroup{}
	for i := 0; i < sheetCount; i++ {
		sheetName := xlsx.GetSheetName(i)
		if strings.HasSuffix(sheetName, "#") {
			continue
		}
		wait.Add(1)
		rows, err := xlsx.GetRows(sheetName)
		if err != nil {
			fmt.Errorf("%v\n", err)
			os.Exit(-1)
		}
		go parseSheet(path, fileName, sheetName, rows, &wait)
	}
	wait.Wait()
	fileRecord[path] = modified
	waitGroup.Done()
}

func parseSheet(path string, fileName string, sheetName string, rows [][]string, wait *sync.WaitGroup) {
	defer wait.Done()
	//纵向解析
	var serverResult []interface{}
	var clientResult []interface{}
	if strings.HasSuffix(sheetName, "_single") {
		var serverIndex = 2
		var clientIndex = 3
		var typeIndex = 4
		var valueIndex = 5
		var start = false
		var clientObj = make(map[string]interface{})
		var serverObj = make(map[string]interface{})
		for i, row := range rows {
			if !start {
				if row[0] == "START" {
					start = true
				}
				continue
			} else {
				if row[0] == "#" {
					continue
				}
				if row[0] == "END" {
					break
				}
				if len(row) <= valueIndex {
					fmt.Printf("error value in %s  sheet:%s  row:%d cellLen:%d \n", path, sheetName, i, valueIndex)
					os.Exit(-1)
				}

				type0 := row[typeIndex]
				if _, ok := supportType[type0]; !ok {
					fmt.Printf("error type in %s  sheet:%s  row:%d cell:%d value:%s\n", path, sheetName, i, typeIndex, type0)
					os.Exit(-1)
				}
				value := row[valueIndex]
				decoder, _ := supportType[type0]

				jsonValue, err := decoder(value)
				if err != nil {
					fmt.Printf("error value in %s  sheet:%s  row:%d cell:%d value:%s\n", path, sheetName, i, valueIndex, value)
					os.Exit(-1)
				}

				clientName := row[clientIndex]
				if clientName != "" {
					clientObj[clientName] = jsonValue
				}
				serverName := row[serverIndex]
				if serverName != "" {
					serverObj[serverName] = jsonValue
				}
			}
		}
		serverResult = append(serverResult, serverObj)
		clientResult = append(clientResult, clientObj)
	} else {
		//TODO 横向解析

	}
	if *jsonDir != "" {
		writeJsonFile(fileName, clientResult, sheetName, serverResult)
	}
	if *luaDir != "" {
		writeLuaFile(fileName, clientResult, sheetName, serverResult)
	}
}

func writeLuaFile(fileName string, clientResult []interface{}, sheetName string, serverResult []interface{}) {
	err := createDir(*luaDir + "/client/" + fileName)
	if err != nil {
		println(err.Error())
		os.Exit(-1)
	}
	clientTable := strings.Builder{}
	clientTable.WriteString("return ")
	writeLuaTableContent(&clientTable, clientResult, 0)
	os.WriteFile(*luaDir+"/client/"+fileName+"/"+sheetName+".lua", []byte(clientTable.String()), fs.ModePerm)

	err = createDir(*luaDir + "/server/" + fileName)
	if err != nil {
		println(err.Error())
		os.Exit(-1)
	}
	serverTable := strings.Builder{}
	serverTable.WriteString("return ")
	writeLuaTableContent(&serverTable, clientResult, 0)
	os.WriteFile(*luaDir+"/server/"+fileName+"/"+sheetName+".lua", []byte(serverTable.String()), fs.ModePerm)
}

func writeJsonFile(fileName string, clientResult []interface{}, sheetName string, severResult []interface{}) {
	err := createDir(*jsonDir + "/client/" + fileName)
	if err != nil {
		println(err.Error())
		os.Exit(-1)
	}
	clientTable, _ := json.Marshal(clientResult)

	os.WriteFile(*jsonDir+"/client/"+fileName+"/"+sheetName+".json", clientTable, fs.ModePerm)

	err = createDir(*jsonDir + "/server/" + fileName)
	if err != nil {
		println(err.Error())
		os.Exit(-1)
	}
	serverTable, _ := json.Marshal(severResult)
	os.WriteFile(*jsonDir+"/server/"+fileName+"/"+sheetName+".json", serverTable, fs.ModePerm)
}

// 创建文件夹
func createDir(dir string) error {
	exist, err := pathExists(dir)
	if err != nil {
		fmt.Printf("create dir error :%s", err.Error())
		os.Exit(-1)
		return err
	}
	if !exist {
		if err := os.MkdirAll(dir, os.ModePerm); err != nil {
			fmt.Printf("create dir error :%s", err.Error())
			os.Exit(-1)
		}
	}
	return nil
}

func pathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}

// 写Lua表内容
func writeLuaTableContent(builder *strings.Builder, data interface{}, idx int) {
	// 如果是指针类型
	if reflect.ValueOf(data).Type().Kind() == reflect.Pointer {
		data = reflect.ValueOf(data).Elem().Interface()
	}
	switch t := data.(type) {
	case int64:
		builder.WriteString(fmt.Sprintf("%d", data))
	case float64:
		builder.WriteString(fmt.Sprintf("%v", data))
	case string:
		builder.WriteString(fmt.Sprintf(`"%s"`, data))
	case []interface{}:
		builder.WriteString("{\n")
		a := data.([]interface{})
		for _, v := range a {
			addTabs(builder, idx)
			writeLuaTableContent(builder, v, idx+1)
			builder.WriteString(",\n")
		}
		addTabs(builder, idx-1)
		builder.WriteString("}")
	case []string:
		builder.WriteString("{\n")
		a := data.([]string)
		sort.Strings(a)
		for _, v := range a {
			addTabs(builder, idx)
			writeLuaTableContent(builder, v, idx+1)
			builder.WriteString(",\n")
		}
		addTabs(builder, idx-1)
		builder.WriteString("}")
	case map[string]interface{}:
		m := data.(map[string]interface{})
		keys := make([]string, 0)
		for k := range m {
			keys = append(keys, k)
		}
		sort.Strings(keys)

		builder.WriteString("{\n")
		for _, k := range keys {
			addTabs(builder, idx)
			builder.WriteString("[")
			writeLuaTableContent(builder, k, idx+1)
			builder.WriteString("] = ")
			writeLuaTableContent(builder, m[k], idx+1)
			builder.WriteString(",\n")
		}
		addTabs(builder, idx-1)
		builder.WriteString("}")
	default:
		builder.WriteString(fmt.Sprintf("%t", data))
		_ = t
	}
}

// 在文件中添加制表符
func addTabs(builder *strings.Builder, idx int) {
	for i := 0; i < idx; i++ {
		builder.WriteString("\t")
	}
}
