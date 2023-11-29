package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"time"

	"github.com/tealeg/xlsx"
)

func main() {
	var storeURL string
	var ok bool
	if storeURL, ok = os.LookupEnv("SHOPIFY_STORE_URL"); !ok {
		log.Fatalf("env var SHOPIFY_STORE_URL is missing")
	}
	apiVersion := "2023-10"

	client := &http.Client{}
	req, err := http.NewRequest("GET", fmt.Sprintf("https://%s/admin/api/%s/products.json?published_status=published&limit=250", storeURL, apiVersion), nil)
	if err != nil {
		fmt.Println(err)
		return
	}

	var shopifyAccessToken string
	if shopifyAccessToken, ok = os.LookupEnv("SHOPIFY_ACCESS_TOKEN"); !ok {
		log.Fatalf("env var SHOPIFY_ACCESS_TOKEN is missing")
	}

	req.Header.Set("X-Shopify-Access-Token", shopifyAccessToken)

	resp, err := client.Do(req)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		fmt.Println(err)
		return
	}

	var productsResp map[string][]map[string]interface{}
	err = json.Unmarshal(body, &productsResp)
	if err != nil {
		fmt.Println(err)
		return
	}

	currentTime := time.Now()

	// Format the time to create the file name
	fileName := "products_" + currentTime.Format("20060102150405") + ".xlsx"

	fileXlsx := xlsx.NewFile()

	sheet, err := fileXlsx.AddSheet("Sheet1")
	if err != nil {
		log.Fatalf("Failed to add sheet: %s", err)
	}

	newRowHeader := sheet.AddRow()
	for _, cellValue := range []string{"status", "tags", "vendor", "title", "handle", "price"} {
		cell := newRowHeader.AddCell()
		cell.Value = cellValue
	}

	for _, productMap := range productsResp["products"] {
		newRow := sheet.AddRow()
		for _, cellValue := range []string{fmt.Sprintf("%v", productMap["status"]), fmt.Sprintf("%v", productMap["tags"]), fmt.Sprintf("%v", productMap["vendor"]), fmt.Sprintf("%v", productMap["title"]), fmt.Sprintf("%v", productMap["handle"]), "0"} {
			cell := newRow.AddCell()
			cell.Value = cellValue
		}
	}

	err = fileXlsx.Save(fileName)
	if err != nil {
		log.Fatalf("Failed to save file: %s", err)
	}

}
