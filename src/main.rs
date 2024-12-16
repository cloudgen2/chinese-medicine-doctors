use reqwest::Client;
use serde::Deserialize;
use std::error::Error;
use xlsxwriter::{Workbook, Format};
use std::fs::File;
use std::io::{Write, BufWriter};

#[derive(Deserialize, Debug)]
struct ResponseData {
    csurname: Option<String>,      // 姓
    cname: Option<String>,         // 名
    esurname: Option<String>,      // 英文姓
    ename: Option<String>,         // 英文名
    ocsurname: Option<String>,     // 可選字段
    ocname: Option<String>,        // 可選字段
    id: Option<String>,            // ID
    ecompany: Option<String>,      // 英文公司
    eaddr1: Option<String>,        // 地址1
    eaddr2: Option<String>,        // 地址2
    eaddr3: Option<String>,        // 地址3
    eaddr4: Option<String>,        // 地址4
    ccompany: Option<String>,      // 中文公司
    caddr1: Option<String>,        // 中文地址1
    caddr2: Option<String>,        // 中文地址2
    caddr3: Option<String>,        // 中文地址3
    caddr4: Option<String>,        // 中文地址4
    spec_date_from: Option<String>, // 特殊日期從
    spec_date_to: Option<String>,   // 特殊日期至
    promote: Option<String>,        // 促銷
    stroke: Option<String>,         // 筆劃
    surnamestroke: Option<String>,  // 姓名筆劃
}

#[tokio::main]
async fn main() -> Result<(), Box<dyn Error>> {
    let client = Client::new();
    let url = "https://www.cmchk.org.hk/cmp/chi/cmp_search_data.php";

    // 創建新的 Excel 檔案
    let workbook = Workbook::new("output.xlsx")?;
    let mut worksheet = workbook.add_worksheet(None)?;

    // 創建 CSV 檔案
    let csv_file = File::create("output.csv")?;
    let mut csv_writer = BufWriter::new(csv_file);

    // 創建格式
    let mut format = Format::new();
    format.set_bold(); // 設置為粗體

    // 寫入表頭到 Excel
    worksheet.write_string(0, 0, "Surname", Some(&format))?;
    worksheet.write_string(0, 1, "Name", Some(&format))?;
    worksheet.write_string(0, 2, "English Surname", Some(&format))?;
    worksheet.write_string(0, 3, "English Name", Some(&format))?;
    worksheet.write_string(0, 4, "ID", Some(&format))?;
    worksheet.write_string(0, 5, "Company (English)", Some(&format))?;
    worksheet.write_string(0, 6, "Address 1 (English)", Some(&format))?;
    worksheet.write_string(0, 7, "Address 2 (English)", Some(&format))?;
    worksheet.write_string(0, 8, "Address 3 (English)", Some(&format))?;
    worksheet.write_string(0, 9, "Address 4 (English)", Some(&format))?;
    worksheet.write_string(0, 10, "Company (Chinese)", Some(&format))?;
    worksheet.write_string(0, 11, "Address 1 (Chinese)", Some(&format))?;
    worksheet.write_string(0, 12, "Address 2 (Chinese)", Some(&format))?;
    worksheet.write_string(0, 13, "Address 3 (Chinese)", Some(&format))?;
    worksheet.write_string(0, 14, "Address 4 (Chinese)", Some(&format))?;
    worksheet.write_string(0, 15, "Spec Date From", Some(&format))?;
    worksheet.write_string(0, 16, "Spec Date To", Some(&format))?;
    worksheet.write_string(0, 17, "Promote", Some(&format))?;
    worksheet.write_string(0, 18, "Stroke", Some(&format))?;
    worksheet.write_string(0, 19, "Surname Stroke", Some(&format))?;

    // 寫入表頭到 CSV
    writeln!(csv_writer, "Surname,Name,English Surname,English Name,ID,Company (English),Address 1 (English),Address 2 (English),Address 3 (English),Address 4 (English),Company (Chinese),Address 1 (Chinese),Address 2 (Chinese),Address 3 (Chinese),Address 4 (Chinese),Spec Date From,Spec Date To,Promote,Stroke,Surname Stroke")?;

    let mut row = 1; // 從第二行開始寫入數據

    for page in 1..=42 {
        let params = [("t", "ladoctor"), ("p", &page.to_string())];

        match client.post(url).form(&params).send().await {
            Ok(response) => {
                if response.status().is_success() {
                    // 獲取原始響應內容以便於錯誤處理
                    let response_text = response.text().await?;
                    // println!("Response Text: {:?}", response_text); // 打印原始響應

                    // 嘗試解析 JSON
                    match serde_json::from_str::<Vec<ResponseData>>(&response_text) {
                        Ok(data) => {
                            if data.is_empty() {
                                eprintln!("第 {} 頁未找到數據", page);
                            }
                            for record in data {
                                // 寫入 Excel
                                worksheet.write_string(row, 0, record.csurname.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 1, record.cname.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 2, record.esurname.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 3, record.ename.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 4, record.id.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 5, record.ecompany.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 6, record.eaddr1.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 7, record.eaddr2.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 8, record.eaddr3.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 9, record.eaddr4.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 10, record.ccompany.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 11, record.caddr1.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 12, record.caddr2.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 13, record.caddr3.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 14, record.caddr4.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 15, record.spec_date_from.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 16, record.spec_date_to.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 17, record.promote.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 18, record.stroke.as_deref().unwrap_or(""), None)?;
                                worksheet.write_string(row, 19, record.surnamestroke.as_deref().unwrap_or(""), None)?;

                                // 寫入 CSV，確保參數與格式化字符串匹配
                                writeln!(csv_writer, "{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}", 
                                         record.csurname.as_deref().unwrap_or(""),
                                         record.cname.as_deref().unwrap_or(""),
                                         record.esurname.as_deref().unwrap_or(""),
                                         record.ename.as_deref().unwrap_or(""),
                                         record.id.as_deref().unwrap_or(""),
                                         record.ecompany.as_deref().unwrap_or(""),
                                         record.eaddr1.as_deref().unwrap_or(""),
                                         record.eaddr2.as_deref().unwrap_or(""),
                                         record.eaddr3.as_deref().unwrap_or(""),
                                         record.eaddr4.as_deref().unwrap_or(""),
                                         record.ccompany.as_deref().unwrap_or(""),
                                         record.caddr1.as_deref().unwrap_or(""),
                                         record.caddr2.as_deref().unwrap_or(""),
                                         record.caddr3.as_deref().unwrap_or(""),
                                         record.caddr4.as_deref().unwrap_or(""),
                                         record.spec_date_from.as_deref().unwrap_or(""),
                                         record.spec_date_to.as_deref().unwrap_or(""),
                                         record.promote.as_deref().unwrap_or(""),
                                         record.stroke.as_deref().unwrap_or(""),
                                         record.surnamestroke.as_deref().unwrap_or(""))?;
                                row += 1;
                            }
                        }
                        Err(err) => {
                            eprintln!("JSON 解析錯誤: {:?}", err);
                            // eprintln!("原始響應: {:?}", response_text); // 輸出原始響應內容
                        }
                    }
                } else {
                    eprintln!("請求失敗: {:?}", response.status());
                }
            }
            Err(err) => {
                eprintln!("請求錯誤: {:?}", err);
            }
        }
    }

    // 關閉工作簿，並保存檔案
    workbook.close().expect("Failed to close the workbook");
    Ok(())
}