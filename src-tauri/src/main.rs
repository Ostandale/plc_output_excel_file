#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]
use chrono::{DateTime, FixedOffset, TimeZone, Utc};
use csv;
use encoding_rs::*;
use serde::Deserialize;
use std::error::Error;
use std::fs;
use std::fs::File;
use std::io;
use std::io::{BufRead, BufReader};
use std::path::{Path, PathBuf};
use tauri::api::file;
use xlsxwriter::*;

#[derive(Debug, Deserialize)]
struct Csv_data {
    id: i32,
    lot_num: &'static str,
    work_num: &'static str,
}

const LOCATE_STD_INFO01_ROW: u32 = 0;
const LOCATE_STD_INFO01_COL: u16 = 0;

const LOCATE_STD_COUNT01_ROW: u32 = 4;
const LOCATE_STD_COUNT01_COL: u16 = 8;

const LOCATE_STD_TIME01_ROW: u32 = 4;
const LOCATE_STD_TIME01_COL: u16 = 0;

const LOCATE_STD_RESIST01_ROW: u32 = 0;
const LOCATE_STD_RESIST01_COL: u16 = 4;

const LOCATE_STD_CP01_ROW: u32 = 0;
const LOCATE_STD_CP01_COL: u16 = 8;
//-------------------------------------------------------------------
// 関数名：
// 引数：
// 戻り値：
//
// 説明：
//
//
//-------------------------------------------------------------------
#[tauri::command]
fn select_csv_dir(path: Option<&str>) -> Result<String, String> {
    //ファイルを選ばずにキャンセルするとフロントにnullが返ってしまうので対策
    if path.is_none() {
        return Err("file select canceled".into());
    }
    let path = path.unwrap();

    let file_list = read_dir(path);
    if file_list.is_err() {
        return Err("error".into());
    }

    let create_dir_name = path.to_string() + "\\analyzed";
    match fs::create_dir(create_dir_name) {
        Ok(_) => {
            println!("ok")
        }
        Err(e) => {
            println!("err{}", e)
        }
    }

    let file_list = file_list.unwrap();

    //ディレクトリの中身が空の場合は戻る
    if file_list.len() == 0 {
        return Err("no file name".into());
    }

    //ファイル一覧を取得したら処理を行う
    for file_name in file_list.iter() {
        //＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
        //ここにワークブックのファイル新規作成を入れて
        //read_csv_fileの中で１行づつ読み込み、編集して書き込む
        let (_, file_name_no_extension, _, _) = isolate_path(file_name);
        read_csv_file(&path.to_string(), &file_name_no_extension);
    }

    Ok("ok".into())
}

//-------------------------------------------------------------------
// 関数名：
// 引数：
// 戻り値：
//
// 説明：
//
//
//-------------------------------------------------------------------
#[tauri::command]
fn file_path(path: Option<&str>) -> Result<String, String> {
    //ファイルを選ばずにキャンセルするとフロントにnullが返ってしまうので対策
    if path.is_none() {
        return Err("file select canceled".into());
    }
    let path = path.unwrap();

    //ファイルのパスからファイル名とパスを分けて取り出す
    let (file_name, file_name_no_extension, file_extension, file_path) = isolate_path(path);

    //拡張子がcsvにマッチしていなければエラー戻り
    if file_extension != "csv" {
        return Err("extension not match".into());
    }

    //ファイル名、パスに切り分ける
    let file_path = file_path + "\\" + "analyzed" + "\\";
    let file_name: &str = &(file_name_no_extension + ".xlsx");
    let file_path: &str = &(file_path + file_name);

    Ok("Done".into())
}

//-------------------------------------------------------------------
// 関数名：
// 引数：
// 戻り値：
//
// 説明：   １ファイル分のCSVデータを読み込んで、エクセルファイルに加工した値を書き出す
//
//
//-------------------------------------------------------------------
//  !   CSVファイルの読込とエクセルファイルへの書き出し
fn read_csv_file(path: &String, file_name: &String) -> Result<(), Box<dyn Error>> {
    let mut csv_file_data: Vec<String> = Vec::new();
    let full_pass = path.to_owned() + "\\" + file_name + ".csv";

    //  todo    シフトJIS対策を行う
    let csv_file_sj = fs::read(&full_pass)?;
    let (csv_file_u8, var_a, var_b) = SHIFT_JIS.decode(&csv_file_sj);

    //  *   CSVからの読込準備
    let mut csv_file = csv::Reader::from_reader(csv_file_u8.as_bytes());

    //  *   エクセルへの書込準備
    let excel_file_name = path.to_owned() + "\\analyzed\\" + file_name + ".xlsx";
    let workbook = Workbook::new(&excel_file_name);

    //  *   エクセルへ書き込むときのフォントサイズ、カラー、背景色などの設定
    let format1 = workbook.add_format().set_font_color(FormatColor::Black);
    let format_ok = workbook
        .add_format()
        .set_font_color(FormatColor::White)
        .set_bg_color(FormatColor::Green);
    let format_ce = workbook
        .add_format()
        .set_font_color(FormatColor::White)
        .set_bg_color(FormatColor::Purple);
    let format_lo = workbook
        .add_format()
        .set_font_color(FormatColor::White)
        .set_bg_color(FormatColor::Orange);
    let format_hi = workbook
        .add_format()
        .set_font_color(FormatColor::White)
        .set_bg_color(FormatColor::Red);
    let format_etc = workbook
        .add_format()
        .set_font_color(FormatColor::White)
        .set_bg_color(FormatColor::Black);

    let format_small_num = workbook
        .add_format()
        .set_font_color(FormatColor::Black)
        .set_bg_color(FormatColor::White)
        .set_num_format("0.000");

    let mut sheet1 = workbook.add_worksheet(Some("分析結果")).unwrap();
    let mut sheet2 = workbook.add_worksheet(Some("ログ")).unwrap();

    //  *   集計用変数
    let mut product_counter = 0;
    let mut ok_counter = 0;
    let mut ce_counter = 0;
    let mut lo_counter = 0;
    let mut hi_counter = 0;
    let mut etc_counter = 0;

    let mut start_time: i64 = 0;
    let mut end_time: i64 = 0;

    //  *   チョコ停計測用
    let mut stop_last_time: i64 = 0;
    let mut stop_count: i64 = 0;
    let mut stop_time_total: i64 = 0;

    //  *   CE LO HI 判定用フラグ
    let mut is_ce: bool = false;
    let mut is_lo: bool = false;
    let mut is_hi: bool = false;

    //  CP値算出用
    let mut cp_total: f64 = 0.0;
    let mut cp_counter: f64 = 0.0;
    let mut cp_measured_value: Vec<f64> = Vec::new();
    let mut cp_differnce: f64 = 0.0;
    let mut cp_up_value: f64 = 0.0;
    let mut cp_down_value: f64 = 0.0;
    let mut most_high_value: f64 = 0.0;
    let mut most_low_value: f64 = 0.0;

    //  *   一番最初の行を抜かして2番目から読込
    //  *   最初の行はheadersで読み込める
    for (index, result) in csv_file.records().enumerate() {
        let rec = result?;
        let data_epoch_time = &rec[0].parse::<i64>().unwrap();
        end_time = *data_epoch_time;
        let data_measured_value = &rec[4].parse::<f64>().unwrap();
        let data_decision = &rec[5];

        if index == 0 {
            //  *   データ取得1回目の情報から号機、製品番号、ロット番号をエクセルシートの1行目に書き込む
            sheet1
                .write_string(
                    LOCATE_STD_INFO01_ROW,
                    LOCATE_STD_INFO01_COL,
                    "号機",
                    Some(&format1),
                )
                .unwrap();
            sheet1
                .write_string(
                    LOCATE_STD_INFO01_ROW + 1,
                    LOCATE_STD_INFO01_COL,
                    &rec[1],
                    Some(&format1),
                )
                .unwrap();

            sheet1
                .write_string(
                    LOCATE_STD_INFO01_ROW + 0,
                    LOCATE_STD_INFO01_COL + 1,
                    "品番",
                    Some(&format1),
                )
                .unwrap();
            sheet1
                .write_string(
                    LOCATE_STD_INFO01_ROW + 1,
                    LOCATE_STD_INFO01_COL + 1,
                    &rec[2],
                    Some(&format1),
                )
                .unwrap();

            sheet1
                .write_string(
                    LOCATE_STD_INFO01_ROW + 0,
                    LOCATE_STD_INFO01_COL + 2,
                    "ロット番号",
                    Some(&format1),
                )
                .unwrap();
            sheet1
                .write_string(
                    LOCATE_STD_INFO01_ROW + 1,
                    LOCATE_STD_INFO01_COL + 2,
                    &rec[3],
                    Some(&format1),
                )
                .unwrap();
            start_time = *data_epoch_time;

            sheet1
                .write_string(
                    LOCATE_STD_RESIST01_ROW + 0,
                    LOCATE_STD_RESIST01_COL + 0,
                    "目標抵抗値",
                    Some(&format1),
                )
                .unwrap();
            sheet1
                .write_string(
                    LOCATE_STD_RESIST01_ROW + 1,
                    LOCATE_STD_RESIST01_COL + 0,
                    &rec[6],
                    Some(&format1),
                )
                .unwrap();

            sheet1
                .write_string(
                    LOCATE_STD_RESIST01_ROW + 0,
                    LOCATE_STD_RESIST01_COL + 1,
                    "許容範囲",
                    Some(&format1),
                )
                .unwrap();
            sheet1
                .write_string(
                    LOCATE_STD_RESIST01_ROW + 1,
                    LOCATE_STD_RESIST01_COL + 1,
                    &rec[7],
                    Some(&format1),
                )
                .unwrap();

            //  *   エクセルシート2枚目の1行目に見出しを加える
            sheet2.write_string(0, 0, "測定値", Some(&format1)).unwrap();
            sheet2.write_string(0, 1, "判定", Some(&format1)).unwrap();
            sheet2
                .write_string(0, 2, "作業時間", Some(&format1))
                .unwrap();

            //  CP計算用
            let difference: f64 = rec[7].parse::<f64>().unwrap();
            let resistance: f64 = rec[6].parse::<f64>().unwrap();
            cp_up_value = resistance * (1.0 + difference / 100.0);
            cp_down_value = resistance * (1.0 - difference / 100.0);
            cp_differnce = cp_up_value - cp_down_value;
            let cp_up_round = ((cp_up_value * 100.0).round()) / 100.0;
            let cp_down_round = ((cp_down_value * 100.0).round()) / 100.0;
            let cp_diff_width = format!("{}-{}", cp_up_round, cp_down_round);

            sheet1
                .merge_range(
                    LOCATE_STD_RESIST01_ROW + 0,
                    LOCATE_STD_RESIST01_COL + 2,
                    LOCATE_STD_RESIST01_ROW + 0,
                    LOCATE_STD_RESIST01_COL + 3,
                    "許容範囲",
                    Some(&format1),
                )
                .unwrap();
            sheet1
                .merge_range(
                    LOCATE_STD_RESIST01_ROW + 1,
                    LOCATE_STD_RESIST01_COL + 2,
                    LOCATE_STD_RESIST01_ROW + 1,
                    LOCATE_STD_RESIST01_COL + 3,
                    &cp_diff_width,
                    Some(&format1),
                )
                .unwrap();

            //  チョコ停　最初の１回だけトータルタイムに入れない
            stop_last_time = *data_epoch_time;
        }

        //  途中停止時間測定
        if *data_epoch_time > stop_last_time + 10000 {
            stop_count += 1;
            stop_time_total += (data_epoch_time - stop_last_time);
        }
        stop_last_time = *data_epoch_time;

        //  OK CE LO HI その他ごとにカウント　CE　LO　HIに関しては異常が出たら2回判定するのでそのあたりも考慮してカウント
        let col1 = match data_decision {
            "OK" => {
                ok_counter += 1;
                product_counter += 1;
                cp_measured_value.push(rec[4].parse().unwrap());
                cp_total += &rec[4].parse().unwrap();
                cp_counter += 1.0;
                is_ce = false;
                is_lo = false;
                is_hi = false;
                &format_ok
            }
            "CE" => {
                if is_ce {
                    ce_counter += 1;
                    product_counter += 1;
                    is_ce = false;
                } else {
                    is_ce = true;
                }
                &format_ce
            }
            "LOW" => {
                if is_lo {
                    lo_counter += 1;
                    product_counter += 1;
                    cp_measured_value.push(rec[4].parse().unwrap());
                    cp_total += &rec[4].parse().unwrap();
                    cp_counter += 1.0;
                    is_lo = false;
                } else {
                    is_lo = true;
                }
                &format_lo
            }
            "HI" => {
                if is_hi {
                    hi_counter += 1;
                    product_counter += 1;
                    cp_measured_value.push(rec[4].parse().unwrap());
                    cp_total += &rec[4].parse().unwrap();
                    cp_counter += 1.0;
                    is_hi = false;
                } else {
                    is_hi = true;
                }
                &format_hi
            }
            _ => {
                etc_counter += 1;
                &format_etc
            }
        };

        //  todo    シート２ページ目　ログ書込み
        //  *   判定に合わせて色付をしながら各データを１つづつ書込
        sheet2
            .write_number(
                (index + 1).try_into().unwrap(),
                0,
                *data_measured_value,
                Some(&format_small_num),
            )
            .unwrap();

        sheet2
            .write_string(
                (index + 1).try_into().unwrap(),
                1,
                data_decision,
                Some(col1),
            )
            .unwrap();

        sheet2
            .write_string(
                (index + 1).try_into().unwrap(),
                2,
                &chrono::Utc
                    .timestamp(*data_epoch_time / 1000 + (9 * 3600), 0)
                    .format("%H:%M:%S")
                    .to_string(),
                Some(&format1),
            )
            .unwrap();

        //       println!("{}  :  {}    :     {}", data_1, data_2, data_3);
    }

    // //  *   最後の列で異常が出ていたらカウント
    // if is_ce {
    //     ce_counter += 1;
    //     product_counter += 1;
    // }

    //  *   最終統計の書込
    //  総個数
    sheet1
        .write_string(
            LOCATE_STD_COUNT01_ROW,
            LOCATE_STD_COUNT01_COL + 0,
            "総個数",
            Some(&format1),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_COUNT01_ROW + 1,
            LOCATE_STD_COUNT01_COL + 0,
            product_counter as f64,
            Some(&format1),
        )
        .unwrap();

    //OK
    sheet1
        .write_string(
            LOCATE_STD_COUNT01_ROW,
            LOCATE_STD_COUNT01_COL + 1,
            "OK",
            Some(&format_ok),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_COUNT01_ROW + 1,
            LOCATE_STD_COUNT01_COL + 1,
            ok_counter as f64,
            Some(&format_ok),
        )
        .unwrap();

    //  CE
    sheet1
        .write_string(
            LOCATE_STD_COUNT01_ROW,
            LOCATE_STD_COUNT01_COL + 2,
            "CE",
            Some(&format_ce),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_COUNT01_ROW + 1,
            LOCATE_STD_COUNT01_COL + 2,
            ce_counter as f64,
            Some(&format_ce),
        )
        .unwrap();

    //  LO
    sheet1
        .write_string(
            LOCATE_STD_COUNT01_ROW,
            LOCATE_STD_COUNT01_COL + 3,
            "LOW",
            Some(&format_lo),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_COUNT01_ROW + 1,
            LOCATE_STD_COUNT01_COL + 3,
            lo_counter as f64,
            Some(&format_lo),
        )
        .unwrap();

    //  HI
    sheet1
        .write_string(
            LOCATE_STD_COUNT01_ROW,
            LOCATE_STD_COUNT01_COL + 4,
            "HI",
            Some(&format_hi),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_COUNT01_ROW + 1,
            LOCATE_STD_COUNT01_COL + 4,
            hi_counter as f64,
            Some(&format_hi),
        )
        .unwrap();

    //  エラー
    sheet1
        .write_string(
            LOCATE_STD_COUNT01_ROW,
            LOCATE_STD_COUNT01_COL + 5,
            "エラー",
            Some(&format_etc),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_COUNT01_ROW + 1,
            LOCATE_STD_COUNT01_COL + 5,
            etc_counter as f64,
            Some(&format_etc),
        )
        .unwrap();

    //  時間

    //  総作業時間
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 0,
            LOCATE_STD_TIME01_COL + 0,
            "総作業時間",
            Some(&format1),
        )
        .unwrap();
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 1,
            LOCATE_STD_TIME01_COL + 0,
            &chrono::Utc
                .timestamp((end_time - start_time) / 1000, 0)
                .format("%H:%M:%S")
                .to_string(),
            Some(&format1),
        )
        .unwrap();

    //  開始時間
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 0,
            LOCATE_STD_TIME01_COL + 2,
            "作業開始時間",
            Some(&format1),
        )
        .unwrap();
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 1,
            LOCATE_STD_TIME01_COL + 2,
            &chrono::Utc
                .timestamp_millis((start_time + (9 * 3600) * 1000))
                .format("%H:%M:%S")
                .to_string(),
            Some(&format1),
        )
        .unwrap();

    //  終了時間
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 0,
            LOCATE_STD_TIME01_COL + 4,
            "作業終了時間",
            Some(&format1),
        )
        .unwrap();
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 1,
            LOCATE_STD_TIME01_COL + 4,
            &chrono::Utc
                .timestamp_millis((end_time + (9 * 3600) * 1000))
                .format("%H:%M:%S")
                .to_string(),
            Some(&format1),
        )
        .unwrap();

    //  チョコ停　カウント
    let format_stop = match stop_count {
        0..=2 => &format_ok,
        2..=6 => &format_lo,
        6..=10 => &format_hi,
        _ => &format_ce,
    };
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 3,
            LOCATE_STD_TIME01_COL + 0,
            "途中停止回数",
            Some(format_stop),
        )
        .unwrap();
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 4,
            LOCATE_STD_TIME01_COL + 0,
            &stop_count.to_string(),
            Some(format_stop),
        )
        .unwrap();

    //  チョコ停　総時間
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 3,
            LOCATE_STD_TIME01_COL + 2,
            "途中停止時間",
            Some(&format1),
        )
        .unwrap();
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 4,
            LOCATE_STD_TIME01_COL + 2,
            &chrono::Utc
                .timestamp_millis(stop_time_total)
                .format("%H:%M:%S")
                .to_string(),
            Some(&format1),
        )
        .unwrap();

    //  １個生産するのにかかる時間
    sheet1
        .write_string(
            LOCATE_STD_TIME01_ROW + 0,
            LOCATE_STD_TIME01_COL + 6,
            "タクトタイム",
            Some(&format1),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_TIME01_ROW + 1,
            LOCATE_STD_TIME01_COL + 6,
            ((end_time - start_time) as f64 / product_counter as f64) / 1000.0,
            Some(&format1),
        )
        .unwrap();

    //  todo    CP値計算
    let cp_average = cp_total / (cp_counter); //  平均　有効取得抵抗値の合計を有効数で割った数値

    let mut cp_total_for_distributed: f64 = 0.0; // 分散　（取得抵抗値ー平均値）^2 を有効数で割った数値
    for index in cp_measured_value {
        cp_total_for_distributed += (index - cp_average) * (index - cp_average);
    }
    let cp_distributed: f64 = cp_total_for_distributed / (cp_counter);

    let standard_deviation = cp_distributed.sqrt(); //  標準偏差　分散の平方根を取った数値

    let cp_value = cp_differnce / (standard_deviation * 6.0); // CP値　目標値±許容差の幅を標準偏差 * 6 で割った数値

    cp_up_value = cp_up_value - cp_average; //  CPK値　平均値と許容差の間で小さい方の数値
    cp_down_value = cp_average - cp_down_value;
    let cpk_value: f64 = if cp_up_value > cp_down_value {
        cp_down_value / (standard_deviation * 3.0) // （目標値＋許容差ー平均値）/標準偏差 * 3
    } else {
        cp_up_value / (standard_deviation * 3.0) // （平均値ー目標値ー許容差）/標準偏差 * 3
    };

    //  平均抵抗値
    sheet1
        .write_string(
            LOCATE_STD_CP01_ROW + 0,
            LOCATE_STD_CP01_COL + 0,
            "平均抵抗値",
            Some(&format1),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_CP01_ROW + 1,
            LOCATE_STD_CP01_COL + 0,
            cp_average,
            Some(&format1),
        )
        .unwrap();

    // CP値色分け
    let format_cp = match cp_value {
        v if v >= 1.67 => &format_ok,
        v if v >= 1.33 => &format_ok,
        v if v >= 1.00 => &format_lo,
        v if v >= 0.67 => &format_hi,
        _ => &format_ce,
    };

    //  CP値
    sheet1
        .write_string(
            LOCATE_STD_CP01_ROW + 0,
            LOCATE_STD_CP01_COL + 2,
            "CP値",
            Some(format_cp),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_CP01_ROW + 1,
            LOCATE_STD_CP01_COL + 2,
            cp_value,
            Some(format_cp),
        )
        .unwrap();

    // CPK値色分け
    let format_cpk = match cpk_value {
        v if v >= 1.67 => &format_ok,
        v if v >= 1.33 => &format_ok,
        v if v >= 1.00 => &format_lo,
        v if v >= 0.67 => &format_hi,
        _ => &format_ce,
    };

    //  CPK値
    sheet1
        .write_string(
            LOCATE_STD_CP01_ROW + 0,
            LOCATE_STD_CP01_COL + 3,
            "CPK値",
            Some(format_cpk),
        )
        .unwrap();
    sheet1
        .write_number(
            LOCATE_STD_CP01_ROW + 1,
            LOCATE_STD_CP01_COL + 3,
            cpk_value,
            Some(format_cpk),
        )
        .unwrap();

    workbook.close().unwrap();
    Ok(())
}

//-------------------------------------------------------------------
// 関数名：
// 引数：
// 戻り値：
//
// 説明：   エクセルファイルに加工した値を書き出す
//
//
//-------------------------------------------------------------------
//エクセルファイルへの書き出し
fn write_excel_data(file_path: &String) -> Result<String, String> {
    //エクセルファイルの書き込み準備
    let workbook = Workbook::new(&file_path);
    //先に書式を設定しておいて、実際の書き込みの際に指定する
    let format1 = workbook.add_format().set_font_color(FormatColor::Black);

    let format2 = workbook
        .add_format()
        .set_font_color(FormatColor::Blue)
        .set_underline(FormatUnderline::Single);

    let format3 = workbook
        .add_format()
        .set_font_color(FormatColor::Green)
        .set_align(FormatAlignment::CenterAcross)
        .set_align(FormatAlignment::VerticalCenter);

    //シートへの書き込み、先に組み上げた書式を指定
    let mut sheet1 = workbook.add_worksheet(None).unwrap();
    sheet1
        .write_string(0, 0, "Red text", Some(&format1))
        .unwrap();
    sheet1.write_number(0, 1, 20., None).unwrap();
    sheet1.write_formula_num(1, 0, "=10+B1", None, 30.).unwrap();
    sheet1
        .write_url(1, 1, "https://github.com/Ostandale/Hello", Some(&format2))
        .unwrap();
    sheet1
        .merge_range(2, 0, 3, 2, "Hello, world", Some(&format3))
        .unwrap();
    sheet1.set_selection(1, 0, 1, 2);
    sheet1.set_tab_color(FormatColor::Cyan);
    match workbook.close() {
        Ok(v) => Ok("Done".into()),
        Err(e) => {
            println!("error:{}", e);
            Err("writing error".into())
        }
    }
    // Ok("Done".into())
}

//-------------------------------------------------------------------
// 関数名：
// 引数：
// 戻り値：
//
// 説明：   ファイル名まで含んでいるパスを渡すと
//          ファイルまでのパス、ファイル名、拡張子に分解してタプルで返す
//
//-------------------------------------------------------------------
//パスからファイル名と拡張子とディレクトリまでの文字列を切り分けてタプルで返す
fn isolate_path(path: &str) -> (String, String, String, String) {
    let path = PathBuf::from(path);

    //拡張子ありのファイル名の切り出し
    let file_name = path.file_name().unwrap().to_string_lossy().into_owned();

    //拡張子なしのファイル名だけの切り出し
    let file_name_no_extension = path.file_stem().unwrap().to_string_lossy().into_owned();

    //拡張子の切り出し
    let file_extension = path.extension().unwrap().to_string_lossy().into_owned();

    //パスの切り出し
    let file_path = path.parent().unwrap().to_string_lossy().into_owned();

    (file_name, file_name_no_extension, file_extension, file_path)
}

//-------------------------------------------------------------------
// 関数名：
// 引数：
// 戻り値：
//
// 説明：   ディレクトリまでのパスを渡すと
//          指定ディレクトリの中のCSVファイルだけのリストを作成する
//
//-------------------------------------------------------------------
fn read_dir<P: AsRef<Path>>(path: P) -> io::Result<Vec<String>> {
    Ok(fs::read_dir(path)?
        .filter_map(|entry| {
            let entry = entry.ok()?;
            let entry_extension = entry.path().extension()?;

            let is_extension_csv = entry.path().extension().unwrap() == "csv"
                || entry.path().extension().unwrap() == "CSV";
            if entry.file_type().ok()?.is_file() && is_extension_csv {
                Some(entry.file_name().to_string_lossy().into_owned())
            } else {
                None
            }
        })
        .collect())
}

//-------------------------------------------------------------------
// 関数名：
// 引数：
// 戻り値：
//
// 説明：
//
//
//-------------------------------------------------------------------
fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![select_csv_dir, file_path,])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
