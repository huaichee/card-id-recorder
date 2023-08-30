use std::fs;

use pcsc::*;
use serde::Deserialize;

#[derive(Deserialize, Debug)]
struct Config {
    is_cepas: bool,
}

fn main() {
    let card_no = read_card().unwrap();

    write_workbook(&card_no);
}

fn get_settings() -> Config {
    let config_file = fs::read_to_string(r"setting.toml");

    let config: Config = match config_file {
        Ok(file) => toml::from_str(&file).unwrap(),
        Err(_) => Config {
            is_cepas: false,
        },
    };

    return config;
}

fn write_workbook(card_no: &str){
    let file = r"test2.xlsx";

    let path = std::path::Path::new(file);
    let mut book = umya_spreadsheet::reader::xlsx::read(path).unwrap();

    let worksheet = book.get_sheet_by_name_mut("data").unwrap();


    let mut row = 1;    
    while row < 101 {
        let card_id_column = worksheet.get_value((6, row));
        let user_id = worksheet.get_value((1, row));

        if user_id.is_empty() {
            println!("No more users available");
            break;
        }

        if !user_id.is_empty() && card_id_column.is_empty() {
            worksheet.get_cell_mut((6,row)).set_value(card_no);

            break;
        }

        // Increment counter
        row += 1;
    }

    let _ = umya_spreadsheet::writer::xlsx::write(&book, path);
    
}   

fn read_card() -> Result<String, Error> {
    let setting = get_settings();

    // Establish a PC/SC context.
    let ctx = match Context::establish(Scope::User) {
        Ok(ctx) => ctx,
        Err(err) => {
            eprintln!("Failed to establish context: {}", err);
            std::process::exit(1);
        }
    };

    // List available readers.
    let mut readers_buf = [0; 2048];
    let mut readers = match ctx.list_readers(&mut readers_buf) {
        Ok(readers) => readers,
        Err(err) => {
            eprintln!("Failed to list readers: {}", err);
            std::process::exit(1);
        }
    };

    // Use the first reader.
    let reader = match readers.next() {
        Some(reader) => reader,
        None => {
            panic!("No readers are connected.");
        }
    };

    // Connect to the card.
    let card = match ctx.connect(reader, ShareMode::Shared, Protocols::ANY) {
        Ok(card) => card,
        Err(Error::NoSmartcard) => {
            panic!("A smartcard is not present in the reader.");
        }
        Err(err) => {
            eprintln!("Failed to connect to card: {}", err);
            std::process::exit(1);
        }
    };

    if setting.is_cepas == true {
        // Send an APDU command.
        let initialize_cepas = b"\x00\xA4\x00\x00\x02\x00\x00";

        let mut cepas_buf = [0; MAX_BUFFER_SIZE];
        let turn_on_cepas = match card.transmit(initialize_cepas, &mut cepas_buf) {
            Ok(rapdu) => rapdu,
            Err(err) => {
                eprintln!("Failed to transmit APDU command to card: {}", err);
                std::process::exit(1);
            }
        };

        println!("Cepas status: {:?}", turn_on_cepas);

        let apdu = b"\x90\x32\x03\x00\x00\x00";

        let mut rapdu_buf = [0; MAX_BUFFER_SIZE];
        let mut rapdu: Vec<_> = match card.transmit(apdu, &mut rapdu_buf) {
            Ok(rapdu) => rapdu.to_vec(),
            Err(err) => {
                eprintln!("Failed to transmit APDU command to card: {}", err);
                std::process::exit(1);
            }
        };

        // remove 144, 0
        rapdu.truncate(rapdu.len() - 2);


        let hex_value: Vec<_> = rapdu
            .iter()
            .map(|x| hex::encode_upper(x.to_be_bytes()))
            .collect();

        let hex_1 = &hex_value[8..16].join("");

        println!("CAN: {:?}", hex_1);

        println!("CSN: {:?}", &hex_value[17..25].join(":"));

        Ok(hex_1.to_string())

    } else {
        let apdu = b"\xFF\xCA\x00\x00\x00";

        let mut rapdu_buf = [0; MAX_BUFFER_SIZE];
        let mut rapdu: Vec<_> = match card.transmit(apdu, &mut rapdu_buf) {
            Ok(rapdu) => rapdu.to_vec(),
            Err(err) => {
                eprintln!("Failed to transmit APDU command to card: {}", err);
                std::process::exit(1);
            }
        };

        rapdu.truncate(rapdu.len() - 2);

        let hex_value: Vec<_> = rapdu
            .iter()
            .map(|x| hex::encode_upper(x.to_be_bytes()))
            .collect();

        let _decimal_value = i64::from_str_radix(&hex_value.join(""), 16);

        // println!("Decimal Method 1: {:?}", decimal_value.unwrap());

         // hex method 1
        let hex_method_1 = hex_value.join(":");
        // println!("Hex Method 1: {:?}", hex_method_1);

        // hex method 2
        let method_two: Vec<_> = hex_value
            .iter()
            .rev()
            .map(|x| x.to_string())
            .collect();

        let _decimal_value_two = i64::from_str_radix(&method_two.join(""), 16);

        // println!("Decimal Method 2: {:?}", decimal_value_two.unwrap());

        // println!("Hex Method 2: {:?}", method_two.join(":"));

        Ok(hex_method_1)
    // }
}
}