use anyhow::{ anyhow, Result };
use calamine::{ open_workbook_auto, DataType, Range, Reader };
use std::fs::{ self, File };
use std::io::{ BufWriter, Write };

mod input;

fn main() -> Result<(), anyhow::Error> {
    let input = input::args();
    
    let dir = if input.input.is_dir() {
        Ok(input.input)
    } else if input.input.is_file() {
        if let Some(dir) = input.input.parent() { Ok(dir.to_path_buf()) } else { Err(anyhow!("unable to read parent directory")) }
    } else {
        Err(anyhow!("unable to read directory/file"))
    };

    for file in fs::read_dir(dir?)? {
        let sce = file?.path();
        match sce.extension().and_then(|s| s.to_str()) {
            Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
            _ => continue,
        }
        
        let mut xl = open_workbook_auto(&sce)?;
        let sheets = xl.sheet_names().to_owned();
        for sheet in sheets {
            match &input.output {
                Some(path) => {
                    fs::create_dir_all(path)?;
                    let path = path.join(format!("{}_{}{}", sce.file_stem().unwrap().to_str().unwrap(), &sheet, ".csv"));
                    let mut dest = BufWriter::new(File::create(path)?);
                    let range = xl.worksheet_range(&sheet).unwrap()?;
                    write_range(&mut dest, &range)?;
                },
                None => {
                    if let Some(path) = sce.parent() {
                        let path = path.join(format!("{}_{}{}", sce.file_stem().unwrap().to_str().unwrap(), &sheet, ".csv"));
                        let mut dest = BufWriter::new(File::create(path)?);
                        let range = xl.worksheet_range(&sheet).unwrap()?;
                        write_range(&mut dest, &range)?;
                    }
                }
            }
        }
    }
    Ok(())
}

fn write_range<W: Write>(dest: &mut W, range: &Range<DataType>) -> std::io::Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match *c {
                DataType::Empty => Ok(()),
                DataType::String(ref s) => write!(dest, "\"{}\"", s.trim()),
                // convert dates to string format
                DataType::Float(ref f) | DataType::DateTime(ref f) => write!(dest, "{}", f),
                DataType::Int(ref i) => write!(dest, "{}", i),
                DataType::Error(ref e) => write!(dest, "{:?}", e),
                DataType::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, ",")?;
            }
        }
        write!(dest, "\r\n")?;
    }
    Ok(())
}
