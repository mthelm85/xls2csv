use std::path::PathBuf;
use structopt::StructOpt;

#[derive(StructOpt, Debug, Clone)]
pub struct Opt {
    #[structopt(parse(from_os_str))]
    pub input: PathBuf,
    #[structopt(short, long, parse(from_os_str))]
    pub output: Option<PathBuf>
}

pub fn args() -> Opt {
    Opt::from_args()
}