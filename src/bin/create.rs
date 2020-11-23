use artydee::{IRtdServer, CLSID_CAT_CLASS};
use com::runtime::{create_instance, init_runtime};
use com::sys::FAILED;
use env_logger;
use log::info;
use std::process::Command;
use std::{env, mem};
use widestring::U16CString;

fn main() {
    if env::var("RUST_LOG").is_err() {
        env::set_var("RUST_LOG", "debug");
    }
    env_logger::init();

    print!("Registering the dll: ");
    info!("reg");
    let output = Command::new("regsvr32")
        .args(&["/s", r"..\artydee.dll"])
        .output()
        .expect("failed to execute process");

    if !output.status.success() {
        println!(
            "regsvr32 doesn't generate any output, but it's error code was {}",
            output.status.code().unwrap()
        );
    }
    println!("done");

    if false {
        println!("create by clsid");
        create_by_clsid()
    } else {
        println!("create by progid");
        create_by_progid()
    }
}

fn create_by_progid() {
    // this calls CoInitialize()
    com::runtime::init_apartment(com::runtime::ApartmentType::Multithreaded).unwrap();
    let mut clsid_cat_class: winapi::shared::guiddef::CLSID = unsafe { mem::zeroed() };
    let rtd_server_progid = U16CString::from_str("Haka.PFX").unwrap();

    unsafe {
        if FAILED(winapi::um::combaseapi::CLSIDFromProgID(
            rtd_server_progid.as_ptr(),
            &mut clsid_cat_class,
        )) {
            panic!("failed for Haka.PFX at CLSIDFromProgID");
        }

        let c = &clsid_cat_class as *const winapi::shared::guiddef::GUID as *const com::sys::GUID;

        let cat = create_instance::<IRtdServer>(&*c).expect("Failed to get a cat");
        use_rtd_server(cat);

        println!("seemed to work out by progid");
    }

    //CLSIDFromProgID();
    //CoCreateInstance();
}

fn create_by_clsid() {
    init_runtime().expect("Failed to initialize COM Library");
    println!("about to create it");

    let cat = create_instance::<IRtdServer>(&CLSID_CAT_CLASS).expect("Failed to get a cat");
    use_rtd_server(cat);

    // TODO: does dropping a com::Interface call IUnknown::Release()?
}

fn use_rtd_server(rtd_server: IRtdServer) {
    unsafe {
        // rtd_server.server_start(std::ptr::null_mut(), std::ptr::null_mut());
        rtd_server.server_terminate();
    }
}
