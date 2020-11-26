use std::{env, fs, path::PathBuf, process::Command};

fn main() {
    let target = env::var("TARGET").unwrap();
    let out_dir = env::var("OUT_DIR").unwrap();

    println!("cargo:rerun-if-changed=haka.pfx.idl");
    println!("cargo:rerun-if-changed=rtd.dll.manifest");
    // let path = env::var("PATH").unwrap();
    // dbg!(path);

    let tool = cc::windows_registry::find_tool(&target, "cl.exe").unwrap();
    let midl = embed_resource::find_tool("midl.exe");
    let mut command = Command::new(midl.unwrap());

    for (k, v) in tool.env() {
        command.env(k.clone(), v.clone());
    }

    command.args(&["/out", &out_dir]);

    if target.starts_with("x86_64-pc-windows") {
        command.args(&["/env", "x64"]);
    } else {
        command.args(&["/env", "win32"]);
    }
    command.arg("haka.pfx.idl");

    let output = command.output().unwrap();

    if !output.status.success() {
        panic!("midl.exe failed: {:?}", output);
    }

    //let _midl = cc::windows_registry::find_tool(&target, "midl.exe").unwrap();
    // 'C:\Program Files (x86)\Windows Kits\10\bin\x64\midl.exe'
    // need vswhom
    // midl.exe needs cl.exe in the path
    // midl /env win32 haka.pf.idl
    let mut rc_file = PathBuf::from(out_dir);
    rc_file.push("haka.pfx.rc");
    fs::copy("haka.pfx.rc", &rc_file).unwrap();

    embed_resource::compile(rc_file);
}
