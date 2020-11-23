use std::{env, fs, path::PathBuf, process::Command};
use vswhom::VsFindResult;

fn main() {
    let target = env::var("TARGET").unwrap();
    let out_dir = env::var("OUT_DIR").unwrap();

    println!("cargo:rerun-if-changed=haka.pfx.idl");
    println!("cargo:rerun-if-changed=rtd.dll.manifest");
    // let path = env::var("PATH").unwrap();
    // dbg!(path);

    // let tool_info = vswhom::
    let tool = cc::windows_registry::find_tool(&target, "cl.exe").unwrap();
    //let midl = cc::windows_registry::find_tool(&target, "midl.exe").unwrap();

    let arch = if target.starts_with("x86_64-pc-windows") {
        Arch::X64
    } else {
        Arch::X86
    };

    let midl = find_with_vswhom(arch).unwrap();
    let mut command = Command::new(midl);

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

#[derive(Debug, Copy, Clone, Hash, PartialEq, Eq, PartialOrd, Ord)]
enum Arch {
    X86,
    X64,
}

fn find_with_vswhom(arch: Arch) -> Option<PathBuf> {
    let res = VsFindResult::search();
    res.as_ref()
        .and_then(|res| res.windows_sdk_root.as_ref())
        .map(PathBuf::from)
        .and_then(|mut root| {
            let ver = root
                .file_name()
                .expect("malformed vswhom-returned SDK root")
                .to_os_string();
            root.pop();
            root.pop();
            root.push("bin");
            root.push(ver);
            try_bin_dir(root, "x86", "x64", arch)
        })
        .and_then(try_midl_exe)
        .or_else(move || {
            res.and_then(|res| res.windows_sdk_root)
                .map(PathBuf::from)
                .and_then(|mut root| {
                    root.pop();
                    root.pop();
                    try_bin_dir(root, "bin/x86", "bin/x64", arch)
                })
                .and_then(try_midl_exe)
        })
}

fn try_bin_dir<R: Into<PathBuf>>(
    root_dir: R,
    x86_bin: &str,
    x64_bin: &str,
    arch: Arch,
) -> Option<PathBuf> {
    try_bin_dir_impl(root_dir.into(), x86_bin, x64_bin, arch)
}

fn try_bin_dir_impl(
    mut root_dir: PathBuf,
    x86_bin: &str,
    x64_bin: &str,
    arch: Arch,
) -> Option<PathBuf> {
    match arch {
        Arch::X86 => root_dir.push(x86_bin),
        Arch::X64 => root_dir.push(x64_bin),
    }

    if root_dir.is_dir() {
        Some(root_dir)
    } else {
        None
    }
}

fn try_midl_exe(mut pb: PathBuf) -> Option<PathBuf> {
    pb.push("midl.exe");
    if pb.exists() {
        Some(pb)
    } else {
        None
    }
}
