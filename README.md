# Rust Excel RTD

This is a basic RTD plugin for Excel written completely in Rust. It's not particularly polished, but it correctly loads and unloads and manages topic subscriptions.

There is a test sheet `sheets/rtdtest.xlsm`. The example sheet displays the time updating every few seconds.

Most Excel installations are still 32 bit. `cargo run --target i686-pc-windows-msvc` will build the 32 bit RTD server and register it.
The 64 bit build will work with 64 bit Excel.

## A bit about internals

It uses the `com` crate for the lowest level COM related needs. Because Excel RTD plugins require IDispatch interfaces, it compiles an IDL file and embeds the resulting TLB. This is embedded into the executable using the `embed-resources` crate.

If you decide to use this as a starting point, you should change the GUID near the top of `lib.rs`. There is a GUID for the type library also, but it is never registered so it doesn't matter much - the typeinfo is loaded by resource number from the dll instead.

The IRtdServer and IRTDUpdateEvent interfaces are small and were hand implemented in Rust. If your COM interface is much bigger, or it will ever change, this could be pretty painful to do manually.

Given that it's half hand-implemented and the interface is stable, it is probably possible to avoid the IDL step completely by using CreateDispTypeInfo, but I'd rather move in the opposite direction if a tool that generates the Rust interfaces becomes available.

Excel won't create COM objects with a progid that has a part shorter than 3 characters for some reason. "Haka.PF" works from the test harness but Excel won't load it. There is probably a way to see logs describing what was going wrong but this took ages to find and fix by comparing with a working C++ implementation. Something to avoid.
