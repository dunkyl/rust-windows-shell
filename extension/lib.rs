// Windows API function names use PascalCase
#![allow(non_snake_case)]
// * const T is a common parameter for COM-related trait functions
#![allow(clippy::not_unsafe_ptr_arg_deref)]

use std::sync::atomic::{AtomicIsize, Ordering};

use contextmenu::MyRustExtension;
use windows::{core::*, Win32::{Foundation::*, System::Registry::*}};
use windows::Win32::UI::Shell::{SHCNE_ASSOCCHANGED, SHChangeNotify, SHCNF_IDLIST};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::Win32::System::LibraryLoader::GetModuleFileNameW;

#[macro_use]
mod strings;
use strings::WinStr;
mod registry;
mod contextmenu;

const CLSID: GUID = GUID::from_u128(0x5AB29D2B_CC1D_45A2_AF6E_6853BF59909B);
const CLSID_STR: &str =             "{5AB29D2B-CC1D-45A2-AF6E-6853BF59909B}";
const NAME: &str = "My Rust Shell Extension";

fn registry_batch() -> registry::RegBatch {
    registry!{ HKEY_LOCAL_MACHINE,
        + ["SOFTWARE\\Classes\\CLSID\\{CLSID_STR}"] {
            ="My Rust Shell Extension Factory"
            + ["InprocServer32"] {
                =(CurrentDllLocation())
                "ThreadingModel"="Apartment"
            }
        }
        + ["SOFTWARE\\Classes\\*\\ShellEx\\ContextMenuHandlers\\{NAME}"] {
            =CLSID_STR
        }
        ["SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"] {
            CLSID_STR=NAME
        }
    }
}
use windows::Win32::Foundation::BOOL;
use windows::Win32::System::Com::{IClassFactory, IClassFactory_Impl};

#[implement(IClassFactory)]
struct MyRustExtensionClassFactory;

impl Default for MyRustExtensionClassFactory {
    fn default() -> Self {
        CLSOBJ_COUNT.fetch_add(1, Ordering::AcqRel);
        MyRustExtensionClassFactory {  }
    }
}

impl IClassFactory_Impl for MyRustExtensionClassFactory {
    fn CreateInstance(&self, pUnkOuter: Option<&IUnknown>, riid: *const GUID, ppvObject: *mut *mut core::ffi::c_void) ->  Result<()> {

        if pUnkOuter.is_some() { return CLASS_E_NOAGGREGATION.ok() }

        let Some(riid) = (unsafe { riid.as_ref() }) else { return E_INVALIDARG.ok() };

        let inst = MyRustExtension::default();

        let result = unsafe { IUnknown::from(inst).query(riid, ppvObject as _) };

        result.ok()
    }

    fn LockServer(&self, _: BOOL) -> Result<()> { Ok(()) }
}

impl Drop for MyRustExtensionClassFactory {
    fn drop(&mut self) {
        CLSOBJ_COUNT.fetch_sub(1, Ordering::AcqRel);
    }
}

fn CurrentDllLocation() -> String {

    let mut filename_buf = vec![0u16; MAX_PATH as usize];

    // number of utf-16 elements of the filepath
    let chars =  unsafe {
        GetModuleFileNameW(HMODULE(INSTANCE.load(std::sync::atomic::Ordering::Acquire)), &mut filename_buf)
    };
    filename_buf.truncate(chars as usize);

    String::from_utf16_lossy(&filename_buf)
}

fn msgBox(title: &WinStr, text: &WinStr,ty: MESSAGEBOX_STYLE) {
    unsafe {
        MessageBoxW(HWND::default(), text.get_pcwstr(), title.get_pcwstr(), ty);
    }
}

static INSTANCE: AtomicIsize = AtomicIsize::new(0);
static CLSOBJ_COUNT: AtomicIsize = AtomicIsize::new(0);

#[no_mangle]
pub extern fn DllMain (hinst: HMODULE, _reason: i32, _: isize) -> bool {
    INSTANCE.store(hinst.0, std::sync::atomic::Ordering::Release);
    true
}


#[no_mangle]
pub extern fn DllCanUnloadNow() -> HRESULT {
    if CLSOBJ_COUNT.load(std::sync::atomic::Ordering::Acquire) > 0 {
        S_FALSE
    } else {
        S_OK
    }
}


#[no_mangle]
pub extern fn DllGetClassObject(
    rclsid: *const GUID,
    riid: *const GUID,
    ppv: *mut *const core::ffi::c_void
) -> HRESULT {

    let Some(riid) = (unsafe { riid.as_ref() }) else {
        return E_INVALIDARG
    };
    let Some(&rclsid) = (unsafe {rclsid.as_ref()}) else {
        return E_INVALIDARG
    };
    if rclsid != CLSID {
        return CLASS_E_CLASSNOTAVAILABLE
    }
    if ppv.is_null() {
        return E_INVALIDARG
    }

    let factory = MyRustExtensionClassFactory::default();

    unsafe {
        Into::<IUnknown>::into(factory).query(riid, ppv)
    }
}


#[no_mangle]
pub extern fn DllRegisterServer() -> HRESULT {

    let result = registry_batch().apply();

    if let Err(e) = result {
        let text = wfmt!("{}", e.message());
        msgBox(&wfmt!("DllRegisterServer"), &text, MB_ICONERROR);
        registry_batch().unapply();
        return e.code();
    }

    unsafe { SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None); }
    S_OK
}


// https://learn.microsoft.com/en-us/windows/win32/shell/debugging-with-the-shell#Unloading
#[no_mangle]
pub extern fn DllUnregisterServer() -> HRESULT {
    
    registry_batch().unapply();

    unsafe { SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None); }
    HRESULT(0)
}

#[no_mangle]
pub extern fn DllInstall(_install: bool) -> HRESULT {
    S_OK
}